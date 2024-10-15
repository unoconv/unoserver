from __future__ import annotations

import argparse
import logging
import os
import shutil
import signal
import socket
import subprocess
import sys
import tempfile
import threading
import time
import platform
import xmlrpc.server
from importlib import metadata
from pathlib import Path

from concurrent import futures

from unoserver import converter, comparer

API_VERSION = "3"
__version__ = metadata.version("unoserver")
logger = logging.getLogger("unoserver")


class XMLRPCServer(xmlrpc.server.SimpleXMLRPCServer):
    def __init__(
        self,
        addr: tuple[str, int],
        allow_none: bool = False,
    ) -> None:
        addr_info = socket.getaddrinfo(addr[0], addr[1], proto=socket.IPPROTO_TCP)

        if len(addr_info) == 0:
            raise RuntimeError(
                f"Could not get interface information for {addr[0]}:{addr[1]}"
            )

        self.address_family = addr_info[0][0]
        self.socket_type = addr_info[0][1]
        super().__init__(addr=addr_info[0][4], allow_none=allow_none)


class UnoServer:
    def __init__(
        self,
        interface="127.0.0.1",
        port="2003",
        uno_interface="127.0.0.1",
        uno_port="2002",
        user_installation=None,
        conversion_timeout=None,
    ):
        self.interface = interface
        self.uno_interface = uno_interface
        self.port = port
        self.uno_port = uno_port
        self.user_installation = user_installation
        self.conversion_timeout = conversion_timeout
        self.libreoffice_process = None
        self.xmlrcp_thread = None
        self.xmlrcp_server = None
        self.intentional_exit = False

    def start(self, executable="libreoffice"):
        logger.info(f"Starting unoserver {__version__}.")

        connection = (
            "socket,host=%s,port=%s,tcpNoDelay=1;urp;StarOffice.ComponentContext"
            % (self.uno_interface, self.uno_port)
        )

        # I think only --headless and --norestore are needed for
        # command line usage, but let's add everything to be safe.
        cmd = [
            executable,
            "--headless",
            "--invisible",
            "--nocrashreport",
            "--nodefault",
            "--nologo",
            "--nofirststartwizard",
            "--norestore",
            f"-env:UserInstallation={self.user_installation}",
            f"--accept={connection}",
        ]

        logger.info("Command: " + " ".join(cmd))
        self.libreoffice_process = subprocess.Popen(cmd)
        self.xmlrcp_thread = threading.Thread(None, self.serve)

        def signal_handler(signum, frame):
            self.intentional_exit = True
            logger.info("Sending signal to LibreOffice")
            try:
                self.libreoffice_process.send_signal(signum)
            except ProcessLookupError as e:
                # 3 means the process is already dead
                if e.errno != 3:
                    raise

            if self.xmlrcp_server is not None:
                self.stop()  # Ensure the server stops

        signal.signal(signal.SIGTERM, signal_handler)
        signal.signal(signal.SIGINT, signal_handler)

        # Signal SIGHUP is available only in Unix systems
        if platform.system() != "Windows":
            signal.signal(signal.SIGHUP, signal_handler)

        time.sleep(10)

        self.xmlrcp_thread.start()

        # Give the thread time to start
        time.sleep(2)
        # Check if it succeeded
        if not self.xmlrcp_thread.is_alive():
            logger.info("Failed to start servers")
            self.stop()
            return None

        return self.libreoffice_process

    def serve(self):
        # Create server
        with XMLRPCServer((self.interface, int(self.port)), allow_none=True) as server:
            self.conv = converter.UnoConverter(
                interface=self.uno_interface, port=self.uno_port
            )
            self.comp = comparer.UnoComparer(
                interface=self.uno_interface, port=self.uno_port
            )

            self.xmlrcp_server = server
            server.register_introspection_functions()

            @server.register_function
            def info():
                import_filters = self.conv.get_filter_names(
                    self.conv.get_available_import_filters()
                )
                export_filters = self.conv.get_filter_names(
                    self.conv.get_available_export_filters()
                )
                return {
                    "unoserver": __version__,
                    "api": API_VERSION,
                    "import_filters": import_filters,
                    "export_filters": export_filters,
                }

            @server.register_function
            def convert(
                inpath=None,
                indata=None,
                outpath=None,
                convert_to=None,
                filtername=None,
                filter_options=[],
                update_index=True,
                infiltername=None,
            ):
                if indata is not None:
                    indata = indata.data

                with futures.ThreadPoolExecutor() as executor:
                    future = executor.submit(
                        self.conv.convert,
                        inpath,
                        indata,
                        outpath,
                        convert_to,
                        filtername,
                        filter_options,
                        update_index,
                        infiltername,
                    )
                    try:
                        return future.result(timeout=self.conversion_timeout)
                    except futures.TimeoutError:
                        logger.error(
                            "Conversion timeout, terminating conversion and exiting."
                        )
                        self.conv.local_context.dispose()
                        self.libreoffice_process.terminate()
                        raise

            @server.register_function
            def compare(
                oldpath=None,
                olddata=None,
                newpath=None,
                newdata=None,
                outpath=None,
                filetype=None,
            ):
                if olddata is not None:
                    olddata = olddata.data
                if newdata is not None:
                    newdata = newdata.data

                with futures.ThreadPoolExecutor() as executor:
                    future = executor.submit(
                        self.comp.compare,
                        oldpath,
                        olddata,
                        newpath,
                        newdata,
                        outpath,
                        filetype,
                    )
                try:
                    return future.result(timeout=self.conversion_timeout)
                except futures.TimeoutError:
                    logger.error(
                        "Comparison timeout, terminating conversion and exiting."
                    )
                    self.conv.local_context.dispose()
                    self.libreoffice_process.terminate()
                    raise

            server.serve_forever()

    def stop(self):

        if self.xmlrcp_server is not None:
            self.xmlrcp_server.shutdown()
            # Make a dummy connection to unblock accept() - otherwise it will
            # hang indefinitely in the accept() call.
            # noinspection PyBroadException
            try:
                with socket.create_connection(
                    (self.interface, int(self.port)), timeout=1
                ):
                    pass
            except Exception:
                pass  # Ignore any except

        if self.xmlrcp_thread is not None:
            self.xmlrcp_thread.join()

        if self.libreoffice_process and self.libreoffice_process.poll() is not None:
            self.libreoffice_process.terminate()
            try:
                self.libreoffice_process.wait(10)
            except subprocess.TimeoutExpired:
                logger.info("Signalling harder...")
                self.libreoffice_process.terminate()


def main():
    logging.basicConfig()
    logger.setLevel(logging.INFO)

    parser = argparse.ArgumentParser("unoserver")
    parser.add_argument(
        "-v",
        "--version",
        action="version",
        help="Display version and exit.",
        version=f"{parser.prog} {__version__}",
    )
    parser.add_argument(
        "--interface",
        default="127.0.0.1",
        help="The interface used by the XMLRPC server",
    )
    parser.add_argument(
        "--uno-interface",
        default="127.0.0.1",
        help="The interface used by the Libreoffice UNO server",
    )
    parser.add_argument(
        "--port", default="2003", help="The port used by the XMLRPC server"
    )
    parser.add_argument(
        "--uno-port", default="2002", help="The port used by the Libreoffice UNO server"
    )
    parser.add_argument("--daemon", action="store_true", help="Deamonize the server")
    parser.add_argument(
        "--executable",
        default=None,
        help="The path to the LibreOffice executable, defaults to looking in the path",
    )
    parser.add_argument(
        "--user-installation",
        default=None,
        help="The path to the LibreOffice user profile",
    )
    parser.add_argument(
        "--libreoffice-pid-file",
        "-p",
        default=None,
        help="If set, unoserver will write the Libreoffice PID to this file. If started "
        "in daemon mode, the file will not be deleted when unoserver exits.",
    )
    parser.add_argument(
        "--conversion-timeout",
        type=int,
        help="Terminate Libreoffice and exit if a conversion does not complete in the "
        "given time (in seconds).",
    )
    args = parser.parse_args()

    if args.daemon:
        cmd = sys.argv
        cmd.remove("--daemon")
        proc = subprocess.Popen(cmd)
        return proc.pid

    with tempfile.TemporaryDirectory() as tmpuserdir:
        user_installation = Path(tmpuserdir).as_uri()

        if args.user_installation is not None:
            user_installation = Path(args.user_installation).as_uri()

        if args.uno_port == args.port:
            raise RuntimeError("--port and --uno-port must be different")

        server = UnoServer(
            args.interface,
            args.port,
            args.uno_interface,
            args.uno_port,
            user_installation,
            args.conversion_timeout,
        )

        if args.executable is not None:
            executable = args.executable
        else:
            # Find the executable automatically. I had problems with
            # LibreOffice using 100% if started with the libreoffice
            # executable, so by default try soffice first. Also throwing
            # ooffice in there as a fallback, I don't think it's used any
            # more, but it doesn't hurt to have it there.
            for name in ("soffice", "libreoffice", "ooffice"):
                if (executable := shutil.which(name)) is not None:
                    break

        # If it's daemonized, this returns the process.
        # It returns 0 of getting killed in a normal way.
        # Otherwise it returns 1 after the process exits.
        process = server.start(executable=executable)
        if process is None:
            return 2
        pid = process.pid

        logger.info(f"Server PID: {pid}")

        if args.libreoffice_pid_file:
            with open(args.libreoffice_pid_file, "wt") as upf:
                upf.write(f"{pid}")

        process.wait()

        if not server.intentional_exit:
            logger.error(f"Looks like LibreOffice died. PID: {pid}")

        # The RPC thread needs to be stopped before the process can exit
        server.stop()
        if args.libreoffice_pid_file:
            # Remove the PID file
            os.unlink(args.libreoffice_pid_file)

        try:
            # Make sure it's really dead
            os.kill(pid, 0)

            if server.intentional_exit:
                return 0
            else:
                return 1
        except OSError as e:
            if e.errno == 3:
                # All good, it was already dead.
                return 0
            raise


if __name__ == "__main__":
    main()
