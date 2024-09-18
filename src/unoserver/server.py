from __future__ import annotations

import argparse
import logging
import os
import signal
import socket
import subprocess
import sys
import tempfile
import threading
import platform
import xmlrpc.server
from importlib import metadata
from pathlib import Path

from unoserver import converter, comparer

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
    ):
        self.interface = interface
        self.uno_interface = uno_interface
        self.port = port
        self.uno_port = uno_port
        self.user_installation = user_installation
        self.libreoffice_process = None
        self.xmlrcp_thread = None
        self.xmlrcp_server = None

    def start(self, executable="libreoffice"):
        import time

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
            logger.info("Sending signal to LibreOffice")
            try:
                self.libreoffice_process.send_signal(signum)
            except ProcessLookupError as e:
                # 3 means the process is already dead
                if e.errno != 3:
                    raise

            if self.xmlrcp_server is not None:
                self.xmlrcp_server.shutdown()

        signal.signal(signal.SIGTERM, signal_handler)
        signal.signal(signal.SIGINT, signal_handler)

        # Signal SIGHUP is available only in Unix systems
        if platform.system() != "Windows":
            signal.signal(signal.SIGHUP, signal_handler)

        time.sleep(10)

        self.xmlrcp_thread.start()

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

                result = self.conv.convert(
                    inpath,
                    indata,
                    outpath,
                    convert_to,
                    filtername,
                    filter_options,
                    update_index,
                    infiltername,
                )
                return result

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

                result = self.comp.compare(
                    oldpath, olddata, newpath, newdata, outpath, filetype
                )
                return result

            server.serve_forever()

    def stop(self):
        if self.libreoffice_process:
            self.libreoffice_process.terminate()
        if self.xmlrcp_server is not None:
            self.xmlrcp_server.shutdown()
        if self.xmlrcp_thread is not None:
            self.xmlrcp_thread.join()


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
        default="libreoffice",
        help="The path to the LibreOffice executable",
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
        )

        # If it's daemonized, this returns the process.
        # It returns 0 of getting killed in a normal way.
        # Otherwise it returns 1 after the process exits.
        process = server.start(executable=args.executable)
        pid = process.pid

        logger.info(f"Server PID: {pid}")

        if args.libreoffice_pid_file:
            with open(args.libreoffice_pid_file, "wt") as upf:
                upf.write(f"{pid}")

        process.wait()

        if args.libreoffice_pid_file:
            # Remove the PID file
            os.unlink(args.libreoffice_pid_file)

        try:
            # Make sure it's really dead
            os.kill(pid, 0)
            # It was killed
            return 0
        except OSError as e:
            if e.errno == 3:
                # All good, it was already dead.
                return 0
            raise


if __name__ == "__main__":
    main()
