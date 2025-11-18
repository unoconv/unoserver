from __future__ import annotations

import argparse
import logging
import os
import platform
import shutil
import signal
import socket
import subprocess
import sys
import tempfile
import threading
import time
import xmlrpc.server
from importlib import metadata
from pathlib import Path

from concurrent import futures

from unoserver import converter, comparer
from com.sun.star.uno import Exception as UnoException

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
        stop_after=None,
    ):
        self.interface = interface
        self.uno_interface = uno_interface
        self.port = port
        self.uno_port = uno_port
        self.user_installation = user_installation
        self.conversion_timeout = conversion_timeout
        self.stop_after = stop_after
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

        time.sleep(5)

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
            logger.info("Starting UnoConverter.")
            attempts = 20
            while attempts > 0:
                try:
                    self.conv = converter.UnoConverter(
                        interface=self.uno_interface, port=self.uno_port
                    )
                    break
                except UnoException as e:
                    # A connection refused just means it hasn't started yet:
                    if "Connection refused" in str(e):
                        logger.debug("Libreoffice is not yet started")
                        time.sleep(2)
                        attempts -= 1
                        continue
                    # This is a different error
                    logger.warning("Error when starting UnoConverter, retrying: %s", e)
                    # These kinds of errors can be retried fewer times
                    attempts -= 4
                    time.sleep(5)
                    continue
            else:
                # We ran out of attempts
                logger.critical("Could not start Libreoffice, exiting.")
                # Make sure it's really dead
                self.libreoffice_process.terminate()
                return

            logger.info("Starting UnoComparer.")
            attempts = 20
            while attempts > 0:
                try:
                    self.comp = comparer.UnoComparer(
                        interface=self.uno_interface, port=self.uno_port
                    )
                    break
                except UnoException as e:
                    # A connection refused just means it hasn't started yet:
                    if "Connection refused" in str(e):
                        logger.debug("Libreoffice is not yet started")
                        attempts -= 1
                        time.sleep(2)
                        continue
                    # This is a different error
                    logger.warning("Error when starting UnoConverter, retrying: %s", e)
                    # These kinds of errors can be retried fewer times
                    attempts -= 4
                    time.sleep(5)
                    continue
            else:
                # We ran out of attempts
                logger.critical("Could not start Libreoffice, exiting.")
                # Make sure it's really dead
                self.libreoffice_process.terminate()
                return

            self.xmlrcp_server = server
            server.register_introspection_functions()

            self.number_of_requests = 0

            def stop_after():
                if self.stop_after is None:
                    return
                self.number_of_requests += 1
                if self.number_of_requests == self.stop_after:
                    logger.info(
                        "Processed %d requests, exiting.",
                        self.stop_after,
                    )
                    self.intentional_exit = True
                    self.libreoffice_process.terminate()

            @server.register_function
            def info():
                """Get server information and available document format filters.

                This XMLRPC endpoint provides version information and lists all available
                import and export filters supported by the LibreOffice instance.

                Returns:
                    dict: Server information containing:
                        - unoserver (str): The version of unoserver
                        - api (str): The API version number
                        - import_filters (list[str]): Names of available import filters
                          for reading various document formats
                        - export_filters (list[str]): Names of available export filters
                          for writing various document formats

                Example:
                    <?xml version="1.0"?>
                    <methodResponse>
                      <params>
                        <param>
                          <value>
                            <struct>
                              <member>
                                <name>unoserver</name>
                                <value><string>2.0.4</string></value>
                              </member>
                              <member>
                                <name>api</name>
                                <value><string>3</string></value>
                              </member>
                              <member>
                                <name>import_filters</name>
                                <value>
                                  <struct>
                                    <member>
                                      <name>Text (encoded)</name>
                                      <value><string>Plain text file with encoding detection</string></value>
                                    </member>
                                    <member>
                                      <name>Microsoft Word 2007/2010/2013/2016</name>
                                      <value><string>Microsoft Word DOCX format</string></value>
                                    </member>
                                  </struct>
                                </value>
                              </member>
                              <member>
                                <name>export_filters</name>
                                <value>
                                  <struct>
                                    <member>
                                      <name>PDF - Portable Document Format</name>
                                      <value><string>Adobe PDF export format</string></value>
                                    </member>
                                    <member>
                                      <name>Microsoft Excel 2007-365</name>
                                      <value><string>Microsoft Excel XLSX format</string></value>
                                    </member>
                                  </struct>
                                </value>
                              </member>
                            </struct>
                          </value>
                        </param>
                      </params>
                    </methodResponse>

                Note:
                    The available filters depend on the LibreOffice installation and version.
                    This endpoint can be used to determine what file formats are supported
                    before attempting conversions.
                """
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
                """Convert documents between different formats using LibreOffice.

                This is an XMLRPC server endpoint that wraps the underlying UnoConverter.convert()
                method. It provides document conversion capabilities with timeout protection,
                thread safety, and automatic request counting for server lifecycle management.

                Args:
                    inpath (str | None): File path to the input document on the local filesystem.
                        Mutually exclusive with indata - exactly one must be provided.
                    indata (bytes | xmlrpc.client.Binary | None): Binary data of the input file
                        for remote conversion without requiring a file path. Mutually exclusive
                        with inpath - exactly one must be provided.
                    outpath (str | None): File path where the converted output should be saved.
                        If provided, the converted file is saved to disk and None is returned.
                        Mutually exclusive with convert_to - exactly one must be provided.
                    convert_to (str | None): Target file extension/format (e.g., "pdf", "xlsx",
                        "docx"). If provided, the converted content is returned as bytes.
                        Mutually exclusive with outpath - exactly one must be provided.
                    filtername (str | None): Specific LibreOffice export filter name to use.
                        If None, the filter is auto-detected based on the target format.
                    filter_options (list): List of filter options as strings in "OptionName=Value"
                        format. Used to customize the export behavior (e.g., PDF quality settings).
                    update_index (bool): Whether to update document indexes (e.g., Table of Contents,
                        cross-references) before conversion. Default is True.
                    infiltername (str | None): Specific LibreOffice import filter name to use.
                        If None, the filter is auto-detected based on the input format.

                Returns:
                    bytes | None: Returns converted file content as bytes if convert_to is specified,
                        or None if outpath is specified (file saved to disk).

                Raises:
                    TimeoutError: If conversion exceeds the configured conversion_timeout.
                        When this occurs, the LibreOffice process is terminated and the server exits.
                    ValueError: If required parameter combinations are not provided (either
                        inpath or indata must be specified, and either outpath or convert_to).

                Examples:
                    Convert file to PDF and save to disk:
                        <?xml version="1.0"?>
                        <methodCall>
                          <methodName>convert</methodName>
                          <params>
                            <param><value><string>/path/to/doc.docx</string></value></param>
                            <param><value><nil/></value></param>
                            <param><value><string>/path/to/output.pdf</string></value></param>
                          </params>
                        </methodCall>

                    Convert binary data and return PDF bytes:
                        <?xml version="1.0"?>
                        <methodCall>
                          <methodName>convert</methodName>
                          <params>
                            <param><value><nil/></value></param>
                            <param><value><base64>UEsDBBQAAAAIAAgA...</base64></value></param>
                            <param><value><nil/></value></param>
                            <param><value><string>pdf</string></value></param>
                          </params>
                        </methodCall>

                    Convert with specific filter options:
                        <?xml version="1.0"?>
                        <methodCall>
                          <methodName>convert</methodName>
                          <params>
                            <param><value><string>doc.docx</string></value></param>
                            <param><value><nil/></value></param>
                            <param><value><string>out.pdf</string></value></param>
                            <param><value><nil/></value></param>
                            <param><value><nil/></value></param>
                            <param>
                              <value>
                                <array>
                                  <data>
                                    <value><string>Quality=90</string></value>
                                    <value><string>CompressImages=true</string></value>
                                  </data>
                                </array>
                              </value>
                            </param>
                          </params>
                        </methodCall>
                """
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
                        result = future.result(timeout=self.conversion_timeout)
                    except futures.TimeoutError:
                        logger.error(
                            "Conversion timeout, terminating conversion and exiting."
                        )
                        self.conv.local_context.dispose()
                        self.libreoffice_process.terminate()
                        raise
                    else:
                        stop_after()
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
                """Compare two documents and generate a comparison result.

                This XMLRPC endpoint wraps the underlying UnoComparer.compare() method
                to provide document comparison capabilities with timeout protection and
                thread safety. It can compare documents by file path or binary data.

                Args:
                    oldpath (str | None): File path to the original/old document on the
                        local filesystem. Mutually exclusive with olddata - exactly one
                        must be provided for the old document.
                    olddata (bytes | xmlrpc.client.Binary | None): Binary data of the
                        original/old document for remote comparison. Mutually exclusive
                        with oldpath - exactly one must be provided for the old document.
                    newpath (str | None): File path to the new/revised document on the
                        local filesystem. Mutually exclusive with newdata - exactly one
                        must be provided for the new document.
                    newdata (bytes | xmlrpc.client.Binary | None): Binary data of the
                        new/revised document for remote comparison. Mutually exclusive
                        with newpath - exactly one must be provided for the new document.
                    outpath (str | None): File path where the comparison result should be
                        saved. If None, the comparison result is returned as binary data.
                    filetype (str | None): Specific file format for the comparison output
                        (e.g., "odt", "docx", "pdf"). If None, format is auto-detected
                        or defaults based on the output path extension.

                Returns:
                    bytes | None: Returns comparison result as bytes if outpath is None,
                        or None if outpath is specified (result saved to disk).

                Raises:
                    TimeoutError: If comparison exceeds the configured conversion_timeout.
                        When this occurs, the LibreOffice process is terminated and the server exits.
                    ValueError: If required parameter combinations are not provided (must
                        specify either oldpath or olddata, and either newpath or newdata).

                Examples:
                    Compare files and save result to disk:
                        <?xml version="1.0"?>
                        <methodCall>
                          <methodName>compare</methodName>
                          <params>
                            <param><value><string>/path/to/v1.docx</string></value></param>
                            <param><value><nil/></value></param>
                            <param><value><string>/path/to/v2.docx</string></value></param>
                            <param><value><nil/></value></param>
                            <param><value><string>/path/to/comparison.odt</string></value></param>
                          </params>
                        </methodCall>

                    Compare binary data and return result:
                        <?xml version="1.0"?>
                        <methodCall>
                          <methodName>compare</methodName>
                          <params>
                            <param><value><nil/></value></param>
                            <param><value><base64>UEsDBBQAAAAIAAgA...</base64></value></param>
                            <param><value><nil/></value></param>
                            <param><value><base64>UEsDBBQAAAAIAAgA...</base64></value></param>
                            <param><value><nil/></value></param>
                            <param><value><string>odt</string></value></param>
                          </params>
                        </methodCall>
                """
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
                    result = future.result(timeout=self.conversion_timeout)
                except futures.TimeoutError:
                    logger.error(
                        "Comparison timeout, terminating conversion and exiting."
                    )
                    self.conv.local_context.dispose()
                    self.libreoffice_process.terminate()
                    raise
                else:
                    stop_after()
                    return result

            logger.info("Started.")
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
    parser.add_argument(
        "--stop-after",
        type=int,
        help="Terminate Libreoffice and exit after the given number of requests.",
    )
    group = parser.add_mutually_exclusive_group()
    group.add_argument(
        "--verbose",
        action="store_true",
        dest="verbose",
        help="Increase informational output to logs.",
    )
    group.add_argument(
        "--quiet",
        action="store_true",
        dest="quiet",
        help="Decrease informational output to logs.",
    )
    parser.add_argument(
        "-f",
        "--logfile",
        dest="logfile",
        help="Write logs to a file (defaults to stderr)",
    )
    args = parser.parse_args()

    if args.verbose:
        log_args = {"level": logging.DEBUG}
    elif args.quiet:
        log_args = {"level": logging.CRITICAL}
    else:
        log_args = {"level": logging.INFO}

    logging.basicConfig(**log_args)

    if args.daemon or args.logfile:
        cmd = sys.argv
        if args.daemon:
            cmd.remove("--daemon")

        if args.logfile:
            cmd.remove("--logfile")
            cmd.remove(args.logfile)

            with open(args.logfile, "ab") as logfile:
                # This is the only way I can find to get logging to a file that
                # will also consistently get exeptions in a thread logged to
                # that file. I could possibly just redirect sys.stderr, but
                # since I need this to daemonize unoserver, I might just as
                # well use it. Just overriding the exception hook will not work
                # file the thread, and I tried overriding the exception hook
                # for the thread as well, but couldn't make it work.
                proc = subprocess.Popen(cmd, stderr=logfile)
        else:
            proc = subprocess.Popen(cmd)

        if args.daemon:
            return proc.pid
        else:
            return proc.wait()

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
            args.stop_after,
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
        logger.info("Stopping Unoserver")
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
