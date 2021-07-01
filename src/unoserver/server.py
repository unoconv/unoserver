import argparse
import logging
import subprocess
import tempfile
from urllib import request

logging.basicConfig()
logger = logging.getLogger("unoserver")
logger.setLevel(logging.INFO)


class UnoServer:
    def __init__(self, interface="127.0.0.1", port="2002"):
        self.interface = interface
        self.port = port

    def start(self, daemon=False, executable="libreoffice"):
        logger.info("Starting unoserver.")

        with tempfile.TemporaryDirectory() as tmpuserdir:

            connection = (
                "socket,host=%s,port=%s,tcpNoDelay=1;urp;StarOffice.ComponentContext"
                % (self.interface, self.port)
            )

            # Store this as an attribute, it helps testing
            self.tmp_uri = "file://" + request.pathname2url(tmpuserdir)

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
                f"-env:UserInstallation={self.tmp_uri}",
                f"--accept={connection}",
            ]

            logger.info("Command: " + " ".join(cmd))
            process = subprocess.Popen(cmd)
            if not daemon:
                process.wait()
            else:
                return process


def main():
    parser = argparse.ArgumentParser("unoserver")
    parser.add_argument(
        "--interface", default="127.0.0.1", help="The interface used by the server"
    )
    parser.add_argument("--port", default="2002", help="The port used by the server")
    parser.add_argument("--daemon", action="store_true", help="Deamonize the server")
    parser.add_argument(
        "--executable",
        default="libreoffice",
        help="The path to the LibreOffice executable",
    )
    args = parser.parse_args()

    server = UnoServer(args.interface, args.port)
    # If it's daemonized, this returns the process.
    # Otherwise it returns None after the process exites.
    return server.start(daemon=args.daemon, executable=args.executable)
