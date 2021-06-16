import argparse
import subprocess
import tempfile
from urllib import request


def server(interface="127.0.0.1", port="2002"):
    print("Starting unoserver.")

    with tempfile.TemporaryDirectory() as tmpuserdir:

        connection = (
            "socket,host=%s,port=%s,tcpNoDelay=1;urp;StarOffice.ComponentContext"
            % (interface, port)
        )
        tmp_uri = "file://" + request.pathname2url(tmpuserdir)

        # I think only --headless and --norestore are needed for
        # command line usage, but let's add everything to be safe.
        cmd = [
            "libreoffice",
            "--headless",
            "--invisible",
            "--nocrashreport",
            "--nodefault",
            "--nologo",
            "--nofirststartwizard",
            "--norestore",
            f"-env:UserInstallation={tmp_uri}",
            f"--accept={connection}",
        ]

        print(cmd)
        subprocess.call(cmd)


def main():
    parser = argparse.ArgumentParser("unoserver")
    parser.add_argument(
        "--interface", default="127.0.0.1", help="The interface used by the server"
    )
    parser.add_argument("--port", default="2002", help="The port used by the server")
    args = parser.parse_args()

    server(args.interface, args.port)


if __name__ == "__main__":
    main()
