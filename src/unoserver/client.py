import argparse
import logging
import os
import sys
import time

from importlib import metadata
from xmlrpc.client import ServerProxy

__version__ = metadata.version("unoserver")
logger = logging.getLogger("unoserver")

API_VERSION = "3"
SFX_FILTER_IMPORT = 1
SFX_FILTER_EXPORT = 2
DOC_TYPES = {
    "com.sun.star.sheet.SpreadsheetDocument",
    "com.sun.star.text.TextDocument",
    "com.sun.star.presentation.PresentationDocument",
    "com.sun.star.drawing.DrawingDocument",
    "com.sun.star.sdb.DocumentDataSource",
    "com.sun.star.formula.FormulaProperties",
    "com.sun.star.script.BasicIDE",
    "com.sun.star.text.WebDocument",  # Supposedly deprecated? But still around.
}


class UnoClient:
    """An RPC client for Unoserver"""

    def __init__(self, server="127.0.0.1", port="2003", host_location="auto"):
        self.server = server
        self.port = port
        if host_location == "auto":
            if server in ("127.0.0.1", "localhost"):
                self.remote = False
            else:
                self.remote = True
        elif host_location == "remote":
            self.remote = True
        elif host_location == "local":
            self.remote = False
        else:
            raise RuntimeError("host_location can be 'auto', 'remote', or 'local'")

    def _connect(self, proxy, retries=5, sleep=10):
        """Check the connection to the proxy multiple times

        Returns the info() data (unoserver version + filters)"""

        while retries > 0:
            try:
                info = proxy.info()
                if not info["api"] == API_VERSION:
                    raise RuntimeError(
                        f"API Version mismatch. Client {__version__} uses API {API_VERSION} "
                        f"while Server {info['unoserver']} uses API {info['api']}."
                    )
                return info
            except ConnectionError as e:
                logger.debug(f"Error {e.strerror}, waiting...")
                retries -= 1
                if retries > 0:
                    time.sleep(sleep)
                    logger.debug("Retrying...")
                else:
                    raise

    def convert(
        self,
        inpath=None,
        indata=None,
        outpath=None,
        convert_to=None,
        filtername=None,
        filter_options=[],
        update_index=True,
        infiltername=None,
    ):
        """Converts a file from one type to another

        inpath: A path (on the local hard disk) to a file to be converted.

        indata: A byte string containing the file content to be converted.

        outpath: A path (on the local hard disk) to store the result, or None, in which case
                 the content of the converted file will be returned as a byte string.

        convert_to: The extension of the desired file type, ie "pdf", "xlsx", etc.

        filtername: The name of the export filter to use for conversion. If None, it is auto-detected.

        update_index: Updates the index before conversion
        """
        if inpath is None and indata is None:
            raise RuntimeError("Nothing to convert.")

        if inpath is not None and indata is not None:
            raise RuntimeError("You can only pass in inpath or indata, not both.")

        if convert_to is None:
            if outpath is None:
                raise RuntimeError(
                    "If you don't specify an output path, you must specify a file-type."
                )
            else:
                convert_to = os.path.splitext(outpath)[-1].strip(os.path.extsep)

        if inpath:
            if self.remote:
                with open(inpath, "rb") as infile:
                    indata = infile.read()
                    inpath = None
            else:
                inpath = os.path.abspath(inpath)

        if outpath:
            outpath = os.path.abspath(outpath)
            if os.path.isdir(outpath):
                raise ValueError("The outpath can not be a directory")

        with ServerProxy(f"http://{self.server}:{self.port}", allow_none=True) as proxy:
            logger.info("Connecting.")
            logger.debug(f"Host: {self.server} Port: {self.port}")
            info = self._connect(proxy)

            if infiltername and infiltername not in info["import_filters"]:
                existing = "\n".join(sorted(info["import_filters"]))
                logger.critical(
                    f"Unknown import filter: {infiltername}. Available filters:\n{existing}"
                )
                raise RuntimeError("Invalid parameter")

            if filtername and filtername not in info["export_filters"]:
                existing = "\n".join(sorted(info["export_filters"]))
                logger.critical(
                    f"Unknown export filter: {filtername}. Available filters:\n{existing}"
                )
                raise RuntimeError("Invalid parameter")

            logger.info("Converting.")
            result = proxy.convert(
                inpath,
                indata,
                None if self.remote else outpath,
                convert_to,
                filtername,
                filter_options,
                update_index,
                infiltername,
            )
            if result is not None:
                # We got the file back over xmlrpc:
                if outpath:
                    logger.info(f"Writing to {outpath}.")
                    with open(outpath, "wb") as outfile:
                        outfile.write(result.data)
                else:
                    # Return the result as a blob
                    logger.info(f"Returning {len(result.data)} bytes.")
                    return result.data
            else:
                logger.info(f"Saved to {outpath}.")

    def compare(
        self,
        oldpath=None,
        olddata=None,
        newpath=None,
        newdata=None,
        outpath=None,
        filetype=None,
    ):
        """Compare two files and convert the result from one type to another.

        newpath: A path (on the local hard disk) to a file to be compared.

        newdata: A byte string containing the file content to be compared.

        oldpath: A path (on the local hard disk) to another file to be compared.

        olddata: A byte string containing the other file content to be compared.

        outpath: A path (on the local hard disk) to store the result, or None, in which case
                 the content of the converted file will be returned as a byte string.

        convert_to: The extension of the desired file type, ie "pdf", "xlsx", etc.
        """
        if (newpath is None and newdata is None) or (
            oldpath is None and olddata is None
        ):
            raise RuntimeError(
                "Nothing to be compared. You mast pass in newpath or newdata and oldpath or olddata."
            )

        if newpath is not None and newdata is not None:
            raise RuntimeError("You can only pass in newpath or newdata, not both.")

        if oldpath is not None and olddata is not None:
            raise RuntimeError("You can only pass in oldpath or olddata, not both.")

        if outpath is None and filetype is None:
            raise RuntimeError(
                "If you don't specify an resulting filepath, you must specify a file-type."
            )
        elif filetype is None:
            filetype = os.path.splitext(outpath)[-1].strip(os.path.extsep)

        if self.remote:
            if oldpath:
                with open(oldpath, "rb") as infile:
                    olddata = infile.read()
                    oldpath = None

            if newpath:
                with open(newpath, "rb") as infile:
                    newdata = infile.read()
                    newpath = None

        if oldpath:
            oldpath = os.path.abspath(oldpath)
        if newpath:
            newpath = os.path.abspath(newpath)

        with ServerProxy(f"http://{self.server}:{self.port}", allow_none=True) as proxy:
            logger.info("Connecting.")
            logger.debug(f"Host: {self.server} Port: {self.port}")
            self._connect(proxy)

            logger.info("Comparing.")
            result = proxy.compare(
                oldpath,
                olddata,
                newpath,
                newdata,
                outpath if self.server in ("127.0.0.1", "localhost") else None,
                filetype,
            )
            if result is not None:
                # We got the file back over xmlrpc:
                if outpath:
                    logger.info(f"Writing to {outpath}.")
                    with open(outpath, "wb") as outfile:
                        outfile.write(result.data)
                else:
                    # Return the result as a blob
                    logger.info(f"Returning {len(result.data)} bytes.")
                    return result.data
            else:
                logger.info(f"Saved to {outpath}.")


def converter_main():
    logging.basicConfig()

    parser = argparse.ArgumentParser("unoconvert")
    parser.add_argument(
        "infile", help="The path to the file to be converted (use - for stdin)"
    )
    parser.add_argument(
        "outfile", help="The path to the converted file (use - for stdout)"
    )
    parser.add_argument(
        "-v",
        "--version",
        action="version",
        help="Display version and exit.",
        version=f"{parser.prog} {__version__}",
    )
    parser.add_argument(
        "--convert-to",
        help="The file type/extension of the output file (ex pdf). Required when using stdout",
    )
    parser.add_argument(
        "--input-filter",
        help="The LibreOffice input filter to use (ex 'writer8'), if autodetect fails",
    )
    parser.add_argument(
        "--output-filter",
        "--filter",
        default=None,
        help="The export filter to use when converting. It is selected automatically if not specified.",
    )
    parser.add_argument(
        "--filter-options",
        "--filter-option",
        default=[],
        action="append",
        help="Pass an option for the export filter, in name=value format. Use true/false for boolean values. "
        "Can be repeated for multiple options.",
    )
    parser.add_argument(
        "--update-index",
        action="store_true",
        help="Updates the indexes before conversion. Can be time consuming.",
    )
    parser.add_argument(
        "--dont-update-index",
        action="store_false",
        dest="update_index",
        help="Skip updating the indexes.",
    )
    parser.set_defaults(update_index=True)
    parser.add_argument(
        "--host", default="127.0.0.1", help="The host the server runs on"
    )
    parser.add_argument("--port", default="2003", help="The port used by the server")
    parser.add_argument(
        "--host-location",
        default="auto",
        choices=["auto", "remote", "local"],
        help="The host location determines the handling of files. If you run the client on the "
        "same machine as the server, it can be set to local, and the files are sent as paths. "
        "If they are different machines, it is remote and the files are sent as binary data. "
        "Default is auto, and it will send the file as a path if the host is 127.0.0.1 or "
        "localhost, and binary data for other hosts.",
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        dest="verbose",
        help="Increase informational output to stderr.",
    )
    parser.add_argument(
        "--quiet",
        action="store_true",
        dest="quiet",
        help="Decrease informational output to stderr.",
    )
    args = parser.parse_args()

    if args.verbose:
        logger.setLevel(logging.DEBUG)
    elif args.quiet:
        logger.setLevel(logging.CRITICAL)
    else:
        logger.setLevel(logging.INFO)
    if args.verbose and args.quiet:
        logger.debug("Make up your mind, yo!")

    client = UnoClient(args.host, args.port, args.host_location)

    if args.outfile == "-":
        # Set outfile to None, to get the data returned from the function,
        # instead of written to a file.
        args.outfile = None

    if args.infile == "-":
        # Get data from stdin
        indata = sys.stdin.buffer.read()
        args.infile = None
    else:
        indata = None

    result = client.convert(
        inpath=args.infile,
        indata=indata,
        outpath=args.outfile,
        convert_to=args.convert_to,
        filtername=args.output_filter,
        filter_options=args.filter_options,
        update_index=args.update_index,
        infiltername=args.input_filter,
    )

    if args.outfile is None:
        sys.stdout.buffer.write(result)


def comparer_main():
    logging.basicConfig()
    logger.setLevel(logging.INFO)

    parser = argparse.ArgumentParser("unocompare")
    parser.add_argument(
        "oldfile",
        help="The path to the original file to be compared with the modified one (use - for stdin)",
    )
    parser.add_argument(
        "newfile",
        help="The path to the modified file to be compared with the original one (use - for stdin)",
    )
    parser.add_argument(
        "outfile",
        help="The path to the result of the comparison and converted file (use - for stdout)",
    )
    parser.add_argument(
        "-v",
        "--version",
        action="version",
        help="Display version and exit.",
        version=f"{parser.prog} {__version__}",
    )
    parser.add_argument(
        "--file-type",
        help="The file type/extension of the result file (ex pdf). Required when using stdout",
    )
    parser.add_argument(
        "--host", default="127.0.0.1", help="The host the server run on"
    )
    parser.add_argument("--port", default="2003", help="The port used by the server")
    parser.add_argument(
        "--host-location",
        default="auto",
        choices=["auto", "remote", "local"],
        help="The host location determines the handling of files. If you run the client on the "
        "same machine as the server, it can be set to local, and the files are sent as paths. "
        "If they are different machines, it is remote and the files are sent as binary data. "
        "Default is auto, and it will send the file as a path if the host is 127.0.0.1 or "
        "localhost, and binary data for other hosts.",
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        dest="verbose",
        help="Increase informational output to stderr.",
    )
    parser.add_argument(
        "--quiet",
        action="store_true",
        dest="quiet",
        help="Decrease informational output to stderr.",
    )
    args = parser.parse_args()

    if args.verbose:
        logger.setLevel(logging.DEBUG)
    elif args.quiet:
        logger.setLevel(logging.CRITICAL)
    else:
        logger.setLevel(logging.INFO)
    if args.verbose and args.quiet:
        logger.debug("Make up your mind, yo!")

    client = UnoClient(args.host, args.port, args.host_location)

    if args.outfile == "-":
        # Set outfile to None, to get the data returned from the function,
        # instead of written to a file.
        args.outfile = None

    if args.oldfile == "-" and args.newfile == "-":
        raise RuntimeError("You can't read both files from stdin")

    if args.oldfile == "-":
        # Get data from stdin
        olddata = sys.stdin.buffer.read()
        newdata = None
        args.oldfile = None

    elif args.newfile == "-":
        newdata = sys.stdin.buffer.read()
        olddata = None
        args.newfile = None
    else:
        olddata = newdata = None

    result = client.compare(
        oldpath=args.oldfile,
        olddata=olddata,
        newpath=args.newfile,
        newdata=newdata,
        outpath=args.outfile,
        filetype=args.file_type,
    )

    if args.outfile is None:
        sys.stdout.buffer.write(result)
