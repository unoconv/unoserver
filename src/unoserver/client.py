import argparse
import logging
import os
import sys

from xmlrpc.client import ServerProxy

logger = logging.getLogger("unoserver")

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

    def __init__(self, server="127.0.0.1", port="2003"):
        self.server = server
        self.port = port

    def convert(
        self,
        inpath=None,
        indata=None,
        outpath=None,
        convert_to=None,
        filtername=None,
        filter_options=[],
        update_index=True,
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

        if inpath and self.server not in ("127.0.0.1", "localhost"):
            with open(inpath, "rb") as infile:
                indata = infile.read()
                inpath = None

        with ServerProxy(f"http://{self.server}:{self.port}", allow_none=True) as proxy:
            result = proxy.convert(
                inpath,
                indata,
                outpath if self.server in ("127.0.0.1", "localhost") else None,
                convert_to,
                filtername,
                filter_options,
                update_index,
            )
            if result is not None:
                # We got the file back over xmlrpc:
                if outpath:
                    with open(outpath, "wb") as outfile:
                        outfile.write(result.data)
                else:
                    # Pipe result to stdout
                    sys.stdout.buffer.write(result.data)

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

        if self.server not in ("127.0.0.1", "localhost"):
            if oldpath:
                with open(oldpath, "rb") as infile:
                    olddata = infile.read()
                    oldpath = None

            if newpath:
                with open(newpath, "rb") as infile:
                    newdata = infile.read()
                    newpath = None

        with ServerProxy(f"http://{self.server}:{self.port}", allow_none=True) as proxy:
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
                    with open(outpath, "wb") as outfile:
                        outfile.write(result.data)
                else:
                    # Pipe result to stdout
                    sys.stdout.buffer.write(result.data)


def converter_main():
    logging.basicConfig()
    logger.setLevel(logging.DEBUG)

    parser = argparse.ArgumentParser("unoconvert")
    parser.add_argument(
        "infile", help="The path to the file to be converted (use - for stdin)"
    )
    parser.add_argument(
        "outfile", help="The path to the converted file (use - for stdout)"
    )
    parser.add_argument(
        "--convert-to",
        help="The file type/extension of the output file (ex pdf). Required when using stdout",
    )
    parser.add_argument(
        "--filter",
        default=None,
        help="The export filter to use when converting. It is selected automatically if not specified.",
    )
    parser.add_argument(
        "--filter-options",
        default=[],
        action="append",
        help="Options for the export filter, in name=value format. Use true/false for boolean values.",
    )
    parser.add_argument(
        "--update-index",
        action="store_true",
        help="Updes the indexes before conversion. Can be time consuming.",
    )
    parser.add_argument(
        "--dont-update-index",
        action="store_false",
        dest="update_index",
        help="Skip updating the indexes.",
    )
    parser.set_defaults(update_index=True)
    parser.add_argument(
        "--interface", default="127.0.0.1", help="The interface used by the server"
    )
    parser.add_argument("--port", default="2003", help="The port used by the server")
    args = parser.parse_args()
    client = UnoClient(args.interface, args.port)

    if args.outfile == "-":
        # Set outfile to None, to get the data returned from the function,
        # instead of written to a file.
        args.outfile = None

    if args.infile == "-":
        # Get data from stdin
        indata = sys.stdin.buffer.read()
        client.convert(
            indata=indata,
            outpath=args.outfile,
            convert_to=args.convert_to,
            filtername=args.filter,
            filter_options=args.filter_options,
            update_index=args.update_index,
        )
    else:
        client.convert(
            inpath=args.infile,
            outpath=args.outfile,
            convert_to=args.convert_to,
            filtername=args.filter,
            filter_options=args.filter_options,
            update_index=args.update_index,
        )


def comparer_main():
    logging.basicConfig()
    logger.setLevel(logging.DEBUG)

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
        "--file-type",
        help="The file type/extension of the result file (ex pdf). Required when using stdout",
    )
    parser.add_argument(
        "--interface", default="127.0.0.1", help="The interface used by the server"
    )
    parser.add_argument("--port", default="2003", help="The port used by the server")
    args = parser.parse_args()

    client = UnoClient(args.interface, args.port)

    if args.outfile == "-":
        # Set outfile to None, to get the data returned from the function,
        # instead of written to a file.
        args.outfile = None

    if args.oldfile == "-":
        # Get data from stdin
        indata = sys.stdin.buffer.read()
        client.compare(
            olddata=indata,
            newpath=args.newfile,
            outpath=args.outfile,
            filetype=args.file_type,
        )
    else:
        client.compare(
            oldpath=args.oldfile,
            newpath=args.newfile,
            outpath=args.outfile,
            filetype=args.file_type,
        )
