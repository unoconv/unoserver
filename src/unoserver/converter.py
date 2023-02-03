try:
    import uno
except ImportError:
    raise ImportError(
        "Could not find the 'uno' library. This package must be installed with a Python "
        "installation that has a 'uno' library. This typically means you should install "
        "it with the same Python executable as your Libreoffice installation uses."
    )

import argparse
import io
import logging
import os
import sys
import unohelper

from pathlib import Path
from com.sun.star.beans import PropertyValue
from com.sun.star.io import XOutputStream

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


def prop2dict(properties):
    return {p.Name: p.Value for p in properties}


def get_doc_type(doc):
    for t in DOC_TYPES:
        if doc.supportsService(t):
            return t

    # LibreOffice opened it, but it's not one of the known document types.
    # This really should only happen if a future version of LibreOffice starts
    # adding document types, which seems unlikely.
    raise RuntimeError(
        "The input document is of an unknown document type. This is probably a bug.\n"
        "Please create an issue at https://github.com/unoconv/unoserver."
    )


class OutputStream(unohelper.Base, XOutputStream):
    def __init__(self):
        self.buffer = io.BytesIO()

    def closeOutput(self):
        pass

    def writeBytes(self, seq):
        self.buffer.write(seq.value)


class UnoConverter:
    def __init__(self, interface="127.0.0.1", port="2002"):
        logger.info("Starting unoconverter.")

        self.local_context = uno.getComponentContext()
        self.resolver = self.local_context.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", self.local_context
        )
        self.context = self.resolver.resolve(
            f"uno:socket,host={interface},port={port};urp;StarOffice.ComponentContext"
        )
        self.service = self.context.ServiceManager
        self.desktop = self.service.createInstanceWithContext(
            "com.sun.star.frame.Desktop", self.context
        )
        self.filter_service = self.service.createInstanceWithContext(
            "com.sun.star.document.FilterFactory", self.context
        )
        self.type_service = self.service.createInstanceWithContext(
            "com.sun.star.document.TypeDetection", self.context
        )

    def find_filter(self, import_type, export_type):
        for export_filter in self.get_available_export_filters():
            if export_filter["DocumentService"] != import_type:
                continue
            if export_filter["Type"] != export_type:
                continue

            # There is only one possible filter per import and export type,
            # so the first one we find is correct
            return export_filter["Name"]

        # No filter found
        return None

    def get_available_export_filters(self):
        # List export filters. You can only search on module, iflags and eflags,
        # so the import and export types we have to test in a loop
        export_filters = self.filter_service.createSubSetEnumerationByQuery(
            "getSortedFilterList():iflags=2"
        )

        while export_filters.hasMoreElements():
            # Filter DocumentService here
            yield prop2dict(export_filters.nextElement())

    def get_available_filter_names(self):
        return [filter["Name"] for filter in self.get_available_export_filters()]

    def convert(
        self, inpath=None, indata=None, outpath=None, convert_to=None, filtername=None
    ):
        """Converts a file from one type to another

        inpath: A path (on the local hard disk) to a file to be converted.

        indata: A byte string containing the file content to be converted.

        outpath: A path (on the local hard disk) to store the result, or None, in which case
                 the content of the converted file will be returned as a byte string.

        convert_to: The extension of the desired file type, ie "pdf", "xlsx", etc.

        filtername: The name of the export filter to use for conversion. If None, it is auto-detected.
        """
        if inpath is None and indata is None:
            raise RuntimeError("Nothing to convert.")

        if inpath is not None and indata is not None:
            raise RuntimeError("You can only pass in inpath or indata, not both.")

        if outpath is None and convert_to is None:
            raise RuntimeError(
                "If you don't specify an output path, you must specify a file-type."
            )

        input_props = (PropertyValue(Name="ReadOnly", Value=True),)

        if inpath:
            # TODO: Verify that inpath exists and is openable, and that outdir exists, because uno's
            # exceptions are completely useless!

            if not Path(inpath).exists():
                raise RuntimeError(f"Path {inpath} does not exist.")

            # Load the document
            import_path = uno.systemPathToFileUrl(os.path.abspath(inpath))
            # This returned None if the file was locked, I'm hoping the ReadOnly flag avoids that.
            logger.info(f"Opening {inpath}")

        elif indata:
            # The document content is passed in as a byte string
            input_stream = self.service.createInstanceWithContext(
                "com.sun.star.io.SequenceInputStream", self.context
            )
            input_stream.initialize((uno.ByteSequence(indata),))
            input_props += (PropertyValue(Name="InputStream", Value=input_stream),)
            import_path = "private:stream"

        document = self.desktop.loadComponentFromURL(
            import_path, "_default", 0, input_props
        )

        # Now do the conversion
        try:
            # Figure out document type:
            import_type = get_doc_type(document)

            if outpath:
                export_path = uno.systemPathToFileUrl(os.path.abspath(outpath))
            else:
                export_path = "private:stream"

            # Figure out the output type:
            if convert_to:
                export_type = self.type_service.queryTypeByURL(
                    f"file:///dummy.{convert_to}"
                )
            else:
                export_type = self.type_service.queryTypeByURL(export_path)

            if not export_type:
                if convert_to:
                    extension = convert_to
                else:
                    extension = os.path.splitext(outpath)[-1]
                raise RuntimeError(
                    f"Unknown export file type, unknown extension '{extension}'"
                )

            if filtername is not None:
                available_filter_names = self.get_available_filter_names()
                if filtername not in available_filter_names:
                    raise RuntimeError(
                        f"'{filtername}' is not a valid filter name. Valid filters are {available_filter_names}"
                    )
            else:
                filtername = self.find_filter(import_type, export_type)
                if filtername is None:
                    raise RuntimeError(
                        f"Could not find an export filter from {import_type} to {export_type}"
                    )

            logger.info(f"Exporting to {outpath}")
            logger.info(f"Using {filtername} export filter")

            output_props = (
                PropertyValue(Name="FilterName", Value=filtername),
                PropertyValue(Name="Overwrite", Value=True),
            )
            if outpath is None:
                output_stream = OutputStream()
                output_props += (
                    PropertyValue(Name="OutputStream", Value=output_stream),
                )
            document.storeToURL(export_path, output_props)

        finally:
            document.close(True)

        if outpath is None:
            return output_stream.buffer.getvalue()
        else:
            return None


def main():
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
        "--interface", default="127.0.0.1", help="The interface used by the server"
    )
    parser.add_argument("--port", default="2002", help="The port used by the server")
    args = parser.parse_args()

    converter = UnoConverter(args.interface, args.port)

    if args.outfile == "-":
        # Set outfile to None, to get the data returned from the function,
        # instead of written to a file.
        args.outfile = None

    if args.infile == "-":
        # Get data from stdin
        indata = sys.stdin.buffer.read()
        result = converter.convert(
            indata=indata,
            outpath=args.outfile,
            convert_to=args.convert_to,
            filtername=args.filter,
        )
    else:
        result = converter.convert(
            inpath=args.infile,
            outpath=args.outfile,
            convert_to=args.convert_to,
            filtername=args.filter,
        )

    if args.outfile is None:
        # Pipe result to stdout
        sys.stdout.buffer.write(result)


if __name__ == "__main__":
    main()
