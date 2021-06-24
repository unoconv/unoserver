try:
    import uno
except ImportError:
    raise ImportError(
        "Could not find the 'uno' library. This package must be installed with a Python "
        "installation that has a 'uno' library. This typically means you should install"
        "it with the same Python executable as your Libreoffice installation uses."
    )

import argparse
import logging
import os
import sys

from com.sun.star.beans import PropertyValue

logging.basicConfig()
logger = logging.getLogger("unoserver")
logger.setLevel(logging.DEBUG)

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
        "Please create an issue at https://github.com/unoconv/unoserver ."
    )


class UnoConverter:
    def __init__(self, interface="127.0.0.1", port="2002"):
        logger.info("Starting unoconverter.")

        self.local_context = uno.getComponentContext()
        self.resolver = self.local_context.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", self.local_context
        )
        self.context = self.resolver.resolve(
            "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext"
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
        # List export filters. You can only search on module, iflags and eflags,
        # so the import and export types we have to test in a loop
        export_filters = self.filter_service.createSubSetEnumerationByQuery(
            "getSortedFilterList():iflags=2"
        )

        while export_filters.hasMoreElements():
            # Filter DocumentService here
            export_filter = prop2dict(export_filters.nextElement())
            if export_filter["DocumentService"] != import_type:
                continue
            if export_filter["Type"] != export_type:
                continue

            # There is only one possible filter per import and export type,
            # so the first one we find is correct
            return export_filter["Name"]

        # No filter found
        logger.error("Unknown file extension {ext}´")
        sys.exit(1)

    def convert(self, infile, outfile):

        # Prepare some things
        export_path = uno.systemPathToFileUrl(os.path.abspath(outfile))
        export_type = self.type_service.queryTypeByURL(export_path)
        if not export_type:
            logger.error(
                f"Unknown export file type, unkown extension {os.path.splitext(outfile)[-1]}"
            )

        # TODO: Verify that infile exists and is openable, and that outdir exists, because uno's
        # exceptions are completely useless!

        # Load the document
        import_path = uno.systemPathToFileUrl(os.path.abspath(infile))
        # This returned None if the file was locked, I'm hoping the ReadOnly flag avoids that.
        logger.info(f"Opening {infile}")
        document = self.desktop.loadComponentFromURL(
            import_path, "_default", 0, (PropertyValue(Name="ReadOnly", Value=True),)
        )

        try:
            # Figure out document type:
            import_type = get_doc_type(document)
            filtername = self.find_filter(import_type, export_type)

            logger.info(f"Exporting to {outfile}")
            logger.info(f"Using {filtername} export filter")

            args = (
                PropertyValue(Name="FilterName", Value=filtername),
                PropertyValue(Name="Overwrite", Value=True),
            )
            document.storeToURL(export_path, args)

        finally:
            document.close(True)


def main():
    parser = argparse.ArgumentParser("unoconverter")
    parser.add_argument(
        "--interface", default="127.0.0.1", help="The interface used by the server"
    )
    parser.add_argument("--port", default="2002", help="The port used by the server")
    parser.add_argument(
        "--infile", required=True, help="The path to the file to be converted"
    )
    parser.add_argument(
        "--outfile", required=True, help="The path to the converted file"
    )
    args = parser.parse_args()

    converter = UnoConverter(args.interface, args.port)
    converter.convert(args.infile, args.outfile)
