import argparse
import os
import sys
import uno

from com.sun.star.beans import PropertyValue

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


class UnoConverter:
    def __init__(self, interface="127.0.0.1", port="2002"):
        print("Starting unoconverter.")

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
        self.filters = self.service.createInstanceWithContext(
            "com.sun.star.document.FilterFactory", self.context
        )
        self.types = self.service.createInstanceWithContext(
            "com.sun.star.document.TypeDetection", self.context
        )

    def get_doc_type(self, doc):
        for t in DOC_TYPES:
            if doc.supportsService(t):
                return t
        return ""

    def find_filter(self, import_type, export_type):
        # List export filters. You can only search on module, iflags and eflags,
        # so the import and export types we have to test in a loop
        export_filters = self.filters.createSubSetEnumerationByQuery(
            "getSortedFilterList():iflags=2"
        )

        candidates = []
        while export_filters.hasMoreElements():
            # Filter DocumentService here
            export_filter = prop2dict(export_filters.nextElement())
            if export_filter["DocumentService"] != import_type:
                continue
            if export_filter["Type"] != export_type:
                continue

            candidates.append(export_filter)

        if len(candidates) == 1:
            return candidates[0]["Name"]

        if len(candidates) == 0:
            print(
                "Unknown file extension {ext}, please specify type with --export-type"
            )
            sys.exit(1)

        print("Ambigous file extension, please specify type with --export-type")
        print("Examples:")
        [print(e.Type) for e in candidates]
        sys.exit(1)

    def convert(self, infile, outfile, export_filter=None):

        ### Prepare some things

        export_path = uno.systemPathToFileUrl(outfile)
        if not export_filter:
            export_type = self.types.queryTypeByURL(export_path)
            if not export_type:
                print(
                    "Unknown export file type, please specify type with --export-type"
                )

        # TODO: Verify that infile exists and is openable, and that outdir exists, because uno's
        # exceptions are completely useless!

        ### Load the document
        import_path = uno.systemPathToFileUrl(infile)
        # This returned None if the file was locked, I'm hoping the ReadOnly flag avoids that.
        print(f"Opening {infile}")
        document = self.desktop.loadComponentFromURL(
            import_path, "_default", 0, (PropertyValue(Name="ReadOnly", Value=True),)
        )

        try:
            # Figure out document type:
            import_type = self.get_doc_type(document)
            filtername = self.find_filter(import_type, export_type)

            print(f"Exporting to {outfile}")
            print(f"Using {filtername} export filter")

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


if __name__ == "__main__":
    main()
