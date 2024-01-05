try:
    import uno
except ImportError:
    raise ImportError(
        "Could not find the 'uno' library. This package must be installed with a Python "
        "installation that has a 'uno' library. This typically means you should install "
        "it with the same Python executable as your Libreoffice installation uses."
    )

import io
import logging
import os
import unohelper

from com.sun.star.beans import PropertyValue
from com.sun.star.io import XOutputStream

logger = logging.getLogger("unoserver")

SFX_FILTER_IMPORT = 1
SFX_FILTER_EXPORT = 2
DOC_TYPES = {
    "com.sun.star.text.TextDocument",  # Only support comparing for writer
}


def prop2dict(properties):
    return {p.Name: p.Value for p in properties}


def get_doc_type(doc):
    for t in DOC_TYPES:
        if doc.supportsService(t):
            return t

    # LibreOffice opened it, but it's not one of the supported document types.
    # This really should only happen if a future version of LibreOffice starts
    # adding document types, which seems unlikely.
    raise RuntimeError(
        "The input document is an unsupported document type for comparing.\n"
        "Please create an issue at https://github.com/unoconv/unoserver."
    )


class OutputStream(unohelper.Base, XOutputStream):
    def __init__(self):
        self.buffer = io.BytesIO()

    def closeOutput(self):
        pass

    def writeBytes(self, seq):
        self.buffer.write(seq.value)


class UnoComparer:
    """The class that performs the comparison

    Don't use this directly, instead use the client.UnoComparer.
    """

    def __init__(self, interface="127.0.0.1", port="2002"):
        logger.info("Starting UnoComparer.")

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

    def is_comparable(self, import_type, importOrg_type):
        # List export filters. You can only search on module, iflags and eflags,
        # so the import and export types we have to test in a loop
        export_filters = self.filter_service.createSubSetEnumerationByQuery(
            "getSortedFilterList():iflags=2"
        )

        while export_filters.hasMoreElements():
            # Filter by Filtername here
            export_filter = prop2dict(export_filters.nextElement())
            if export_filter["Type"] != importOrg_type:
                continue
            if export_filter["DocumentService"] != import_type:
                continue

            # There is only one possible DocumentService per import and export type,
            # so the first one we find is correct
            return True

        # No DocumentService found
        return False

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
        return None

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

        inpath: A path (on the local hard disk) to a file to be compared.
        inOrgpath: A path (on the local hard disk) to another file to be compared.

        indata: A byte string containing the file content to be compared.

        outpath: A path (on the local hard disk) to store the result, or None, in which case
                 the content of the converted file will be returned as a byte string.

        filetype: The extension of the desired file type, ie "pdf", "xlsx", etc.
        """
        new_props = (PropertyValue(Name="Hidden", Value=True),)

        if newpath:
            # TODO: Verify that inpath exists and is openable, and that outdir exists, because uno's
            # exceptions are completely useless!

            # Load the document
            logger.info(f"Opening file {newpath}")
            newpath = uno.systemPathToFileUrl(os.path.abspath(newpath))
            # This returned None if the file was locked, I'm hoping the ReadOnly flag avoids that.

        elif newdata:
            # The document content is passed in as a byte string
            new_stream = self.service.createInstanceWithContext(
                "com.sun.star.io.SequenceInputStream", self.context
            )
            new_stream.initialize((uno.ByteSequence(newdata),))
            new_props += (PropertyValue(Name="InputStream", Value=new_stream),)
            newpath = "private:stream"

        new_document = self.desktop.loadComponentFromURL(
            newpath, "_blank", 0, new_props
        )
        new_type = get_doc_type(new_document)

        old_props = (PropertyValue(Name="Hidden", Value=True),)

        if oldpath:
            # TODO: Verify that inpath exists and is openable, and that outdir exists, because uno's
            # exceptions are completely useless!

            # Load the document
            logger.info(f"Opening file {oldpath}")
            oldpath = uno.systemPathToFileUrl(os.path.abspath(oldpath))
            old_props += (PropertyValue(Name="URL", Value=oldpath),)
            # This returned None if the file was locked, I'm hoping the ReadOnly flag avoids that.
            old_type = self.type_service.queryTypeByURL(oldpath)

        elif olddata:
            # The document content is passed in as a byte string
            old_stream = self.service.createInstanceWithContext(
                "com.sun.star.io.SequenceInputStream", self.context
            )
            old_stream.initialize((uno.ByteSequence(newdata),))
            old_props += (PropertyValue(Name="InputStream", Value=new_stream),)
            old_props += (PropertyValue(Name="URL", Value="private:stream"),)
            old_type = self.type_service.queryTypeByDescriptor(old_props, False)[0]

        old_props += (PropertyValue(Name="NoAcceptDialog", Value=True),)

        logger.info(f"Opening original file {oldpath}")

        # Now do the comparison, then the conversion
        try:
            # Figure out document type of import file:
            # Figure out document type of original import file:
            # check that the two type is same
            isComparable = self.is_comparable(new_type, old_type)

            if not isComparable:
                raise RuntimeError("Cannot compare two different type of document!")

            dispatch_helper = self.service.createInstanceWithContext(
                "com.sun.star.frame.DispatchHelper", self.context
            )
            dispatch_helper.executeDispatch(
                new_document.getCurrentController().getFrame(),
                ".uno:CompareDocuments",
                "",
                0,
                old_props,
            )

            if outpath:
                export_path = uno.systemPathToFileUrl(os.path.abspath(outpath))
            else:
                export_path = "private:stream"

            # Figure out the output type:
            if filetype:
                export_type = self.type_service.queryTypeByURL(
                    f"file:///dummy.{filetype}"
                )
            else:
                export_type = self.type_service.queryTypeByURL(export_path)

            if not export_type:
                if filetype:
                    extension = filetype
                else:
                    extension = os.path.splitext(outpath)[-1]
                raise RuntimeError(
                    f"Unknown export file type, unknown extension '{extension}'"
                )

            filtername = self.find_filter(new_type, export_type)
            if filtername is None:
                raise RuntimeError(
                    f"Could not find an export filter from {new_type} to {export_type}"
                )

            logger.info(f"Exporting to {outpath}")
            logger.info(
                f"Using {filtername} export filter from {new_type} to {export_type}"
            )

            output_props = (
                PropertyValue(Name="FilterName", Value=filtername),
                PropertyValue(Name="Overwrite", Value=True),
            )
            if outpath is None:
                output_stream = OutputStream()
                output_props += (
                    PropertyValue(Name="OutputStream", Value=output_stream),
                )
            new_document.storeToURL(export_path, output_props)
            new_document.dispose()

        finally:
            new_document.close(True)

        if outpath is None:
            return output_stream.buffer.getvalue()
        else:
            return None
