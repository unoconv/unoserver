import os
import pytest
import tempfile

from unoserver import converter

TEST_DOCS = os.path.join(os.path.abspath(os.path.split(__file__)[0]), "documents")


def test_conversion(server_fixture):
    conv = converter.UnoConverter()
    infile = os.path.join(TEST_DOCS, "simple.odt")

    with tempfile.NamedTemporaryFile(suffix=".pdf") as outfile:
        # Let Libreoffice write to the file and close it.
        conv.convert(infile, outfile.name)
        # We now open it to check it, we can't use the outfile object,
        # it won't reflect the external changes.
        with open(outfile.name, "rb") as testfile:
            start = testfile.readline()
            assert start == b"%PDF-1.5\n"
