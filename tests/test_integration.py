"""Tests that start a real unoserver and does real things"""

import io
import os
import pytest
import sys
import tempfile
import time

from unoserver import converter, server


TEST_DOCS = os.path.join(os.path.abspath(os.path.split(__file__)[0]), "documents")


@pytest.mark.parametrize("filename", ["simple.odt", "simple.xlsx"])
def test_pdf_conversion(server_fixture, filename):
    infile = os.path.join(TEST_DOCS, filename)

    with tempfile.NamedTemporaryFile(suffix=".pdf") as outfile:
        # Let Libreoffice write to the file and close it.
        sys.argv = ["unoconverter", infile, outfile.name]
        converter.main()

        # We now open it to check it, we can't use the outfile object,
        # it won't reflect the external changes.
        with open(outfile.name, "rb") as testfile:
            start = testfile.readline()
            assert start == b"%PDF-1.5\n"


class FakeStdio(io.BytesIO):
    """A BytesIO with a buffer attribute, usable to send binary stdin data"""

    @property
    def buffer(self):
        return self


@pytest.mark.parametrize("filename", ["simple.odt", "simple.xlsx"])
def test_stdin_stdout(server_fixture, monkeypatch, filename):

    with open(os.path.join(TEST_DOCS, filename), "rb") as infile:
        infile_stream = FakeStdio(infile.read())

    outfile_stream = FakeStdio()

    monkeypatch.setattr("sys.stdin", infile_stream)
    monkeypatch.setattr("sys.stdout", outfile_stream)

    sys.argv = ["unoconverter", "-", "-", "--convert-to", "pdf"]
    converter.main()

    outfile_stream.seek(0)
    start = outfile_stream.readline()
    assert start == b"%PDF-1.5\n"


def test_csv_conversion(server_fixture):
    conv = converter.UnoConverter()
    infile = os.path.join(TEST_DOCS, "simple.xlsx")

    with tempfile.NamedTemporaryFile(suffix=".csv") as outfile:
        # Let Libreoffice write to the file and close it.
        conv.convert(inpath=infile, outpath=outfile.name)
        # We now open it to check it, we can't use the outfile object,
        # it won't reflect the external changes.
        with open(outfile.name, "rb") as testfile:
            contents = testfile.readline()
            assert contents == b"This,Is,A,Simple,Excel,File\n"
            contents = testfile.readline()
            assert contents == b"1,2,3,4,5,6\n"


def test_impossible_conversion(server_fixture):
    conv = converter.UnoConverter()
    infile = os.path.join(TEST_DOCS, "simple.odt")

    with tempfile.NamedTemporaryFile(suffix=".xls") as outfile:
        # Let Libreoffice write to the file and close it.
        with pytest.raises(RuntimeError) as e:
            conv.convert(inpath=infile, outpath=outfile.name)
            assert "Could not find an export filter" in e


def test_multiple_servers(server_fixture):
    # The server fixture should already have started a server.
    # Make sure we can start a second one.
    sys.argv = ["unoserver", "--daemon"]
    process = server.main()
    try:
        # Wait for it to start
        time.sleep(5)
        # Make sure the process is still running, meaning return_code is None
        assert process.returncode is None
    finally:
        # Now kill the process
        process.terminate()
        # Wait for it to terminate
        process.wait()
        # And verify that it was killed
        assert process.returncode == 255


def test_unknown_outfile_type(server_fixture):
    infile = os.path.join(TEST_DOCS, "simple.odt")

    with tempfile.NamedTemporaryFile(suffix=".bog") as outfile:
        sys.argv = ["unoconverter", infile, outfile.name]
        # Type detection should fail, as it's not a .doc file:
        with pytest.raises(RuntimeError):
            converter.main()
