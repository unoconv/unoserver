"""Tests that start a real unoserver and does real things"""

import io
import os
import pytest
import re
import socket
import subprocess
import sys
import tempfile
import time

from xmlrpc.client import Fault
from unoserver import client


TEST_DOCS = os.path.join(os.path.abspath(os.path.split(__file__)[0]), "documents")


@pytest.mark.parametrize("filename", ["simple.odt", "simple.xlsx"])
def test_pdf_conversion(server_fixture, filename):
    infile = os.path.join(TEST_DOCS, filename)

    with tempfile.NamedTemporaryFile(suffix=".pdf") as outfile:
        # Let Libreoffice write to the file and close it.
        sys.argv = ["unoconverter", infile, outfile.name]
        client.converter_main()

        # We now open it to check it, we can't use the outfile object,
        # it won't reflect the external changes.
        with open(outfile.name, "rb") as testfile:
            start = testfile.readline()
            assert start.startswith(b"%PDF-1.")


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
    client.converter_main()

    outfile_stream.seek(0)
    start = outfile_stream.readline()
    assert start.startswith(b"%PDF-1.")


def test_csv_conversion(server_fixture):
    conv = client.UnoClient()
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
    conv = client.UnoClient()
    infile = os.path.join(TEST_DOCS, "simple.odt")

    with tempfile.NamedTemporaryFile(suffix=".xls") as outfile:
        # Let Libreoffice write to the file and close it.
        with pytest.raises(Fault) as e:
            conv.convert(inpath=infile, outpath=outfile.name)
            assert "Could not find an export filter" in e


def test_multiple_servers(server_fixture):
    # The server fixture should already have started a server.
    # Make sure we can start a second one.
    cmd = ["unoserver", "--uno-port=2102", "--port=2103"]
    process = subprocess.Popen(cmd)
    try:
        # Wait for it to start
        time.sleep(5)
        # Make sure the process is still running, meaning return_code is None
        assert process.returncode is None

        # Make a conversion
        conv = client.UnoClient(port="2103")
        infile = os.path.join(TEST_DOCS, "simple.odt")
        with tempfile.NamedTemporaryFile(suffix=".pdf") as outfile:
            conv.convert(inpath=infile, outpath=outfile.name)

    finally:
        # Now kill the process
        process.terminate()
        # Wait for it to terminate
        process.wait(30)
        # And verify that it was killed
        assert process.returncode == 0


def test_unknown_outfile_type(server_fixture):
    infile = os.path.join(TEST_DOCS, "simple.odt")

    with tempfile.NamedTemporaryFile(suffix=".bog") as outfile:
        sys.argv = ["unoconverter", infile, outfile.name]
        # Type detection should fail, as it's not a .doc file:
        with pytest.raises(Fault) as e:
            client.converter_main()
        assert "Unknown export file type" in e.value.faultString


@pytest.mark.parametrize("filename", ["simple.odt", "simple.xlsx"])
def test_explicit_export_filter(server_fixture, filename):
    infile = os.path.join(TEST_DOCS, filename)

    # We use an extension that's not .pdf to verify that the converter does not auto-detect filter based on extension
    with tempfile.NamedTemporaryFile(suffix=".csv") as outfile:
        sys.argv = [
            "unoconverter",
            "--filter",
            "writer_pdf_Export",
            infile,
            outfile.name,
        ]
        client.converter_main()

        # We now open it to check it, we can't use the outfile object,
        # it won't reflect the external changes.
        with open(outfile.name, "rb") as testfile:
            start = testfile.readline()
            assert start.startswith(b"%PDF-1.")


@pytest.mark.parametrize("filename", ["simple.odt", "simple.xlsx"])
def test_invalid_explicit_export_filter_prints_available_filters(
    caplog, server_fixture, filename
):
    infile = os.path.join(TEST_DOCS, filename)

    # We use an extension that's not .pdf to verify that the converter does not auto-detect filter based on extension
    with tempfile.NamedTemporaryFile(suffix=".csv") as outfile:
        sys.argv = ["unoconverter", "--filter", "asdasdasd", infile, outfile.name]
        try:
            client.converter_main()
        except RuntimeError:
            errstr = caplog.text
            print(errstr)
            print("=" * 30)
            assert "Office Open XML Text" in errstr
            assert "writer8" in errstr
            assert "writer_pdf_Export" in errstr


def test_update_index(server_fixture):
    infile = os.path.join(TEST_DOCS, "index-with-fields.odt")

    with tempfile.NamedTemporaryFile(suffix=".rtf") as outfile:
        # Let Libreoffice write to the file and close it.
        sys.argv = ["unoconverter", infile, outfile.name]
        client.converter_main()

        # We now open it to check it, we can't use the outfile object,
        # it won't reflect the external changes.
        with open(outfile.name, "rb") as testfile:
            # The timestamp in Header 2 should appear exactly twice after update
            matches = re.findall(b"13:18:27", testfile.read())
            assert len(matches) == 2

        with tempfile.NamedTemporaryFile(suffix=".rtf") as outfile:
            # Let Libreoffice write to the file and close it.
            sys.argv = ["unoconverter", "--dont-update-index", infile, outfile.name]
            client.converter_main()

            with open(outfile.name, "rb") as testfile:
                # The timestamp in Header 2 should appear exactly once
                matches = re.findall(b"13:18:27", testfile.read())
                assert len(matches) == 1


def test_convert_not_local():
    hostname = socket.gethostname()
    cmd = ["unoserver", "--uno-port=2104", "--port=2105", f"--interface={hostname}"]
    process = subprocess.Popen(cmd)
    try:
        # Wait for it to start
        time.sleep(5)
        # Make sure the process is still running, meaning return_code is None
        assert process.returncode is None

        # Make a conversion
        infile = os.path.join(TEST_DOCS, "simple.odt")
        with tempfile.NamedTemporaryFile(suffix=".pdf") as outfile:
            sys.argv = [
                "unoconverter",
                "--host",
                hostname,
                "--port=2105",
                infile,
                outfile.name,
            ]
            client.converter_main()

            with open(outfile.name, "rb") as testfile:
                start = testfile.readline()
                assert start.startswith(b"%PDF-1.")

    finally:
        # Now kill the process
        process.terminate()
        # Wait for it to terminate
        process.wait(30)
        # And verify that it was killed
        assert process.returncode == 0


# This currently does not work on Ubuntu 20.04.
def skip_test_compare_not_local():
    hostname = socket.gethostname()
    cmd = ["unoserver", "--uno-port=2104", "--port=2105", f"--interface={hostname}"]
    process = subprocess.Popen(cmd)
    try:
        # Wait for it to start
        time.sleep(5)
        # Make sure the process is still running, meaning return_code is None
        assert process.returncode is None

        # Make a comparison
        infile1 = os.path.join(TEST_DOCS, "simple.odt")
        infile2 = os.path.join(TEST_DOCS, "index-with-fields.odt")
        with tempfile.NamedTemporaryFile(suffix=".pdf") as outfile:
            sys.argv = [
                "unoconverter",
                "--host",
                hostname,
                "--port=2105",
                "--input-filter=odt",
                infile1,
                infile2,
                outfile.name,
            ]
            client.comparer_main()

            with open(outfile.name, "rb") as testfile:
                start = testfile.readline()
                assert start.startswith(b"%PDF-1.")

    finally:
        # Now kill the process
        process.terminate()
        # Wait for it to terminate
        process.wait(30)
        # And verify that it was killed
        assert process.returncode == 0
