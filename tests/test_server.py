"""Unoserver unit tests"""
import os

from unittest import mock
from unoserver import server

TEST_DOCS = os.path.join(os.path.abspath(os.path.split(__file__)[0]), "documents")


@mock.patch("subprocess.Popen")
def test_server_params(popen_mock):
    srv = server.UnoServer()
    srv.start()
    popen_mock.assert_called_with(
        [
            "libreoffice",
            "--headless",
            "--invisible",
            "--nocrashreport",
            "--nodefault",
            "--nologo",
            "--nofirststartwizard",
            "--norestore",
            f"-env:UserInstallation={srv.tmp_uri}",
            "--accept=socket,host=127.0.0.1,port=2002,tcpNoDelay=1;urp;StarOffice.ComponentContext",
        ]
    )
