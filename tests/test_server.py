"""Unoserver unit tests"""
import os

from unittest import mock
from unoserver import server

TEST_DOCS = os.path.join(os.path.abspath(os.path.split(__file__)[0]), "documents")


@mock.patch("threading.Thread")
@mock.patch("subprocess.Popen")
def test_server_params(popen_mock, thread_mock):
    srv = server.UnoServer(port="2203", uno_port="2202")
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
            f"-env:UserInstallation={srv.user_installation}",
            "--accept=socket,host=127.0.0.1,port=2202,tcpNoDelay=1;urp;StarOffice.ComponentContext",
        ]
    )


@mock.patch("threading.Thread")
@mock.patch("subprocess.Popen")
def test_server_ipv6_params(popen_mock, thread_mock):
    srv = server.UnoServer(interface="::", port="2203", uno_port="2202")
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
            f"-env:UserInstallation={srv.user_installation}",
            "--accept=socket,host=127.0.0.1,port=2202,tcpNoDelay=1;urp;StarOffice.ComponentContext",
        ]
    )
