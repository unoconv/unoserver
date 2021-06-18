import pytest

from unoserver import server


@pytest.fixture(scope="session")
def unoserver():
    srvr = server.UnoServer()
    process = srvr.start(daemon=True)
    yield process  # provide the fixture value
    print("Teardown Unoserver")
    process.terminate()
