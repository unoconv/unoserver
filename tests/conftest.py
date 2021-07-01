import pytest
import time

from unoserver import server


@pytest.fixture(scope="session")
def server_fixture():
    srvr = server.UnoServer()
    process = srvr.start(daemon=True)
    # Give libreoffice a chance to start
    time.sleep(8)
    yield process  # provide the fixture value
    print("Teardown Unoserver")
    process.terminate()
