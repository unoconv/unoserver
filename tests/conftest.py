import pytest
import time
import tempfile
from pathlib import Path

from unoserver import server


@pytest.fixture(scope="session")
def server_fixture():
    with tempfile.TemporaryDirectory() as tmpuserdir:
        user_installation = Path(tmpuserdir).as_uri()
        srvr = server.UnoServer(user_installation=user_installation)
        process = srvr.start()
        # Give libreoffice a chance to start
        time.sleep(8)
        yield process  # provide the fixture value
        print("Teardown Unoserver")
        srvr.stop()
