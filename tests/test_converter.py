"""Unoconverter unit tests"""
import os
import pytest
import uno

from unoserver import converter

TEST_DOCS = os.path.join(os.path.abspath(os.path.split(__file__)[0]), "documents")


def test_get_doc_type():
    ctx = uno.getComponentContext()
    smgr = ctx.ServiceManager
    # Pass in the service manager, it doesn't support any known document type
    with pytest.raises(RuntimeError):
        converter.get_doc_type(smgr)
