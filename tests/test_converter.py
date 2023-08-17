"""Unoconverter unit tests"""
import os
import pytest
import sys
import uno

from unoserver import converter, client

TEST_DOCS = os.path.join(os.path.abspath(os.path.split(__file__)[0]), "documents")


def test_get_doc_type():
    ctx = uno.getComponentContext()
    smgr = ctx.ServiceManager
    # Pass in the service manager, it doesn't support any known document type
    with pytest.raises(RuntimeError):
        converter.get_doc_type(smgr)


old_import = __builtins__["__import__"]


def new_import(name, *optargs, **kwargs):
    if name == "uno":
        raise ImportError
    else:
        return old_import(name, *optargs, **kwargs)


def test_no_uno(monkeypatch):
    # Patch the import. Mock doesn't work here, because it refuses to
    # deal with __builtins__ as a module, so We can't use mock.patch().
    # Pytests monkeypatch will unmonkeypatch at the end of the test.
    monkeypatch.setitem(__builtins__, "__import__", new_import)
    del sys.modules["uno"]
    del sys.modules["unoserver"]
    del sys.modules["unoserver.converter"]

    # This should now raise an import error
    with pytest.raises(ImportError) as e:
        from unoserver import converter  # noqa: F401

    assert "This package must be installed with a Python" in str(e.value)


def test_wrong_arguments(monkeypatch):
    conv = client.UnoClient()

    with pytest.raises(RuntimeError):
        # You need to pass in an infile, or data
        conv.convert()

    with pytest.raises(RuntimeError):
        # But not both
        conv.convert(inpath="somesortoffile.xls", indata="Shoobx rules!")

    with pytest.raises(RuntimeError):
        # You need to pass in an outfile or a file type
        conv.convert(inpath="somesortoffile.xls")
