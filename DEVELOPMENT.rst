Development Guide
=================

Requirements
------------

In addition to Python 3 and normal development tools, ie, the `build-essentials` package on Ubuntu,
XCode on OS X, etc, you also need to install pre-commit.




Installing for development
--------------------------

Unoserver uses a Makefile as a shortcut for common development tasks. To install Unoserver
for development, simply clone the repository, and then make the development environment.

    $ git clone git@github.com:unoconv/unoserver.git
    $ cd unoserver
    $ virtualenv ve
    $ ve/bin/pip install -e .[devenv]


Code quality
------------




Running tests
-------------

    $ make test


Releasing
---------

For releases we use zest.releaser to release, package and upload to PyPI.
Make sure you have a correct .pypirc so that you can upload packages to PyPI.
The run the `fullrelease` command from zest.releaser, and it will guide you through the process.

    $ fullrelease
