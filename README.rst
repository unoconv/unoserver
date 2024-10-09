unoserver
=========

Using LibreOffice as a server for converting documents.

Overview
--------

Using LibreOffice to convert documents is easy, you can use a command like this to
convert a file to PDF, for example::

    $ libreoffice --headless --convert-to pdf ~/Documents/MyDocument.odf

However, that will load LibreOffice into memory, convert a file and then exit LibreOffice,
which means that the next time you convert a document LibreOffice needs to be loaded into
memory again.

To avoid that, LibreOffice has a listener mode, where it can listen for commands via a port,
and load and convert documents without exiting and reloading the software. This lowers the
CPU load when converting many documents with somewhere between 50% and 75%, meaning you can
convert somewhere between two and four times as many documents in the same time using a listener.

Unoserver contains three commands to help you do this, `unoserver` which starts a listener on the
specified IP interface and port, and `unoconverter` which will connect to a listener and ask it
to convert a document, as well as `unocompare` which will connect to a listener and ask it
to compare two documents and convert the result document.


Installation
------------

NB! Windows and Mac support is as of yet untested.

Unoserver needs to be installed by and run with the same Python installation that LibreOffice uses,
to properly run the `unoserver` command. For client/server installations, see below.

On Unix this usually means you can just install it with::

   $ sudo -H pip install unoserver

If you have multiple versions of LibreOffice installed, you need to install it for each one.
Usually each LibreOffice install will have it's own `python` executable and you need to run
`pip` with that executable::

  $ sudo -H /full/path/to/python -m pip install unoserver

To find all Python installations that have the relevant LibreOffice libraries installed,
you can run a script called `find_uno.py`::

  wget -O find_uno.py https://gist.githubusercontent.com/regebro/036da022dc7d5241a0ee97efdf1458eb/raw/find_uno.py
  python3 find_uno.py

This should give an output similar to this::

  Trying python found at /usr/bin/python3... Success!
  Trying python found at /opt/libreoffice7.1/program/python... Success!
  Found 2 Pythons with Libreoffice libraries:
  /usr/bin/python3
  /opt/libreoffice7.1/program/python

The `/usr/bin/python3` binary will be the system Python used for versions of
Libreoffice installed by the system package manager. The Pythons installed
under `/opt/` will be Python versions that come with official LibreOffice
distributions.

To install on such distributions, do the following::

  $ wget https://bootstrap.pypa.io/get-pip.py
  $ sudo /path/to/python get-pip.py
  $ sudo /path/to/python -m pip install unoserver

You can also install it in a virtualenv, if you are using the system Python
for that virtualenv, and specify the ``--system-site-packages`` parameter::

  $ virtualenv --python=/usr/bin/python3 --system-site-packages virtenv
  $ virtenv/bin/pip install unoserver

Windows and Mac installs aren't officially supported yet, but on Windows the
paths to the LibreOffice Python executable are usually in locations such as
`C:\\Program Files (x86)\\LibreOffice\\python.exe`. On Mac it can be for
example `/Applications/LibreOffice.app/Contents/python` or
`/Applications/LibreOffice.app/Contents/Resources/python`.


Usage
-----

Installing unoserver installs three scripts, `unoserver`, `unoconverter` and `unocompare`.
The server can also be run as a module with `python3 -m unoserver.server`, with the same
arguments as the main script, which can be useful as it must be run with the LibreOffice
provided Python.


Unoserver
~~~~~~~~~

.. code::

  unoserver [-h] [-v] [--interface INTERFACE] [--uno-interface UNO_INTERFACE] [--port PORT] [--uno-port UNO_PORT]
            [--daemon] [--executable EXECUTABLE] [--user-installation USER_INSTALLATION]
            [--libreoffice-pid-file LIBREOFFICE_PID_FILE]

* `--interface`: The interface used by the XMLRPC server, defaults to "127.0.0.1"
* `--port`: The port used by the XMLRPC server, defaults to "2003"
* `--uno-interface`: The interface used by the LibreOffice server, defaults to "127.0.0.1"
* `--uno-port`: The port used by the LibreOffice server, defaults to "2002"
* `--daemon`: Deamonize the server
* `--executable`: The path to the LibreOffice executable
* `--user-installation`: The path to the LibreOffice user profile, defaults to a dynamically created temporary directory
* `--libreoffice-pid-file`: If set, unoserver will write the Libreoffice PID to this file.
  If started in daemon mode, the file will not be deleted when unoserver exits.
* `--conversion-timeout`: Terminate Libreoffice and exit if a conversion does not complete in the given time (in seconds).
* `-v, --version`: Display version and exit.

Unoconvert
~~~~~~~~~~

.. code::

  unoconvert [-h] [-v] [--convert-to CONVERT_TO] [--input-filter INPUT_FILTER] [--output-filter OUTPUT_FILTER]
             [--filter-options FILTER_OPTIONS] [--update-index] [--dont-update-index] [--host HOST] [--port PORT]
             [--host-location {auto,remote,local}] infile outfile

* `infile`: The path to the file to be converted (use - for stdin)
* `outfile`: The path to the converted file (use - for stdout)
* `--convert-to`: The file type/extension of the output file (ex pdf). Required when using stdout
* `--input-filter`: The LibreOffice input filter to use (ex 'writer8'), if autodetect fails
* `--output-filter`: The export filter to use when converting. It is selected automatically if not specified.
* `--filter`: Deprecated alias for `--output-filter`
* `--filter-option`: Pass an option for the export filter, in name=value format. Use true/false for boolean values. Can be repeated for multiple options.
* `--filter-options`: Deprecated alias for `--filter-option`.
* `--host`: The host used by the server, defaults to "127.0.0.1"
* `--port`: The port used by the server, defaults to "2003"
* `--host-location`: The host location determines the handling of files. If you run the client on the
  same machine as the server, it can be set to local, and the files are sent as paths. If they are
  different machines, it is remote and the files are sent as binary data. Default is auto, and it will
  send the file as a path if the host is 127.0.0.1 or localhost, and binary data for other hosts.
* `-v, --version`: Display version and exit.

Example for setting PNG width/height::

  unoconvert infile.odt outfile.png --filter-options PixelWidth=640 --filter-options PixelHeight=480


Unocompare
~~~~~~~~~~

.. code::

  unocompare [-h] [-v] [--file-type FILE_TYPE] [--host HOST] [--port PORT] [--host-location {auto,remote,local}]
             oldfile newfile outfile

* `oldfile`: The path to the older file to be compared with the original one (use - for stdin)
* `newfile`: The path to the newer file to be compared with the modified one (use - for stdin)
* `outfile`: The path to the result of the comparison and converted file (use - for stdout)
* `--file-type`: The file type/extension of the result output file (ex pdf). Required when using stdout
* `--host`: The host used by the server, defaults to "127.0.0.1"
* `--port`: The port used by the server, defaults to "2003"
* `--host-location`: The host location determines the handling of files. If you run the client on the
  same machine as the server, it can be set to local, and the files are sent as paths. If they are
  different machines, it is remote and the files are sent as binary data. Default is auto, and it will
  send the file as a path if the host is 127.0.0.1 or localhost, and binary data for other hosts.
* `-v, --version`: Display version and exit.


Client/Server installations
---------------------------

If you are installing Unoserver on a dedicated machine (virtual or not) to do the conversions and
are running the commands from a different machine, or if you want to call the convert/compare commands
from Python directly, the clients do not need access to Libreoffice. You can therefore follow the
instructions above to make Unoserver have access to the LibreOffice library, but on the client
side you can simply install Unoserver as any other Python library, with `python -m pip install unoserver`
using the Python you want to use as the client executable.

Please note that there is no security on either ports used, and as a result Unoserver is vulnerable
to DDOS attacks, and possibly worse. The ports used **must not** be accessible to anything outside the
server stack being used.

Unoserver is designed to be started by some service management software, such as Supervisor or similar,
that will restart the service should it crash. Unoserver does not try to restart LibreOffice if it
crashes, but should instead also stop in that sitution. The ``--conversion-timeout`` argument will
teminate LibreOffice if it takes to long to convert a document, and that termination will also result
in Unoserver quitting. Because of this service monitoring software should be set up to restart
Unoserver when it exits.


Development and Testing
-----------------------

1. Clone the repo from `https://github.com/unoconv/unoserver`.

2. Setup a virtualenv::

    $ virtualenv --system-site-packages ve
    $ ve/bin/pip install -e .[devenv]

3. Run tests::

    $ ve/bin/pytest tests

4. Run `flake8` linting:

    $ ve/bin/flake8 src tests


Comparison with `unoconv`
-------------------------

Unoserver started as a rewrite, and hopefully a replacement to `unoconv`, a module with support
for using LibreOffice as a listener to convert documents.

Differences for the user
~~~~~~~~~~~~~~~~~~~~~~~~

* Easier install for system versions of LibreOffice. On Linux, the packaged versions of LibreOffice
  typically uses the system Python, making it easy to install `unoserver` with a simple
  `sudo pip install unoserver` command.

* Separate commands for server and client. The client no longer tries to start a listener and then
  close it after conversion if it can't find a listener. Instead the new `unoconverter` client
  requires the `unoserver` to be started. This makes it less practical for one-off converts,
  but as mentioned that can easily be done with LibreOffice itself.

* The `unoserver` listener does not prevent you from using LibreOffice as a normal user, while the
  `unoconv` listener would block you from starting LibreOffice to open a document normally.

* You should be able to on a multi-core machine run several `unoservers` with different ports.
  There is however no support for any form of load balancing in `unoserver`, you would have to
  implement that yourself in your usage of `unoconverter`. For performant multi-core scaling, it
  is necessary to specify unique values for each `unoserver`'s `--port` and `--uno-port` options.

* Only LibreOffice is officially supported. Other variations are untested.


Differences for the maintainer
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

* It's a complete and clean rewrite, supporting only Python 3, with easier to understand and
  therefore easier to maintain code, hopefully meaning more people can contribute.

* It doesn't rely on internal mappings of file types and export filters, but asks LibreOffice
  for this information, which will increase compatibility with different LibreOffice versions,
  and also lowers maintenance.
