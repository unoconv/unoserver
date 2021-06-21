unoserver
=========

Using LibreOffice as a server for converting documents.

Overview
--------

Using LibreOffice to convert documents is easy, you can use a command like this to
convert a file to PDF, for example:

    $ libreoffice --headless --convert-to pdf ~/Documents/MyDocument.odf

However, that will load LibreOffice into memory, convert a file and then exit LibreOffice,
which means that the next time you convert a document LibreOffice needs to be loaded into
memory again.

To avoid that, LibreOffice has a listener mode, where it can listen for commands via a port,
and load and convert documents without exiting and reloading the software. This lowers the
CPU load when converting many documents with somewhere between 50% and 75%, meaning you can
convert somewhere between two and four times as many documents in the same time using a listener.

Unoserver contains two commands to help you do this, `unoserver` which starts a listener on the
specified IP interface and port, and `unoconverter` which will connect to a listener and ask it
to convert a document.


Installation
------------

NB! Windows and Mac support is as of yet untested.

Unoserver needs to be installed by and run with the same Python installation that LibreOffice uses.
On Unix this usually means you can just install it with:

   $ sudo pip install unoserver

If you have multiple versions of LibreOffice installed, you need to install it for each one.
Usually each LibreOffice install will have it's own `python` executable and you need to run
`pip` with that executable:

  $ sudo /full/path/to/python -m pip install unoserver

On Windows the paths to the LibreOffice Python executable are usually in locations such as
`C:\Program Files (x86)\LibreOffice\python.exe`. On Mac it can be for example
`/Applications/LibreOffice.app/Contents/python`.


Usage
-----

Unoserver
~~~~~~~~~

`unoserver [-h] [--interface INTERFACE] [--port PORT]`

  * `--interface`: The interface used by the server, defaults to "localhost"
  * `--port`: The port used by the server, defaults to "2002"

Unoconvert
~~~~~~~~~~

`unoconverter [-h] [--interface INTERFACE] [--port PORT] --infile INFILE --outfile OUTFILE`

  * `--interface`: The interface used by the server, defaults to "localhost"
  * `--port`: The port used by the server, defaults to "2002"
  * `--infile`: The path to the file to be converted
  * `--outfile`: The path to the converted file


Comparison with `unoconv`
-------------------------

Unoserver started as a rewrite, and hopefully a replacement to `unoconv`, a module with support
for using LibreOffice as a listener to convert documents.

Differences for the user
~~~~~~~~~~~~~~~~~~~~~~~~

* Easier install for system versions of LibreOffice. On Linux, the apckaged versions of LibreOffice
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
  implement that yourself in your usage of `unoconverter`.

* Only LibreOffice is officially supported. Other variations are untested.


Differences for the maintainer
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

* It's a complete and clean rewrite, supporting only Python 3, with easier to understand and
  therefore easier to maintain code, hopefully meaning more people can contribute.

* It doesn't rely on internal mappings of file types and export filters, but asks LibreOffice
  for this information, which will increase compatibility with different LibreOffice versions,
  and also lowers maintenance.