import argparse
import uno


from com.sun.star.beans import PropertyValue


def converter(infile, outfile, interface="127.0.0.1", port="2002"):
    print("Starting unoconverter.")

    # Get the uno component context from the PyUNO runtime
    localContext = uno.getComponentContext()
    resolver = localContext.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", localContext
    )

    ctx = resolver.resolve(
        "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext"
    )

    # Declare the ServiceManager
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

    path = uno.systemPathToFileUrl(infile)
    document = desktop.loadComponentFromURL(path, "_default", 0, ())

    print(f"To {path}")
    path = uno.systemPathToFileUrl(outfile)

    args = (
        PropertyValue(Name="FilterName", Value="writer_pdf_Export"),
        PropertyValue(Name="Overwrite", Value=True),
    )
    document.storeToURL(path, args)


def main():
    parser = argparse.ArgumentParser("unoserver")
    parser.add_argument(
        "--interface", default="127.0.0.1", help="The interface used by the server"
    )
    parser.add_argument("--port", default="2002", help="The port used by the server")
    parser.add_argument(
        "--infile", required=True, help="The path to the file to be converted"
    )
    parser.add_argument(
        "--outfile", required=True, help="The path to the converted file"
    )
    args = parser.parse_args()

    converter(args.infile, args.outfile, args.interface, args.port)


if __name__ == "__main__":
    main()
