#!/opt/libreoffice5.2/program/python
# -*- coding: utf-8 -*-
def macro():
    ctx = XSCRIPTCONTEXT.getComponentContext()
    zf = ctx.getServiceManager().createInstanceWithContext("com.sun.star.packages.zip.ZipFileAccess", ctx)
    import unoinsp
    ins = unoinsp.ObjInsp(XSCRIPTCONTEXT)
    ins.tree(zf)
    
    
if __name__ == "__main__":
    import sys
    import unopy
    XSCRIPTCONTEXT = unopy.connect()
    if not XSCRIPTCONTEXT:
        print("Failed to connect.")
        sys.exit(0)
    sys.exit(macro())