#!/opt/libreoffice5.2/program/python
# -*- coding: utf-8 -*-
def macro():
    ctx = XSCRIPTCONTEXT.getComponentContext()
    zf = ctx.getServiceManager().createInstanceWithContext("com.sun.star.packages.zip.ZipFileAccess", ctx)
    import unoinsp
    ins = unoinsp.ObjInsp(XSCRIPTCONTEXT)
#     ins.tree(zf,("core",))  # coreインターフェイスを出力しない。
#     ins.tree(zf)
#     ins.wtree(zf)
#     ins.tree(zf,("core",))
#     ins.wtree(zf,tuple(["core"]))  # coreインターフェイスを出力しない。
    imp = ctx.getServiceManager().createInstanceWithContext("mytools.basicide.IATest", ctx)
    ins.wtree(imp)


    
    
if __name__ == "__main__":
    import sys
    import unopy
    XSCRIPTCONTEXT = unopy.connect()
    if not XSCRIPTCONTEXT:
        print("Failed to connect.")
        sys.exit(0)
    sys.exit(macro())