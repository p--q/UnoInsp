#!/opt/libreoffice5.2/program/python
# -*- coding: utf-8 -*-
import unohelper
from com.sun.star.io import XOutputStream

g_ImplementationHelper = unohelper.ImplementationHelper()

class TupleOutputStream( unohelper.Base, XOutputStream ):
    # The component must have a ctor with the component context as argument.
    def __init__( self, ctx ):
        self.t = ()
        self.closed = 0

    # idl void closeOutput();
    def closeOutput(self):
        self.closed = 1

    # idl void writeBytes( [in] sequence<byte>seq );
    def writeBytes( self, seq ):
        self.t = self.t + seq      # simply add the incoming tuple to the member

    # idl void flush();
    def flush( self ):
        pass

    # convenience function to retrieve the tuple later (no UNO function, may
    # only be called from python )
    def getTuple( self ):
        return self.t

# add the TupleOutputStream class to the implementation container,
# which the loader uses to register/instantiate the component.
g_ImplementationHelper.addImplementation( \
    TupleOutputStream,"org.openoffice.pyuno.PythonOutputStream",
                        ("com.sun.star.io.OutputStream",),)