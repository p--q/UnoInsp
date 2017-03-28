#!/opt/libreoffice5.2/program/python
# -*- coding: utf-8 -*-
import unohelper
 
# interfaces
from com.sun.star.task import XJob
from com.sun.star.lang import XServiceInfo
 
ImplName = "mytools.basicide.IATest"
ServiceNames = ("com.sun.star.task.Job",)
 
class IATest(unohelper.Base, XJob, XServiceInfo):
    
    def __init__(self,ctx):
        self.ctx = ctx
        print("initialized")
    
    # XJob
    def execute(self, args):
        print("job executed.")
    
    # XServiceInfo
    def getImplementationName(self):
        return ImplName
    
    def supportsService(self, name):
        return name in ServiceNames
    
    def getSupportedServiceNames(self):
        return ServiceNames
 
 
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(
    IATest, ImplName, ServiceNames)