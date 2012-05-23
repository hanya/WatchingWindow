# -*- coding: utf-8 -*-
"""
Copyright (c) 2011 Tsutomu Uchino All rights reserved.

Permission is hereby granted, free of charge, to any person obtaining 
a copy of this software and associated documentation files (the "Software"), 
to deal in the Software without restriction, including without limitation 
the rights to use, copy, modify, merge, publish, distribute, sublicense, 
and/or sell copies of the Software, and to permit persons to whom the Software 
is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included 
in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS 
OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, 
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL 
THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES 
OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, 
ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE 
OR OTHER DEALINGS IN THE SOFTWARE.
"""

import unohelper

from com.sun.star.ui import XUIElementFactory

IMPL_NAME = "mytools.calc.WatchWindow"
RESOURCE_NAME = "private:resource/toolpanel/mytools.calc/WatchWindow"

class WatchingWindowFactory(unohelper.Base, XUIElementFactory):
    """ Factory for watching window. """
    def __init__(self, ctx):
        self.ctx = ctx
    
    # XUIElementFactory
    def createUIElement(self, name, args):
        element = None
        if name == RESOURCE_NAME:
            frame = None
            parent = None
            for arg in args:
                if arg.Name == "Frame":
                    frame = arg.Value
                elif arg.Name == "ParentWindow":
                    parent = arg.Value
            if frame and parent:
                # should be checked what kind of document is
                try:
                    import pyww.ww
                    element = pyww.ww.WatchingWindow(self.ctx, frame, parent)
                except Exception as e:
                    print(e)
        return element
    
    # XServiceInfo
    def getImplementationName(self):
        return IMPL_NAME
    def supportsService(self, name):
        return IMPL_NAME == name
    def supportedServiceNames(self):
        return (IMPL_NAME,)


g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(
    WatchingWindowFactory, IMPL_NAME, (IMPL_NAME,),)
