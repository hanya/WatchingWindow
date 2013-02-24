
import unohelper
from com.sun.star.lang import XComponent
from com.sun.star.ui import XUIElement, XToolPanel
from com.sun.star.ui.UIElementType import TOOLPANEL as UET_TOOLPANEL

class WatchingWindowUIElement(unohelper.Base, XUIElement, XToolPanel, XComponent):
    """ Manages UI elemnt of Watching Window. """
    
    def __init__(self, ctx, frame, parent):
        from pyww.ww import WatchingWindow
        from pyww.view import WatchingWindowView
        try:
            self.model = WatchingWindow(ctx)
            self.view = WatchingWindowView(
                ctx, self.model, frame, parent)
            self.model.set_view(self.view)
        except Exception as e:
            print(e)
            import traceback
            traceback.print_exc()
    
    # XComponent
    def dispose(self):
        self.view.dispose()
        self.model.dispose()
        self.view = None
        self.model = None
    
    def addEventListener(self, ev): pass
    
    def removeEventListener(self, ev): pass
    
    # XUIElement
    def getRealInterface(self):
        return self
    
    @property
    def Frame(self):
        if self.view:
            return self.view.frame
    
    @property
    def ResourceURL(self):
        from pyww import RESOURCE_NAME
        return RESOURCE_NAME
    
    @property
    def Type(self):
        return UET_TOOLPANEL
    
    # XToolPanel
    def createAccessible(self, parent):
        if self.view:
            return self.view.cont.getAccessibleContext()
    
    @property
    def Window(self):
        if self.view:
            return self.view.cont
