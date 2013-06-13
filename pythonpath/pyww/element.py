
import unohelper
from com.sun.star.lang import XComponent
from com.sun.star.ui import XUIElement, XToolPanel
from com.sun.star.ui.UIElementType import TOOLPANEL as UET_TOOLPANEL

from pyww import RESOURCE_NAME

try:
    from com.sun.star.ui import XSidebarPanel, LayoutSize
except:
    class XSidebarPanel(object):
        """ Dummy class for environments do not have sidebar support. """
        pass


class WatchingWindowUIElement(unohelper.Base, XSidebarPanel, XUIElement, XToolPanel, XComponent):
    """ Manages UI elemnt of Watching Window. """
    
    def __init__(self, ctx, frame, parent, is_sidebar=True):
        from pyww.ww import WatchingWindow
        from pyww.view import WatchingWindowView
        try:
            if is_sidebar:
                import pyww.container
                model_container = pyww.container.get_model_container()
                model = model_container.get(frame)
                if not model:
                    model = WatchingWindow(ctx, frame)
                    model_container.add(model, frame)
            else:
                model = WatchingWindow(ctx)
            self.model = model
            self.view = WatchingWindowView(ctx, self.model, frame, parent)
        except Exception as e:
            print(e)
            import traceback
            traceback.print_exc()
    
    # XComponent
    def dispose(self):
        self.view.dispose()
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
    
    # XSidebarPanel
    def getHeightForWidth(self, width):
        return LayoutSize(0, -1, 0)
