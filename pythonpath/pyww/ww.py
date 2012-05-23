
import unohelper

import threading

from com.sun.star.lang import XComponent
from com.sun.star.ui import XUIElement, XToolPanel
from com.sun.star.ui.UIElementType import TOOLPANEL as UET_TOOLPANEL
from com.sun.star.beans import PropertyValue
from com.sun.star.beans.PropertyState import DIRECT_VALUE as PS_DIRECT_VALUE

from pyww import EXT_ID
from pyww import grid
from pyww.settings import Settings, SettingsRDF
from pyww.helper import get_cell_references

class CurrentStringResource(object):
    """ Keeps current string resource. """
    
    Current = None
    
    def get(ctx):
        klass = CurrentStringResource
        if klass.Current is None:
            import pyww.helper
            from pyww import RES_DIR, RES_FILE
            res = pyww.helper.get_current_resource(
                ctx, RES_DIR, RES_FILE)
            klass.Current = pyww.helper.StringResource(res)
        return klass.Current
    
    get = staticmethod(get)


class StopIteration(Exception):
    pass

class WatchingWindow(unohelper.Base, XUIElement, XToolPanel, XComponent):
    """ Watching window model. """
    
    WarnCells = 10
    
    def __init__(self, ctx, frame, parent):
        self.ctx = ctx
        self.frame = frame
        self.parent = parent
        
        self.view = None
        self.window = None
        self.data_model = None
        self.store_watches = False
        
        self._stop_iteration = False
        self.thread = None
        
        try:
            from pyww.view import WatchingWindowView
            
            res = CurrentStringResource.get(ctx)
            self.rdf = SettingsRDF(self.ctx, frame.getController().getModel())
            
            settings = Settings(ctx, res)
            show_input_line = settings.get("InputLine")
            self.WarnCells = settings.get("WarnNumberOfCells")
            if len(self.rdf):
                self.store_watches = True
            else:
                self.store_watches = settings.get("StoreWatches")
            
            view = WatchingWindowView(ctx, self, frame, parent, res, show_input_line)
            self.view = view
            self.data_model = grid.Rows(view.get_data_model(), res)
            self.window = view.cont
            
            self.update_buttons_state()
            
            def _focus_back():
                self.frame.getContainerWindow().setFocus()
            # wait UI creation is finished
            threading.Timer(0.3, _focus_back).start()
            threading.Timer(0.8, self._init).start()
        except Exception as e:
            print(e)
    
    # XComponent
    def dispose(self):
        self.data_model.clear()
        self.data_model = None
        
        self.ctx = None
        self.frame = None
        self.parent = None
        self.view = None
        self.window = None
    
    def addEventListener(self, ev): pass
    def removeEventListener(self, ev): pass
    
    # XUIElement
    def getRealInterface(self):
        return self
    @property
    def Frame(self):
        return self.frame
    @property
    def ResourceURL(self):
        return RESOURCE_NAME
    @property
    def Type(self):
        return UET_TOOLPANEL
    
    # XToolPanel
    def createAccessible(self, parent):
        return self.window.getAccessibleContext()
    @property
    def Window(self):
        return self.window
    
    def set_button_enable(self, name, state):
        """ update enabled state of named button. """
        self.view.update_buttons_state(self, name, state)
    
    def select_entry(self, index):
        """ select specific entry by index. """
        self.view.select_entry(index)
    
    def is_selected_entry_moveable(self, up):
        """ check selected entry is moveable. """
        return self.view.is_selected_entry_moveable(up)
    
    def get_selected_entry_index(self):
        """ get selected. -1 will be returned if not selected. """
        return self.view.get_selected_entry_index()
    
    def update_buttons_state(self):
        """ request to update buttons state. """
        self.view.update_buttons_state()
    
    def add_entry(self):
        """ request to add current selected object. """
        model = self.frame.getController().getModel()
        obj = model.getCurrentSelection()
        if obj.supportsService("com.sun.star.sheet.SheetCell"):
            self.add_cell(obj)
        elif obj.supportsService("com.sun.star.sheet.SheetCellRange"):
            if not self.thread:
                self.thread = threading.Thread(
                    target=self.add_cell_range, args=(obj, False))
                self.thread.start()
            #self.add_cell_range(obj, False)
        elif obj.supportsService("com.sun.star.sheet.SheetCellRanges"):
            self.add_cell_ranges(obj)
        self.update_buttons_state()
    
    
    def _add_entries_started(self):
        self.view.enable_add_watch(False)
        self.view.spinner_start()
        # ToDo disable add button
    
    def _add_entries_finished(self):
        self.view.enable_add_watch(True)
        self.view.spinner_stop()
        # ToDo enable add button
        self._stop_iteration = False
        self.thread = None
    
    def stop_iteration(self):
        self._stop_iteration = True
    
    def add_cell(self, obj, add_to_order=True):
        """ watch cell. """
        n = self.data_model.add_watch(obj)
        if add_to_order and self.store_watches:
            self.watch_added(n)
    
    def add_cell_range(self, obj, checked=False):
        """ watch all cells in the range. """
        addr = obj.getRangeAddress()
        if not checked:
            cells = (addr.EndColumn - addr.StartColumn + 1) * (addr.EndRow - addr.StartRow + 1)
            if cells >= self.WarnCells and self.view.confirm_warn_cells(cells) == 0:
                return
        self._add_entries_started()
        data_model = self.data_model
        try:
            n = 0
            names = []
            for i in range(addr.EndColumn - addr.StartColumn + 1):
                for j in range(addr.EndRow - addr.StartRow + 1):
                    cell = obj.getCellByPosition(i, j)
                    data_model.reserve_watch(cell)
                    names.append(cell.AbsoluteName)
                    n += 1
                    if n == 100:
                        data_model.add_reserved()
                        n = 0
                    if self._stop_iteration:
                        raise StopIteration()
        except StopIteration:
            pass
        except Exception, e:
            print(e)
        data_model.add_reserved()
        self.watch_add_list(names)
        self._add_entries_finished()
    
    def add_cell_ranges(self, obj):
        """ add to watch all cells in the ranges. """
        cells = 0
        ranges = obj.getRangeAddresses()
        for addr in ranges:
            cells += (addr.EndColumn - addr.StartColumn + 1) * (addr.EndRow - addr.StartRow + 1)
        if cells >= self.WarnCells and self.view.confirm_warn_cells(cells) == 0:
            return
        try:
            enume = obj.createEnumeration()
            while enume.hasMoreElements():
                self.add_cell_range(enume.nextElement(), True)
                if self._stop_iteration:
                    raise Exception()
        except Exception as e:
            print(e)
    
    def remove_entry(self):
        """ remove current selected one. """
        index = self.view.get_selected_entry_index()
        if index >= 0:
            name = self.data_model.get(index).get_header()
            self.data_model.remove_watch(index)
            self.update_buttons_state()
            self.view.update_view()
            # select same row or last row
            if self.data_model.getRowCount() <= index:
                index -= 1
            self.view.select_entry(index)
            if self.store_watches:
                self.watch_removed(name)
    
    def remove_all_entries(self):
        """ remove all. """
        self.data_model.remove_all_watch()
        self.update_buttons_state()
        self.view.update_view()
        self.watch_cleared()
    
    def update_all(self):
        """ request to update all. """
        self.data_model.update_all_watch()
    
    def goto_cell(self, desc=""):
        """ move cell cursor to the cell of the selected watch. """
        if not desc:
            index = self.view.get_selected_entry_index()
            if index >= 0:
                row = self.data_model.get(index)
                desc = row.get_header()
        if desc:
            self.goto_point(desc)
    
    def move_entry(self, up):
        """ move selected entry. """
        index = self.view.get_selected_entry_index()
        if up and index > 0:
            n = index -1
            self.data_model.exchange_watches(index, n)
            self.select_entry(n)
        elif not up and index < (self.data_model.getRowCount() - 1):
            n = index +1
            self.data_model.exchange_watches(index, n)
            self.select_entry(n)
            
        self.watch_moved(n)
    
    def goto_point(self, desc):
        """ move cursor to the specified address. """
        arg = PropertyValue("ToPoint", 0, desc, PS_DIRECT_VALUE)
        self.dispatch(".uno:GoToCell", (arg,))
        self.frame.getComponentWindow().setFocus()
    
    def dispatch(self, cmd, args):
        """ dispatch with arguments. """
        helper = self.ctx.getServiceManager().createInstanceWithContext(
            "com.sun.star.frame.DispatchHelper", self.ctx)
        helper.executeDispatch(self.frame, cmd, "_self", 0, args)
        
    
    def get_cell_references(self, cell):
        """ get cell references which are used in the formula. """
        return get_cell_references(self.frame.getController().getModel(), cell)
    
    
    def update_row(self, formula):
        """ update selected row with the formula. """
        index = self.get_selected_entry_index()
        if index >= 0:
            row = self.data_model.get(index)
            row.set_formula(formula)
    
    def set_input_line(self):
        """ set formula of the selected row into the input line. """
        index = self.get_selected_entry_index()
        if index >= 0:
            row = self.data_model.get(index)
            self.view.set_input_line(row.get_formula())
    
    def hidden(self):
        self.data_model.enable_update(False)
    
    def shown(self):
        self.data_model.enable_update(True)
    
    def switch_store_state(self):
        """ Change store state. """
        self.store_watches = not self.store_watches
        if self.store_watches:
            self.rdf.add_list_to_order(self.data_model.get_row_names())
        else:
            self.rdf.clear()
    
    def _init(self):
        self.rdf.load()
        if not len(self.rdf):
            return
        self._add_entries_started()
        n = 0
        data_model = self.data_model
        model = self.frame.getController().getModel()
        sheets = model.getSheets()
        try:
            for name in self.rdf.order:
                try:
                    ranges = sheets.getCellRangesByName(name)
                    if ranges:
                        cell = ranges[0].getCellByPosition(0, 0)
                        data_model.reserve_watch(cell)
                        n += 1
                        if n == 100:
                            data_model.add_reserved()
                            n = 0
                        
                except Exception, e:
                    print(e)
                if self._stop_iteration:
                    raise StopIteration()
        except StopIteration:
            pass
            # ToDo remove non processed
        data_model.add_reserved()
        self._add_entries_finished()
    
    def watch_added(self, n):
        if n >= 0:
            row = self.data_model.get(n)
            if row:
                self.rdf.add(row.get_header())
    
    def watch_add_list(self, names):
        self.rdf.add_list_to_order(names)
    
    def watch_removed(self, name):
        if name:
            self.rdf.remove(name)
    
    def watch_moved(self, n):
        if n >= 0:
            row = self.data_model.get(n)
            if row:
                self.rdf.move(n, row.get_header())
    
    def watch_cleared(self):
        self.rdf.clear()
