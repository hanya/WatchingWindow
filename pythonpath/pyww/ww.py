
import threading

from pyww import EXT_ID
from pyww import grid
from pyww.settings import Settings, SettingsRDF
from pyww.helper import get_cell_references


class StopIteration(Exception):
    pass

class WatchingWindow(object):
    """ Watching window model. """
    
    def __init__(self, ctx):
        self.ctx = ctx
        self.view = None
        self.rows = None
        self.rdf = None
        
        self._stop_iteration = False
        self._iterating = False
        
        self.store_watches = False
        self.load_settings(init=True)
    
    def set_view(self, view):
        self.view = view
        self.rdf = SettingsRDF(self.ctx, view.get_doc())
        if len(self.rdf):
            self.store_watches = True
        self.rows = grid.Rows(view.get_data_model(), view.res)
        self.update_buttons_state()
        # wait UI creation is finished
        threading.Timer(0.3, self.view.focus_to_doc).start()
        threading.Timer(0.8, self._init).start()
    
    def dispose(self):
        self.rows.clear()
        self.rows = None
        self.ctx = None
        self.rdf = None
        self.view = None
    
    def get_watches_count(self):
        """ Returns number of watches. """
        return self.rows.getRowCount()
    
    def load_settings(self, init=False):
        """ Load configuration. """
        settings = Settings(self.ctx)
        self.warn_cells = settings.get("WarnNumberOfCells")
        if init:
            self.store_watches = settings.get("StoreWatches")
    
    def settings_changed(self):
        """ Request to load settings. """
        self.load_settings()
    
    def select_entry(self, index):
        """ select specific entry by index. """
        self.view.select_entry(index)
    
    def update_buttons_state(self):
        """ request to update buttons state. """
        self.view.update_buttons_state()
    
    def add_entry(self):
        """ request to add current selected object. """
        obj = self.view.get_current_selection()
        if obj.supportsService("com.sun.star.sheet.SheetCell"):
            self.add_cell(obj)
        elif obj.supportsService("com.sun.star.sheet.SheetCellRange"):
            if not self._iterating:
                threading.Timer(
                    0.05, self.add_cell_range, args=(obj, False)).start()
        elif obj.supportsService("com.sun.star.sheet.SheetCellRanges"):
            self.add_cell_ranges(obj)
        self.update_buttons_state()
    
    
    def _add_entries_started(self):
        self.view.enable_add_watch(False)
        self.view.spinner_start()
    
    def _add_entries_finished(self):
        self.view.enable_add_watch(True)
        self.view.spinner_stop()
        self._stop_iteration = False
        self._iterating = False
    
    def stop_iteration(self):
        self._stop_iteration = True
    
    def add_cell(self, obj, add_to_order=True):
        """ watch cell. """
        n = self.rows.add_watch(obj)
        if add_to_order and self.store_watches:
            self.watch_added(n)
    
    def add_cell_range(self, obj, checked=False):
        """ watch all cells in the range. """
        addr = obj.getRangeAddress()
        if not checked:
            cells = (addr.EndColumn - addr.StartColumn + 1) * (addr.EndRow - addr.StartRow + 1)
            if cells >= self.warn_cells and \
                self.view.confirm_warn_cells(cells) == 0:
                return
        self._add_entries_started()
        rows = self.rows
        try:
            n = 0
            names = []
            for i in range(addr.EndColumn - addr.StartColumn + 1):
                for j in range(addr.EndRow - addr.StartRow + 1):
                    cell = obj.getCellByPosition(i, j)
                    rows.reserve_watch(cell)
                    names.append(cell.AbsoluteName)
                    n += 1
                    if n == 100:
                        rows.add_reserved()
                        n = 0
                    if self._stop_iteration:
                        raise StopIteration()
        except StopIteration:
            pass
        except Exception, e:
            print(e)
        rows.add_reserved()
        self.watch_add_list(names)
        self._add_entries_finished()
    
    def add_cell_ranges(self, obj):
        """ Add to watch all cells in the ranges. """
        cells = 0
        ranges = obj.getRangeAddresses()
        for addr in ranges:
            cells += (addr.EndColumn - addr.StartColumn + 1) * (addr.EndRow - addr.StartRow + 1)
        if cells >= self.warn_cells and \
            self.view.confirm_warn_cells(cells) == 0:
            self._iterating = False
            return
        try:
            enume = obj.createEnumeration()
            while enume.hasMoreElements():
                self.add_cell_range(enume.nextElement(), True)
                if self._stop_iteration:
                    raise StopIteration()
        except StopIteration as e:
            print(e)
        except:
            pass
    
    def remove_entry(self, index):
        """ Remove entry. """
        if index >= 0:
            name = self.rows.get(index).get_header()
            self.rows.remove_watch(index)
            self.update_buttons_state()
            self.view.update_view(1)
            # select same row or last row
            if self.rows.getRowCount() <= index:
                index -= 1
            self.select_entry(index)
            if self.store_watches:
                self.watch_removed(name)
    
    def remove_all_entries(self):
        """ Remove all. """
        self.rows.remove_all_watch()
        self.update_buttons_state()
        self.view.update_view()
        self.watch_cleared()
    
    def update_all(self):
        """ Request to update all. """
        self.rows.update_all_watch()
    
    def move_entry(self, index, up):
        """ Move entry. """
        if up and index > 0:
            n = index -1
            self.rows.exchange_watches(index, n)
            self.select_entry(n)
        elif not up and index < (self.rows.getRowCount() - 1):
            n = index +1
            self.rows.exchange_watches(index, n)
            self.select_entry(n)
            
        self.watch_moved(n)
    
    def get_cell_by_name(self, addr):
        r = self.view.get_doc().getSheets().getCellRangesByName(addr)
        if r:
            return r[0].getCellByPosition(0, 0)
        return None
    
    def get_cell_references(self, addr):
        """ get cell references which are used in the formula. """
        cell = self.get_cell_by_name(addr)
        if cell:
            return get_cell_references(self.view.get_doc(), cell)
        return None
    
    def get_formula(self, addr):
        cell = self.get_cell_by_name(addr)
        if cell:
            return cell.getFormula()
        return None
    
    def update_row(self, index, formula):
        """ update selected row with the formula. """
        if index >= 0:
            row = self.rows.get(index)
            row.set_formula(formula)
    
    def stop_watching(self):
        """ Stop all watching, when the view is hidden. """
        self.rows.enable_update(False)
    
    def start_watching(self):
        """ Start all watching, when the view is shown. """
        self.rows.enable_update(True)
    
    def switch_store_state(self):
        """ Change store state. """
        self.store_watches = not self.store_watches
        if self.store_watches:
            self.rdf.add_list_to_order(self.rows.get_row_names())
        else:
            self.rdf.clear()
    
    def _init(self):
        self.rdf.load()
        if not len(self.rdf):
            return
        self._add_entries_started()
        n = 0
        rows = self.rows
        model = self.view.get_doc()
        sheets = model.getSheets()
        try:
            for name in self.rdf.order:
                try:
                    ranges = sheets.getCellRangesByName(name)
                    if ranges:
                        cell = ranges[0].getCellByPosition(0, 0)
                        rows.reserve_watch(cell)
                        n += 1
                        if n == 100:
                            rows.add_reserved()
                            n = 0
                        
                except Exception, e:
                    print(e)
                if self._stop_iteration:
                    raise StopIteration()
        except StopIteration:
            pass
            # ToDo remove non processed
        rows.add_reserved()
        self._add_entries_finished()
    
    def watch_added(self, n):
        if n >= 0:
            row = self.rows.get(n)
            if row:
                self.rdf.add(row.get_header())
    
    def watch_add_list(self, names):
        self.rdf.add_list_to_order(names)
    
    def watch_removed(self, name):
        if name:
            self.rdf.remove(name)
    
    def watch_moved(self, n):
        if n >= 0:
            row = self.rows.get(n)
            if row:
                self.rdf.move(n, row.get_header())
    
    def watch_cleared(self):
        self.rdf.clear()
