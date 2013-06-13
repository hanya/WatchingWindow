
import threading

from pyww import EXT_ID
from pyww import grid
from pyww.settings import Settings, SettingsRDF
from pyww.helper import get_cell_references
from pyww.resource import CurrentStringResource


class StopIteration(Exception):
    pass

class WatchingWindow(object):
    """ Watching window model. """
    
    def __init__(self, ctx, frame):
        self.ctx = ctx
        self.frame = frame
        self._stop_iteration = False
        self._iterating = False
        self.store_watches = False
        
        res = CurrentStringResource.get(ctx)
        import pyww.helper
        # Data model for the grid control will be reused. 
        # Grid data listeners added by the control that already dead 
        # are removed by disposed exception thrown by the listeners.
        data_model = pyww.helper.create_service(
                        ctx, "com.sun.star.awt.grid.DefaultGridDataModel")
        self.rows = grid.Rows(data_model, res)
        self.rdf = SettingsRDF(self.ctx, self.get_doc())
        self.load_settings()
        self._init()
    
    def get_data_model(self):
        """ Returns grid data model for the grid control. """
        return self.rows.get_data_model()
    
    def dispose(self):
        """ Called when the document model is going to be disposed. """
        self.stop_iteration()
        self.rows.clear()
        self.rows = None
        self.frame = None
        self.ctx = None
        self.rdf = None
    
    def get_doc(self):
        return self.frame.getController().getModel()
    
    def get_watches_count(self):
        """ Returns number of watches. """
        return self.rows.get_row_count()
    
    def load_settings(self):
        """ Load configuration. """
        settings = Settings(self.ctx)
        self.store_watches = settings.get("StoreWatches")
    
    def stop_iteration(self):
        self._stop_iteration = True
    
    def get_cells_count(self):
        """ Returns number of cells in the current selection. """
        try:
            obj = self.get_doc().getCurrentSelection()
            if obj.supportsService("com.sun.star.sheet.SheetCell"):
                return 1
            elif obj.supportsService("com.sun.star.sheet.SheetCellRange"):
                addr = obj.getRangeAddress()
                return (addr.EndColumn - addr.StartColumn + 1) * \
                        (addr.EndRow - addr.EndColumn + 1)
            elif obj.supportsService("com.sun.star.sheet.SheetCellRanges"):
                count = 0
                ranges = obj.getRangeAddresses()
                for addr in ranges:
                    count += (addr.EndColumn - addr.StartColumn + 1) * \
                        (addr.EndRow - addr.EndColumn + 1)
                return count
        except:
            return 0
    
    def add_entry(self):
        """ request to add current selected object. """
        try:
            obj = self.get_doc().getCurrentSelection()
        except:
            return
        if not obj: return
        if obj.supportsService("com.sun.star.sheet.SheetCell"):
            self.add_cell(obj)
        elif obj.supportsService("com.sun.star.sheet.SheetCellRange"):
            self.add_cell_range(obj, False)
        elif obj.supportsService("com.sun.star.sheet.SheetCellRanges"):
            self.add_cell_ranges(obj)
    
    def add_cell(self, obj, add_to_order=True):
        """ watch cell. """
        n = self.rows.add_watch(obj)
        if add_to_order and self.store_watches:
            self.rdf_watch_added(n)
    
    def add_cell_range(self, obj, checked=False):
        """ watch all cells in the range. """
        addr = obj.getRangeAddress()
        rows = self.rows
        try:
            names = []
            for i in range(addr.EndColumn - addr.StartColumn + 1):
                for j in range(addr.EndRow - addr.StartRow + 1):
                    cell = obj.getCellByPosition(i, j)
                    rows.reserve_watch(cell)
                    names.append(cell.AbsoluteName)
                    if self._stop_iteration:
                        raise StopIteration()
            rows.add_reserved()
            self.rdf_watch_add_list(names)
        except StopIteration:
            pass
        except:
            pass
    
    def add_cell_ranges(self, obj):
        """ Add to watch all cells in the ranges. """
        ranges = obj.getRangeAddresses()
        try:
            enume = obj.createEnumeration()
            while enume.hasMoreElements():
                self.add_cell_range(enume.nextElement(), True)
                if self._stop_iteration:
                    raise StopIteration()
        except StopIteration:
            pass
        except:
            pass
    
    def remove_entry(self, index):
        """ Remove entry. """
        if index >= 0:
            name = self.rows.get(index).get_header()
            self.rows.remove_watch(index)
            if self.store_watches:
                self.rdf_watch_removed(name)
    
    def remove_all_entries(self):
        """ Remove all. """
        self.rows.remove_all_watch()
        self.rdf_watch_cleared()
    
    def update_all(self):
        """ Request to update all. """
        self.rows.update_all_watch()
    
    def move_entry(self, index, up):
        """ Move entry. """
        if up and index > 0:
            n = index -1
            self.rows.exchange_watches(index, n)
        elif not up and index < (self.rows.get_row_count() - 1):
            n = index +1
            self.rows.exchange_watches(index, n)
            
        self.rdf_watch_moved(n)
    
    def get_address(self, index):
        """ Returns cell address for specified index. """
        return self.rows.get_row_header(index)
    
    def get_cell_by_name(self, addr):
        r = self.get_doc().getSheets().getCellRangesByName(addr)
        if r:
            return r[0].getCellByPosition(0, 0)
        return None
    
    def get_cell_references(self, addr):
        """ get cell references which are used in the formula. """
        cell = self.get_cell_by_name(addr)
        if cell:
            return get_cell_references(self.get_doc(), cell)
        return None
    
    def get_formula(self, addr):
        cell = self.get_cell_by_name(addr)
        if cell:
            return cell.FormulaLocal
        return None
    
    def update_row(self, index, formula):
        """ update selected row with the formula. """
        if index >= 0:
            row = self.rows.get(index)
            if row:
                row.set_formula(formula)
    
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
        n = 0
        rows = self.rows
        sheets = self.get_doc().getSheets()
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
                except Exception as e:
                    print(e)
                if self._stop_iteration:
                    raise StopIteration()
            rows.add_reserved()
        except StopIteration:
            pass
    
    # RDF specific functions
    
    def rdf_watch_added(self, n):
        if n >= 0:
            try:
                row = self.rows.get(n)
            except:
                return
            if row:
                self.rdf.add(row.get_header())
    
    def rdf_watch_add_list(self, names):
        self.rdf.add_list_to_order(names)
    
    def rdf_watch_removed(self, name):
        if name:
            self.rdf.remove(name)
    
    def rdf_watch_moved(self, n):
        if n >= 0:
            try:
                row = self.rows.get(n)
            except:
                return
            if row:
                self.rdf.move(n, row.get_header())
    
    def rdf_watch_cleared(self):
        self.rdf.clear()
