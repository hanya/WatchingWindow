
import unohelper

from com.sun.star.util import XModifyListener
from com.sun.star.table.CellContentType import FORMULA as CCT_FORMULA


class Rows(object):
    """ Row container. """
    
    def __init__(self, data_model, res):
        self._rows = []
        self._names = set()
        self.data_model = data_model
        
        self._tooltip = "%s: %%s\n%s: %%s\n%s: %%s\n%s: %%s" % (
            res["Sheet"], res["Cell"], res["Value"], res["Formula"])
        
        self._reserved = []
    
    def getRowCount(self):
        """ Returns number of rows kept by this container. """
        return self.data_model.RowCount
    
    def get_row_names(self):
        """ Return list of cell names. """
        return [row.get_header() for row in self._rows]
    
    def clear(self):
        """ Remove all rows from this container. """
        for row in self._rows:
            row.removed()
        self._rows = []
        self.listeners = []
    
    def enable_update(self, state):
        """ Enable watching on all rows. """
        for row in self._rows:
            row.enable_watching(state)
    
    def get(self, index):
        """ Get row by index. """
        return self._rows[index]
    
    def get_row_header(self, i):
        """ Get row heading by index. """
        if i >= 0 and i < len(self._rows):
            return self._rows[i].get_header()
        return ""
    
    def _tooltip_from_data(self, data):
        """ Get tooltip text from data. """
        return self._tooltip % data
    
    def _broadcast_added(self, index, row):
        data_model = self.data_model
        data = row.get_data()
        data_model.addRow(row.get_header(), data)
        data_model.updateRowToolTip(
            self.getRowCount() -1, self._tooltip_from_data(data))
    
    def _broadcast_removed(self, index):
        self.data_model.removeRow(index)
    
    
    def _broadcast_reserved_added(self, index, rows):
        data_model = self.data_model
        data = tuple([row.get_data() for row in rows])
        data_model.addRows(
            tuple([row.get_header() for row in rows]), 
            data
        )
        for i, d in enumerate(data):
            data_model.updateRowToolTip(
                i + index, self._tooltip_from_data(d))
    
    def reserve_watch(self, cell):
        """ Reserve to add watch. Reserved cells are added by add_reserved method. """
        if not cell.AbsoluteName in self._names:
            row = GridRow(cell, self)
            self._reserved.append(row)
    
    def add_reserved(self):
        """ Add reserved rows to the list. """
        if self._reserved:
            n = len(self._rows)
            self._rows[n:n + len(self._reserved)] = self._reserved
            self._names.update(
                [row.get_header() for row in self._reserved])
            self._broadcast_reserved_added(
                len(self._rows) - len(self._reserved), self._reserved)
        self._reserved[:] = []
    
    
    def add_watch(self, cell):
        """ Add new cell to watch. """
        if not cell.AbsoluteName in self._names:
            row = GridRow(cell, self)
            self._names.add(row.get_header())
            self._rows.append(row)
            i = len(self._rows) - 1
            self._broadcast_added(i, row)
            return i
    
    def remove_watch(self, index):
        """ Remove watch by index. """
        if index >= 0 and index < len(self._rows):
            try:
                row = self._rows.pop(index)
                self._names.remove(row.get_header())
                row.removed()
                self._broadcast_removed(index)
            except Exception as e:
                print("remove_watch: %s" % str(e))
    
    def remove_all_watch(self):
        """ Remove all watches. """
        for row in self._rows:
            try:
                self._names.remove(row.get_header())
            except:
                pass
            row.removed()
        
        self._rows[:] = []
        self._names.clear()
        self.data_model.removeAllRows()
    
    def update_watch(self, row):
        """ Force to update specific row. """
        try:
            i = self._rows.index(row) # ToDo make this faster
            self._broadcast_changed(i, row)
            # ToDo update input line if selected in the view
        except Exception as e:
            print("update_watch: %s" % str(e))
    
    def update_all_watch(self):
        """ Update all rows. """
        try:
            # cell address might be changed, so update _names
            names = []
            for i, row in enumerate(self._rows):
                self._broadcast_changed(i, row)
                names.append(row.get_header())
            self._names.clear()
            self._names.update(names)
        except Exception as e:
            print("update_all_watch: %s" % str(e))
    
    def _broadcast_changed(self, index, row):
        data_model = self.data_model
        data = row.get_data()
        data_model.updateRowData((0, 1, 2, 3), index, data)
        data_model.updateRowHeading(index, row.get_header())
        data_model.updateRowToolTip(index, self._tooltip_from_data(data))
    
    def exchange_watches(self, index_a, index_b):
        """ Exchange two rows specified by indexes. """
        if index_a >= 0 and index_b >= 0 and \
                index_a < len(self._rows) and index_b < len(self._rows):
            row_a = self._rows[index_a]
            row_b = self._rows[index_b]
            
            self._rows[index_a] = row_b
            self._rows[index_b] = row_a
            
            self._broadcast_changed(index_a, row_b)
            self._broadcast_changed(index_b, row_a)


class GridRow(unohelper.Base, XModifyListener):
    """ Row data of the grid, which keeps watched cell reference. """
    
    def __init__(self, cell, data_model):
        self._removed = False
        self.watching = False
        self.cell = cell
        self.data_model = data_model
        self.enable_watching(True)
    
    def __eq__(self, other):
        if isinstance(other, GridRow):
            return self.cell == other.get_cell()
        else:
            try:
                addr2 = other.getCellAddress()
                addr1 = self.cell.getCellAddress()
                return addr1.Sheet == addr2.Sheet and \
                    addr1.Row == addr2.Row and \
                    addr1.Column == addr2.Column
            except: pass
        return False
    
    def get_cell(self):
        """ get cell which is kepy by the row. """
        return self.cell
    
    def add_modify_listener(self):
        """ set modify listener. """
        self.cell.addModifyListener(self)
    
    def remove_modify_listener(self):
        """ remove modify listener. """
        self.cell.removeModifyListener(self)
    
    def enable_watching(self, state):
        """ enable watching all. """
        if self.watching and not state:
            self.remove_modify_listener()
        elif not self.watching and state:
            self.add_modify_listener()
        self.watching = state
    
    def removed(self):
        """ the row have been removed. """
        if not self._removed:
            self.remove_modify_listener()
            self._removed = False
            self.watching = False
            self.cell = None
            self.data_model = None
    
    def get_header(self):
        """ get address as header string. """
        return self.cell.AbsoluteName
    
    def get_data(self):
        """ get cell data. """
        if self.cell:
            addr = self.cell.AbsoluteName
            
            n = addr.rfind(".")
            sheet_name = addr[1:n]
            if sheet_name.startswith("'"):
                sheet_name = sheet_name[1:len(sheet_name)-1].replace("''", "'")
            
            return (
                sheet_name, addr[n + 1:].replace("$", ""), 
                self.cell.getString(), 
                self.cell.getFormula() if self.cell.getType() == CCT_FORMULA else ""
            )
        else:
            return ("", "", "internal", "error")
    
    def get_sheet_name(self):
        """ Get name of the sheet. """
        ret = self.cell.AbsoluteName
        sheet_name = ret[0:ret.rfind('.')]
        if sheet_name.startswith("'"):
            ret = sheet_name[1:len(sheet_name)-1].replace("''", "'")
        else:
            ret = sheet_name
        return ret
    
    def get_range_name(self):
        """ Get range name of the cell. """
        ret = self.cell.AbsoluteName
        return ret[ret.rfind('.') + 1:].replace("$", "")
    
    def get_content_type(self):
        """ get cell content type. """
        return self.cell.getType()
    
    def get_formula(self):
        """ get formula of the cell. """
        return self.cell.getFormula()
    
    def set_formula(self, text):
        """ set formula to the cell. """
        self.cell.setFormula(text)
    
    # XEventListener
    def disposing(self, ev):
        pass
    
    # XModifyListener
    def modified(self, ev):
        if self.watching:
            self.data_model.update_watch(self)

