
import threading
import unohelper

from com.sun.star.awt import XWindowListener, XKeyHandler, \
    XFocusListener, XActionListener, XMouseListener
from com.sun.star.awt.grid import XGridSelectionListener, XGridRowSelection
from com.sun.star.view.SelectionType import SINGLE as ST_SINGLE
from com.sun.star.style.HorizontalAlignment import RIGHT as HA_RIGHT
from com.sun.star.awt import Rectangle
from com.sun.star.awt.MouseButton import LEFT as MB_LEFT, RIGHT as MB_RIGHT
from com.sun.star.awt.Key import RETURN as K_RETURN, \
    UP as K_UP, DOWN as K_DOWN, HOME as K_HOME, END as K_END, \
    DELETE as K_DELETE, CONTEXTMENU as K_CONTEXTMENU
from com.sun.star.awt.PosSize import X as PS_X, Y as PS_Y, \
    WIDTH as PS_WIDTH, HEIGHT as PS_HEIGHT, SIZE as PS_SIZE
from com.sun.star.awt.MenuItemStyle import CHECKABLE as MIS_CHECKABLE
from com.sun.star.awt.MessageBoxButtons import \
    BUTTONS_OK_CANCEL as MBB_BUTTONS_OK_CANCEL, \
    DEFAULT_BUTTON_CANCEL as MBB_DEFAULT_BUTTON_CANCEL
from com.sun.star.beans import PropertyValue
from com.sun.star.beans.PropertyState import DIRECT_VALUE as PS_DIRECT_VALUE

from pyww.settings import Settings
from pyww.helper import create_control, create_container, create_controls, \
    PopupMenuWrapper, MenuEntry, messagebox
import pyww.resource


class WatchingWindowView(object):
    """ Watching window view. """
    
    LEFT_MARGIN = 3
    RIGHT_MARGIN = 3
    TOP_MARGIN = 3
    BUTTON_SEP = 2
    
    BUTTON_WIDTH = 28
    BUTTON_HEIGHT = 28
    
    INPUT_LINE_HEIGHT = 23
    
    def __init__(self, ctx, model, frame, parent):
        self.ctx = ctx
        self.frame = frame
        self.res = pyww.resource.CurrentStringResource.get(ctx)
        self.model = model
        self.grid = None
        self._context_menu = None
        self.input_line_shown = Settings(ctx).get("InputLine")
        self._key_handler = self.KeyHandler(self)
        self._grid_key_handler = self.GridKeyHandler(self)
        self._focus_listener = self.FocusListener(self, False)
        self._create_view(ctx, parent, self.res, 
                            model.get_data_model(), self.input_line_shown)
        parent.addWindowListener(self.WindowListener(self))
        self.update_buttons_state()
    
    def focus_to_doc(self):
        """ Move focus to current document. """
        self.frame.getContainerWindow().setFocus()
    
    # Grid functions
    
    def grid_select_entry(self, index):
        """ Select specific row by index. """
        if 0 <= index < self.model.get_watches_count():
            self.grid.selectRow(index)
    
    def grid_is_entry_selected(self):
        """ Check any entry is selected. """
        return self.grid.hasSelectedRows()
    
    def get_selected_entry_index(self):
        """ Get selected entry index. """
        if self.grid.hasSelectedRows():
            return self.grid.getSelectedRows()[0]
        return -1
    
    def get_selected_entry_heading(self):
        """ Get row heading of selected entry. """
        index = self.get_selected_entry_index()
        if 0 <= index < self.model.get_watches_count():
            return self.model.get_address(index)
        return ""
    
    def grid_select_current(self):
        """ Select cursor row. """
        self.grid_select_entry(self.grid.getCurrentRow())
    
    def grid_deselect_all(self):
        self.grid.deselectAllRows()
    
    def grid_goto_row(self, index):
        self.grid.goToCell(0, index)
    
    def is_selected_entry_moveable(self, up):
        """ Check selected entry is moveable. """
        i = self.get_selected_entry_index()
        if up and i > 0:
            return True
        elif not up and i < (self.model.get_watches_count() - 1):
            return True
        return False
    
    def dispose(self):
        self.cont = None
        self.model = None
        self.grid = None
        self.res = None
        self._context_menu = None
    
    class ListenerBase(unohelper.Base):
        def __init__(self, act):
            self.act = act
        
        # XEventListener
        def disposing(self, ev):
            self.act = None
    
    class MouseListener(ListenerBase, XMouseListener):
        def mouseEntered(self, ev): pass
        def mouseExited(self, ev): pass
        def mousePressed(self, ev):
            if ev.Buttons == MB_RIGHT and ev.ClickCount == 1:
                self.act.context_menu(ev.X, ev.Y)
        
        def mouseReleased(self, ev):
            if ev.Buttons == MB_LEFT and ev.ClickCount == 2:
                self.act.cmd_goto()
    
    class ButtonListener(ListenerBase, XActionListener):
        def actionPerformed(self, ev):
            self.act.execute_cmd(ev.ActionCommand)
    
    # Command processing
    
    def execute_cmd(self, command):
        try:
            getattr(self, "cmd_" + command)()
        except:
            pass
    
    def cmd_add(self):
        if 1000 < self.model.get_cells_count():
            self.errorbox(self.res["Too many cells are selected."], 
                self.res["Watching Window"])
            return
        self.model.add_entry()
    
    def cmd_delete(self):
        index = self.get_selected_entry_index()
        if 0 <= index:
            self.model.remove_entry(index)
            self.grid_select_current()
    
    def cmd_update(self):
        self.model.update_all()
    
    def cmd_goto(self, addr=None):
        if addr is None:
            index = self.get_selected_entry_index()
            if index < 0: return
            try:
                addr = self.model.get_address(index)
            except:
                pass
        if addr:
            self.goto_point(addr)
    
    def cmd_up(self):
        index = self.get_selected_entry_index()
        if 0 <= index:
            self.model.move_entry(index, True)
    
    def cmd_down(self):
        index = self.get_selected_entry_index()
        if 0 <= index:
            self.model.move_entry(index, False)
    
    def cmd_clear(self):
        self.model.remove_all_entries()
    
    def cmd_settings(self):
        from pyww.settings import Settings
        try:
            Settings(self.ctx).configure(self.res)
        except Exception as e:
            print(e)
    
    def cmd_switch_inputline(self):
        self.switch_input_line()
    
    def cmd_option(self):
        ps = self.cont.getControl("btn_option").getPosSize()
        self.option_popup(ps.X, ps.Y + ps.Height)
    
    def cmd_switch_store(self):
        self.model.switch_store_state()
    
    def cmd_about(self):
        from pyww import EXT_ID, EXT_DIR
        from pyww.dialogs import AboutDialog
        from pyww.helper import get_package_info, get_text_content
        name, version = get_package_info(self.ctx, EXT_ID)
        text = get_text_content(self.ctx, EXT_DIR + "LICENSE")
        text = "\n".join(text.split("\n")[20:])
        
        AboutDialog(self.ctx, self.res, 
            name=name, version=version, text=text, 
        ).execute()
    
    
    def goto_point(self, desc):
        """ move cursor to the specified address. """
        self.dispatch(
            ".uno:GoToCell", 
            (PropertyValue("ToPoint", 0, desc, PS_DIRECT_VALUE),))
        self.frame.getComponentWindow().setFocus()
    
    def dispatch(self, cmd, args):
        """ dispatch with arguments. """
        self.ctx.getServiceManager().createInstanceWithContext(
            "com.sun.star.frame.DispatchHelper", self.ctx).\
                executeDispatch(self.frame, cmd, "_self", 0, args)
    
    class GridSelectionListener(ListenerBase, XGridSelectionListener):
        def selectionChanged(self, ev):
            try:
                self.act.update_buttons_state()
                self.act.update_input_line()
            except Exception as e:
                print(e)
    
    class FocusListener(ListenerBase, XFocusListener):
        def __init__(self, act, is_grid):
            self.act = act
            self._is_grid = is_grid
        
        def focusGained(self, ev):
            self.act.focus_gained(self._is_grid)
        
        def focusLost(self, ev):
            self.act.focus_lost(self._is_grid)
    
    def focus_gained(self, is_grid):
        """ Set key handler to the toolkit when the focus gained into 
            input field or grid field. 
            This handler is required to consum some key events. """
        self.frame.getContainerWindow().getToolkit().addKeyHandler(
            self._grid_key_handler if is_grid else self._key_handler)
    
    def focus_lost(self, is_grid):
        """ Remove key handler has been set by focus gained event. """
        self.frame.getContainerWindow().getToolkit().removeKeyHandler(
            self._grid_key_handler if is_grid else self._key_handler)
    
    class KeyHandler(ListenerBase, XKeyHandler):
        RETURN = K_RETURN
        def keyPressed(self, ev):
            if ev.KeyCode == self.__class__.RETURN:
                self.act.update_row_formula()
                return True
            return False
        
        def keyReleased(self, ev):
            return True
    
    def update_row_formula(self):
        """ Update cell formula from input line. """
        index = self.get_selected_entry_index()
        if 0 <= index:
            self.model.update_row(index, self.get_input_text())
    
    class GridKeyHandler(ListenerBase, XKeyHandler):
        RETURN = K_RETURN
        def keyPressed(self, ev):
            code = ev.KeyCode
            if code == K_RETURN:
                self.act.cmd_goto()
                return True
            elif code in (K_UP, K_DOWN, K_HOME, K_END):
                self.act.grid_cmd_cursor(code, ev.Modifiers & 0b11)
                return True
            return False
        
        def keyReleased(self, ev):
            if ev.KeyCode == K_CONTEXTMENU:
                self.act.grid_cmd_contextmenu()
            return True
    
    def grid_cmd_contextmenu(self):
        # ToDo current selection is not shown
        index = self.get_selected_entry_index()
        if 0 <= index:
            # ToDo calculate better location
            self.context_menu(0, 0)
            # How to move focus to the popup menu? it seems no way
    
    def grid_cmd_cursor(self, key, mod):
        """ Move cursor by key event. """
        index = self.get_selected_entry_index()
        if key == K_UP:
            index -= 1
        elif key == K_DOWN:
            index += 1
        elif key == K_HOME:
            index = 0
        elif key == K_END:
            index = self.model.get_watches_count() - 1
        if index < 0 or self.model.get_watches_count() <= index:
            return
        self.grid_deselect_all()
        self.grid_goto_row(index)
        self.grid_select_current()
    
    def update_buttons_state(self):
        """ Update state of buttons by current situation. """
        ubs = self.update_button_state
        delete_state = False
        goto_state = False
        update_state = False
        up_state = False
        down_state = False
        if self.model.get_watches_count() == 0:
            pass
        else:
            if self.grid_is_entry_selected():
                delete_state = True
                goto_state = True
                
                up_state = self.is_selected_entry_moveable(True)
                down_state = self.is_selected_entry_moveable(False)
            update_state = True
        ubs("btn_delete", delete_state)
        ubs("btn_goto", goto_state)
        ubs("btn_update", update_state)
        
        ubs("btn_up", up_state)
        ubs("btn_down", down_state)
    
    def update_input_line(self):
        addr = self.get_selected_entry_heading()
        if addr:
            self.set_input_line(self.model.get_formula(addr))
    
    def set_input_line(self, text):
        """ Set text to input line. """
        self.cont.getControl("edit_input").getModel().Text = text
    
    def get_input_text(self):
        """ Get text from input line. """
        return self.cont.getControl("edit_input").getModel().Text
    
    def enable_add_watch(self, state):
        """ Request to change state of add button. """
        self.update_button_state("btn_add", state)
    
    def context_menu(self, x, y):
        """ Show context menu at the coordinate. """
        index = self.get_selected_entry_index()
        if index < 0: return
        _ = self.res.get
        popup = self._context_menu
        if popup is None:
            popup = PopupMenuWrapper(self.ctx, 
                (
                    MenuEntry(_("Go to Cell"), 4, 0, "goto"), 
                    MenuEntry(_("Go to"), 6, 1, "gotocell"), 
                    MenuEntry(_("Remove"), 8, 2, "delete"), 
                    MenuEntry("", -1, 3), 
                    MenuEntry(_("Up"), 10, 4, "up"), 
                    MenuEntry(_("Down"), 11, 5, "down")
                ), True)
            self._context_menu = popup
        
        if popup:
            addr = self.get_selected_entry_heading()
            state = False
            if addr:
                refs = self.model.get_cell_references(addr)
                if refs:
                    popup.setPopupMenu(
                        6, 
                        PopupMenuWrapper(
                            self.ctx, 
                            [MenuEntry(ref, i + 1000, i, "") 
                                for i, ref in enumerate(refs)], 
                            False
                        )
                    )
                    state = True
            popup.enableItem(6, state)
            popup.enableItem(10, self.is_selected_entry_moveable(True))
            popup.enableItem(11, self.is_selected_entry_moveable(False))
            
            ps = self.grid.getPosSize()
            n = popup.execute(self.cont.getPeer(), x + ps.X, y + ps.Y)
            if n > 0 and n < 1000:
                self.execute_cmd(popup.getCommand(n))
            elif n >= 1000:
                addr = refs[n - 1000]
                self.cmd_goto(addr)
    
    def option_popup(self, x, y):
        """ Show popup menu for option button. """
        _ = self.res.get
        popup = PopupMenuWrapper(self.ctx, 
            (
                MenuEntry(_("Clear"), 32, 0, "clear"), 
                MenuEntry("", -1, 1, ""), 
                MenuEntry(_("Input line"), 1024, 2, "switch_inputline", style=MIS_CHECKABLE), 
                MenuEntry(_("Store watches"), 2048, 3, "switch_store", style=MIS_CHECKABLE), 
                MenuEntry("", -1, 4, ""), 
                MenuEntry(_("Settings..."), 512, 5, "settings"), 
                MenuEntry(_("About"), 4096, 6, "about"), 
            ), True)
        popup.checkItem(1024, self.input_line_shown)
        popup.checkItem(2048, self.model.store_watches)
        
        n = popup.execute(self.cont.getPeer(), x, y)
        if n > 0:
            self.execute_cmd(popup.getCommand(n))
    
    def _messagebox(self, message, title, message_type, buttons):
        """ Show message in message box. """
        return messagebox(self.ctx, self.frame.getContainerWindow(), 
            message, title, message_type, buttons)
    
    def message(self, message, title):
        """ Shows message with title. """
        return self._messagebox(message, title, "messbox", 1)
    
    def errorbox(self, message, title):
        """ Shows error message with title. """
        return self._messagebox(message, title, "errorbox", 1)
    
    def update_button_state(self, name, state):
        """ Update state of specific button. """
        ctrl = self.cont.getControl(name)
        if ctrl.isEnabled() != state:
            ctrl.setEnable(state)
    
    def switch_input_line(self, new_state=None):
        """ Switch to show/hide input line. """
        height = self.cont.getPosSize().Height
        btn_height = self.BUTTON_HEIGHT
        if new_state is None:
            new_state = not self.input_line_shown
        if new_state:
            self.grid.setPosSize(
                0, self.TOP_MARGIN * 3 + btn_height + self.INPUT_LINE_HEIGHT, 
                0, height - (self.TOP_MARGIN * 3 + btn_height + self.INPUT_LINE_HEIGHT), 
                PS_Y + PS_HEIGHT)
            self.cont.getControl("edit_input").addFocusListener(self._focus_listener)
        else:
            self.grid.setPosSize(0, self.TOP_MARGIN * 2 + btn_height, 
            0, height - (self.TOP_MARGIN * 2 + btn_height), PS_Y + PS_HEIGHT)
            self.cont.getControl("edit_input").removeFocusListener(self._focus_listener)
        self.cont.getControl("edit_input").setVisible(new_state)
        self.input_line_shown = new_state
    
    class WindowListener(ListenerBase, XWindowListener):
        def windowMoved(self, ev): pass
        def windowHidden(self, ev):
            pass#self.model.stop_watching()
        def windowShown(self, ev):
            pass#self.model.start_watching()
        def windowResized(self, ev):
            self.act.window_resized(ev.Width, ev.Height)
    
    def window_resized(self, width, height):
        gc = self.cont.getControl
        btn_width = self.BUTTON_WIDTH
        btn_height = self.BUTTON_HEIGHT
        right_margin = self.RIGHT_MARGIN
        self.cont.setPosSize(0, 0, width, height, PS_SIZE)
        
        gc("btn_option").setPosSize(
            width - btn_width - right_margin, 0, 0, 0, PS_X)
        gc("btn_update").setPosSize(
            width - btn_width * 2 - right_margin - self.BUTTON_SEP, 0, 0, 0, PS_X)
        
        if self.input_line_shown:
            gc("grid").setPosSize(
                0, 0, width, 
                height - btn_height - self.TOP_MARGIN * 3 - self.INPUT_LINE_HEIGHT, 
                PS_SIZE)
        else:
            gc("grid").setPosSize(
                0, 0, width, 
                height - btn_height - self.TOP_MARGIN * 2, PS_SIZE)
        gc("edit_input").setPosSize(
            0, 0, width - self.LEFT_MARGIN - right_margin, 0, PS_WIDTH)
    
    def _create_view(self, ctx, parent, res, data_model, show_input_line=False):
        from pyww import ICONS_DIR
        LEFT_MARGIN = self.LEFT_MARGIN
        RIGHT_MARGIN = self.RIGHT_MARGIN
        TOP_MARGIN = self.TOP_MARGIN
        BUTTON_SEP = self.BUTTON_SEP
        BUTTON_WIDTH = self.BUTTON_WIDTH
        BUTTON_HEIGHT = self.BUTTON_HEIGHT
        INPUT_LINE_HEIGHT = self.INPUT_LINE_HEIGHT
        
        cont = create_container(ctx, parent, (), ())
        self.cont = cont
        background_color = cont.StyleSettings.DialogColor
        
        button_listener = self.ButtonListener(self)
        ps = parent.getPosSize()
        create_controls(ctx, cont, 
            (
                ("Button", "btn_add", 
                    LEFT_MARGIN, TOP_MARGIN, BUTTON_WIDTH, BUTTON_HEIGHT, 
                    ("FocusOnClick", "HelpText", "HelpURL", "ImageURL"), 
                    (False, res["Add"], "", ICONS_DIR + "add_16.png"), 
                    {"ActionCommand": "add", "ActionListener": button_listener}), 
                ("Button", "btn_delete", 
                    LEFT_MARGIN + BUTTON_SEP + BUTTON_WIDTH, TOP_MARGIN, 
                    BUTTON_WIDTH, BUTTON_HEIGHT, 
                    ("FocusOnClick", "HelpText", "HelpURL", "ImageURL"), 
                    (False, res["Remove"], "", ICONS_DIR + "delete_16.png"), 
                    {"ActionCommand": "delete", "ActionListener": button_listener}), 
                ("Button", "btn_goto", 
                    LEFT_MARGIN + (BUTTON_SEP + BUTTON_WIDTH) * 2, TOP_MARGIN, 
                    BUTTON_WIDTH, BUTTON_HEIGHT, 
                    ("FocusOnClick", "HelpText", "HelpURL", "ImageURL"), 
                    (False, res["Go to Cell"], "", ICONS_DIR + "goto_16.png"), 
                    {"ActionCommand": "goto", "ActionListener": button_listener}), 
                ("Button", "btn_up", 
                    LEFT_MARGIN + (BUTTON_SEP + BUTTON_WIDTH) * 3, TOP_MARGIN, 
                    BUTTON_WIDTH, BUTTON_HEIGHT, 
                    ("FocusOnClick", "HelpText", "HelpURL", "ImageURL"), 
                    (False, res["Up"], "", ICONS_DIR + "up_16.png"), 
                    {"ActionCommand": "up", "ActionListener": button_listener}), 
                ("Button", "btn_down", 
                    LEFT_MARGIN + (BUTTON_SEP + BUTTON_WIDTH) * 4, TOP_MARGIN, 
                    BUTTON_WIDTH, BUTTON_HEIGHT, 
                    ("FocusOnClick", "HelpText", "HelpURL", "ImageURL"), 
                    (False, res["Down"], "", ICONS_DIR + "down_16.png"), 
                    {"ActionCommand": "down", "ActionListener": button_listener}), 
                ("Button", "btn_update", 
                    LEFT_MARGIN + (BUTTON_SEP + BUTTON_WIDTH) * 5, TOP_MARGIN, 
                    BUTTON_WIDTH, BUTTON_HEIGHT, 
                    ("FocusOnClick", "HelpText", "HelpURL", "ImageURL"), 
                    (False, res["Update All"], "", ICONS_DIR + "update_16.png"), 
                    {"ActionCommand": "update", "ActionListener": button_listener}), 
                ("Button", "btn_option", 
                    LEFT_MARGIN + (BUTTON_WIDTH + BUTTON_SEP) * 6, TOP_MARGIN, 
                    BUTTON_WIDTH, BUTTON_HEIGHT, 
                    ("FocusOnClick", "HelpText", "HelpURL", "ImageURL"), 
                    (False, res["Option"], "", ICONS_DIR + "tune_16.png"), 
                    {"ActionCommand": "option", "ActionListener": button_listener}), 
                ("Edit", "edit_input", 
                    LEFT_MARGIN, TOP_MARGIN * 2 + BUTTON_HEIGHT, 
                    ps.Width, INPUT_LINE_HEIGHT, 
                    ("HelpText", "HelpURL"), 
                    (res["Input line"], ""))
            )
        )
        
        grid_y = TOP_MARGIN + BUTTON_HEIGHT + TOP_MARGIN + \
            ((TOP_MARGIN + INPUT_LINE_HEIGHT) if show_input_line else 0)
        
        grid = create_control(ctx, "grid.UnoControlGrid", 
            0, grid_y, ps.Width, ps.Height - grid_y, 
            ("BackgroundColor", "Border", "GridDataModel", "EvenRowBackgroundColor", 
                "HScroll", "SelectionModel", "ShowColumnHeader", 
                "ShowRowHeader", "VScroll"), 
            (background_color, 0, data_model, 0xeeeeee, 
                False, ST_SINGLE, True, False, False))
        grid_model = grid.getModel()
        self.grid = grid
        
        column_model = grid_model.ColumnModel
        for title in [res[_title] 
                for _title in ("Sheet", "Cell", "Value", "Formula")]:
            column = column_model.createColumn()
            column.Title = title
            column_model.addColumn(column)
        column_model.getColumn(2).HorizontalAlign = HA_RIGHT
        
        cont.addControl("grid", grid)
        grid.addMouseListener(self.MouseListener(self))
        grid.addSelectionListener(self.GridSelectionListener(self))
        grid.addFocusListener(self.FocusListener(self, True))
        
        edit_input = cont.getControl("edit_input")
        self.switch_input_line(show_input_line)

