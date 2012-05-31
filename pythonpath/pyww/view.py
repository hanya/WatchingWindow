
import unohelper

from com.sun.star.awt import XWindowListener, XKeyHandler, \
    XFocusListener, XActionListener, XMouseListener
from com.sun.star.awt.grid import XGridSelectionListener, XGridRowSelection
from com.sun.star.view.SelectionType import SINGLE as ST_SINGLE
from com.sun.star.style.HorizontalAlignment import RIGHT as HA_RIGHT
from com.sun.star.awt.PopupMenuDirection import EXECUTE_DEFAULT as PMD_EXECUTE_DEFAULT
from com.sun.star.awt import Rectangle
from com.sun.star.awt.MouseButton import LEFT as MB_LEFT, RIGHT as MB_RIGHT
from com.sun.star.awt.Key import RETURN as K_RETURN
from com.sun.star.awt.InvalidateStyle import UPDATE as IS_UPDATE, CHILDREN as IS_CHILDREN
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
    get_backgroundcolor, create_popup, MenuEntry


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


class SpinnerMouseListener(unohelper.Base, XMouseListener):
    """ Mouse listener for spinner to stop current iteration. """
    
    def __init__(self, act):
        self.act = act
    
    def disposing(self, ev):
        self.act = None
    
    def mouseEntered(self, ev): pass
    def mouseExited(self, ev): pass
    def mouseReleased(self, ev): pass
    def mousePressed(self, ev):
        self.act.request_to_stop()


class WatchingWindowView(unohelper.Base, XWindowListener, XActionListener, 
    XMouseListener, XGridSelectionListener, XFocusListener, XKeyHandler):
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
        self.res = CurrentStringResource.get(ctx)
        self.model = model
        self.grid = None
        self._context_menu = None
        settings = Settings(ctx)
        self.input_line_shown = settings.get("InputLine")
        self._create_view(ctx, parent, self.res, self.input_line_shown)
        parent.addWindowListener(self)
    
    def focus_to_doc(self):
        """ Move focus to current document. """
        self.frame.getContainerWindow().setFocus()
    
    def get_doc(self):
        """ Get document model. """
        return self.frame.getController().getModel()
    
    def get_current_selection(self):
        """ Get current selection in the document. """
        return self.get_doc().getCurrentSelection()
    
    def get_data_model(self):
        """ Returns data model. """
        return self.grid.getModel().GridDataModel
    
    
    def select_entry(self, index):
        """ Select specific row by index. """
        if index >= 0 and index < self.model.get_watches_count():
            self.grid.selectRow(index)
    
    def is_entry_selected(self):
        """ Check any entry is selected. """
        return self.grid.hasSelectedRows()
    
    def get_selected_entry_index(self):
        """ Get selected entry index. """
        if self.grid.hasSelectedRows():
            return self.grid.getSelectedRows()[0]
        return -1
    
    def is_selected_entry_moveable(self, up):
        """ Check selected entry is moveable. """
        i = self.get_selected_entry_index()
        if up and i > 0:
            return True
        elif not up and i < (self.model.get_watches_count() - 1):
            return True
        return False
    
    def get_selected_entry_heading(self):
        """ Get row heading of selected entry. """
        index = self.get_selected_entry_index()
        data_model = self.get_data_model()
        if 0 <= index < data_model.RowCount:
            return data_model.getRowHeading(index)
        return ""
    
    def dispose(self):
        self.cont = None
        self.model = None
        self.grid = None
        self.res = None
        self._context_menu = None
    
    # XEventListener
    def disposing(self, ev):
        pass
    
    # XMouseListener
    def mouseEntered(self, ev): pass
    def mouseExited(self, ev): pass
    def mousePressed(self, ev):
        if ev.Buttons == MB_RIGHT and ev.ClickCount == 1:
            try:
                self.context_menu(ev.X, ev.Y)
            except Exception as e:
                print(e)
    
    def mouseReleased(self, ev):
        if ev.Buttons == MB_LEFT and ev.ClickCount == 2:
            try:
                self.cmd_goto()
            except Exception as e:
                print(e)
    
    # XActionListener
    def actionPerformed(self, ev):
        self.execute_cmd(ev.ActionCommand)
    
    def execute_cmd(self, command):
        try:
            getattr(self, "cmd_" + command)()
        except:
            pass
    
    def cmd_add(self):
        # ToDo if selected cell is already added, select it in the grid
        self.model.add_entry()
    
    def cmd_delete(self):
        index = self.get_selected_entry_index()
        if 0 <= index:
            self.model.remove_entry(index)
    
    def cmd_update(self):
        self.model.update_all()
    
    def cmd_goto(self, addr=None):
        if addr is None:
            addr = self.get_selected_entry_heading()
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
        settings = Settings(self.ctx)
        if settings.configure(self.res):
            self.model.settings_changed()
    
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
        from helper import get_package_info, get_text_content
        name, version = get_package_info(self.ctx, EXT_ID)
        text = get_text_content(self.ctx, EXT_DIR + "LICENSE")
        text = "\n".join(text.split("\n")[20:])
        
        AboutDialog(self.ctx, self.res, 
            name=name, version=version, text=text, 
        ).execute()
    
    
    def request_to_stop(self):
        """ Stop current task to add watches. """
        self.model.stop_iteration()
    
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
    
    # XGridSelectionListener
    def selectionChanged(self, ev):
        try:
            self.update_buttons_state()
            self.update_input_line()
        except Exception, e:
            print(e)
    
    # XFocusListener
    def focusGained(self, ev):
        self.frame.getController().addKeyHandler(self)
    
    def focusLost(self, ev):
        self.frame.getController().removeKeyHandler(self)
    
    # XKeyHandler
    def keyPressed(self, ev):
        if ev.KeyCode == K_RETURN:
            index = self.get_selected_entry_index()
            if 0 <= index:
                self.model.update_row(index, self.get_input_text())
            return True
        return False
    
    def keyReleased(self, ev):
        return True
    
    def update_buttons_state(self):
        """ Update state of buttons by current situation. """
        ubs = self.update_button_state
        if self.model.get_watches_count() == 0:
            ubs("btn_delete", False)
            ubs("btn_goto", False)
            ubs("btn_update", False)
            
            ubs("btn_up", False)
            ubs("btn_down", False)
        else:
            if self.is_entry_selected():
                ubs("btn_delete", True)
                ubs("btn_goto", True)
                
                ubs("btn_up", self.is_selected_entry_moveable(True))
                ubs("btn_down", self.is_selected_entry_moveable(False))
            else:
                ubs("btn_delete", False)
                ubs("btn_goto", False)
                
                ubs("btn_up", False)
                ubs("btn_down", False)
            ubs("btn_update", True)
    
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
        _ = self.res.get
        popup = self._context_menu
        if popup is None:
            popup = create_popup(self.ctx, 
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
                        create_popup(
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
            n = popup.execute(
                    self.cont.getPeer(), 
                    Rectangle(x + ps.X, y + ps.Y, 0, 0), 
                    PMD_EXECUTE_DEFAULT)
            if n > 0 and n < 1000:
                self.execute_cmd(popup.getCommand(n))
            elif n >= 1000:
                addr = refs[n - 1000]
                self.cmd_goto(addr)
    
    def option_popup(self, x, y):
        """ Show popup menu for option button. """
        _ = self.res.get
        popup = create_popup(self.ctx, 
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
        
        n = popup.execute(
                self.cont.getPeer(), Rectangle(x, y, 0, 0), PMD_EXECUTE_DEFAULT)
        if n > 0:
            self.execute_cmd(popup.getCommand(n))
    
    
    def messagebox(self, message, title, message_type, buttons):
        """ Show message in message box. """
        parent = self.frame.getContainerWindow()
        msgbox = parent.getToolkit().createMessageBox(
            parent, Rectangle(), message_type, buttons, title, message)
        n = msgbox.execute()
        msgbox.dispose()
        return n
    
    def message(self, message, title):
        """ Shows message with title. """
        return self.messagebox(message, title, "messbox", 1)
    
    def confirm(self, message, title):
        """ Show confirm dialog ."""
        return self.messagebox(
            message, title, "messbox", 
            MBB_BUTTONS_OK_CANCEL + MBB_DEFAULT_BUTTON_CANCEL)
    
    def confirm_warn_cells(self, cells):
        """ Show warn cells dialog. """
        return self.confirm(
            self.res["Do you want to add %s cells?"] % cells, 
            self.res["Watching Window"])
    
    def spinner_start(self):
        """ Start and show spinner. """
        spinner = self.cont.getControl("spinner")
        spinner.setVisible(True)
        spinner.startAnimation()
    
    def spinner_stop(self):
        """ Stop and hide spinner. """
        spinner = self.cont.getControl("spinner")
        spinner.setVisible(False)
        spinner.stopAnimation()
    
    def update_view(self, rows=None):
        """ Force redraw the grid. """
        # ToDo invalidate for rows removed if specified
        self.cont.getPeer().invalidateRect(
            self.grid.getPosSize(), IS_UPDATE)
    
    def update_button_state(self, name, state):
        """ Update state of specific button. """
        self.cont.getControl(name).setEnable(state)
    
    def switch_input_line(self, new_state=None):
        """ Switch to show/hide input line. """
        ps = self.cont.getPosSize()
        height = ps.Height
        btn_height = self.BUTTON_HEIGHT
        if new_state is None:
            new_state = not self.input_line_shown
        if new_state:
            self.grid.setPosSize(
                0, self.TOP_MARGIN * 3 + btn_height + self.INPUT_LINE_HEIGHT, 
                0, 
                height - (self.TOP_MARGIN * 3 + btn_height + self.INPUT_LINE_HEIGHT), 
                PS_Y + PS_HEIGHT)
            self.cont.getControl("edit_input").addFocusListener(self)
        else:
            self.grid.setPosSize(0, self.TOP_MARGIN * 2 + btn_height, 
            0, height - (self.TOP_MARGIN * 2 + btn_height), PS_Y + PS_HEIGHT)
            self.cont.getControl("edit_input").removeFocusListener(self)
            self.update_view()
        self.cont.getControl("edit_input").setVisible(new_state)
        self.input_line_shown = new_state
    
    # XWindowListener
    def windowMoved(self, ev): pass
    def windowHidden(self, ev):
        self.model.stop_watching()
    
    def windowShown(self, ev):
        self.model.start_watching()
    
    def windowResized(self, ev):
        gc = self.cont.getControl
        ps = ev.Source.getPosSize()
        width = ps.Width
        height = ps.Height
        btn_width = self.BUTTON_WIDTH
        btn_height = self.BUTTON_HEIGHT
        right_margin = self.RIGHT_MARGIN
        self.cont.setPosSize(0, 0, width, height, PS_SIZE)
        
        gc("spinner").setPosSize(
            width - btn_width * 2 - right_margin, 0, 0, 0, PS_X)
        gc("btn_option").setPosSize(
            width - btn_width - right_margin, 0, 0, 0, PS_X)
        
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
    
    def _create_view(self, ctx, parent, res, show_input_line=False):
        from pyww import ICONS_DIR
        LEFT_MARGIN = self.LEFT_MARGIN
        RIGHT_MARGIN = self.RIGHT_MARGIN
        TOP_MARGIN = self.TOP_MARGIN
        BUTTON_SEP = self.BUTTON_SEP
        BUTTON_WIDTH = self.BUTTON_WIDTH
        BUTTON_HEIGHT = self.BUTTON_HEIGHT
        INPUT_LINE_HEIGHT = self.INPUT_LINE_HEIGHT
        
        cont = create_container(ctx, parent, 
            ("BackgroundColor",), (get_backgroundcolor(parent),))
        self.cont = cont
        
        ps = parent.getPosSize()
        create_controls(ctx, cont, 
            (
                ("Button", "btn_add", 
                    LEFT_MARGIN, TOP_MARGIN, BUTTON_WIDTH, BUTTON_HEIGHT, 
                    ("FocusOnClick", "HelpText", "HelpURL", "ImageURL"), 
                    (False, res["Add"], "", ICONS_DIR + "add_16.png"), 
                    {"ActionCommand": "add", "ActionListener": self}), 
                ("Button", "btn_delete", 
                    LEFT_MARGIN + BUTTON_SEP + BUTTON_WIDTH, TOP_MARGIN, 
                    BUTTON_WIDTH, BUTTON_HEIGHT, 
                    ("FocusOnClick", "HelpText", "HelpURL", "ImageURL"), 
                    (False, res["Remove"], "", ICONS_DIR + "delete_16.png"), 
                    {"ActionCommand": "delete", "ActionListener": self}), 
                ("Button", "btn_update", 
                    LEFT_MARGIN + (BUTTON_SEP + BUTTON_WIDTH) * 2, TOP_MARGIN, 
                    BUTTON_WIDTH, BUTTON_HEIGHT, 
                    ("FocusOnClick", "HelpText", "HelpURL", "ImageURL"), 
                    (False, res["Update All"], "", ICONS_DIR + "update_16.png"), 
                    {"ActionCommand": "update", "ActionListener": self}), 
                ("Button", "btn_goto", 
                    LEFT_MARGIN + (BUTTON_SEP + BUTTON_WIDTH) * 3, TOP_MARGIN, 
                    BUTTON_WIDTH, BUTTON_HEIGHT, 
                    ("FocusOnClick", "HelpText", "HelpURL", "ImageURL"), 
                    (False, res["Go to Cell"], "", ICONS_DIR + "goto_16.png"), 
                    {"ActionCommand": "goto", "ActionListener": self}), 
                ("Button", "btn_up", 
                    LEFT_MARGIN + (BUTTON_SEP + BUTTON_WIDTH) * 4, TOP_MARGIN, 
                    BUTTON_WIDTH, BUTTON_HEIGHT, 
                    ("FocusOnClick", "HelpText", "HelpURL", "ImageURL"), 
                    (False, res["Up"], "", ICONS_DIR + "up_16.png"), 
                    {"ActionCommand": "up", "ActionListener": self}), 
                ("Button", "btn_down", 
                    LEFT_MARGIN + (BUTTON_SEP + BUTTON_WIDTH) * 5, TOP_MARGIN, 
                    BUTTON_WIDTH, BUTTON_HEIGHT, 
                    ("FocusOnClick", "HelpText", "HelpURL", "ImageURL"), 
                    (False, res["Down"], "", ICONS_DIR + "down_16.png"), 
                    {"ActionCommand": "down", "ActionListener": self}), 
                ("Button", "btn_option", 
                    LEFT_MARGIN + (BUTTON_WIDTH + BUTTON_SEP) * 6, TOP_MARGIN, 
                    BUTTON_WIDTH, BUTTON_HEIGHT, 
                    ("FocusOnClick", "HelpText", "HelpURL", "ImageURL"), 
                    (False, res["Option"], "", ICONS_DIR + "tune_16.png"), 
                    {"ActionCommand": "option", "ActionListener": self}), 
                ("Edit", "edit_input", 
                    LEFT_MARGIN, TOP_MARGIN * 2 + BUTTON_HEIGHT, 
                    ps.Width, INPUT_LINE_HEIGHT, 
                    ("HelpText", "HelpURL"), 
                    (res["Input line"], ""))
            )
        )
        # ToDo use SpinningProgressControlModel if fixed
        _spinner_image = ICONS_DIR + "spin_16_%02d.png"
        spinner = create_control(ctx, "AnimatedImagesControl", 
            LEFT_MARGIN + (BUTTON_WIDTH + BUTTON_SEP) * 6, 
            TOP_MARGIN + 6, 
            16, 16, 
            (), (), 
        )
        cont.addControl("spinner", spinner)
        spinner_model = spinner.getModel()
        spinner_model.insertImageSet(0, 
            tuple([_spinner_image % i for i in range(1, 6)]))
        spinner_model.AutoRepeat = True
        spinner_model.ScaleMode = 0
        spinner.setVisible(False)
        spinner.addMouseListener(SpinnerMouseListener(self))
        
        grid_y = TOP_MARGIN + BUTTON_HEIGHT + TOP_MARGIN + \
            ((TOP_MARGIN + INPUT_LINE_HEIGHT) if show_input_line else 0)
        
        grid = create_control(ctx, "grid.UnoControlGrid", 
            0, grid_y, ps.Width, ps.Height - grid_y, 
            ("Border", "EvenRowBackgroundColor", "HScroll",
                "SelectionModel", "ShowColumnHeader", 
                "ShowRowHeader", "VScroll"), 
            (0, 0xeeeeee, False, ST_SINGLE, True, False, True))
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
        grid.addMouseListener(self)
        grid.addSelectionListener(self)
        
        edit_input = cont.getControl("edit_input")
        self.switch_input_line(show_input_line)

