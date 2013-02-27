
import uno
import unohelper

from com.sun.star.awt.PosSize import POSSIZE as PS_POSSIZE

def create_service(ctx, name, args=None):
    """ Create service with args if required. """
    smgr = ctx.getServiceManager()
    if args:
        return smgr.createInstanceWithArgumentsAndContext(name, args, ctx)
    else:
        return smgr.createInstanceWithContext(name, ctx)

def create_control(ctx, control_type, x, y, width, height, names, values, model_name=None):
    """ create a control. """
    smgr = ctx.getServiceManager()
    
    ctrl = smgr.createInstanceWithContext(
        "com.sun.star.awt." + control_type, ctx)
    if not model_name:
        model_name = "com.sun.star.awt." + control_type + "Model"
    ctrl_model = smgr.createInstanceWithContext(model_name, ctx)
    
    if len(names) > 0:
        ctrl_model.setPropertyValues(names, values)
    ctrl.setModel(ctrl_model)
    ctrl.setPosSize(x, y, width, height, PS_POSSIZE)
    return ctrl


def create_container(ctx, parent, names, values):
    """ create control container. """
    cont = create_control(ctx, "UnoControlContainer", 0, 0, 0, 0, names, values)
    cont.createPeer(parent.getToolkit(), parent)
    return cont


def create_controls(ctx, container, controls):
    """
    ((TYPE, NAME, x, y, width, height, PROP_NAMES, PROP_VALUES, OPTIONS), ())
    """
    smgr = ctx.getServiceManager()
    for defs in controls:
        c = smgr.createInstanceWithContext(
            "com.sun.star.awt.UnoControl" + defs[0], ctx)
        cm = smgr.createInstanceWithContext(
            "com.sun.star.awt.UnoControl" + defs[0] + "Model", ctx)
        cm.setPropertyValues(defs[6], defs[7])
        c.setModel(cm)
        c.setPosSize(defs[2], defs[3], defs[4], defs[5], PS_POSSIZE)
        
        container.addControl(defs[1], c)
        if len(defs) == 9:
            options = defs[8]
            if defs[0] == "Button":
                if "ActionCommand" in options:
                    c.setActionCommand(options["ActionCommand"])
                if "ActionListener" in options:
                    c.addActionListener(options["ActionListener"])
            elif defs[0] == "Combo":
                if "TextListener" in options:
                    c.addTextListener(options["TextListener"])
        
    return container


def get_backgroundcolor(window):
    """ Get background color through accesibility api. """
    try:
        return window.getAccessibleContext().getBackground()
    except:
        pass
    return 0xeeeeee


from com.sun.star.task import XInteractionHandler
 
class DummyHandler(unohelper.Base, XInteractionHandler):
    """ dummy XInteractionHanlder interface for 
        the StringResouceWithLocation """
    def __init__(self): pass
    def handle(self,request): pass


def get_resource(ctx, location, name, locale):
    """ load from resource and returns them as a dictionary. """
    try:
        resolver = ctx.getServiceManager().createInstanceWithContext(
            "com.sun.star.resource.StringResourceWithLocation", ctx)
        resolver.initialize((location, True, locale, name, "", DummyHandler()))
        return resolver
    except Exception as e:
        print(e)
    return None


def load_resource_as_dict(ctx, dir_url, file_name, locale, include_id=False):
    """ Load resource as dict. """
    res = get_resource(ctx, dir_url, file_name, locale)
    strings = {}
    if res:
        default_locale = res.getDefaultLocale()
        for id in res.getResourceIDs():
            str_id = res.resolveStringForLocale(id, default_locale)
            resolved = res.resolveString(id)
            strings[str_id] = resolved
            if include_id:
                strings[id] = resolved
    return strings


def get_current_resource(ctx, dir_url, file_name):
    """ Get resource for current locale. """
    return load_resource_as_dict(ctx, dir_url, file_name, get_ui_locale(ctx))


class StringResource(object):
    """ Keeps string resource. """
    
    def __init__(self, res):
        self._res = res
    
    def get(self, name):
        try:
            return self._res[name]
        except:
            return name
    
    def __getitem__(self, name):
        try:
            return self._res[name]
        except:
            return name


from com.sun.star.lang import Locale
from com.sun.star.beans import PropertyValue
from com.sun.star.beans.PropertyState import DIRECT_VALUE as PS_DIRECT_VALUE


def get_config_access(ctx, nodepath, updatable=False):
    """ get configuration access. """
    arg = PropertyValue("nodepath", 0, nodepath, PS_DIRECT_VALUE)
    cp = ctx.getServiceManager().createInstanceWithContext(
        "com.sun.star.configuration.ConfigurationProvider", ctx)
    if updatable:
        return cp.createInstanceWithArguments(
        "com.sun.star.configuration.ConfigurationUpdateAccess", (arg,))
    else:
        return cp.createInstanceWithArguments(
        "com.sun.star.configuration.ConfigurationAccess", (arg,))


def get_config_value(ctx, nodepath, name):
    """ get configuration value. """
    cua = get_config_access(ctx, nodepath)
    return cua.getPropertyValue(name)


def get_ui_locale(ctx):
    """ get UI locale as css.lang.Locale struct. """
    locale = get_config_value(ctx, "/org.openoffice.Setup/L10N", "ooLocale")
    parts = locale.split("-")
    lang = parts[0]
    country = ""
    if len(parts) == 2:
        country = parts[1]
    return Locale(lang, country, "")


def create_dialog(ctx, url):
    """ create dialog from url. """
    try:
        return ctx.getServiceManager().createInstanceWithContext(
            "com.sun.star.awt.DialogProvider", ctx).createDialog(url)
    except:
        return None


def get_extension_package(ctx, ext_id):
    """ Get extension package for extension id. """
    repositories = ("user", "shared", "bundle")
    manager_name = "/singletons/com.sun.star.deployment.ExtensionManager"
    manager = None
    if ctx.hasByName(manager_name):
        # 3.3 is required
        manager = ctx.getByName(manager_name)
    else:
        return None
    package = None
    for repository in repositories:
        package = manager.getDeployedExtension(repository, ext_id, "", None)
        if package:
            break
    return package


def get_package_info(ctx, ext_id):
    """ Returns package name and version. """
    package = get_extension_package(ctx, ext_id)
    if package:
        return package.getDisplayName(), package.getVersion()
    return "", ""


def get_text_content(ctx, file_url, encoding="utf-8"):
    sfa = create_service(ctx, "com.sun.star.ucb.SimpleFileAccess")
    if sfa.exists(file_url):
        textio = create_service(ctx, "com.sun.star.io.TextInputStream")
        try:
            io = sfa.openFileRead(file_url)
            textio.setInputStream(io)
            textio.setEncoding(encoding)
            lines = []
            while not textio.isEOF():
                lines.append(textio.readLine())
            io.closeInput()
            return "\n".join(lines)
        except:
            pass
    return None


def check_method_parameter(ctx, interface_name, method_name, param_index, param_type):
    """ Check the method has specific type parameter at the specific position. """
    cr = create_service(ctx, "com.sun.star.reflection.CoreReflection")
    try:
        idl = cr.forName(interface_name)
        m = idl.getMethod(method_name)
        if m:
            info = m.getParameterInfos()[param_index]
            return info.aType.getName() == param_type
    except:
        pass
    return False


from com.sun.star.awt.MenuItemStyle import CHECKABLE as MIS_CHECKABLE

class MenuEntry(object):
    """ an entry of the popup menu for create_popup function. """
    def __init__(self, label, id, position, command="", style=0):
        self.label = label
        self.id = id
        self.position = position
        self.style = style
        self.command = command


def create_popup(ctx, items, hide_disabled=False):
    """ creates new popup menu. """
    popup = ctx.getServiceManager().createInstanceWithContext(
        "com.sun.star.awt.PopupMenu", ctx)
    popup.hideDisabledEntries(hide_disabled)
    for item in items:
        if item.id == -1:
            popup.insertSeparator(item.position)
        else:
            popup.insertItem(item.id, item.label, item.style, item.position)
            popup.setCommand(item.id, item.command)
    return popup

from com.sun.star.awt import Point, Rectangle
from com.sun.star.awt.PopupMenuDirection import EXECUTE_DEFAULT as PMD_EXECUTE_DEFAULT

class PopupMenuWrapper(object):
    """ Wrapps popup menu for compatibility. """
    
    def __init__(self, ctx, items, hide_disabled=False):
        self._popup = create_popup(ctx, items, hide_disabled)
        self._use_point = check_method_parameter(
            ctx, "com.sun.star.awt.XPopupMenu", "execute", 
            1, "com.sun.star.awt.Point")
    
    def execute(self, peer, x, y, direction=PMD_EXECUTE_DEFAULT):
        if self._use_point:
            pos = Point(x, y)
        else:
            pos = Rectangle(x, y, 0, 0)
        return self._popup.execute(peer, pos, direction)
    
    def getCommand(self, id):
        return self._popup.getCommand(id)
    
    def setPopupMenu(self, id, sub_popup):
        if isinstance(sub_popup, PopupMenuWrapper):
            sub_popup = sub_popup._popup
        self._popup.setPopupMenu(id, sub_popup)
    
    def checkItem(self, id, state):
        self._popup.checkItem(id, state)
    
    def enableItem(self, id, state):
        self._popup.enableItem(id, state)


def messagebox(ctx, parent, message, title, message_type, buttons):
    """ Show message in message box. """
    toolkit = parent.getToolkit()
    older_imple = check_method_parameter(
        ctx, "com.sun.star.awt.XMessageBoxFactory", "createMessageBox", 
        1, "com.sun.star.awt.Rectangle")
    if older_imple:
        msgbox = toolkit.createMessageBox(
            parent, Rectangle(), message_type, buttons, title, message)
    else:
        message_type = uno.getConstantByName("com.sun.star.awt.MessageBoxType." + {
            "messbox": "MESSAGEBOX", "infobox": "INFOBOX", 
            "warningbox": "WARNINGBOX", "errorbox": "ERRORBOX", 
            "querybox": "QUERYBOX"}[message_type])
        msgbox = toolkit.createMessageBox(
            parent, message_type, buttons, title, message)
    n = msgbox.execute()
    msgbox.dispose()
    return n


from com.sun.star.sheet import SingleReference, ComplexReference
from com.sun.star.table import CellRangeAddress
from com.sun.star.table.CellContentType import FORMULA as CCT_FORMULA

def get_cell_references(doc, cell):
    """ get references in the cell formula. """
    if not cell.getType() == CCT_FORMULA:
        return ()
    
    addr = cell.getCellAddress()
    tokens = cell.getTokens()
    addresses = []
    
    for token in tokens:
        data = token.Data
        if token.OpCode == 0 and data != None:
            if isinstance(data, SingleReference):
                ref = data
                addresses.append(CellRangeAddress(addr.Sheet + ref.RelativeSheet, 
                    addr.Column + ref.RelativeColumn, addr.Row + ref.RelativeRow, 
                    addr.Column + ref.RelativeColumn, addr.Row + ref.RelativeRow))
                
            elif isinstance(data, ComplexReference):
                ref1 = data.Reference1
                ref2 = data.Reference2
                
                addresses.append(CellRangeAddress(addr.Sheet + ref1.RelativeSheet, 
                    addr.Column + ref1.RelativeColumn, addr.Row + ref1.RelativeRow, 
                    addr.Column + ref1.RelativeColumn, addr.Row + ref1.RelativeRow))
                
                addresses.append(CellRangeAddress(addr.Sheet + ref2.RelativeSheet, 
                    addr.Column + ref2.RelativeColumn, addr.Row + ref2.RelativeRow, 
                    addr.Column + ref2.RelativeColumn, addr.Row + ref2.RelativeRow))
    
    ranges = doc.createInstance(
        "com.sun.star.sheet.SheetCellRanges")
    ranges.addRangeAddresses(tuple(addresses), False)
    return ranges.getRangeAddressesAsString().split(";")

