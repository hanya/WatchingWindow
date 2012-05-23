
from pyww import DIALOG_DIR
from pyww.helper import create_dialog


class DialogBase(object):
    
    def __init__(self, ctx, res, **kwds):
        self.ctx = ctx
        self.res = res
        self.kwds = kwds
    
    def _(self, name):
        return self.res.get(name)
    
    def execute(self):
        result = None
        self._create()
        self._init()
        n = self._execute()
        if n:
            result = self._result()
        self._dispose()
        return result


class Dialog(DialogBase):
    
    URI = ""
    
    def create(self, name):
        return self.ctx.getServiceManager().createInstanceWithContext(
            name, self.ctx)
    
    def _create(self):
        dialog = create_dialog(self.ctx, self.URI)
        self._translate(dialog)
        self.dialog = dialog
    
    def _translate(self, container):
        _ = self._
        container_model = container.getModel()
        if hasattr(container_model, "Title"):
            container_model.Title = _(container_model.Title)
        for c in container.getControls():
            m = c.getModel()
            if hasattr(m, "Label"):
                m.Label = _(m.Label)
    
    def _dispose(self):
        if self.dialog:
            self.dialog.dispose()
        self.dialog = None
    
    def _execute(self):
        if self.dialog:
            return self.dialog.execute()
    
    def _result(self):
        pass
    
    def _init(self):
        pass
    
    def get(self, name):
        return self.dialog.getControl(name)
    
    def set_label(self, name, label):
        self.get(name).getModel().Label = label
    
    def set_text(self, name, text):
        self.get(name).getModel().Text = text
    
    def get_state(self, name):
        return self.get(name).getModel().State == 1
    
    def set_state(self, name, state):
        self.get(name).getModel().State = int(state)
    
    def set_properties(self, name, names, props):
        c = self.get(name)
        if c:
            c.getModel().setPropertyValues(names, props)


class SettingsDialog(Dialog):
    
    URI = DIALOG_DIR + "Settings.xdl"
    
    def get_value(self, name):
        return self.get(name).getModel().Value
    
    def set_value(self, name, value):
        self.get(name).getModel().Value = value
    
    def _result(self):
        result = {}
        result["warn_cells"] = int(self.get_value("num_AddCells"))
        result["show_input_line"] = self.get_state("check_InputLine")
        result["store"] = self.get_state("check_StoreWatches")
        return result
    
    def _init(self):
        kwds = self.kwds
        self.set_value("num_AddCells", kwds["warn_cells"])
        self.set_state("check_InputLine", kwds["show_input_line"])
        self.set_state("check_StoreWatches", kwds["store"])


class AboutDialog(Dialog):
    
    URI = DIALOG_DIR + "About.xdl"
    
    def _init(self):
        self.set_properties("edit_text", ("PaintTransparent",), (True,))
        kwds = self.kwds
        self.set_label("label_name", kwds["name"])
        self.set_label("label_version", kwds["version"])
        self.set_text("edit_text", kwds["text"])

