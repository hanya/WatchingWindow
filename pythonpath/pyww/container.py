
import unohelper
from com.sun.star.lang import XEventListener


class ModelContainer(object):
    """ Keeps model that keeps data of ww to make live until the 
        document is living.
    """
    
    class EventListener(unohelper.Base, XEventListener):
        def __init__(self, act, model):
            self.act = act
            self.model = model
        
        def disposing(self, ev):
            try:
                self.act.remove(self.model)
            except Exception as e:
                print("Error while removeing the model: ")
                print(e)
            self.act = None
    
    def __init__(self):
        self._models = [] # each element is tuple of doc model and ww model
    
    def get(self, frame):
        _doc_model = frame.getController().getModel()
        for doc_model, model in self._models:
            if _doc_model == doc_model:
                return model
        return None
    
    def add(self, model, frame):
        doc_model = frame.getController().getModel()
        self._models.append((doc_model, model))
        doc_model.com_sun_star_lang_XComponent_addEventListener(
                            self.EventListener(self, model))
    
    def remove(self, model):
        try:
            n = None
            for i, (doc_model, _model) in enumerate(self._models):
                if model == _model:
                    n = i
                    break
            if not n is None:
                model = self._models.pop(n)
                try:
                    model[1].dispose()
                except:
                    pass
                model = None
        except Exception as e:
            print(e)


_container = None

def get_model_container():
    global _container
    if _container is None:
        _container = ModelContainer()
    return _container
