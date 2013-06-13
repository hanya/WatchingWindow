
from pyww import CONFIG_NODE
from pyww.helper import get_config_access


class Settings(object):
    """ Load and set configuration values. """
    
    def __init__(self, ctx):
        self.ctx = ctx
        self.res = None
        self._loaded = False
        self.check_input = None
    
    def configure(self, res):
        from pyww.dialogs import SettingsDialog
        cua = get_config_access(self.ctx, CONFIG_NODE, True)
        
        show_input_line = cua.getPropertyValue("InputLine")
        store = cua.getPropertyValue("StoreWatches")
        
        dialog = SettingsDialog(
            self.ctx, res, 
            show_input_line=show_input_line, 
            store=store, 
        )
        result = dialog.execute()
        if result:
            self.check_input = result["show_input_line"]
            self.store = result["store"]
            try:
                cua.setPropertyValue("InputLine", self.check_input)
                cua.setPropertyValue("StoreWatches", self.store)
                cua.commitChanges()
            except Exception as e:
                print(e)
            return True
        return False
    
    def _load(self):
        cua = get_config_access(self.ctx, CONFIG_NODE)
        self.check_input = cua.getPropertyValue("InputLine")
        self.store = cua.getPropertyValue("StoreWatches")
    
    def get(self, name):
        """ get specified value. """
        if not self._loaded:
            self._load()
        if name == "InputLine":
            return self.check_input
        elif name == "StoreWatches":
            return self.store


from com.sun.star.rdf.URIs import RDF_VALUE


class SettingsRDF(object):
    """ Store settings in rdf. """
    
    BASE_TYPE = "http://mytools.calc/watchingwindow/1.0"
    TYPE_NAME = "Settings"
    
    FILE_NAME = "mytools_calc_watchingwindow/settings.rdf"
    
    ORDER_NAMESPACE = "watches:"
    ORDER_SUBJECT = "order"
    
    ORDER_SEP = "\t"
    
    def __init__(self, ctx, doc):
        self.ctx = ctx
        self.doc = doc
        self.order = []
        self.graph = self.get_graph()
    
    def __len__(self):
        return len(self.order)
    
    def create(self, name, args):
        return self.ctx.getServiceManager().\
            createInstanceWithArgumentsAndContext(name, args, self.ctx)
    
    def create_uri(self, n, l=None):
        """ Create new URI from namespace and localname or known one. """
        if l is None:
            a = (n,)
        else:
            a = (n, l)
        return self.create("com.sun.star.rdf.URI", a)
    
    def create_literal(self, v):
        """ Create literal from string. """
        return self.create("com.sun.star.rdf.Literal", (v,))
    
    def get_type_uri(self):
        return self.create_uri(self.BASE_TYPE + "/" + self.TYPE_NAME)
    
    def load_order(self, s):
        self.order = s.split(self.ORDER_SEP)
    
    def write_order(self):
        self.update(
            self.ORDER_SUBJECT, 
            self.ORDER_SEP.join(self.order), 
            self.ORDER_NAMESPACE)
    
    def add_list_to_order(self, names):
        if self.graph:
            n = len(self.order)
            self.order[n:n+len(names)] = names
            self.write_order()
    
    def add_to_order(self, name):
        if name in self.order:
            return
        self.order.append(name)
        self.write_order()
    
    def remove_from_order(self, name, store=True):
        try:
            del self.order[self.order.index(name)]
            if store:
                self.write_order()
        except:
            pass
    
    def insert_to_order(self, pos, name, store=True):
        self.order.insert(pos, name)
        if store:
            self.write_order()
    
    def clear_order(self):
        self.order[:] = []
    
    def load(self):
        """ Load statements from rdf. """
        if self.graph:
            value_uri = self.create_uri(RDF_VALUE)
            try:
                enume = self.graph.getStatements(None, value_uri, None)
            except:
                return
            while enume.hasMoreElements():
                st = enume.nextElement()
                s = st.Subject.StringValue
                if s.startswith(self.ORDER_NAMESPACE):
                    self.load_order(st.Object.StringValue)
    
    def add(self, name):
        """ Add name, value pair as rdf statement. """
        if self.graph:
            self.add_to_order(name)
    
    def move(self, pos, name):
        if self.graph:
            self.remove_from_order(name, False)
            self.insert_to_order(pos, name)
    
    def get_graph(self):
        """ Get existing graph or create new one if not found. """
        doc = self.doc
        repo = doc.getRDFRepository()
        type_uri = self.get_type_uri()
        
        graph_names = doc.getMetadataGraphsWithType(type_uri)
        if len(graph_names) == 0:
            graph_name = doc.addMetadataFile(self.FILE_NAME, (type_uri,))
        else:
            graph_name = graph_names[0]
        
        return repo.getGraph(graph_name)
    
    def find_statement(self, name, namespace):
        """ Find statement by name returns first statement found. """
        if self.graph:
            _name = namespace + name
            value_uri = self.create_uri(RDF_VALUE)
            try:
                enume = self.graph.getStatements(None, value_uri, None)
            except:
                self.graph = self.get_graph()
                enume = self.graph.getStatements(None, value_uri, None)
            while enume.hasMoreElements():
                st = enume.nextElement()
                if st.Subject.StringValue == _name:
                    return st
        return None
    
    def update(self, name, value, namespace):
        """ Update specific subject with value. """
        st = self.find_statement(name, namespace)
        if st:
            try:
                self.graph.removeStatements(st.Subject, None, None)
            except:
                pass
        self.graph.addStatement(
            self.create_uri(namespace, name), 
            self.create_uri(RDF_VALUE), 
            self.create_literal(value))
    
    def remove(self, name, pop=True):
        """ Remove statement by its name. """
        if self.graph:
            self.remove_from_order(name)
    
    def clear(self):
        """ Remove all statement. """
        if self.graph:
            self.graph.clear()
            self.doc.removeMetadataFile(self.graph)
        
        self.clear_order()

