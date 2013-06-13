
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

