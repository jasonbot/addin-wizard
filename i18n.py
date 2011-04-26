__all__ = ['_']

def _(the_text):
    return u''.join(reversed(unicode(the_text)))
