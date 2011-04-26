__all__ = ['_']

import json
import os

try:
    translations_dict = json.loads(open(os.path.join('resources', 'resource_strings.json'), 'rb').read().decode('utf-8'))
except:
    print "Can't load translation strings file."

def _(the_text):
    return translations_dict.get(the_text, the_text)
