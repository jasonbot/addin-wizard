import datetime
import itertools
import random
import os
import shutil
import time
import uuid
import xml.etree.ElementTree
import xml.dom.minidom

NAMESPACE = "{http://schemas.esri.com/Desktop/AddIns}"

def makeid(prefix="id", seen=set()):
    if prefix[-1].isdigit():
        newnum = int(''.join(char for char in prefix if char.isdigit()))
        prefix = ''.join(char for char in prefix if not char.isdigit())
        seen.add(newnum)
    num = 1
    while num in seen:
        num += 1
    seen.add(num)
    return prefix + str(num)

class XMLSerializable(object):
    __registry__ = {}
    def xmlNode(self, parent_node):
        raise NotImplementedError("Method not implemented for %r" % 
                                                            self.__class__)
    @classmethod
    def loadNode(cls, node, id_cache=None):
        tagname = node.tag[len(NAMESPACE):]
        if 'refID' in node.attrib and id_cache and node.attrib['refID'] in id_cache:
            return id_cache[node.attrib['refID']]
        elif tagname in cls.__registry__:
            return cls.__registry__[tagname].fromNode(node, id_cache)
        raise NotImplementedError("Deserialization not implemented for %r" % cls)
    @classmethod
    def registerType(cls, klass):
        cls.__registry__[klass.__name__] = klass
        return klass

class XMLAttrMap(XMLSerializable):
    def addAttrMap(self, node):
        if hasattr(self, '__attr_map__'):
            for attr_node, attr_name in self.__attr_map__.iteritems():
                value = getattr(self, attr_name, '')
                if isinstance(value, bool):
                    value = str(value).lower()
                else:
                    value = unicode(value)
                node.attrib[attr_node] = value
        node.tail = "\n        "
    @classmethod
    def fromNode(cls, node, id_cache=None):
        instance = cls()
        for attrib, mapping_attrib in getattr(cls, '__attr_map__', {}).iteritems():
            val = node.attrib.get(attrib, '')
            if val in ('true', 'false'): # Coerce bool?
                val = True if val == 'true' else False
            else: # Coerce int?
                try:
                    val = int(val)
                except ValueError:
                    pass
            setattr(instance, mapping_attrib, val)
        if hasattr(instance, 'items'):
            item_nodes = node.find(NAMESPACE+"Items")
            for item in item_nodes.getchildren() if item_nodes is not None else []:
                instance.items.append(XMLSerializable.loadNode(item, id_cache))
        if 'id' in node.attrib and isinstance(id_cache, dict):
            id_cache[node.attrib['id']] = instance
            makeid(node.attrib['id'])
        if 'class' in node.attrib:
            makeid(node.attrib['class'])
        help_node = node.find(NAMESPACE+"Help")
        if help_node is not None:
            instance.help_heading = help_node.attrib.get('heading', '')
            instance.help_string = (help_node.text or '').strip()
        return instance

class HasPython(object):
    @property
    def python(self):
        methods = getattr(self, '__python_methods__', [])
        method_string = "\n".join(
            "    def {0}({1}):\n{2}        pass".format(method, 
                                                        ", ".join(args),
                                                        "        {0}\n".format(repr(doc)) 
                                                                               if doc
                                                                               else '')
                for method, doc, args in methods
        )
        init_code = getattr(self, '__init_code__', '')
        if init_code:
            method_string = "    def __init__(self):\n{0}\n{1}".format(
                    "\n".join("        "+ line 
                        for line in init_code), 
                    method_string)
        if not method_string:
            method_string = "    pass"
        return "class {0}(object):\n{1}".format(self.klass, method_string)

class Command(object):
    def refNode(self, parent):
        return xml.etree.ElementTree.SubElement(parent,
                                                self.__class__.__name__, 
                                                {'refID': self.id})

class UIControl(Command, XMLSerializable, HasPython):
    pass

@XMLSerializable.registerType
class Extension(XMLAttrMap, HasPython):
    "Extension"
    __attr_map__ = {'name': 'name',
                    'description': 'description',
                    'class': 'klass',
                    'id': 'id',
                    'category': 'category',
                    'showInExtensionDialog': 'show_in_dialog',
                    'autoLoad': 'auto_load'}
    __python_methods__ = [('startup', '', ['self']),
                          ('shutdown', '', ['self']),
                          ('activeViewChanged', '', ['self']),
                          ('mapsChanged', '', ['self']),
                          ('newDocument', '', ['self']),
                          ('openDocument', '', ['self']),
                          ('beforeCloseDocument', '', ['self']),
                          ('closeDocument', '', ['self']),
                          ('beforePageIndexExtentChange', '', ['self', 'old_id']),
                          ('pageIndexExtentChanged', '', ['self', 'new_id']),
                          ('contentsChanged', '', ['self']),
                          ('contentsCleared', '', ['self']),
                          ('focusMapChanged', '', ['self']),
                          ('spatialReferenceChanged', '', ['self']),
                          ]
    def __init__(self, name=None, description=None, klass=None, id=None, category=None):
        self.name = name or 'New Extension'
        self.description = description or ''
        self.klass = klass or makeid("ExtensionClass")
        self.id = id or makeid("extension")
        self.category = category or ''
        self.show_in_dialog = True
        self.auto_load = True
    def xmlNode(self, parent):
        newnode = xml.etree.ElementTree.SubElement(parent, 
                                                   self.__class__.__name__)
        self.addAttrMap(newnode)
        return newnode

class ControlContainer(XMLAttrMap):
    def __init__(self):
        self.items = []
    def addItemsToNode(self, parent_node):
        elt = xml.etree.ElementTree.SubElement(parent_node, 'Items')
        for item in self.items:
            if hasattr(item, 'refNode'):
                item.refNode(elt)
            else:
                item.xmlNode(elt)
    def xmlNode(self, parent):
        newnode = xml.etree.ElementTree.SubElement(parent, 
                                                   self.__class__.__name__)
        self.addAttrMap(newnode)
        self.addItemsToNode(newnode)

@XMLSerializable.registerType
class Menu(ControlContainer):
    "Menu"
    __attr_map__ = {'caption': 'caption', 
                    'isRootMenu': 'top_level',
                    'isShortcutMenu': 'shortcut_menu',
                    'separator': 'separator',
                    'category': 'category',
                    'id': 'id'}
    def __init__(self, caption='Menu', top_level=False, shortcut_menu=False, separator=False, category=None, id=None):
        super(Menu, self).__init__()
        self.caption = caption or ''
        self.top_level = bool(top_level)
        self.shortcut_menu = bool(shortcut_menu)
        self.separator = bool(separator)
        self.category = category or ''
        self.id = id or makeid("menuitem")

@XMLSerializable.registerType
class ToolPalette(ControlContainer, Command):
    "Tool Palette"
    __attr_map__ = {'columns': 'columns',
                    'canTearOff': 'tearoff',
                    'isMenuStyle': 'menu_style',
                    'category': 'category',
                    'id': 'id'}
    def __init__(self, caption=None, columns=2, tearoff=False, menu_style=False, category=None, id=None):
        super(ToolPalette, self).__init__()
        self.caption = caption or 'Palette'
        self.columns = columns
        self.tearoff = bool(tearoff)
        self.menu_style = bool(menu_style)
        self.category = category or ''
        self.id = id or makeid("tool_palette")

@XMLSerializable.registerType
class Toolbar(ControlContainer):
    "Toolbar"
    __attr_map__ = {'caption': 'caption',
                    'category': 'category',
                    'id': 'id',
                    'showInitially': 'show_initially'}
    def __init__(self, id=None, caption=None, category=None, show_initially=True):
        super(Toolbar, self).__init__()
        self.id = id or makeid("toolbar")
        self.caption = caption or 'Toolbar'
        self.category = category or ''
        self.show_initially = bool(show_initially)

@XMLSerializable.registerType
class Button(XMLAttrMap, UIControl):
    "Button"
    __python_methods__ = [('onClick', '', ['self'])]
    __init_code__ = ['self.enabled = True', 'self.checked = False']
    __attr_map__ = {'caption': 'caption' ,
                    'class': 'klass',
                    'category': 'category',
                    'image': 'image' ,
                    'tip': 'tip' ,
                    'message': 'message' ,
                    'id': 'id',
                    'message': 'message'}
    def __init__(self, caption=None, klass=None, category=None, image=None,
                 tip=None, message=None, id=None):
        self.caption = caption or "Button"
        self.klass = klass or makeid("ButtonClass")
        self.category = category or ''
        self.image = image or ''
        self.tip = tip or ''
        self.message = message or ''
        self.id = id or makeid("button")
        self.help_heading = ''
        self.help_string = ''
    def xmlNode(self, parent):
        newnode = xml.etree.ElementTree.SubElement(parent,
                                                self.__class__.__name__)
        self.addAttrMap(newnode)
        help = xml.etree.ElementTree.SubElement(newnode, 'Help', {'heading': self.help_heading})
        help.text = self.help_string
        return newnode

@XMLSerializable.registerType
class ComboBox(Button):
    "Combo Box"
    __attr_map__ = {'caption': 'caption',
                    'category': 'category',
                    'id': 'id',
                    'class': 'klass',
                    'tip': 'tip',
                    'message': 'message',
                    'sizeString': 'size_string',
                    'itemSizeString': 'item_size_string',
                    'hintText': 'hint_text',
                    'rows': 'rows'}
    @property
    def __init_code__(self):
        return ['self.items = ["item1", "item2"]', 'self.editable = %r' % self.editable, 'self.enabled = True']
    __python_methods__ = [('onSelChange',  '', ['self', 'selection']),
                          ('onEditChange', '', ['self', 'text']),
                          ('onFocus', '', ['self', 'focused']),
                          ('onEnter', '', ['self'])
                         ]
    def __init__(self, klass=None, id=None, category=None, caption=None):
        self.klass = klass or makeid("ComboBoxClass")
        self.id = id or makeid("combo_box")
        self.category = category or ''
        self.caption = caption or "ComboBox"
        self.message = ''
        self.tip = ''
        self.help_heading = ''
        self.help_string = ''
        self.size_string = "WWWWWW"
        self.item_size_string = "WWWWWW"
        self.hint_text = ''
        self.editable = True
        self.rows = 4

@XMLSerializable.registerType
class Tool(Button):
    "Python Tool"
    __init_code__ = ['self.enabled = True']
    __python_methods__ = [('onMouseDown', '', ['self', 'x', 'y', 'button', 'shift']),
                          ('onMouseDownMap', '', ['self', 'x', 'y', 'button', 'shift']),
                          ('onMouseUp', '', ['self', 'x', 'y', 'button', 'shift']),
                          ('onMouseUpMap', '', ['self', 'x', 'y', 'button', 'shift']),
                          ('onMouseMove', '', ['self', 'x', 'y', 'button', 'shift']),
                          ('onMouseMoveMap', '', ['self', 'x', 'y', 'button', 'shift']),
                          ('onDblClick', '', ['self']),
                          ('onKeyDown', '', ['self', 'keycode', 'shift']),
                          ('onKeyUp', '', ['self', 'keycode', 'shift']),
                          ('deactivate', '', ['self'])
                          ]
    def __init__(self, caption=None, klass=None, category=None, image=None,
                 tip=None, message=None, id=None):
        super(Tool, self).__init__(caption, klass, category, image, tip, message, id)
        self.caption = caption or 'Tool'
        self.klass = klass or makeid("ToolClass")
        self.id = id or makeid("tool")
        self.image = image or ''
        self.help_heading = ''
        self.help_string = ''

@XMLSerializable.registerType
class MultiItem(XMLAttrMap, UIControl):
    "MultiItem"
    __attr_map__ = {'id': 'id',
                    'class': 'klass',
                    'hasSeparator': 'separator'}
    __python_methods__ = [('onItemClick', '', ['self'])]
    __init_code__ = ['self.items = ["item1", "item2"]']
    def __init__(self, klass=None, id=None, separator=None):
        self.klass = klass or makeid("MultiItemClass")
        self.id = id or makeid("multi_item")
        self.caption = 'MultiItem'
        self.separator = bool(separator)
    def xmlNode(self, parent):
        newnode = xml.etree.ElementTree.SubElement(parent,
                                                self.__class__.__name__)
        self.addAttrMap(newnode)
        return newnode

class PythonAddin(object):
    __apps__ = ('ArcMap', 'ArcCatalog', 'ArcGlobe', 'ArcScene')
    def __init__(self, name, description, namespace, author='Untitled',
                 company='Untitled', version='0.1', image='', app='ArcMap'):
        self.name = name
        self.description = description
        self.namespace = namespace
        self.author = author
        self.company = company
        assert app in self.__apps__
        self.app = app
        self.guid = "{%s}" % uuid.uuid4()
        self.version = version
        self.image = image
        self.addinfile = self.namespace + '.py'
        self.items = []
    def remove(self, target_item):
        def rm_(container_object, target_item):
            items = getattr(container_object, 'items', [])
            if target_item in items:
                container_object.items.pop(container_object.items.index(target_item))
                return True
            else:
                for item in (i for i in items if isinstance(i, ControlContainer)):
                    if rm_(item, target_item):
                        return True
            return False
        return rm_(self, target_item)
    @property
    def commands(self):
        ids = set()
        for container in [i for i in self if isinstance(i, ControlContainer)]:
            for item in container.items:
                if isinstance(item, Command):
                    if item.id not in ids:
                        ids.add(item.id)
                        yield item
    @property
    def menus(self):
        return [menu for menu in self.items if isinstance(menu, Menu)]
    @property
    def toolbars(self):
        return [toolbar for toolbar in self.items if isinstance(toolbar, Toolbar)]
    @property
    def extensions(self):
        return [extension for extension in self.items if isinstance(extension, Extension)]
    @classmethod
    def fromXML(cls, xmlfile):
        new_addin = cls('unknown', 'unknown', 'unknown')
        id_cache = {}
        doc = xml.etree.ElementTree.parse(xmlfile)
        root = doc.getroot()
        assert root.tag == NAMESPACE+'ESRI.Configuration', root.tag
        new_addin.name = (root.find(NAMESPACE+"Name").text or '').strip()
        new_addin.guid = (root.find(NAMESPACE+"AddInID").text or '').strip()
        new_addin.description = (root.find(NAMESPACE+"Description").text or '').strip()
        new_addin.version = (root.find(NAMESPACE+"Version").text or '').strip()
        new_addin.image = (root.find(NAMESPACE+"Image").text or '').strip()
        new_addin.author = (root.find(NAMESPACE+"Author").text or '').strip()
        new_addin.company = (root.find(NAMESPACE+"Company").text or '').strip()
        addin_node = root.find(NAMESPACE+"AddIn")
        new_addin.addinfile = addin_node.attrib.get('library', '')
        projectpath = os.path.join(os.path.dirname(xmlfile), 'Install', os.path.dirname(new_addin.addinfile))
        if os.path.isfile(os.path.join(projectpath, new_addin.addinfile)):
            addin_py = os.path.join(projectpath, new_addin.addinfile)
            path, ext = os.path.splitext(addin_py)

            new_index = 1
            new_file = path + "_" + str(new_index) + ext
            while os.path.exists(new_file):
                new_index += 1
                new_file = path + "_" + str(new_index) + ext
            #shutil.copyfile(addin_py, new_file)
            new_addin.backup_data = (addin_py, new_file)
            new_addin.warning = u"Python script {0} already exists. Will create a backup as {1} upon first save.".format(addin_py, new_file)
        app_node = addin_node.getchildren()[0]
        new_addin.app = app_node.tag[len(NAMESPACE):]
        for command in app_node.find(NAMESPACE+"Commands").getchildren():
            XMLSerializable.loadNode(command, id_cache)
        for tag in ("Extensions", "Toolbars", "Menus"):
            for item in app_node.find(NAMESPACE+tag).getchildren():
                new_addin.items.append(XMLSerializable.loadNode(item, id_cache))
        return new_addin
    def backup(self):
        backup_data = getattr(self, 'backup_data', None)
        if backup_data:
            shutil.copyfile(*backup_data)
        self.backup_data = None
    @property
    def xml(self):
        root = xml.etree.ElementTree.Element('ESRI.Configuration',
                {'xmlns': "http://schemas.esri.com/Desktop/AddIns",
                 'xmlns:xsi': "http://www.w3.org/2001/XMLSchema-instance"})
        xml.etree.ElementTree.SubElement(root, 'Name').text = self.name
        xml.etree.ElementTree.SubElement(root, 'AddInID').text = self.guid
        xml.etree.ElementTree.SubElement(root, 'Description').text = self.description
        xml.etree.ElementTree.SubElement(root, 'Version').text = self.version
        xml.etree.ElementTree.SubElement(root, 'Image').text = self.image
        xml.etree.ElementTree.SubElement(root, 'Author').text = self.author
        xml.etree.ElementTree.SubElement(root, 'Company').text = self.company
        xml.etree.ElementTree.SubElement(root, 'Date').text = datetime.datetime.now().strftime("%m/%d/%Y")
        targets = xml.etree.ElementTree.SubElement(root, 'Targets')
        target = xml.etree.ElementTree.SubElement(targets, 'Target', {'name': "Desktop", 'version': "10.1"})
        addinnode = xml.etree.ElementTree.SubElement(root, 'AddIn', {'language': 'PYTHON', 
                                                                     'library': self.addinfile,
                                                                     'namespace': self.namespace})
        appnode = xml.etree.ElementTree.SubElement(addinnode, self.app)
        appnode.text = "\n    "
        commandnode = xml.etree.ElementTree.SubElement(appnode, 'Commands')
        commandnode.text = "\n        "
        for command in self.commands:
            command.xmlNode(commandnode)
        commandnode.tail = "\n    "
        extensionnode = xml.etree.ElementTree.SubElement(appnode, 'Extensions')
        extensionnode.text = "\n        "
        for extension in self.extensions:
            extension.xmlNode(extensionnode)
        extensionnode.tail = "\n    "
        toolbarnode = xml.etree.ElementTree.SubElement(appnode, 'Toolbars')
        toolbarnode.text = "\n        "
        for toolbar in self.toolbars:
            toolbar.xmlNode(toolbarnode)
        toolbarnode.tail = "\n    "
        menunode = xml.etree.ElementTree.SubElement(appnode, 'Menus')
        menunode.text = "\n        "
        for menu in self.menus:
            menu.xmlNode(menunode)
        menunode.tail = "\n    "
        markup = xml.etree.ElementTree.tostring(root).encode("utf-8")
        return markup
        #return xml.dom.minidom.parseString(markup).toprettyxml("    ")
    def __iter__(self):
        def ls_(item):
            if hasattr(item, 'items'):
                for item_ in item.items:
                    for item__ in ls_(item_):
                        yield item__
            yield item
        for bitem in self.items:
            for aitem in ls_(bitem):
                yield aitem
    @property
    def python(self):
        return "import arcpy\n\n" + "\n\n".join(x.python for x in self if hasattr(x, 'python'))

class PythonAddinProjectDirectory(object):
    def __init__(self, path):
        self._path = path
        if not os.path.exists(path):
            raise IOError(u"{0} does not exist. Please select a directory that exists.".format(path))
        listing = os.listdir(path)
        if listing:
            if not all(item in listing for item in ('config.xml', 'Install', 'Images')):
                raise ValueError(u"{0} is not empty. Please select an empty directory to host this new addin.".format(path))
            else:
                #raise ValueError("{0} exists and is a valid project directory, but loading doesn't work yet.")
                self.addin = PythonAddin.fromXML(os.path.join(self._path, 'config.xml'))
                self.warning = getattr(self.addin, 'warning', None)
        else:
            addin_name = ''.join(x for x in os.path.basename(path) if x.isalpha())
            self.addin = PythonAddin("Python Addin", "New Addin", (addin_name if addin_name else 'python') + "_addin")
    def save(self):
        # Make install/images dir
        install_dir = os.path.join(self._path, 'Install')
        images_dir = os.path.join(self._path, 'Images')
        if not os.path.exists(install_dir):
            os.mkdir(install_dir)
        if not os.path.exists(images_dir):
            os.mkdir(images_dir)

        # Copy packaging\* into project dir
        initial_dirname = os.path.dirname(os.path.abspath(__file__))
        if os.path.isfile(initial_dirname):
            print "UP A DIR"
            initial_dirname = os.path.dirname(initial_dirname)
        packaging_dir = os.path.join(initial_dirname,
                                     'packaging')
        for filename in os.listdir(packaging_dir):
            if not os.path.exists(os.path.join(self._path, filename)):
                shutil.copyfile(os.path.join(packaging_dir, filename), os.path.join(self._path, filename))

        # For consolidating images
        seen_images = set([self.addin.image])

        # Auto-fill category
        for item in self.addin:
            # Auto-category
            if hasattr(item, 'category'):
                item.category = self.addin.name or self.addin.description
            # Collect for later
            if getattr(item, 'image', ''):
                seen_images.add(item.image)

        # Consolidate images
        seen_images = filter(bool, seen_images)
        full_path = dict((image, 
                          os.path.abspath(
                              os.path.join(self._path, image))) 
                          for image in seen_images)
        to_relocate = set(image_file for (image_name, image_file) 
                            in full_path.iteritems()
                                if os.path.dirname(image_file) != images_dir)
        relocated_images = {}
        image_files = set(x.lower() for x in os.listdir(images_dir))
        for image_file in to_relocate:
            new_filename = os.path.basename(image_file)
            if new_filename.lower() in image_files:
                num = 1
                fn, format = os.path.splitext(new_filename)
                while new_filename.lower() in image_files:
                    new_filename = fn + "_" + str(num) + format
                    num += 1
            relocated_images[image_file] = new_filename
            shutil.copyfile(image_file, os.path.join(self._path, 'Images', new_filename))
        for item_with_image in (item for item in 
                                    ([self.addin] + list(self.addin))
                                        if getattr(item, 'image', '')):
            item_image = item_with_image.image
            if item_image in relocated_images:
                item_with_image.image = os.path.join('Images', relocated_images[item_image])
            else:
                item_with_image.image = os.path.join('Images', os.path.basename(item_image))

        # Back up .py file if necessary
        self.addin.backup()
        # Output XML and Python stub
        with open(os.path.join(self._path, 'config.xml'), 'wb') as out_handle:
            out_handle.write(self.addin.xml)
        with open(os.path.join(install_dir, self.addin.addinfile), 'wb') as out_python:
            out_python.write(self.addin.python)

if __name__ == "__main__":
    myaddin = PythonAddin("My Addin", "This is a new addin", "myaddin")
    toolbar = Toolbar()
    toolbar.items.append(Button("Hello there", "HelloButton"))
    myaddin.items.append(toolbar)
    print myaddin.xml
