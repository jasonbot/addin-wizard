import datetime
import itertools
import random
import os
import uuid
import xml.etree.ElementTree

def makeid(prefix="id"):
    return "%s%05x" % (prefix, random.randint(1, 32000))

class XMLSerializable(object):
    def xmlNode(self, parent_node):
        raise NotImplementedError("Method not implemented for %r" % 
                                                            self.__class__)
    @classmethod
    def fromNode(cls, node, id_cache=None):
        raise NotImplementedError("Method not implemented for %r" % cls)

class UIControl(XMLSerializable):
    def refNode(self, parent):
        return xml.etree.ElementTree.SubElement(parent,
                                                self.__class__.__name__, 
                                                {'refID': self.id})

class ControlContainer(XMLSerializable):
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
        self.addItemsToNode(newnode)

class Menu(ControlContainer):
    pass

class Toolbar(ControlContainer):
    def __init__(self, id=None, caption=None):
        self.id = id or makeid("toolbar")
        self.caption = caption
        super(Toolbar, self).__init__()

class Button(UIControl):
    def __init__(self, caption, klass, category=None, image=None,
                 tip=None, message=None, id=None):
        self.caption = caption
        self.klass = klass
        self.category = category
        self.image = image
        self.tip = tip
        self.message = message
        self.id = id or makeid("button")
    def xmlNode(self, parent):
        return xml.etree.ElementTree.SubElement(parent,
                                                self.__class__.__name__, 
                                                {'caption': self.caption or '',
                                                 'class': self.klass,
                                                 'category': self.category or '',
                                                 'image': self.image or '',
                                                 'tip': self.tip or '',
                                                 'message': self.message or '',
                                                 'id': self.id})

class PythonAddin(object):
    def __init__(self, name, description, namespace, author='Untitled',
                 company='Untitled', version='0.1', image='', app='ArcMap'):
        self.name = name
        self.description = description
        self.namespace = namespace
        self.author = author
        self.company = company
        assert app in ('ArcMap', 'ArcCatalog', 'ArcGlobe', 'ArcScene')
        self.app = app
        self.guid = uuid.uuid4()
        self.version = version
        self.image = image
        self.addinfile = self.namespace + '.py'
        self.items = []
    @property
    def commands(self):
        ids = set()
        for container in itertools.chain(self.menus, self.toolbars):
            for item in container.items:
                if isinstance(item, UIControl):
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
        return []
    @property
    def xml(self):
        root = xml.etree.ElementTree.Element('ESRI.Configuration',
                {'xmlns': "http://schemas.esri.com/Desktop/AddIns",
                 'xmlns:xsi': "http://www.w3.org/2001/XMLSchema-instance"})
        root.text = "\n"
        xml.etree.ElementTree.SubElement(root, 'Name').text = self.name
        xml.etree.ElementTree.SubElement(root, 'AddInID').text = "{%s}" % self.guid
        xml.etree.ElementTree.SubElement(root, 'Description').text = self.description
        xml.etree.ElementTree.SubElement(root, 'Version').text = self.version
        xml.etree.ElementTree.SubElement(root, 'Image').text = self.image
        xml.etree.ElementTree.SubElement(root, 'Author').text = self.author
        xml.etree.ElementTree.SubElement(root, 'Company').text = self.company
        xml.etree.ElementTree.SubElement(root, 'Date').text = datetime.datetime.now().strftime("%M/%d/%Y")
        targets = xml.etree.ElementTree.SubElement(root, 'Targets')
        target = xml.etree.ElementTree.SubElement(targets, 'Target', {'name': "Desktop", 'version': "10.0"})
        targets.tail = "\n"
        addinnode = xml.etree.ElementTree.SubElement(root, 'AddIn', {'language': 'PYTHON', 
                                                                     'library': self.addinfile,
                                                                     'namespace': self.namespace})
        addinnode.text = "\n    "
        addinnode.tail = "\n"
        appnode = xml.etree.ElementTree.SubElement(addinnode, self.app)
        appnode.tail = "\n"
        commandnode = xml.etree.ElementTree.SubElement(appnode, 'Commands')
        for command in self.commands:
            command.xmlNode(commandnode)
        extensionnode = xml.etree.ElementTree.SubElement(appnode, 'Extensions')
        for extension in self.extensions:
            extension.xmlNode(extensionnode)
        toolbarnode = xml.etree.ElementTree.SubElement(appnode, 'Toolbars')
        for toolbar in self.toolbars:
            toolbar.xmlNode(toolbarnode)
        menunode = xml.etree.ElementTree.SubElement(appnode, 'Menus')
        for menu in self.menus:
            meni.xmlNode(menunode)
        return xml.etree.ElementTree.tostring(root).encode("utf-8")

if __name__ == "__main__":
    myaddin = PythonAddin("My Addin", "This is my starting addin", "myaddin")
    toolbar = Toolbar()
    toolbar.items.append(Button("Hello there", "HelloButton"))
    myaddin.items.append(toolbar)
    print myaddin.xml
