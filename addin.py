import datetime
import os
import uuid
import xml.etree.ElementTree

class XMLSerializable(object):
    def xmlNode(self, parent_node):
        raise NotImplementedError("Method not implemented for %r" % self.__class__)
    @classmethod
    def fromNode(cls, node, id_cache=None):
        raise NotImplementedError("Method not implemented for %r" % cls)

class UIControl(XMLSerializable):
    def refNode(self, parent):
        return xml.etree.ElementTree.SubElement(self.__class__.__name__, 
                                                {'refID': self.id})

class ControlContainer(XMLSerializable):
    def __init__(self):
        self._items = []
    def addItemsToNode(self, parent_node):
        elt = xml.etree.ElementTree.SubElement(parent_node, 'Items')
        for item in self._items:
            item.refNode(elt)

class Menu(UIControl, ControlContainer):
    pass

class ToolBar(UIControl, ControlContainer):
    pass

class PythonAddin(object):
    def __init__(self, name, description, namespace, author='Untitled', company='Untitled', version='0.1', image='', app='ArcMap'):
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
    @property
    def commands(self):
        return []
    @property
    def menus(self):
        return []
    @property
    def toolbars(self):
        return []
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
    print myaddin.xml
