import addin
import addin_ui
import os
import re
import sys
import wx

class AddinMakerAppWindow(addin_ui.AddinMakerWindow):
    # Sentinel/singletons for the treeview
    class _extensiontoplevel(object):
        "Extensions"
        pass
    class _menutoplevel(object):
        "Toplevel Menus"
        pass
    class _toolbartoplevel(object):
        "Toolbars"
        pass
    def __init__(self, *args, **kws):
        super(AddinMakerAppWindow, self).__init__(*args, **kws)
        self.Bind(wx.EVT_CLOSE, self.OnClose)

        self._selected_data = None
        self.contents_tree.Bind(wx.EVT_RIGHT_DOWN, self.TreePopupRClick)
        self.contents_tree.Bind(wx.EVT_CONTEXT_MENU, self.TreePopup)
        self.SelectFolder(None)

        self._newextensionid = wx.NewId()
        self._newmenuid = wx.NewId()
        self._newtoolbarid = wx.NewId()

    def AddExtension(self, event):
        extension = addin.Extension()
        self.project.addin.items.append(extension)
        menuitem = self.contents_tree.AppendItem(self.extensionsroot, extension.name)
        self.contents_tree.SetItemPyData(menuitem, extension)
        self.contents_tree.SelectItem(menuitem, True)

    def AddMenu(self, event):
        menu = addin.Menu("Menu", self._selected_data is self._menutoplevel)
        self.project.addin.items.append(menu)
        menuitem = self.contents_tree.AppendItem(self.menusroot, menu.caption)
        self.contents_tree.SetItemPyData(menuitem, menu)
        self.contents_tree.SelectItem(menuitem, True)

    def AddToolbar(self, event):
        toolbar = addin.Toolbar()
        self.project.addin.items.append(toolbar)
        toolbaritem = self.contents_tree.AppendItem(self.toolbarsroot, toolbar.caption)
        self.contents_tree.SetItemPyData(toolbaritem, toolbar)
        self.contents_tree.SelectItem(toolbaritem, True)

    @property
    def extensionmenu(self):
        extensionmenu = wx.Menu()
        cmd = extensionmenu.Append(self._newextensionid, "New Extension")
        extensionmenu.Bind(wx.EVT_MENU, self.AddExtension, cmd)
        return extensionmenu

    @property
    def menumenu(self):
        menumenu = wx.Menu()
        cmd = menumenu.Append(self._newmenuid, "New Menu")
        menumenu.Bind(wx.EVT_MENU, self.AddMenu, cmd)
        return menumenu

    @property
    def toolbarmenu(self):
        toolbarmenu = wx.Menu()
        cmd = toolbarmenu.Append(self._newtoolbarid, "New Toolbar")
        toolbarmenu.Bind(wx.EVT_MENU, self.AddToolbar, cmd)
        return toolbarmenu

    @property
    def controlcontainermenu(self):
        class ItemAppender(object):
            def __init__(self, tree, selection, item, cls, save_button):
                self.tree = tree
                self.selection = selection
                self.item = item
                self.cls = cls
                self.save_button = save_button
            def __call__(self, event):
                try:
                    new_item = self.cls()
                    self.item.items.append(new_item)
                    toolbaritem = self.tree.AppendItem(self.selection, getattr(new_item, 'caption', str(new_item)))
                    self.tree.SetItemPyData(toolbaritem, new_item)
                    self.tree.SelectItem(toolbaritem, True)
                    self.save_button.Enable(True)
                except Exception as e:
                    print e

        tree = self.contents_tree
        selection = self.contents_tree.GetSelection()
        item = self._selected_data

        controlcontainermenu = wx.Menu()
        buttoncmd = controlcontainermenu.Append(-1, "New Button")
        controlcontainermenu.Bind(wx.EVT_MENU, ItemAppender(tree, selection, item, addin.Button, self.save_button), buttoncmd)
        toolcmd = controlcontainermenu.Append(-1, "New Tool")
        controlcontainermenu.Bind(wx.EVT_MENU, ItemAppender(tree, selection, item, addin.Tool, self.save_button), toolcmd)
        if not isinstance(item, addin.ToolPalette):
            menucmd = controlcontainermenu.Append(-1, "New Menu")
            controlcontainermenu.Bind(wx.EVT_MENU, ItemAppender(tree, selection, item, addin.Menu, self.save_button), menucmd)
        if isinstance(item, addin.Menu):
            multiitemcmd = controlcontainermenu.Append(-1, "New MultiItem")
            controlcontainermenu.Bind(wx.EVT_MENU, ItemAppender(tree, selection, item, addin.MultiItem, self.save_button), multiitemcmd)
        if not isinstance(item, (addin.ToolPalette, addin.Menu)):
            palettecmd = controlcontainermenu.Append(-1, "New Tool Palette")
            controlcontainermenu.Bind(wx.EVT_MENU, ItemAppender(tree, selection, item, addin.ToolPalette, self.save_button), palettecmd)
        if isinstance(item, addin.Toolbar):
            comboboxcmd = controlcontainermenu.Append(-1, "New Combo Box")
            controlcontainermenu.Bind(wx.EVT_MENU, ItemAppender(tree, selection, item, addin.ComboBox, self.save_button), comboboxcmd)
        return controlcontainermenu

    def loadTreeView(self):
        # Set up treeview control
        self.contents_tree.DeleteAllItems()
        self.treeroot = self.contents_tree.AddRoot("Root")
        self.extensionsroot = self.contents_tree.AppendItem(self.treeroot, "EXTENSIONS")
        self.contents_tree.SetItemPyData(self.extensionsroot, self._extensiontoplevel)
        self.menusroot = self.contents_tree.AppendItem(self.treeroot, "MENUS")
        self.contents_tree.SetItemPyData(self.menusroot, self._menutoplevel)
        self.toolbarsroot = self.contents_tree.AppendItem(self.treeroot, "TOOLBARS")
        self.contents_tree.SetItemPyData(self.toolbarsroot, self._toolbartoplevel)
        self.contents_tree.SetItemBold(self.extensionsroot, True)
        self.contents_tree.SetItemBold(self.menusroot, True)
        self.contents_tree.SetItemBold(self.toolbarsroot, True)

    def loadProject(self):
        # Repopulate the tree view control et al with settings from the loaded project
        self.loadTreeView()
        # Set up metadata text entry
        self.project_name.SetLabel(self.project.addin.name)
        self.project_version.SetLabel(self.project.addin.version)
        self.project_company.SetLabel(self.project.addin.company)
        self.project_description.SetLabel(self.project.addin.description)
        self.project_author.SetLabel(self.project.addin.author)

    def OnClose(self, event):
        if self.save_button.IsEnabled():
            self.SaveProject(event)
        self.Destroy()

    def SelectFolder(self, event):
        dlg = wx.DirDialog(self, "Choose a directory to use as an AddIn project root:", 
                           style=wx.DD_DEFAULT_STYLE)
        dlg.SetPath(wx.StandardPaths.Get().GetDocumentsDir())
        if dlg.ShowModal() == wx.ID_OK:
            self.path = dlg.GetPath()
            try:
                self.project = addin.PythonAddinProjectDirectory(self.path)
            except Exception as e:
                errdlg = wx.MessageDialog(self, e.message, 'Error initializing addin', wx.OK | wx.ICON_ERROR)
                errdlg.ShowModal()
                errdlg.Destroy()
            else:
                self.folder_button.SetLabel(self.path)
                self.loadProject()
                return
            self.SelectFolder(event)
        elif event is None:
            sys.exit(0)
        else:
            return

    def ProjectNameText(self, event):
        newvalue = self.project_name.GetLabel()
        if not newvalue:
            self.project_name.SetLabel(self.project.addin.name)
        else:
            self.project.addin.name = newvalue
        self.save_button.Enable(True)
        event.Skip()

    def ProjectCompanyText(self, event):
        newvalue = self.project_company.GetLabel()
        self.project.addin.company = newvalue
        self.save_button.Enable(True)
        event.Skip()

    def ProjectDescriptionText(self, event):
        newvalue = self.project_description.GetLabel()
        self.project.addin.description = newvalue
        self.save_button.Enable(True)
        event.Skip()

    def ProjectAuthorText(self, event):
        newvalue = self.project_author.GetLabel()
        self.project.addin.author = newvalue
        self.save_button.Enable(True)
        event.Skip()

    def ProjectVersionText(self, event):
        newvalue = self.project_version.GetLabel()
        if not newvalue:
            self.project_version.SetLabel(self.project.addin.version)
        else:
            self.project.addin.version = newvalue
        self.save_button.Enable(True)
        event.Skip()

    def ComboBox(self, event):
        self.project.addin.app = self.product_combo_box.GetValue()
        self.save_button.Enable(True)
        event.Skip()

    def setupPropsDialog(self):
        sizer = self.item_property_panel.GetSizer()
        sizer.Clear(True)
        if hasattr(self._selected_data, '__doc__') and self._selected_data.__doc__:
            st = wx.StaticText(self.item_property_panel, -1, str(self._selected_data.__doc__))
            st.SetFont(wx.Font(16, wx.DEFAULT, wx.NORMAL, wx.NORMAL, 0, ""))
            sizer.Add(st, 0, wx.ALL, 8)
        pythonliteral = re.compile("^[_A-Za-z][_A-Za-z0-9]*$").match
        def isinteger(val):
            try:
                int(val)
                return True
            except:
                return False
        proplist = [p for p in (('name', 'Name', str, None), 
                                ('caption', 'Caption', str, None), 
                                ('klass', 'Class Name', str, pythonliteral), 
                                ('id', 'ID (Variable Name)', str, pythonliteral),
                                ('description', 'Description', str, None),
                                ('tip', 'Tooltip', str, None),
                                ('message', 'Message', str, None),
                                ('hint_text', 'Hint Text', str, None),
                                ('help_heading', 'Help Heading', str, None),
                                ('help_string', 'Help Content', str, None),
                                ('editable', 'Editable', bool, None),
                                ('separator', 'Has Separator', bool, None),
                                ('show_initially', 'Show Initially', bool, None),
                                ('auto_load', 'Load Automatically', bool, None),
                                ('tearoff', 'Can Tear Off', bool, None),
                                ('menu_style', 'Menu Style', bool, None),
                                ('columns', 'Column Count', str, isinteger),
                                ('image', 'Image for Control', wx.Bitmap, None)) 
                                    if hasattr(self._selected_data, p[0])]
        for prop, caption, datatype, validator in proplist:
            # This is all kind of hairy, sorry
            newsizer = wx.BoxSizer(wx.HORIZONTAL)
            # Text entry
            if datatype in (str, int):
                st = wx.StaticText(self.item_property_panel, -1, caption + ":", style=wx.ALIGN_RIGHT)
                st.SetMinSize((100, 16))
                newsizer.Add(st, 0, wx.ALL|wx.ALIGN_CENTER_VERTICAL, 2)
                text = wx.TextCtrl(self.item_property_panel, -1, str(getattr(self._selected_data, prop, '')) or '')
                class edittext(object):
                    def __init__(self, edit_object, command, app, propname, validator, datatype):
                        self.edit_object = edit_object
                        self.command = command
                        self.app = app
                        self.propname = propname
                        self.validator = validator
                        self.datatype = datatype
                    def __call__(self, event):
                        newvalue = self.command.GetLabel()
                        if self.validator is None or self.validator(newvalue):
                            try:
                                setattr(self.edit_object, self.propname, self.datatype(newvalue))
                            except Exception as e:
                                print e
                            self.app.contents_tree.SetItemText(self.app.contents_tree.GetSelection(), 
                                                               getattr(self.edit_object, 'caption', 
                                                                   getattr(self.edit_object, 'name', 
                                                                       str(self.edit_object))))
                            self.app.save_button.Enable(True)
                        else:
                            self.command.SetLabel(str(getattr(self.edit_object, self.propname, '')))
                        event.Skip()
                self.Bind(wx.EVT_TEXT, edittext(self._selected_data, text, self, prop, validator, datatype), text)
                newsizer.Add(text, 1, wx.RIGHT, 8)
            # Checkbox
            elif datatype is bool:
                class toggle(object):
                    def __init__(self, edit_object, propname, control, app):
                        self.edit_object = edit_object
                        self.propname = propname
                        self.control = control
                        self.app = app
                    def __call__(self, event):
                        setattr(self.edit_object, self.propname, self.control.GetValue())
                        self.app.save_button.Enable(True)
                boolcheck = wx.CheckBox(self.item_property_panel, -1, caption)
                boolcheck.SetValue(getattr(self._selected_data, prop))
                self.Bind(wx.EVT_CHECKBOX, toggle(self._selected_data, prop, boolcheck, self), boolcheck)
                newsizer.Add(boolcheck, 1, wx.LEFT, 100)
            # Image selection
            elif datatype is wx.Bitmap:
                class pickbitmap(object):
                    def __init__(self, edit_object, propname, control, app):
                        self.edit_object = edit_object
                        self.propname = propname
                        self.control = control
                        self.app = app
                    def __call__(self, event):
                        images_path = os.path.join(self.app.path, 'Images')
                        potentialdir = getattr(self.edit_object, self.propname, '')
                        if potentialdir:
                            images_path = os.path.join(self.app.path, 
                                                       os.path.dirname(potentialdir))
                        default_path = (images_path 
                                            if os.path.exists(images_path)
                                            else self.app.path)
                        dlg = wx.FileDialog(
                            self.app, message="Choose an image file for control",
                            defaultDir=default_path, 
                            defaultFile="",
                            wildcard="All Files (*.*)|*.*|"
                                     "GIF images (*.gif)|*.gif|"
                                     "PNG images (*.png)|*.png|"
                                     "BMP images (*.bmp)|*.bmp",
                            style=wx.OPEN)
                        if dlg.ShowModal() == wx.ID_OK:
                            image_file = dlg.GetPath()
                            bitmap = wx.Bitmap(image_file, wx.BITMAP_TYPE_ANY)
                            self.app.save_button.Enable(True)
                            self.control.SetBitmapLabel(bitmap)
                            setattr(self.edit_object, self.propname, image_file)
                            self.app.Fit()
                            self.app.Layout()
                            self.app.Refresh()

                st = wx.StaticText(self.item_property_panel, -1, "Image for control:", style=wx.ALIGN_RIGHT)
                st.SetMinSize((100, 16))
                newsizer.Add(st, 0, wx.ALL|wx.ALIGN_CENTER_VERTICAL, 2)
                if getattr(self._selected_data, prop, ''):
                    bitmap = wx.Bitmap(os.path.join(self.path, getattr(self._selected_data, prop)), wx.BITMAP_TYPE_ANY)
                else:
                    bitmap = wx.EmptyBitmap(16, 16, 32)
                    bitmap.SetMaskColour((0, 0, 0))
                    if bitmap.HasAlpha():
                        print "ALPHA"
                        for xc in xrange(bitmap.GetWidth()):
                            for yc in xrange(bitmap.GetHeight()):
                                bitmap.SetAlpha(xc, yc, 255)
                choosefilebutton = wx.BitmapButton(self.item_property_panel, -1, bitmap)
                self.Bind(wx.EVT_BUTTON, pickbitmap(self._selected_data, prop, choosefilebutton, self), choosefilebutton)
                newsizer.Add(choosefilebutton, 1, wx.ALL|wx.EXPAND, 0)
            # WHO KNOWS!
            else:
                newsizer.Add(wx.StaticText(self.item_property_panel, -1, caption + ": " + str(getattr(self._selected_data, prop))), 0, wx.EXPAND)
            sizer.Add(newsizer, 0, wx.EXPAND|wx.BOTTOM, 2)
        #self.item_property_panel.SetSizerAndFit(sizer, True)
        sizer.Layout()
        self.Refresh()

    def SelChanged(self, event):
        try:
            self._selected_data = self.contents_tree.GetItemPyData(self.contents_tree.GetSelection())
        except:
            self._selected_data = None
        self.setupPropsDialog()

    def DeleteItem(self, event):
        pass

    def ChangeTab(self, event):
        pass

    def SelectProjectImage(self, event):
        dlg = wx.FileDialog(
            self, message="Choose an image file",
            defaultDir=(os.path.join(self.path, 
                                     os.path.dirname(self.project.addin.image))
                            if self.project.addin.image
                                else self.path), 
            defaultFile="",
            wildcard="All Files (*.*)|*.*|"
                     "GIF images (*.gif)|*.gif|"
                     "PNG images (*.png)|*.png|"
                     "BMP images (*.bmp)|*.bmp",
            style=wx.OPEN)
        if dlg.ShowModal() == wx.ID_OK:
            image_file = dlg.GetPath()
            self.project.addin.image = image_file
            bitmap = wx.Bitmap(image_file, wx.BITMAP_TYPE_ANY)
            self.icon_bitmap.SetBitmap(bitmap)
            self.Fit()
            self.Refresh()
        self.save_button.Enable(True)
        event.Skip()

    def SaveProject(self, event):
        try:
            print self.project.addin.xml
            print
            print self.project.addin.python
            self.project.save()
            self.save_button.Enable(False)
        except Exception as e:
            print e

    def TreePopupRClick(self, event):
        id = self.contents_tree.HitTest(event.GetPosition())[0]
        self.contents_tree.SelectItem(id, True) # Set right-clicked item as selection for popups

    def TreePopup(self, event):
        menu = None
        sd = self._selected_data
        if sd is self._extensiontoplevel:
            menu = self.extensionmenu
        elif sd is self._menutoplevel:
            menu = self.menumenu
        elif sd is self._toolbartoplevel:
            menu = self.toolbarmenu
        elif isinstance(sd, addin.ControlContainer):
            menu = self.controlcontainermenu
        elif isinstance(sd, (addin.UIControl, addin.XMLAttrMap)):
            menu = wx.Menu()
        else:
            print sd
        if menu:
            if sd not in (self._extensiontoplevel, self._menutoplevel, self._toolbartoplevel):
                if menu.GetMenuItemCount():
                    menu.AppendSeparator()
                removecmd = menu.Append(-1, "Remove")
                def remove(event):
                    if self.project.addin.remove(sd):
                        self.contents_tree.Delete(self.contents_tree.GetSelection())
                        self.save_button.Enable(True)
                menu.Bind(wx.EVT_MENU, remove, removecmd)
            self.PopupMenu(menu)
            menu.Destroy()

if __name__ == "__main__":
    app = wx.PySimpleApp(0)
    wx.InitAllImageHandlers()
    addin_window = AddinMakerAppWindow(None, -1, "")
    app.SetTopWindow(addin_window)
    addin_window.Show()
    app.MainLoop()
