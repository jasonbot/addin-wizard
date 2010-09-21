import addin
import addin_ui
import os
import sys
import wx

class AddinMakerAppWindow(addin_ui.AddinMakerWindow):
    class _extensiontoplevel(object):
        pass
    class _menutoplevel(object):
        pass
    class _toolbartoplevel(object):
        pass
    def __init__(self, *args, **kws):
        super(AddinMakerAppWindow, self).__init__(*args, **kws)
        self.Bind(wx.EVT_CLOSE, self.OnClose)
        self.contents_tree.Bind(wx.EVT_RIGHT_DOWN, self.TreePopupRClick)
        self.contents_tree.Bind(wx.EVT_CONTEXT_MENU, self.TreePopup)
        self.SelectFolder(None)
    def loadProject(self):
        # Repopulate the tree view control et al with settings from the loaded project
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
        dlg = wx.DirDialog(self, "Choose a directory to use as an AddIn project root", 
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
    def SelChanged(self, event):
        try:
            self._selected_data = self.contents_tree.GetItemPyData(self.contents_tree.GetSelection())
        except:
            self._selected_data = None
    def ChangeTab(self, event):
        pass
    def SelectProjectImage(self, event):
        dlg = wx.FileDialog(
            self, message="Choose an image file",
            defaultDir=(os.path.abspath(os.path.dirname(self.project.addin.image))
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
            self.Layout()
            self.SetSize(self.GetSize())
            self.Refresh()
        self.save_button.Enable(True)
        event.Skip()
    def SaveProject(self, event):
        self.save_button.Enable(False)
    def TreePopupRClick(self, event):
        id = self.contents_tree.HitTest(event.GetPosition())[0]
        self.contents_tree.ToggleItemSelection(id)
    def TreePopup(self, event):
        menu = None
        sd = self._selected_data
        if sd is self._extensiontoplevel:
            print "EXTN MENU"
        elif sd is self._menutoplevel:
            print "MENU MENU"
        elif sd is self._toolbartoplevel:
            print "TOOLBAR MENU"
        else:
            print sd
        #menu = wx.Menu()
        #menu.Append(wx.NewId(), "One")
        #menu.Append(wx.NewId(), "Two")
        if menu:
            self.PopupMenu(menu)
            menu.Destroy()

if __name__ == "__main__":
    app = wx.PySimpleApp(0)
    wx.InitAllImageHandlers()
    addin_window = AddinMakerAppWindow(None, -1, "")
    app.SetTopWindow(addin_window)
    addin_window.Show()
    app.MainLoop()
