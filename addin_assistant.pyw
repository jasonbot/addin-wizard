import addin
import addin_ui
import os
import sys
import wx

class AddinMakerAppWindow(addin_ui.AddinMakerWindow):
    def __init__(self, *args, **kws):
        super(AddinMakerAppWindow, self).__init__(*args, **kws)
        self.contents_tree.Bind(wx.EVT_CONTEXT_MENU, self.TreePopup)
        self.SelectFolder(None)
    def loadProject(self):
        # Repopulate the tree view control et al with settings from the loaded project
        self.contents_tree.DeleteAllItems()
        self.treeroot = self.contents_tree.AddRoot("Root")
        self.extensionsroot = self.contents_tree.AppendItem(self.treeroot, "EXTENSIONS")
        self.menusroot = self.contents_tree.AppendItem(self.treeroot, "MENUS")
        self.toolbarsroot = self.contents_tree.AppendItem(self.treeroot, "TOOLBARS")
        self.contents_tree.SetItemBold(self.extensionsroot, True)
        self.contents_tree.SetItemBold(self.menusroot, True)
        self.contents_tree.SetItemBold(self.toolbarsroot, True)

        self.project_name.SetLabel(self.project.addin.name)
        self.project_version.SetLabel(self.project.addin.version)
        self.project_company.SetLabel(self.project.addin.company)
        self.project_description.SetLabel(self.project.addin.description)
        self.project_author.SetLabel(self.project.addin.author)
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
        event.Skip()
    def ProjectCompanyText(self, event):
        newvalue = self.project_company.GetLabel()
        self.project.addin.company = newvalue
        event.Skip()
    def ProjectDescriptionText(self, event):
        newvalue = self.project_description.GetLabel()
        self.project.addin.description = newvalue
        event.Skip()
    def ProjectAuthorText(self, event):
        newvalue = self.project_author.GetLabel()
        self.project.addin.author = newvalue
        event.Skip()
    def ProjectVersionText(self, event):
        newvalue = self.project_version.GetLabel()
        if not newvalue:
            self.project_version.SetLabel(self.project.addin.version)
        else:
            self.project.addin.version = newvalue
        event.Skip()
    def ComboBox(self, event):
         self.project.addin.app = self.product_combo_box.GetValue()
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
            self.Refresh()
    def TreePopup(self, event):
        print self.contents_tree.GetSelection()
        menu = wx.Menu()
        menu.Append(wx.NewId(), "One")
        menu.Append(wx.NewId(), "Two")
        self.PopupMenu(menu)
        menu.Destroy()

if __name__ == "__main__":
    app = wx.PySimpleApp(0)
    wx.InitAllImageHandlers()
    addin_window = AddinMakerAppWindow(None, -1, "")
    app.SetTopWindow(addin_window)
    addin_window.Show()
    app.MainLoop()
