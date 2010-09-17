import addin
import addin_ui
import wx

class AddinMakerAppWindow(addin_ui.AddinMakerWindow):
    def __init__(self, *args, **kws):
        super(AddinMakerAppWindow, self).__init__(*args, **kws)
        self.contents_tree.Bind(wx.EVT_CONTEXT_MENU, self.TreePopup)
        self.SelectFolder(None)
    def loadProject(self):
        self.contents_tree.DeleteAllItems()
        self.treeroot = self.contents_tree.AddRoot("Root")
        self.extensionsroot = self.contents_tree.AppendItem(self.treeroot, "EXTENSIONS")
        self.menusroot = self.contents_tree.AppendItem(self.treeroot, "MENUS")
        self.toolbarsroot = self.contents_tree.AppendItem(self.treeroot, "TOOLBARS")
        self.contents_tree.SetItemBold(self.extensionsroot, True)
        self.contents_tree.SetItemBold(self.menusroot, True)
        self.contents_tree.SetItemBold(self.toolbarsroot, True)
    def SelectFolder(self, event):
        dlg = wx.DirDialog(self, "Choose a directory:",
                            style=wx.DD_DEFAULT_STYLE
                                  | wx.DD_DIR_MUST_EXIST
                           )
        if dlg.ShowModal() == wx.ID_OK:
            self.folder_button.SetLabel(dlg.GetPath())
            self.loadProject()
    def TreePopup(self, event):
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
