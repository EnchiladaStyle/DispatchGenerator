import wx
import os
from CreateDataModel import createDataModel
from VRP import VRP
from format_solution import format_solution
from createDispatch import createDispatch
from MasterTemplate import generateTemplate
from loadTemplate import selectTemplates, loadTemplate
from FileDropTarget import FileDropTarget
from MainMenu import MainMenu

class MainFrame(wx.Frame):
    def __init__(self, *args, **kw):
        super(MainFrame, self).__init__(*args, **kw)
        drop_target = FileDropTarget(self)

        mainPanel = wx.Panel(self)
        hbox = wx.BoxSizer(wx.HORIZONTAL)

        menuPanel = MainMenu(mainPanel, drop_target)
        right_panel = wx.Panel(mainPanel)

        right_panel.SetBackgroundColour(wx.Colour(250, 200, 200))

        hbox.Add(menuPanel, 2, wx.EXPAND | wx.ALL, 5)
        hbox.Add(right_panel, 3, wx.EXPAND | wx.ALL, 5)

        mainPanel.SetSizer(hbox)

        self.SetSize((600, 400))
        self.Centre()

        right_panel.SetDropTarget(drop_target)
        self.right_panel = right_panel

        self.label = wx.StaticText(self.right_panel, label="Drop an Excel file here...", pos=(50, 80))

        self.Bind(wx.EVT_CLOSE, self.onClose)

    def updateStatus(self, message):
        self.label.SetLabel(message)

    def onClose(self, event):
        self.Destroy()
        os._exit(0)
        
class MyApp(wx.App):
    def OnInit(self):
        self.frame = MainFrame(None, title="Easy Dispatch")
        self.frame.Show()
        return True

if __name__ == "__main__":
    app = MyApp()
    app.MainLoop()
    print("ending")
    app.Destroy()

