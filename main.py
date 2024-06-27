import wx
import os
from openpyxl import load_workbook
from CreateDataModel import createDataModel
from VRP import VRP
from format_solution import format_solution
from createDispatch import createDispatch
from MasterTemplate import generateTemplate

class FileDropTarget(wx.FileDropTarget):
    def __init__(self, window):
        wx.FileDropTarget.__init__(self)
        self.window = window
        self.solution = None

    def OnDropFiles(self, x, y, filenames):

        self.filename = filenames[-1]
        print(self.filename)
        if self.filename.endswith('.xlsx'):
            try:
                message = f"Loaded '{os.path.basename(self.filename)}'"
            except Exception as e:
                message = f"Failed to load '{os.path.basename(self.filename)}': {str(e)}"
        else:
            message = f"The file '{os.path.basename(self.filename)}' is not an Excel file."
        self.window.updateStatus(message)
        return True
    
    def retrieveTemplate(self, event):
        generateTemplate(self.filename)
    
    def generate(self, event):
        data = createDataModel(self.filename)
        self.solution = VRP(data)
        if self.solution[3]:
            formattedSolution = format_solution(self.solution[0], self.solution[1], self.solution[2], self.solution[3])
            createDispatch(formattedSolution, self.filename)
            self.window.updateStatus("Dispatch Successfully Generated")
            
        else:
            self.window.updateStatus("Insufficient Resources")

class MainFrame(wx.Frame):
    def __init__(self, *args, **kw):
        super(MainFrame, self).__init__(*args, **kw)

        mainPanel = wx.Panel(self)
        hbox = wx.BoxSizer(wx.HORIZONTAL)

        menuPanel = GeneratorManager(mainPanel)
        right_panel = wx.Panel(mainPanel)

        right_panel.SetBackgroundColour(wx.Colour(250, 200, 200))

        hbox.Add(menuPanel, 2, wx.EXPAND | wx.ALL, 5)
        hbox.Add(right_panel, 1, wx.EXPAND | wx.ALL, 5)

        mainPanel.SetSizer(hbox)

        self.SetSize((600, 600))
        self.Centre()



    '''def __init__(self):
        super().__init__(None, title='Dispatch Generator', size=(600, 600))
        self.panel = None
        self.sizer = wx.BoxSizer(wx.VERTICAL)
        self.SetSizer(self.sizer)
        self.showGeneratorManager()

    def showGeneratorManager(self):
        if self.panel:
            self.panel.Destroy()
        self.panel = GeneratorManager(self)
        self.sizer.Add(self.panel, 1, wx.EXPAND)
        self.Layout()

    def showTemplateManager(self):
        if self.panel:
            self.panel.Destroy()
        self.panel = TemplateManager(self)
        self.sizer.Add(self.panel, 1, wx.EXPAND)
        self.Layout()


    def on_close(self, event):
        os._exit(0)'''


class TemplateManager(wx.Panel):
    def __init__(self, parent, *args, **kw):
        super(TemplateManager, self).__init__(parent, *args, **kw)
        self.parent = parent
        self.InitUI()

    def InitUI(self):
        
        self.homeButton = wx.Button(self, label="Home")
        self.homeButton.Bind(wx.EVT_BUTTON, self.goHome)

        self.sizer = wx.BoxSizer(wx.HORIZONTAL)
        self.sizer.Add(self.homeButton, 0, wx.ALL, 5)

        self.SetSizer(self.sizer)
        self.Layout()

    def goHome(self, event):
        self.parent.showGeneratorManager()




class GeneratorManager(wx.Panel):
    def __init__(self, parent, *args, **kw):
        super(GeneratorManager, self).__init__(parent, *args, **kw)
        self.parent = parent
        self.InitUI()
    
    def InitUI(self):
        dropTarget = FileDropTarget(self)
        self.SetDropTarget(dropTarget)
        self.label = wx.StaticText(self, label="Drop an Excel file here...", pos=(50, 80))

        self.generateButton = wx.Button(self, label="Generate Dispatch")
        self.generateButton.Bind(wx.EVT_BUTTON, dropTarget.generate)

        self.templateButton = wx.Button(self, label="Retrieve Template")
        self.templateButton.Bind(wx.EVT_BUTTON, dropTarget.retrieveTemplate)

        self.templateManagerButton = wx.Button(self, label="Template Manager")
        self.templateManagerButton.Bind(wx.EVT_BUTTON, self.onOpenTemplateManager)

        self.sizer = wx.BoxSizer(wx.HORIZONTAL)
        self.sizer.Add(self.generateButton, 0, wx.ALL, 5)
        self.sizer.Add(self.templateButton, 0, wx.ALL, 5)
        self.sizer.Add(self.templateManagerButton, 0, wx.ALL, 5)

        self.SetSizer(self.sizer)
        self.Layout()


    def onOpenTemplateManager(self, event):
        self.parent.showTemplateManager()
        

    def updateStatus(self, message):
        self.label.SetLabel(message)
        


class MyApp(wx.App):
    def OnInit(self):
        self.frame = MainFrame(None, title="oh yeah")
        self.frame.Show()
        return True

if __name__ == "__main__":
    app = MyApp()
    app.MainLoop()

