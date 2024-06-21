import wx
import os
from openpyxl import load_workbook
from CreateDataModel import createDataModel
from VRP import VRP
from format_solution import format_solution
from createDispatch import createDispatch
from generateTemplate import generateTemplate

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

class MyFrame(wx.Frame):
    def __init__(self):
        super().__init__(None, title='Dispatch Generator', size=(400, 200))
        panel = wx.Panel(self)
        self.label = wx.StaticText(panel, label="Drop an Excel file here...", pos=(50, 80))
        
        # Set up the file drop target
        dropTarget = FileDropTarget(self)
        self.SetDropTarget(dropTarget)

        self.generateButton = wx.Button(panel, label="Generate Dispatch")
        self.generateButton.Bind(wx.EVT_BUTTON, dropTarget.generate)

        self.templateButton = wx.Button(panel, label="Retrieve Template")
        self.templateButton.Bind(wx.EVT_BUTTON, dropTarget.retrieveTemplate)

        self.sizer = wx.BoxSizer(wx.HORIZONTAL)
        self.sizer.Add(self.generateButton, 0, wx.ALL, 5)
        self.sizer.Add(self.templateButton, 0, wx.ALL, 5)


        panel.SetSizerAndFit(self.sizer)
        self.Bind(wx.EVT_CLOSE, self.on_close)


    def updateStatus(self, message):
        self.label.SetLabel(message)

    def on_close(self, event):
        os._exit(0)



class MyApp(wx.App):
    def OnInit(self):
        frame = MyFrame()
        frame.Show()
        return True

if __name__ == "__main__":
    app = MyApp()
    app.MainLoop()
