import wx
import os
from openpyxl import load_workbook
from CreateDataModel import createDataModel
from VRP import VRP
from format_solution import format_solution
from createDispatch import createDispatch
from MasterTemplate import generateTemplate
from loadTemplate import selectTemplates, loadTemplate
from saveTemplate import saveTemplate, deleteTemplateFromDatabase

class MainMenu(wx.Panel):
    def __init__(self, parent, dropTarget, *args, **kw):
        super(MainMenu, self).__init__(parent, *args, **kw)
        self.parent = parent
        self.dropTarget = dropTarget
        self.InitUI()
        
    
    def InitUI(self):

        self.currentTemplate = ("Master Template", None)
        self.currentTemplateLabel = wx.StaticText(self, label=f"   {self.currentTemplate[0]}")
        self.templateNamesAndIds = selectTemplates()

        self.generateButton = wx.Button(self, label="Generate Dispatch")
        self.generateButton.Bind(wx.EVT_BUTTON, self.generate)

        self.templateButton = wx.Button(self, label="Load Template")
        self.templateButton.Bind(wx.EVT_BUTTON, self.retrieveTemplate)

        self.newTemplateButton = wx.Button(self, label="Save New Template")
        self.newTemplateButton.Bind(wx.EVT_BUTTON, self.saveNewTemplate)

        self.deleteTemplateButton = wx.Button(self, label="Delete Template")
        self.deleteTemplateButton.Bind(wx.EVT_BUTTON, self.deleteTemplate)

        verticalHousing = wx.BoxSizer(wx.VERTICAL)
        topMenu = wx.BoxSizer(wx.VERTICAL)

        topMenu.Add(self.generateButton, 0, wx.ALL, 5)
        topMenu.Add(self.templateButton, 0, wx.ALL, 5)
        topMenu.Add(self.newTemplateButton, 0, wx.ALL, 5)
        topMenu.Add(self.deleteTemplateButton, 0, wx.ALL, 5)
        verticalHousing.Add(topMenu, 0, wx.EXPAND | wx.ALL, 5)

        verticalHousing.AddSpacer(20)
        verticalHousing.Add(wx.StaticText(self, label="   Current Template:"))
        verticalHousing.Add(self.currentTemplateLabel)

        scrollPanel = wx.ScrolledWindow(self, style=wx.VSCROLL)
        scrollPanel.SetScrollRate(5, 5)
        bottomMenu = wx.BoxSizer(wx.VERTICAL)
        scrollPanel.SetSizer(bottomMenu)

        btn = wx.Button(scrollPanel, label="Master Template")
        btn.Bind(wx.EVT_BUTTON, self.onSelectTemplate)
        btn.template = ("Master Template", None)
        bottomMenu.Add(btn, 0, wx.EXPAND | wx.ALL, 5)
        for template in self.templateNamesAndIds:
            btn = wx.Button(scrollPanel, label=template[0])
            btn.Bind(wx.EVT_BUTTON, self.onSelectTemplate)
            btn.template = template
            bottomMenu.Add(btn, 0, wx.EXPAND | wx.ALL, 5)
        verticalHousing.Add(scrollPanel, 1, wx.EXPAND | wx.ALL, 5)

        self.SetSizer(verticalHousing)
        self.Layout()


    def saveNewTemplate(self, event):

        nameBox = wx.Dialog(self, title="Enter Name", size=(250, 150))
        nameBox.CentreOnParent()
        askNamePanel = wx.Panel(nameBox)

        vbox = wx.BoxSizer(wx.VERTICAL)
        self.text_ctrl = wx.TextCtrl(askNamePanel)
        vbox.Add(self.text_ctrl, flag=wx.EXPAND | wx.ALL, border=10)

        submitButton = wx.Button(askNamePanel, label="Submit")
        vbox.Add(submitButton, flag=wx.ALIGN_CENTER | wx.ALL, border=10)

        askNamePanel.SetSizer(vbox)
        submitButton.Bind(wx.EVT_BUTTON, lambda evt: self.onSubmitName(evt, nameBox))

        nameBox.ShowModal()

    def onSubmitName(self, event, dialog):
        name = self.text_ctrl.GetValue()
        saveTemplate(self.dropTarget.filename, name)
        dialog.EndModal(wx.ID_OK)
        self.DestroyChildren()
        self.InitUI()
        

    def retrieveTemplate(self, event):
        if self.currentTemplate[0] == "Master Template":
            generateTemplate(self.dropTarget.filename)
        else:
            print(f"data: {self.currentTemplate[1]}")
            loadTemplate(self.dropTarget.filename, int(self.currentTemplate[1]))

    def generate(self, event):
        data = createDataModel(self.dropTarget.filename)
        self.solution = VRP(data)
        if self.solution[3]:
            formattedSolution = format_solution(self.solution[0], self.solution[1], self.solution[2], self.solution[3])
            createDispatch(formattedSolution, self.dropTarget.filename)
                
        else:
            self.window.updateStatus("Insufficient Resources")

    def onSelectTemplate(self, event):
        button = event.GetEventObject()
        self.currentTemplateLabel.SetLabel(f"   {button.template[0]}")
        self.currentTemplate = button.template

    def deleteTemplate(self, event):
        print("trying to delete template")
        self.confirmationPanel = wx.Panel(self)
        confirmationText = wx.StaticText(self.confirmationPanel, label="Are you sure you want to delete this template?")

        yesButton = wx.Button(self.confirmationPanel, label="Yes")
        yesButton.Bind(wx.EVT_BUTTON, self.onConfirmDelete)
        noButton = wx.Button(self.confirmationPanel, label="No")
        noButton.Bind(wx.EVT_BUTTON, self.onCancelDelete)

        vbox = wx.BoxSizer(wx.VERTICAL)
        hbox = wx.BoxSizer(wx.HORIZONTAL)

        vbox.Add(confirmationText, flag=wx.ALL | wx.CENTER, border=10)
        hbox.Add(yesButton, flag=wx.ALL | wx.CENTER, border=10)
        hbox.Add(noButton, flag=wx.ALL, border=5)
        vbox.Add(hbox, flag=wx.ALIGN_CENTER)

        self.confirmationPanel.SetSizer(vbox)
        self.GetSizer().Add(self.confirmationPanel, flag=wx.EXPAND | wx.ALL, border=10)
        self.Layout()
        

    def onConfirmDelete(self, event):
        templateId = self.currentTemplate[1]
        
        deleteTemplateFromDatabase(templateId)
        self.DestroyChildren()
        self.InitUI()
        self.Layout()

    def onCancelDelete(self, event):
        self.confirmationPanel.Destroy()
        self.Layout()