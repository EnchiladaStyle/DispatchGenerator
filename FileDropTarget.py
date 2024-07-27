import wx
import os
from VRP import VRP
from createDispatch import createDispatch

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