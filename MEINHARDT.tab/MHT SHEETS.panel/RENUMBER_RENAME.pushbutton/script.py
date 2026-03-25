# -*- coding: utf-8 -*-
__title__ = "NUMBER/NAME"  # Name of the button displayed in Revit
__author__ = "GM"
__version__ = '2.0'
__doc__ = """Version: 2.0
Date    = 2026-03-21
_____________________________________________________________________
Description:

Rename multiple sheets at once with Find/Replace/Suffix/Prefix logic.
You can select sheets in Project Browser or if nothing selected
you will get a menu to select your sheets.
_____________________________________________________________________
How-to:

-> Select sheets in ProjectBrowser (optional)
-> Click the button
-> Set your criterias
-> Rename
_____________________________________________________________________
Author: GM"""

# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╚═╝ IMPORTS
# ==================================================================
from Autodesk.Revit.DB import *
from Autodesk.Revit.Exceptions import ArgumentException

#pyRevit
from pyrevit import forms

# .NET IMPORTS
from clr import AddReference
AddReference("System")

# ╦  ╦╔═╗╦═╗╦╔═╗╔╗ ╦  ╔═╗╔═╗
# ╚╗╔╝╠═╣╠╦╝║╠═╣╠╩╗║  ║╣ ╚═╗
#  ╚╝ ╩ ╩╩╚═╩╩ ╩╚═╝╩═╝╚═╝╚═╝ VARIABLES
# ==================================================================
uidoc   = __revit__.ActiveUIDocument
doc     = __revit__.ActiveUIDocument.Document


def get_selected_sheets(uidoc, title="Select Sheets", label="Select Sheet"):
    """Get selected sheets from Project Browser or prompt user."""
    selection = uidoc.Selection
    selected_ids = selection.GetElementIds()
    
    # Try to get ViewSheet objects from selection
    selected_sheets = []
    for elem_id in selected_ids:
        try:
            elem = doc.GetElement(elem_id)
            if isinstance(elem, ViewSheet):
                selected_sheets.append(elem)
        except:
            pass
    
    if selected_sheets:
        return selected_sheets
    
    # If nothing selected, get all sheets and prompt user
    collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Sheets).WhereElementIsNotElementType()
    all_sheets = [sheet for sheet in collector if isinstance(sheet, ViewSheet)]
    
    if not all_sheets:
        forms.alert("No sheets found in the project.", title=title)
        return []
    
    if len(all_sheets) == 1:
        return all_sheets
    
    selected = forms.SelectFromList.show(all_sheets, 
                                        multiselect=True,
                                        name_attr='Name',
                                        title=title)
    return selected if selected else []


selected_sheets = get_selected_sheets(uidoc=uidoc, title=__title__, label='Select Sheets to Rename')

if not selected_sheets:
    forms.alert("Please select at least one sheet.", title=__title__)
    __revit__.Exit(True)

# ╔═╗╦ ╦╔╗╔╔═╗╔╦╗╦╔═╗╔╗╔╔═╗
# ╠╣ ║ ║║║║║   ║ ║║ ║║║║╚═╗
# ╚  ╚═╝╝╚╝╚═╝ ╩ ╩╚═╝╝╚╝╚═╝ FUNCTIONS
# ==================================================================
def update_project_browser():
    """Function to close and reopen ProjectBrowser so changes to Sheetnumber would become visible."""
    from Autodesk.Revit.UI import DockablePanes, DockablePane
    project_browser_id = DockablePanes.BuiltInDockablePanes.ProjectBrowser
    project_browser = DockablePane(project_browser_id)
    project_browser.Hide()
    project_browser.Show()

# ╔═╗╦  ╔═╗╔═╗╔═╗╔═╗╔═╗
# ║  ║  ╠═╣╚═╗╚═╗║╣ ╚═╗
# ╚═╝╩═╝╩ ╩╚═╝╚═╝╚═╝╚═╝ CLASSES
# ==================================================================

class MyWindow(forms.WPFWindow):
    """GUI for ViewSheet renaming tool."""
    def __init__(self, xaml_file_name):
        self.form = forms.WPFWindow.__init__(self, xaml_file_name)
        self.main_title.Text = __title__


    def rename(self):
        t = Transaction(doc, __title__)
        t.Start()
        self.rename_sheet_name()
        self.rename_sheet_number()
        update_project_browser()
        t.Commit()


    def rename_sheet_name(self):
        """Function to rename SheetName if it is different to current one."""

        for sheet in selected_sheets:
            sheet_name_new = self.sheet_name_prefix + sheet.Name.replace(self.sheet_name_find, self.sheet_name_replace) + self.sheet_name_suffix
            fail_count = 0

            while fail_count < 5:
                fail_count += 1

                try:
                    if sheet.Name != sheet_name_new:
                        sheet.Name = sheet_name_new
                        break
                except ArgumentException:
                    sheet_name_new += "*"
                except:
                    sheet_name_new += "_"


    def rename_sheet_number(self):
        for sheet in selected_sheets:
            sheet_number_new = self.sheet_number_prefix + sheet.SheetNumber.replace(self.sheet_number_find, self.sheet_number_replace) + self.sheet_number_suffix
            fail_count = 0
            while fail_count < 5:
                fail_count += 1
                try:
                    if sheet.SheetNumber != sheet_number_new:
                        sheet.SheetNumber = sheet_number_new
                        break
                except ArgumentException:
                    sheet_number_new += "*"
                except:
                    sheet_number_new += "_"


    ### GUI PROPERTIES
    # SHEETNUMBER PROPERTIES
    @property
    def sheet_number_find(self):
        return self.input_sheet_number_find.Text

    @property
    def sheet_number_replace(self):
        return self.input_sheet_number_replace.Text

    @property
    def sheet_number_prefix(self):
        return self.input_sheet_number_prefix.Text

    @property
    def sheet_number_suffix(self):
        return self.input_sheet_number_suffix.Text

    # SHEETNAME PROPERTIES
    @property
    def sheet_name_find(self):
        return self.input_sheet_name_find.Text

    @property
    def sheet_name_replace(self):
        return self.input_sheet_name_replace.Text

    @property
    def sheet_name_prefix(self):
        return self.input_sheet_name_prefix.Text

    @property
    def sheet_name_suffix(self):
        return self.input_sheet_name_suffix.Text

    # GUI EVENT HANDLERS:
    def button_close(self,sender,e):
        """Stop application by clicking on a <Close> button in the top right corner."""
        self.Close()

    def button_run(self, sender, e):
        """Button action: Rename sheets with given criteria."""
        self.rename()
        forms.alert("Sheets renamed successfully.", title=__title__)

# ╔╦╗╔═╗╦╔╗╔
# ║║║╠═╣║║║║
# ╩ ╩╩ ╩╩╝╚╝ MAIN
# ==================================================================
if __name__ == '__main__':
    MyWindow("Script.xaml").ShowDialog()
