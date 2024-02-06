import pandas as pd
import openpyxl as opxl
import shutil
import os
from pptx import Presentation 

wb = opxl.load_workbook("Membership Proposal Details.xlsx")

class company:
    def __init__(self):
        self.fullName = ""
        self.companyName = ""
        self.email = ""
        self.phoneNumber = ""
        self.membershipType = ""
        self.space = ""
        self.description = ""
        self.term = ""
        self.startDate = ""
        self.endDate = ""
        self.rate = ""
        self.additionalFees = ""
        self.discountPromo = ""
        self.misc = ""
        self.link = ""
    

def gatherInfo(obj, sheetname):
    wb = opxl.load_workbook(sheetname)
    wb._active_sheet_index = 0
    ws = wb.active

    obj.fullName = ws["B1"].value
    obj.companyName = ws["B2"].value
    obj.email = ws["B3"].value
    obj.phoneNumber = ws["B4"].value

    wb._active_sheet_index = 1
    ws = wb.active

    obj.membershipType = ws["B1"].value
    obj.space = ws["B3"].value
    obj.description = ws["B4"].value
    obj.term = ws["B5"].value
    obj.startDate = ws["B6"].value
    obj.endDate = ws["B7"].value
    obj.rate = ws["B8"].value
    obj.additionalFees = ws["B9"].value
    obj.discountPromo = ws["B10"].value
    obj.misc = ws["B12"].value
    obj.link = ws["B17"].value

def copy_ppt(source_ppt, destination_ppt):
    shutil.copyfile(source_ppt, destination_ppt)


path = os.getcwd()
obj = company()
sheetname = "Membership Proposal Details.xlsx"

gatherInfo(obj, sheetname)

sourcePPT = "VX Proposal Template Updated .pptx"

destinationPPT = f"{path}/{obj.companyName} Proposal Template.pptx"

copy_ppt(sourcePPT, destinationPPT)
