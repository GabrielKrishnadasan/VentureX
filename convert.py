#!/usr/bin/env python3

import pandas as pd
import openpyxl as opxl
import shutil
import os
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

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
    """
    Takes a comapny object and a specifit excel spreadsheet, and gathers the data from the excel and places it into the company object
    """
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
    """
    Might not need this
    """
    shutil.copyfile(source_ppt, destination_ppt)

def edit_ppt_text(pptFile, obj, path):
    """
    Takes the template pptx, company object, and current path/working directory and updates the template pptx with
    the data from the excel sheet and makes a copy of the template to store in the same directory
    """
    editDict = {}
    editDict["space1"] = obj.space
    editDict["description1"] = obj.description
    editDict["term1"] = obj.term
    editDict["price1"] = obj.rate


    presentation = Presentation(pptFile)

    for item in presentation.slides[0].shapes:
        if hasattr(item, "text"):
            if item.text == "Company name":
                item.text = obj.companyName + " Proposal"      
                for paragraph in item.text_frame.paragraphs:
                    for run in paragraph.runs:
                        run.font.name = "Bebas Neue"
                        run.font.size = Pt(72)
                        run.font.color.rgb = RGBColor(255,255,255)

    for item in presentation.slides[5].shapes:
        if hasattr(item, "text"):
            if item.text == "Quote for XYZ Company":
                item.text = "Quote for " + obj.companyName + " Company"
                for paragraph in item.text_frame.paragraphs:
                    for run in paragraph.runs:
                        run.font.name = "Bebas Neue"
                        run.font.size = Pt(66)
                        run.font.color.rgb = RGBColor(0,0,0)
            if item.text in editDict:
                item.text = editDict[item.text]
                for paragraph in item.text_frame.paragraphs:
                    paragraph.alignment = PP_ALIGN.CENTER
                    for run in paragraph.runs:
                        run.font.name = "Open Sans"
                        run.font.size = Pt(22)
                        run.font.color.rgb = RGBColor(0,0,0)

    presentation.save(f"{path}/{obj.companyName} VX Proposal.pptx")    


#Gets the directory location of the script
script_dir = os.path.dirname(os.path.realpath(__file__))
#Changes current directory to that of the script
os.chdir(script_dir)
#Saves the path to the script in path
path = os.getcwd()
#Gets the template pptx
sourcePPT = f"{path}/VX Proposal Template.pptx"
#Gets the data excel sheet
sheetname = f"{path}/Membership Proposal Details.xlsx"

#Creates an instance of company class and populates its variables with gather info
obj = company()
gatherInfo(obj, sheetname)

#Calls the edit_ppt_text method
edit_ppt_text(sourcePPT, obj, path)
