import pandas as pd
import openpyxl as opxl

wb = opxl.load_workbook("Membership Proposal Details.xlsx")




class company:
    def __init__(self, fullName, companyName, email, phoneNumber, membershipType, space, description, term, startDate, endDate, rate, additionalFees, discountPromo, misc, link):
        self.fullName = fullName
        self.companyName = companyName
        self.email = email
        self.phoneNumber = phoneNumber
        self.membershipType = membershipType
        self.space = space
        self.description = description
        self.term = term
        self.startDate = startDate
        self.endDate = endDate
        self.rate = rate
        self.additionalFees = additionalFees
        self.discountPromo = discountPromo
        self.misc = misc
        self.link = link
    

def gatherInfo():
    pass

