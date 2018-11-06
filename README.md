# OT_NS
process ONTS trackers for ON_NS summary

# -*- coding: utf-8 -*-
"""
Created on Fri Nov  2 16:19:01 2018

@author: kdn0mqp
"""
from gooey import Gooey, GooeyParser
import pandas as pd
import os
import xlwings as xw


@Gooey(program_name="Process Trackers into OT_NS", program_description="@author:kdn0mqp,Alan", default_size=(1000, 700))
def parse_args():
    parser = GooeyParser()
    parser.add_argument("fileFolderPath", metavar="File Folder Path",
                        help=" you can drag folder to below \n or you can choose folder by clicking> Browser",
                        widget="DirChooser")
    parser.add_argument("fileofOTNS", metavar="OT_NS Summary File",
                        help="please choose your OT_NS Summary file\n", widget="FileChooser")
    args = parser.parse_args()
    return args


args = parse_args()

fileFolderPath = args.fileFolderPath
trackerList = os.listdir(fileFolderPath)
otnsFile = args.fileofOTNS
namelistDF = pd.DataFrame()
databaseDF = pd.DataFrame()
for ph in trackerList:
    path = os.path.join(fileFolderPath, ph)
    print(path)

    detailName = pd.read_excel(path, sheet_name="Namelist")
    detailName = detailName.iloc[1:, 3:10]
    detailName = detailName.loc[detailName["Unnamed: 4"].notnull(), :]
    detailName = detailName.iloc[1:, :]
    namelistDF = namelistDF.append(detailName)

    detailDatabase = pd.read_excel(path, sheet_name="Database")
    detailDatabase = detailDatabase.iloc[3:, 11:23]
    detailDatabase = detailDatabase.loc[detailDatabase["Unnamed: 11"].notnull(), :]
    detailDatabase = detailDatabase.iloc[1:, :]
    databaseDF = databaseDF.append(detailDatabase)

    # print(databaseDF["Unnamed: 11"].dtypes)

writer = pd.ExcelWriter(r"D:\For OT_NS.xlsx")
namelistDF.to_excel(writer, sheet_name="Sheet1")
databaseDF.to_excel(writer, sheet_name="Sheet2")
writer.save()
detail1 = pd.read_excel("D:\\For OT_NS.xlsx", sheet_name="Sheet1")
detail2 = pd.read_excel("D:\\For OT_NS.xlsx", sheet_name="Sheet2")
# detail2=detail2.iloc[1:,1:]
# print(detail2)


app = xw.App(visible=True, add_book=False)
app.display_alerts = False
app.screen_updating = True

wb2 = app.books.open(otnsFile)
wb2.sheets["Namelist"].range("d6:j207").clear_contents()
wb2.sheets["Database"].range("l8:w16702").clear_contents()
wb2.sheets["Namelist"].range("d6").value = detail1.values
wb2.sheets["Database"].range("l8").value = detail2.values
# only by .values, can you paste without header and index
