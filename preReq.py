import urllib.request
import xlrd
import tkinter as tk
import csv
from tkinter import filedialog

# ############################################
# PUC Dependency Graph Viewer
# ############################################
# Guilherme Dantas, 7/23/2019
# ############################################
# How to use:
# 1 - Log in your PUC Online account
# 2 - Goto to "Falta Cursar"
# 3 - Download the excel spreadsheet
# 4 - Download xlrd pip dependecy:
#     $ pip install xlrd
# 5 - Run this script
# 6 - Open the spreadsheet file
# 7 - Wait a little bit (download html...)
# 8 - Done. The graph should be printed out.
# ############################################


def get_html(mat):
    # Extracted from https://stackoverflow.com/a/30890016/5696107
    """ reads from course code are returns html """
    with urllib.request.urlopen(
            "https://www.puc-rio.br/ferramentas/ementas/ementa.aspx?cd="+mat) as fp:
        htmlbytes = fp.read()
        htmltext = htmlbytes.decode("ISO-8859-1")
        return htmltext


def read_col(filepath, col, header):
    # Extracted from https://stackoverflow.com/a/36235630/5696107
    """ reads column from spreadsheet and returns array """
    workbook = xlrd.open_workbook(filepath)
    sheets = workbook.sheet_names()
    sheet = workbook.sheet_by_name(sheets[0])
    x = []
    for rownum in range(header+1, sheet.nrows):
        x.append(sheet.cell(rownum, col).value)
    return x


def open_sheet(title):
    # Extracted from https://stackoverflow.com/a/14119223/5696107
    # and https://pythonspot.com/tk-file-dialogs/
    # and https://stackoverflow.com/a/31732248/5696107
    """ Opens file dialog and returns spreadsheet file """
    root = tk.Tk()
    root.withdraw()
    return filedialog.askopenfilename(
        title=title,
        filetypes=(("Excel spreadsheets", "*.xls"),)
    )


def save_to_csv(path, *args):
    # Extracted from https://realpython.com/python-csv/
    """ Reads several arrays and save them to a csv file """
    with open(path, mode='w') as f:
        writer = csv.writer(
            f, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
        for i in range(len(args[0])):
            row = []
            for arg in args:
                row.append(arg[i])
            writer.writerow(row)


def save_mats(path):
    """ Saves courses ids in mat.csv from spreadsheet """
    mats = read_col(path, 0, 2)
    save_to_csv("mat.csv", mats)


path = open_sheet("Open the courses spreadsheet")
save_mats(path)