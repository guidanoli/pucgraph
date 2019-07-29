import urllib.request
import xlrd
import tkinter as tk
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


def open_file():
    # Extracted from https://stackoverflow.com/a/14119223/5696107
    # and https://pythonspot.com/tk-file-dialogs/
    # and https://stackoverflow.com/a/31732248/5696107
    """ Opens file dialog and returns spreadsheet file """
    root = tk.Tk()
    root.withdraw()
    return filedialog.askopenfilename(
        title="Selecione o arquivo de Falta Cursar",
        filetypes=(("Excel spreadsheets", "*.xls"),)
    )


def get_prereq_dict(filepath):
    """ reads from spreadsheet and returns dictionary """
    mats = read_col(filepath, 2, 1)
    prereq = dict()
    for m in mats:
        prereq[m] = list()
        html = get_html(m)
        for other_m in mats:
            if other_m != m:
                index = html.find(other_m)
                if index != -1:
                    prereq[m].append(other_m)
    return prereq


def print_graph(dependecy_dict):
    """ Reads dependecy dictionary and prins graph accordingly """
    for mat in dependecy_dict.keys():
        posreq = list()
        for other_mat, dependecies in dependecy_dict.items():
            if mat in dependecies:
                posreq.append(other_mat)
        print(" -> ".join([mat, " -> ".join(posreq)]))


path = open_file()
if path == "":
    print("No file selected.")
else:
    prereq = get_prereq_dict(path)
    print_graph(prereq)
