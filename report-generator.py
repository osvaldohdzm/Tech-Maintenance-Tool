import win32com.client as win32
import win32clipboard as clip
import os
import json
import docx
import os.path
import numpy as np
import pdb

from os.path import abspath
from win32com.client import constants
from docx import Document
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
from docx.shared import RGBColor
from inspect import getmembers
import argparse
import logging

from datetime import datetime

#pdb.set_trace()



#dt_string = datetime.now().strftime("%d-%m-%Y %H-%M-%S")
#logging.basicConfig(filename='logs\\{} script.log'.format(dt_string), encoding='utf-8', level=logging.DEBUG)
#logger = logging.getLogger(__name__)
visible_mode_win32com = True

def count_inrange(list1, l, r):
     
    # x for x in list1 is same as traversal in the list
    # the if condition checks for the number of numbers in the range
    # l to r
    # the return is stored in a list
    # whose length is the answer
    return len(list(x for x in list1 if l <= x <= r))
 

def merge_docx2(files, final_docx):
    output = wordapp.Documents.Add()
    output.Application.CutCopyMode = False
    for fn in files:
        output.Application.Selection.InsertFile(os.path.join(dn,fn) )        
    output.SaveAs(os.path.join(dn,'output.docx')  )
    output.Close(False)

parser = argparse.ArgumentParser(description='Web vulnerabilities report generator.')
parser.add_argument("-j", help='file JSON name', required=True)
args = vars(parser.parse_args())
analysis_filename = args['j']

# Create target directory & all intermediate directories if don't exists
if not os.path.isfile(analysis_filename):
    print('No existe ese archivo.')
    exit()

# Opening JSON file
with open(analysis_filename, encoding='utf-8') as json_file:
    data = json.load(json_file)

dn = os.path.dirname(os.path.realpath(__file__))
template_file_path = os.path.join(dn,data['<<template_name>>'])

name_file = data['<<report_name>>'] + ' - ' + data['<<client_name>>'] +".docx"

wordapp = win32.gencache.EnsureDispatch("Word.Application")
wordapp.Visible = visible_mode_win32com
wordapp.DisplayAlerts = False

doc = wordapp.Documents.Open(template_file_path)
doc.Activate()

wordapp.ActiveDocument.TrackRevisions = False  # Maybe not need this (not really but why not)
wordapp.Selection.GoTo(win32.constants.wdGoToPage, win32.constants.wdGoToAbsolute, "2")

for From in data.keys():
    try:
        wordapp.ActiveWindow.ActivePane.View.SeekView =win32.constants.wdSeekMainDocument
        wordapp.Selection.Find.Execute(From, False, False, False, False, False, True, win32.constants.wdFindContinue, False, data[From], win32.constants.wdReplaceAll) 
        wordapp.ActiveWindow.ActivePane.View.SeekView = win32.constants.wdSeekCurrentPageHeader
        wordapp.Selection.Find.Execute(From, False, False, False, False, False, True, win32.constants.wdFindContinue, False, data[From], win32.constants.wdReplaceAll) 
    except Exception as e:
        print(e)    

software_list = []
for software in data['<<software_list>>']:      
    software_list.append(software) 

wordapp.Selection.HomeKey(Unit=win32.constants.wdStory)
wordapp.Selection.Find.Execute('<<software_list>>') 
wordapp.Selection.Text = '\n'.join(software_list)

problem_list = []
for problem in data['<<problems_list>>']:      
    problem_list.append(problem) 
wordapp.Selection.HomeKey(Unit=win32.constants.wdStory)
wordapp.Selection.Find.Execute('<<problems_list>>') 
wordapp.Selection.Text = '\n'.join(problem_list)


tec_list = []
for tec in data['<<tecnic_details>>']:      
    tec_list.append(tec) 
wordapp.Selection.HomeKey(Unit=win32.constants.wdStory)
wordapp.Selection.Find.Execute('<<tecnic_details>>') 
wordapp.Selection.Text = '\n'.join(tec_list)


recomendations_list = []
for recomendation in data['<<recomendations>>']:      
    recomendations_list.append(recomendation) 
wordapp.Selection.HomeKey(Unit=win32.constants.wdStory)
wordapp.Selection.Find.Execute('<<recomendations>>') 
wordapp.Selection.Text = '\n'.join(recomendations_list)

print(doc.Tables.Count)

index = 3
solutions_list = []
for diagnostic in data['<<diagnostics>>']:
    solutions_list = []
    for solution in diagnostic['<<solutions>>']:  
        solutions_list.append(solution) 
    doc.Tables(5).Cell(index, 1).Range.Text = diagnostic['<<problem>>']
    doc.Tables(5).Cell(index, 2).Range.Text =  '\n'.join(solutions_list)
    index = index + 1 
    doc.Tables(5).Rows.Add()
doc.Tables(5).Cell(index, 1).Select() 
wordapp.Selection.SelectRow() 
wordapp.Selection.Cells.Delete(ShiftCells=win32.constants.wdDeleteCellsEntireRow)

index = 2
product_list = []
subtotals = []
for product in data['<<budget_product>>']:
    doc.Tables(7).Cell(index, 1).Range.Text = product['<<concept>>']
    doc.Tables(7).Cell(index, 2).Range.Text =  product['<<cantidad>>']
    doc.Tables(7).Cell(index, 3).Range.Text =  "${:,.2f}".format(float(product['<<price>>']))
    if float(product['<<discount>>']) == 0:
        doc.Tables(7).Cell(index, 4).Range.Text =  "-"
    else:
        doc.Tables(7).Cell(index, 4).Range.Text =  "{0:.0%}".format(float(product['<<discount>>']))
    mount = float(product['<<price>>']) * (1 - float(product['<<discount>>'])) * float(product['<<cantidad>>']) 
    doc.Tables(7).Cell(index, 5).Range.Text =  "${:,.2f}".format(mount)
    subtotals.append(mount)
    index = index + 1 
    doc.Tables(7).Rows.Add()
doc.Tables(7).Cell(index, 1).Select() 
wordapp.Selection.SelectRow() 
wordapp.Selection.Cells.Delete(ShiftCells=win32.constants.wdDeleteCellsEntireRow)

subtotal_products = sum(map(float,subtotals))
if data["<<budget_include_iva>>"]:
    iva_products = 0
else:
    iva_products = subtotal_products * 0.16
total_products = subtotal_products + iva_products
doc.Tables(8).Cell(1, 5).Range.Text = "${:,.2f}".format(subtotal_products)
doc.Tables(8).Cell(2, 5).Range.Text = "${:,.2f}".format(iva_products )
doc.Tables(8).Cell(3, 5).Range.Text = "${:,.2f}".format(total_products) 


index = 2
services_list = []
services_subtotals = []
for service in data['<<budget_service>>']:
    doc.Tables(9).Cell(index, 1).Range.Text = service['<<concept>>']
    doc.Tables(9).Cell(index, 2).Range.Text =  "${:,.2f}".format(float(service['<<price>>']))
    if float(service['<<discount>>']) == 0:
        doc.Tables(9).Cell(index, 3).Range.Text =  "-"
    else:
        doc.Tables(9).Cell(index, 3).Range.Text =  "{0:.0%}".format(float(service['<<discount>>']))
    mount = float(service['<<price>>']) * (1 - float(service['<<discount>>']))

    doc.Tables(9).Cell(index, 4).Range.Text = "${:,.2f}".format(mount)
    services_subtotals.append(mount)
    index = index + 1 
    doc.Tables(9).Rows.Add()
doc.Tables(9).Cell(index, 1).Select() 
wordapp.Selection.SelectRow() 
wordapp.Selection.Cells.Delete(ShiftCells=win32.constants.wdDeleteCellsEntireRow)

subtotal_services = sum(map(float,services_subtotals))
iva_services = subtotal_services * 0.16
total_services = subtotal_services + iva_services
doc.Tables(10).Cell(1, 4).Range.Text =  "${:,.2f}".format(subtotal_services)
doc.Tables(10).Cell(2, 4).Range.Text ="${:,.2f}".format(iva_services )
doc.Tables(10).Cell(3, 4).Range.Text = "${:,.2f}".format(total_services)


print("Saving in {}".format(os.path.join(dn,name_file)))

wordapp.Selection.GoTo(win32.constants.wdGoToPage, win32.constants.wdGoToAbsolute, "2")
wordapp.ActiveDocument.SaveAs(os.path.join(dn,name_file))
wordapp.Application.Quit()
