#! python3
# https://pbpython.com/python-word-template.html

from __future__ import print_function
from mailmerge import MailMerge
from datetime import date

import pandas as pd
import os

location = (
    r'D:\path\path\looooooooongpath\FOLDER')

os.chdir(location)

for i in range(len(data)):
    template = "template.docx" # with MergeFields inserted
    document = MailMerge(template)

    data = pd.read_excel('excel.xlsx') # excel template with headings to populate in row 1

    document.merge(
    MergeField_1 = str(data['COLUMN 1'][i]), # order is not important, only the names
    MergeField_2 = str(data['COLUMN 2'][i]),
    MegreField_3 = str(data['COLUMN 3'][i])
    )

    document.write('DOC_'+str(i+1)+'.docx')


