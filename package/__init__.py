print('Initialising package...\n')

import pandas as pd
import requests
import openpyxl
from pypdf import PdfReader
import os

cont = input("Please ensure you have moved the Excel files (in .xlsx format) to the directory of this module. Leave blank and press enter to continue... (otherwise input anything and press enter to exit)")
dlpath = os.path.join(os.getcwd(),'downloads')
print('\nPDFs would be downloaded to: ' + dlpath)