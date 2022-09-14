import merge_and_release
import openpyxl
import pandas as pd
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font
from openpyxl.styles import colors
from openpyxl import Workbook
from openpyxl.worksheet import pagebreak
import openpyxl.worksheet.header_footer
import datetime
merge_and_release.merge_and_release(['PO_44635688.xlsx'],'test.xlsx')
