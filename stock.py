import openpyxl
from openpyxl.styles import Font, Color
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from xlwt import Workbook
from xlsxwriter import Workbook    
from nsetools import Nse
from bsedata.bse import BSE
import smtplib
import pandas as pd
from nsepy import get_history
from datetime import date
import matplotlib.pyplot as plt
import nse
import bse
import sentmail

def stock_data():
    sentmail.sent_mail()
    
stock_data()


    
  



