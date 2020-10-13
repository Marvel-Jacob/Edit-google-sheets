import gspread
from oauth2client.service_account import ServiceAccountCredentials
from pprint import pprint
import pandas as pd
import os
import glob
import time
import pygsheets
import numpy as np
from datetime import datetime



file_urls = {}
ae_wbs = pd.read_csv('https://docs.google.com/spreadsheets/d/e/2PACX-1vQZdinq3iY9rciqmFHSyseCMoHAdu/pub?output=csv')
for ae, link in zip(list(ae_wbs['name']),list(ae_wbs['link'])):
    file_urls.update({ae:link})
    file_urls = dict(sorted(file_urls.items()))

date = "July 2020"

mo_template = pd.read_excel('template.xlsx', sheet_name="MO", header=None, dtype=str).head(33).fillna('')
mo_template = mo_template.values.transpose().tolist()
bs_template = pd.read_excel('template.xlsx', sheet_name="BS", header=None, dtype=str).head(39).fillna('')
bs_template = bs_template.values.transpose().tolist()

scope_pygsheets = ['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/spreadsheets']
gc = pygsheets.authorize(client_secret = "client_secret_pygsheet.json", scopes = scope_pygsheets)

for name, url in list(file_urls.items())[53:]:
    print("-----------",name,"------------")
    print(url)
    workbook = gc.open_by_url(url)
    wb_info = workbook.to_json()
    nsheets = len(wb_info['sheets'])

    register = {}
    for i in range(nsheets):
        isheet = wb_info['sheets'][i]['properties']['index']
        sheet_name = wb_info['sheets'][i]['properties']['title']
        register.update({isheet:sheet_name})
    for i in register:
        print(register[i])
        sheet_title = register[i].lower()
        if "month" in sheet_title:
            sheet = workbook[i]
            data = pd.DataFrame(sheet.get_as_df())
            si = list(data.columns).index(date)
            ei = si+6
            sheet.update_dimensions_visibility(si,ei, dimension='columns', hidden=False)
            sheet.update_dimensions_visibility(si-7,si, dimension='columns', hidden=True)
        elif "sample" not in sheet_title and "sheet" not in sheet_title and "copy" not in sheet_title:
            sheet = workbook[i]
            data = pd.DataFrame(sheet.get_as_df(start='a7'))
            try:
                si = list(data.columns).index(date)
            except:
                continue
            ei = si + 7
            sheet.update_dimensions_visibility(si,ei, dimension='columns', hidden=False)
            sheet.update_dimensions_visibility(si-8,si, dimension='columns', hidden=True)
    