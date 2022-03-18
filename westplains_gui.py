import re
from tkinter.filedialog import askdirectory, askopenfilename
from tkinter import Menubutton, Tk, StringVar, Text
from tkinter import PhotoImage
from tkinter.font import Font
from tkinter.ttk import Label
from tkinter import Button
from tkinter.ttk import Frame, Style
from tkinter.ttk import OptionMenu
from tkinter import Label as label
from tkcalendar import DateEntry
from tkinter import messagebox
# from typing import Text
from pandas import json_normalize
from pandas.core import frame 
import requests, json
from datetime import date, datetime
import numpy as np
import glob, time
import threading
from tkinter.messagebox import showerror
import pandas as pd
from requests.models import to_key_val_list
from tqdm import trange
import win32com.client
import os
import xlwings as xw
from tabula import read_pdf
# import PyPDF2
from collections import defaultdict
import xlwings.constants as win32c
import threading, sys




# path = r'C:\Users\imam.khan\OneDrive - BioUrja Trading LLC\Documents\Revelio'
path = r'J:\WEST PLAINS\REPORT\Westplains_gui'
today = datetime.strftime(date.today(), format = "%d%m%Y")



root = Tk()
root.title('Westplains App')
root.geometry('648x696')
photo = PhotoImage(file = path + '\\'+'biourjaLogo.png')
root.iconphoto(False, photo)
root["bg"]= "white"


frame_title = Frame(root)
frame_options = Frame(root)
frame_folder = Frame(root)
frame_submit = Frame(root)
frame_msg = Frame(root)
s = Style(frame_options)
s.configure("TMenubutton", background="#f5fcfc",width=19, font=("Book Antiqua", 12))
s.configure("TMenu", width=19)
s.configure("TFrame", background="white")


class MyDateEntry(DateEntry):
    def __init__(self, master=None, **kw):
        DateEntry.__init__(self, master=master, date_pattern='mm.dd.yyyy',**kw)
        # add black border around drop-down calendar
        self._top_cal.configure(bg='black', bd=1)
        # add label displaying today's date below
        label(self._top_cal, bg='gray90', anchor='w',
                 text='Today: %s' % date.today().strftime('%x')).pack(fill='both', expand=1)



def moc_get_df_from_input_excel(input_xl, mtm_file, open_ap_file, open_ar_file,unsettled_pay_file, unsettled_recev_file):
    """This function returns the dataframe that will be used the MOC allocment process"""
    try:
        req_dict = {}
        key_list = ['Open A/R','Inventory','Unsettled A/R','Unsettled A/P','Adjustments if required',
                    'Deferred Payments','Accounts Payable']

        req_dict = req_dict.fromkeys(key_list)       
        
       
        """This is the code for Inventory MTM Excel Report"""
        try:
            inner_keys = ['Alliance/Hay Springs','Gering','Omaha','Johnstown','KC','BROWNSVILL']
            inner_dict = {}.fromkeys(inner_keys)
            wb_mtm = xw.Book(mtm_file,update_links=False)
            ws_mtm = wb_mtm.sheets['MTM Excel Summary']
            last_row = ws_mtm.range(f'A'+ str(ws_mtm.cells.last_cell.row)).end('up').row
            first_row  = ws_mtm.range(f"A{last_row}").end('up').last_cell.row
            req_index = first_row + 1
            df_mtm = pd.read_excel(mtm_file,sheet_name='MTM Excel Summary', usecols="A,B", skiprows=req_index)   
            new_dict = dict(zip(df_mtm.iloc[:,0],df_mtm.iloc[:,1]))
            inner_dict['Alliance/Hay Springs'] = new_dict['HS']
            inner_dict['Gering'] = new_dict.get('GER')
            inner_dict['Omaha'] = new_dict.get('OM')
            inner_dict['Johnstown'] = new_dict.get('JT')
            inner_dict['KC'] = new_dict.get('KANSAS CTY')
            inner_dict['BROWNSVILL'] = new_dict.get('BR')
            req_dict['Inventory'] = inner_dict
        except Exception as e:
            print(e)
            print("The format of input file is wrong for MTM inventory or the file does not exist. Please enter the correct format")
        finally:
            try:
                wb_mtm.app.quit()
            except Exception as e:
                print(e)
        
        """"This is the code for Open AP files"""
        try:
            inner_keys = ['Alliance/Hay Springs','Gering','Omaha','Johnstown','KC','BROWNSVILL']
            inner_dict = {}.fromkeys(inner_keys)
            # df_ap = pd.read_excel(open_ap_file,sheet_name='For allocation entry',usecols="A,B", skiprows=2)
            df_ap = pd.read_excel(open_ap_file,sheet_name = 0, usecols="A,B", skiprows=2)

            new_dict = dict(zip(df_ap.iloc[:,0],df_ap.iloc[:,1]))
            inner_dict['Alliance/Hay Springs'] = new_dict['HAYSPRG']
            inner_dict['Gering'] = new_dict.get('GERING')
            inner_dict['Omaha'] = new_dict.get('TERMINAL')
            inner_dict['Johnstown'] = new_dict.get('OMA COMM') + new_dict.get('JTELEV')
            inner_dict['KC'] = new_dict.get('KANSAS CTY')
            inner_dict['BROWNSVILL'] = new_dict.get('BROWNSVILL')
            req_dict['Accounts Payable'] = inner_dict
        except Exception as e:
            print(e)
            print("The format of input file is wrong for Open AP or the file does not exist. Please enter the correct format")
        
        """"This is the code for Open AR files"""
        try:
            inner_keys = ['Alliance/Hay Springs','Gering','Omaha','Johnstown','KC','BROWNSVILL']
            inner_dict = {}.fromkeys(inner_keys)
            # df_ar = pd.read_excel(open_ar_file, sheet_name='For allocation entry',usecols="A,B", skiprows=2)
            df_ar = pd.read_excel(open_ar_file, sheet_name = 0, usecols="A,B", skiprows=2)
            new_dict = dict(zip(df_ar.iloc[:,0],df_ar.iloc[:,1]))
            inner_dict['Alliance/Hay Springs'] = new_dict['HAYSPRG']
            inner_dict['Gering'] = new_dict.get('GERING')
            inner_dict['Omaha'] = new_dict.get('TERMINAL')
            inner_dict['Johnstown'] = new_dict.get('OMA COMM') + new_dict.get('JTELEV')
            inner_dict['KC'] = new_dict.get('KANSAS CTY')
            inner_dict['BROWNSVILL'] = new_dict.get('BROWNSVILL')
            req_dict['Open A/R'] = inner_dict
        except Exception as e:
            print(e)
            print("The format of input file is wrong for Open AR or the file does not exist. Please enter the correct format")
        
        """This is the code for Unsettled Payables files"""
        try:
            inner_keys = ['Alliance/Hay Springs','Gering','Omaha','Johnstown','KC','BROWNSVILL']
            inner_dict = {}.fromkeys(inner_keys)
            # df_pay = pd.read_excel(unsettled_pay_file, sheet_name = 'For allocation entry', usecols="A,B", skiprows=2)
            df_pay = pd.read_excel(unsettled_pay_file, sheet_name = 0, usecols="A,B", skiprows=2)
            new_dict = dict(zip(df_pay.iloc[:,0],df_pay.iloc[:,1]))
            inner_dict['Alliance/Hay Springs'] = new_dict['HAY SPRINGS - WEST PLAINS, LLC']
            inner_dict['Gering'] = new_dict.get('GERING - WEST PLAINS, LLC')
            inner_dict['Omaha'] = new_dict.get('OMAHA TERMINAL ELEVATOR - WEST PLAINS, LLC')
            inner_dict['Johnstown'] = new_dict.get('OMAHA COMM - WEST PLAINS, LLC') + new_dict.get('JOHNSTOWN - WEST PLAINS, LLC')
            inner_dict['KC'] = new_dict.get('KANSAS CTY')
            inner_dict['BROWNSVILL'] = new_dict.get('BROWNSVILLE - WEST PLAINS, LLC')
            req_dict['Unsettled A/P'] = inner_dict
        except Exception as e:
            print(e)
            print("The format of input file is wrong for Unsettled A/P or the file does not exist. Please enter the correct format")
            
        """This is the code for Unsettled Receivables"""
        try:
            inner_keys = ['Alliance/Hay Springs','Gering','Omaha','Johnstown','KC','BROWNSVILL']
            inner_dict = {}.fromkeys(inner_keys)
            # df_recev = pd.read_excel(unsettled_recev_file, sheet_name = 'For allocation entry', usecols="A,B", skiprows=2)
            df_recev = pd.read_excel(unsettled_recev_file, sheet_name = 0, usecols="A,B", skiprows=2)
            new_dict = dict(zip(df_recev.iloc[:,0],df_recev.iloc[:,1]))
            inner_dict['Alliance/Hay Springs'] = new_dict['HAY SPRINGS - WEST PLAINS, LLC']
            inner_dict['Gering'] = new_dict.get('GERING - WEST PLAINS, LLC')
            inner_dict['Omaha'] = new_dict.get('OMAHA TERMINAL ELEVATOR - WEST PLAINS, LLC')
            inner_dict['Johnstown'] = new_dict.get('OMAHA COMM - WEST PLAINS, LLC') + new_dict.get('JOHNSTOWN - WEST PLAINS, LLC')
            inner_dict['KC'] = new_dict.get('KANSAS CTY')
            inner_dict['BROWNSVILL'] = new_dict.get('BROWNSVILLE - WEST PLAINS, LLC')
            req_dict['Unsettled A/R'] = inner_dict
        except Exception as e:
            print(e)
            print("The format of input file is wrong for Unsettled A/R or the file does not exist. Please enter the correct format")
            
        
        main_df = pd.DataFrame(req_dict)
        print("Main dataframe created")
        return main_df
    except Exception as e:
        print(e)
    finally:
        pass


def update_moc_excel(main_df,template_dir,output_dir, input_date):
    """This fucntion genereates the out put file for MOC Allocment in the output files folder"""
    try:
        for file in os.listdir(template_dir):
            if 'West Plains Interest Allocation' in file:
                wb_alloc = xw.Book(template_dir + '\\' + file, update_links=False)
                ws_alloc = wb_alloc.sheets['LOC Interest Allocation']

                ws_alloc.range('A3').value = datetime.strptime(input_date,"%m.%d.%Y").date()
                ws_alloc.range('E9:E15').options(transpose=True).value = main_df.values[0]
                ws_alloc.range('F9:F15').options(transpose=True).value = main_df.values[1]
                ws_alloc.range('G9:G15').options(transpose=True).value = main_df.values[2]
                ws_alloc.range('I9:I15').options(transpose=True).value = main_df.values[3]
                ws_alloc.range('J9:J15').options(transpose=True).value = main_df.values[4]
                ws_alloc.range('P9:P15').options(transpose=True).value = main_df.values[5]

                ws_alloc.range('E9:p15').api.NumberFormat = '_("$"* #,##0_);_("$"* (#,##0);_("$"* "-"??_);_(@_)'

                # ws_alloc.range('E17:p17').formula = '=+E9+E10+E11-E12-E13-E14-E15'
                # ws_alloc.range('E19:p19').formula = '=E17/$Q$17'
                # ws_alloc.range('E20:p20').formula = '=E19*$E$62'
                ws_alloc_totals = ws_alloc.range('E17:p17').value
                ws_alloc_totals_lst = ['E17','F17','G17','H17','I17','J17','K17','L17','M17','N17','O17','P17']
                ws_total_dict = dict(zip(ws_alloc_totals_lst, ws_alloc_totals))
                neg_dict = {key:val for key,val in ws_total_dict.items() if val <0}

                if len(neg_dict) > 0:
                    for key,val in neg_dict.items():
                        if key == 'E17':
                            ws_alloc.range('E29:E35').options(transpose=True).value = main_df.values[0]
                        elif key == 'F17':
                            ws_alloc.range('F29:F35').options(transpose=True).value = main_df.values[1]
                        elif key == 'G17':
                            ws_alloc.range('G29:G35').options(transpose=True).value = main_df.values[2]
                        elif key == 'I17':
                            ws_alloc.range('I29:I35').options(transpose=True).value = main_df.values[3]
                        elif key == 'J17':
                            ws_alloc.range('J29:J35').options(transpose=True).value = main_df.values[4]
                        elif key == 'P17':
                            ws_alloc.range('P29:P35').options(transpose=True).value = main_df.values[5]
                else:        
                    ws_alloc.range('E29:E35').options(transpose=True).value = main_df.values[0]
                    ws_alloc.range('F29:F35').options(transpose=True).value = main_df.values[1]
                    ws_alloc.range('G29:G35').options(transpose=True).value = main_df.values[2]
                    ws_alloc.range('I29:I35').options(transpose=True).value = main_df.values[3]
                    ws_alloc.range('J29:J35').options(transpose=True).value = main_df.values[4]
                    ws_alloc.range('P29:P35').options(transpose=True).value = main_df.values[5]

                # ws_alloc.range('E37:p37').formula = '=+E29+E30+E31-E32-E33-E34-E35'
                # ws_alloc.range('E39:p39').formula = '=E37/$Q$37'
                # ws_alloc.range('E40:p40').formula = '=E39*$E$62'

                ws_alloc.range('E29:p35').api.NumberFormat = '_("$"* #,##0_);_("$"* (#,##0);_("$"* "-"??_);_(@_)'
                wb_alloc.save(output_dir + '\\' + file.replace(file.split('_')[1],input_date) + '.xls')                
                print(f"MOC Allocment file generated for {input_date}")
    except Exception as e:
        print("Template file was not found or some other issue occured")
        print(e)
    finally:
        try:
            wb_alloc.app.quit()
        except Exception as e:
            print(e)



def getColumnName(n):
    try:
        result = ''
        while n > 0:
            index = (n - 1) % 26
            result += chr(index + ord('A'))
            n = (n - 1) // 26
    
        return result[::-1]
    except Exception as e:
        raise e


def num_to_col_letters(num):
    try:
        letters = ''
        while num:
            mod = (num - 1) % 26
            letters += chr(mod + 65)
            num = (num - 1) // 26
        return ''.join(reversed(letters))
    except Exception as e:
        raise e

def mtm_pdf_data_extractor(input_date, f, hrw_pdf_loc, yc_pdf_loc):
    try:
        # reader = PyPDF2.PdfFileReader(open(f, mode='rb' ))
        # n = reader.getNumPages() 
        inp_month_year = datetime.strptime(input_date,"%m.%d.%Y").replace(day=1)
        # data_list = []
        for loc in [hrw_pdf_loc, yc_pdf_loc]:
            df = read_pdf(loc, pages = 1, guess = False, stream = True ,
                                    pandas_options={'header':0}, area = ["700,70,1000,1200"], columns=['150','480','550','650', '700','800','900'])
            df = pd.concat(df, ignore_index=True)
            df = df[["MONTH","SETTLE"]]
            form_dict = {"'6":"75", "'4":"50", "'2":"25", "'0":"0"}
            for month in range(len(df)):
                if "JLY" in df["MONTH"][month]:
                    df["MONTH"][month] = df["MONTH"][month].replace("JLY","JUL")
                if inp_month_year == datetime.strptime(df["MONTH"][month], "%b %y"):
                    settle_price = df.loc[:,'SETTLE'][month+1]
                    for key in form_dict:
                        if key in settle_price:
                            if 'HRW' in loc.upper():
                                hrw_fut = int(settle_price.replace(key,form_dict[key]))/10000  
                            elif 'YC' in loc.upper():
                                yc_fut =  int(settle_price.replace(key,form_dict[key]))/10000
                            break
                    break
                elif inp_month_year < datetime.strptime(df["MONTH"][month], "%b %y"):
                    settle_price = df.loc[:,'SETTLE'][month]
                    for key in form_dict:
                        if key in settle_price:
                            if 'HRW' in loc.upper():
                                hrw_fut = int(settle_price.replace(key,form_dict[key]))/10000  
                            elif 'YC' in loc.upper():
                                yc_fut =  int(settle_price.replace(key,form_dict[key]))/10000
                            break
                    break
                
                

        date_df = read_pdf(f, pages = 1, guess = False, stream = True ,
                        pandas_options={'header':None}, area = ["20,40,40,800"])
        print(date_df)
        # pdf_date = date_df[0][0][0].split()[-1]

        com_loc  = read_pdf(f, pages = 'all', guess = False, stream = True ,
                        pandas_options={'header':None}, area = ["30,15,50,120"])
        com_loc = pd.concat(com_loc, ignore_index=True)

        com_loc = list(com_loc[0].str.split('Commodity: ',expand=True)[1])
        # loc_dict = dict(zip(com_loc, [[]]*len(com_loc)))
        loc_dict = defaultdict(list)
        for page in range(1,len(com_loc)+1):
            df = read_pdf(f, pages = page, guess = False, stream = True ,
                            pandas_options={'header':0}, area = ["75,10,580,850"], columns=["50,85, 180,225, 260, 280,300,360,400,430,480,525,570,620,665,720"])
            df = pd.concat(df, ignore_index=True)
            ########logger.info("Filtering only required columns")
            df = df.iloc[:,[0,1,2,3,-2,-1]]
            # df = df[df['Offsite Name Cont. No.'].str.contains("Company Owned Risk:"),df['Offsite Name Cont. No.'].str.contains("Unpriced Sales:")]
            df = df[(df['Offsite Name Cont. No.'].str.contains("Company Owned Risk:")) | (df['Offsite Name Cont. No.'].str.contains("priced Sales:"))]
            # for i in df.loc[:,"Offsite Name Cont. No."]:

            df["Quantity.5"].fillna(0, inplace=True)
            df["Value.5"].fillna(0, inplace=True)

            df["Quantity.5"] = df["Quantity.5"].astype(str).str.replace("(","-").str.replace(",","").str.replace(")","").astype(float).astype(int)
            df["Value.5"] = df["Value.5"].astype(str).str.replace("(","-").str.replace(",","").str.replace(")","").astype(float).astype(int)

            for i in range(len(df)):
                print(df.iloc[i,2]) #2 for "Offsite Name Cont. No."
                if "priced Sales" in df.iloc[i,2]:
                    print("Unprised Value found")
                    if df.iloc[-2,2] == 'Unpriced Sales:' and df.iloc[-2,-2]==0: #pd.isna(df.iloc[-2,-1]):
                        pass
                    else:
                        df.iloc[i+1,-2] = df.iloc[i+1,-2] - df.iloc[i,-2]
                        df.iloc[i+1,-1] = df.iloc[i+1,-1] - df.iloc[i,-1]

            # n_df[n_df.iloc[:,2].str.contains("Company Owned Risk:")] #Another way
            
            
            
            loc_dict[com_loc[page-1]].append(df)
            

            # print(df)

            ########logger.info("keeping online required columns")
        repl = {"(":"-",")":"",",":""}
        for key, value in loc_dict.items():
            if len(value)>1:
                print(key)
                key_value = []
                key_value.append(pd.concat(value, ignore_index=True))
                loc_dict[key] = key_value
                # print(len(value))
                # print()
        
        
        return loc_dict, hrw_fut, yc_fut
    except Exception as e:
        raise e

def mtm_excel(input_date,input_xl,loc_dict,loc_sheet, output_location, hrw_fut, yc_fut):
    try:
        monthYear = datetime.strftime(datetime.strptime(input_xl.split("_")[-1].split(".xlsx")[0],"%m.%d.%Y"), "%d-%b")
        
        retry = 0
        while retry<10:
            try:
                wb = xw.Book(input_xl, update_links=True)

                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry==10:
                    raise e
        retry = 0
        while retry<10:
            try:
                m_sht = wb.sheets[0]

                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry==10:
                    raise e
        

        last_row=m_sht.range(f'A' + str(m_sht.cells.last_cell.row)).end('up').row
        main_loc = m_sht.range(f"A1:A{last_row}").value
        hrw_value=0
        yc_value = 0
        m_sht.range(f"A3").value = datetime.strptime(input_date,"%m.%d.%Y")
        for i in range(len(main_loc)):
            # if main_loc[i] == "Eligible Inventory, held in warehouse or in-transit":
            #     main_loc[i+1] == datetime.strptime(input_date,"%m.%d.%Y")
                
            if main_loc[i]=="HRW" and hrw_value==0:
                hrw = f"B{i+1}"
                hrw_value+=1
            elif main_loc[i]=="HRW" and hrw_value==1:
                hrw_2 = i+3
                hrw_value+=1
            elif main_loc[i]=="YC" and yc_value == 0:
                yc = f"B{i+1}"
                yc_value+=1

            elif main_loc[i]=="Commodity":
                other_loc = f"A{i+3}"
            elif main_loc[i] == "FW":
                other_loc_2 = f"A{i+1}"
            elif main_loc[i] == "Sunflowers":
                sunflwr = f"{i+1}"



        # hrw_locations = m_sht.range("B7").expand('down').value
        hrw_locations = m_sht.range(hrw).expand('down').value
        ########logger.info("Updating lists")
        df = pd.read_excel(loc_sheet, sheet_name='HRW')
        columns = df.set_index(['Column'])["Name"].to_dict()

        loc_abbr = df.set_index(['Column'])["Name"].to_dict()
        # locations[locations.index('Hay Springs')] = 'HAYSPRG'
        # locations[locations.index('Johnstown')] = 'JTELEV'
        # locations[locations.index('Lisco')] = 'LISCO - W'
        # locations[locations.index('Merriman')] = 'MERRIMA'
        # locations[locations.index('Mirage Flats')] = 'MIRAGE F'
        # locations[locations.index('Omaha Terminal')] = 'TERMINAL'
        # locations[locations.index('North Greeley')] = 'NGREEL'
        # locations[locations.index('North Greeley')] = 'NGREEL'
        # locations[locations.index('North Greeley')] = 'NGREEL'
        loc_dict["HRW"][0].set_index('Location', inplace=True) #DF re_idct[loc_abbr[location]]
        i = int(hrw.replace("B", ""))
        start=int(hrw.replace("B", ""))
        for location in hrw_locations:
            try:
                if location == "Lisco":
                    # m_sht.range(f"F127").value = loc_dict["HRW"][0].loc[loc_abbr[location]]["Price"]
                    m_sht.range(f"F{hrw_2+1}").value = loc_dict["HRW"][0].loc[loc_abbr[location]]["Value.5"]/loc_dict["HRW"][0].loc[loc_abbr[location]]["Quantity.5"] # loc_dict["HRW"][0].loc[loc_abbr[location]]["Price"]
                    # m_sht.range(f"C127").value = loc_dict["HRW"][0].loc[loc_abbr[location]]["Quantity.5"]
                    m_sht.range(f"C{hrw_2+1}").value = loc_dict["HRW"][0].loc[loc_abbr[location]]["Quantity.5"]
                    m_sht.range(f"F{i}").value = 0
                    m_sht.range(f"C{i}").value = 0
                
                elif location == "Mirage Flats":
                    # m_sht.range(f"F128").value = loc_dict["HRW"][0].loc[loc_abbr[location]]["Price"]
                    m_sht.range(f"F{hrw_2+2}").value = loc_dict["HRW"][0].loc[loc_abbr[location]]["Value.5"]/loc_dict["HRW"][0].loc[loc_abbr[location]]["Quantity.5"] #loc_dict["HRW"][0].loc[loc_abbr[location]]["Price"]
                    # m_sht.range(f"C128").value = loc_dict["HRW"][0].loc[loc_abbr[location]]["Quantity.5"]
                    m_sht.range(f"C{hrw_2+2}").value = loc_dict["HRW"][0].loc[loc_abbr[location]]["Quantity.5"]
                    m_sht.range(f"F{i}").value = 0
                    m_sht.range(f"C{i}").value = 0
                
                else:
                    m_sht.range(f"F{i}").value = loc_dict["HRW"][0].loc[loc_abbr[location]]["Value.5"]/loc_dict["HRW"][0].loc[loc_abbr[location]]["Quantity.5"] #loc_dict["HRW"][0].loc[loc_abbr[location]]["Price"]
                    m_sht.range(f"C{i}").value = loc_dict["HRW"][0].loc[loc_abbr[location]]["Quantity.5"]
                    m_sht.range(f"I{i}").value = hrw_fut
                i+=1
            except:
                m_sht.range(f"F{i}").value = 0
                m_sht.range(f"C{i}").value = 0
                i+=1
                pass
        end = i-1
        hrw_basis_loc = m_sht.range(f"B{start}:B{end}").value
        hrw_basis = m_sht.range(f"J{start}:J{end}").value
        hrw_basis = [0.0000 if d is None else d for d in hrw_basis]

        hrw_basis_dict = dict(zip(hrw_basis_loc, hrw_basis))

        
        retry = 0
        while retry<10:
            try:
                hrw_sht = wb.sheets["HRW MTM Basis"]

                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry==10:
                    raise e
        
        last_col_num = hrw_sht.range("A3").expand("right").last_cell.column
        last_col = num_to_col_letters(last_col_num+1)
        hrw_sht.range(f"{last_col}3").value = monthYear
        hrw_sht.range(f"{last_col}3").color = "#FFCC99"
        hrw_sht.range(f"{last_col}4").value = "basis"
        hrw_sht.range(f"{last_col}4").api.Font.Underline = True
        hrw_sht.range(f"{last_col}4").color = "#FFCC99"
        hrw_sht.range(f"{last_col}4").api.HorizontalAlignment = win32c.HAlign.xlHAlignRight

        hrw_basis_sht_loc = hrw_sht.range(f"A5").expand("down").value
        i=5
        for location in hrw_basis_sht_loc:
            try:
                hrw_sht.range(f"{last_col}{i}").value = hrw_basis_dict[location]
                i+=1
            except:
                hrw_sht.range(f"{last_col}{i}").value = 0
                hrw_sht.range(f"{last_col}{i}").value = 0
                i+=1
                pass

        # hrw_sht.range(f"{last_col}5").options(transpose=True).value = hrw_basis


        hrw_sht.range(f"{last_col}5").expand("down").api.NumberFormat = "0.0000_);[Red](0.0000)"


        ########logger.info("now updating yc data")
        loc_dict["YC"][0].set_index('Location', inplace=True)
        yc_locations = m_sht.range(yc).expand('down').value
        i=int(yc.replace("B", ""))
        start = int(yc.replace("B", ""))
        for location in yc_locations:
            try:
                m_sht.range(f"F{i}").value = loc_dict["YC"][0].loc[loc_abbr[location]]["Value.5"]/loc_dict["YC"][0].loc[loc_abbr[location]]["Quantity.5"] #loc_dict["YC"][0].loc[loc_abbr[location]]["Price"]
                m_sht.range(f"C{i}").value = loc_dict["YC"][0].loc[loc_abbr[location]]["Quantity.5"]
                m_sht.range(f"I{i}").value = yc_fut
                i+=1
            except:
                m_sht.range(f"F{i}").value = 0
                m_sht.range(f"C{i}").value = 0
                i+=1
                pass
        end = i-1

        yc_basis_loc = m_sht.range(f"B{start}:B{end}").value
        yc_basis = m_sht.range(f"J{start}:J{end}").value
        yc_basis = [0.0000 if d is None else d for d in yc_basis]

        yc_basis_dict = dict(zip(yc_basis_loc, yc_basis))
        retry = 0
        while retry<10:
            try:
                yc_sht = wb.sheets["YC MTM Basis"]

                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry==10:
                    raise e
        
        last_col_num = yc_sht.range("A3").expand("right").last_cell.column
        last_col = num_to_col_letters(last_col_num+1)
        yc_sht.range(f"{last_col}3").value = monthYear
        yc_sht.range(f"{last_col}3").color = "#FFCC99"
        yc_sht.range(f"{last_col}4").value = "basis"
        yc_sht.range(f"{last_col}4").api.Font.Underline = True
        yc_sht.range(f"{last_col}4").color = "#FFCC99"
        yc_sht.range(f"{last_col}4").api.HorizontalAlignment = win32c.HAlign.xlHAlignRight


        yc_basis_sht_loc = yc_sht.range(f"A5").expand("down").value
        i=5
        for location in yc_basis_sht_loc:
            try:
                yc_sht.range(f"{last_col}{i}").value = yc_basis_dict[location]
                i+=1
            except:
                yc_sht.range(f"{last_col}{i}").value = 0
                yc_sht.range(f"{last_col}{i}").value = 0
                i+=1
                pass

        # yc_sht.range(f"{last_col}5").options(transpose=True).value = yc_basis
        yc_sht.range(f"{last_col}5").expand("down").api.NumberFormat = "0.0000_);[Red](0.0000)"

        other_loc_lst = m_sht.range(other_loc).expand('down').value
        i=int(other_loc.replace("A", ""))
        for location in other_loc_lst:
            try:
                # m_sht.range(f"F{i}").value = loc_dict[loc_abbr[location]][0].iloc[-1,-2] #Price
                m_sht.range(f"F{i}").value = loc_dict[loc_abbr[location]][0].iloc[-1,-1]/loc_dict[loc_abbr[location]][0].iloc[-1,-2]
                # m_sht.range(f"C{i}").value = loc_dict[loc_abbr[location]][0].iloc[-1,-1]
                m_sht.range(f"C{i}").value = loc_dict[loc_abbr[location]][0].iloc[-1,-2] #Quantity
                i+=1
            except:
                m_sht.range(f"F{i}").value = 0
                m_sht.range(f"C{i}").value = 0
                i+=1
        ########logger.info("Updating sunflower prices")
        # m_sht.range(f"F113").value = loc_dict['SUNFLWR'][0].iloc[-1,-2]
        m_sht.range(f"F{sunflwr}").value = loc_dict['SUNFLWR'][0].iloc[-1,-1]/loc_dict['SUNFLWR'][0].iloc[-1,-2]  #Price
        # m_sht.range(f"C113").value = loc_dict['SUNFLWR'][0].iloc[-1,-1]
        m_sht.range(f"C{sunflwr}").value = loc_dict['SUNFLWR'][0].iloc[-1,-2]  #Quantity

        other_loc_2_lst = m_sht.range(other_loc_2).expand('down').value
        i=int(other_loc_2.replace("A", ""))
        for location in other_loc_2_lst:
            try:
                m_sht.range(f"F{i}").value = loc_dict[loc_abbr[location]][0].iloc[-1,-1]/loc_dict[loc_abbr[location]][0].iloc[-1,-2] #Price
                # m_sht.range(f"F{i}").value = loc_dict[loc_abbr[location]][0].iloc[-1,-2]
                m_sht.range(f"C{i}").value = loc_dict[loc_abbr[location]][0].iloc[-1,-2] #Quantity
                # m_sht.range(f"C{i}").value = loc_dict[loc_abbr[location]][0].iloc[-1,-1]
                i+=1
            except:
                m_sht.range(f"F{i}").value = 0
                m_sht.range(f"C{i}").value = 0
                i+=1
        
        print()
        wb.save(output_location)
    except Exception as e:
        raise e
    finally:
        wb.app.quit()
    

def bbr(input_date, output_date):
    pass


def cpr(input_date, output_date):
    cpr_file_date = input_date.replace('.','-')
    output_cpr  = r'J:\WEST PLAINS\REPORT\CPR reports\Output files'+'\\Counter Party Risk Consolidated '+cpr_file_date+'.xlsx'
    output_cpr_copy  = r'J:\WEST PLAINS\REPORT\CPR reports\Output files'+'\\Counter Party Risk Consolidated '+cpr_file_date+' Report Copy.xlsx'
    
    input_cpr = r'J:\WEST PLAINS\REPORT\CPR reports\Raw Files\Counter Party Risk Consolidated '+cpr_file_date+'.xlsx'

    input_cpr_copy = r'J:\WEST PLAINS\REPORT\CPR reports\Raw Files\Counter Party Risk Consolidated '+cpr_file_date+' Report Copy.xlsx'

    UnsettledRec_book = r'J:\WEST PLAINS\REPORT\Unsettled Receivables\Output files\Unsettled Receivables _'+input_date+'.xlsx'

    UnsettledPay_book = r'J:\WEST PLAINS\REPORT\Unsettled Payables\Output files\Unsettled Payables _'+input_date+'.xlsx'

    Open_AR_book = r'J:\WEST PLAINS\REPORT\Open AR\Output files\Open AR _'+input_date+' - Production.xlsx'

    Open_AP_book = r'J:\WEST PLAINS\REPORT\Open AP\Output files\Open AP _'+input_date+'.xlsx'

    CTM_book = r'J:\WEST PLAINS\REPORT\CTM Combined report\Output files\CTM Combined _'+input_date+'.xlsx'

    if not os.path.exists(input_cpr):
            return(f"{input_cpr} Excel file not present for date {cpr_file_date}")

    if not os.path.exists(input_cpr_copy):
            return(f"{input_cpr_copy}Excel file not present for date {cpr_file_date}")

    if not os.path.exists(UnsettledRec_book):
            return(f"{UnsettledRec_book}Excel file not present for date {input_date}")

    if not os.path.exists(UnsettledPay_book):
            return(f"{UnsettledPay_book}Excel file not present for date {input_date}")

    if not os.path.exists(Open_AR_book):
            return(f"{Open_AR_book}Excel file not present for date {input_date}")
    
    if not os.path.exists(Open_AP_book):
            return(f"{Open_AP_book}Excel file not present for date {input_date}")

    if not os.path.exists(CTM_book):
            return(f"{CTM_book}Excel file not present for date {input_date}")


    try:
        
        # input_file = f'{book_name} {sheet_date}.xlsx'

        retry = 0
        while retry<10:
            try:
                wb = xw.Book(input_cpr, update_links=False)
                break
            except:
                time.sleep(5)
                retry+=1
        
        retry = 0
        while retry<10:
            try:
                
                ws1 = wb.sheets[f'Data {input_date}']
                break
            except:
                time.sleep(5)
                retry+=1

        num_row = ws1.range('A1').end('down').row
        num_col = ws1.range('A1').end('right').column

        # ws1.range(f'2:{num_row}').delete()
        
        # Opening Unsettled Receivables Workbook
        ####logger.info('Opening Unsettled Receivables Workbook')
        retry = 0
        while retry<10:
            try:
                UnsettledRec_wb = xw.Book(UnsettledRec_book,update_links=False)
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry==10:
                    raise e
        retry=0
        while retry<10:
            try:
                UnsettledRec_ws = UnsettledRec_wb.sheets['Excl Macq & IC']
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry==10:
                    raise e

        column_lst =  UnsettledRec_ws.range('A1').expand('right').value
        name_col = column_lst.index('Customer/Vendor Name')
        Net_col = column_lst.index('Net')
        UnsettledRec_CustomerName = UnsettledRec_ws.range(f'{getColumnName(name_col+1)}2').expand('down').value
        UnsettledRec_Net = UnsettledRec_ws.range(f'{getColumnName(Net_col+1)}2').expand('down').value
        ws1.range('A2').options(transpose = True).value = UnsettledRec_CustomerName
        ws1.range('C2').options(transpose = True).value = UnsettledRec_Net


        # Opening Unsettled Payables Workbook
        ####logger.info('Opening Unsettled Payables Workbook')
        retry = 0
        while retry<10:
            try:
                UnsettledPay_wb = xw.Book(UnsettledPay_book, update_links=False)
                
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry==10:
                    raise e
        retry=0
        while retry<10:
            try:
                
                UnsettledPay_ws = UnsettledPay_wb.sheets['Excl Macq & IC']
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry==10:
                    raise e

        column_lst =  UnsettledPay_ws.range('A1').expand('right').value
        name_col = column_lst.index('Customer/Vendor Name')
        Net_col = column_lst.index('Net')
        UnsettledPay_CustomerName = UnsettledPay_ws.range(f'{getColumnName(name_col+1)}2').expand('down').value
        UnsettledPay_Net = UnsettledPay_ws.range(f'{getColumnName(Net_col+1)}2').expand('down').value
        num_row = ws1.range('A1').end('down').row
        ws1.range(f'A{num_row+1}').options(transpose = True).value = UnsettledPay_CustomerName
        ws1.range(f'D{num_row+1}').options(transpose = True).value = UnsettledPay_Net

        
        retry=0
        while retry<10:
            try:
                OpenAR_wb = xw.Book(Open_AR_book,update_links=False)
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry==10:
                    raise e
        retry=0
        while retry<10:
            try:
                OpenAR_ws = OpenAR_wb.sheets['Eligible']
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry==10:
                    raise e
        OpenAR_ws = OpenAR_wb.sheets['Eligible']
        column_lst =  OpenAR_ws.range('A1').expand('right').value
        name_col = column_lst.index('Customer Name')
        Balance_col = column_lst.index('Balance')
        OpenAR_CustomerName = OpenAR_ws.range(f'{getColumnName(name_col+1)}2').expand('down').value
        OpenAR_Balance =  OpenAR_ws.range(f'{getColumnName(Balance_col+1)}2').expand('down').value
        num_row = ws1.range('A1').end('down').row
        ws1.range(f'A{num_row+1}').options(transpose = True).value = OpenAR_CustomerName
        ws1.range(f'E{num_row+1}').options(transpose = True).value =OpenAR_Balance

        # Opening Open AP Workbook
        #logger.info('Opening Open AP Workbook')
        retry = 0
        while retry<10:
            try:
                OpenAP_wb = xw.Book(Open_AP_book,update_links=False)
                
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry==10:
                    raise e
        retry = 0
        while retry<10:
            try:
                
                OpenAP_ws = OpenAP_wb.sheets['Excl Macq & IC']
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry==10:
                    raise e
        column_lst =  OpenAP_ws.range('A1').expand('right').value
        name_col = column_lst.index('Vendor')
        Balance_col = column_lst.index('Invoice Balance')
        OpenAP_Vendor = OpenAP_ws.range(f'{getColumnName(name_col+1)}2').expand('down').value
        OpenAP_Balance =  OpenAP_ws.range(f'{getColumnName(Balance_col+1)}2').expand('down').value
        num_row = ws1.range('A1').end('down').row
        ws1.range(f'A{num_row+1}').options(transpose = True).value = OpenAP_Vendor
        ws1.range(f'F{num_row+1}').options(transpose = True).value = OpenAP_Balance

        retry = 0
        while retry<10:
            try:
                CTM_wb = xw.Book(CTM_book,update_links=False)
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry==10:
                    raise e
        retry = 0
        while retry<10:
            try:
                CTM_ws = CTM_wb.sheets['Excl Macq & IC']
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry==10:
                    raise e

        column_lst = CTM_ws.range('A1').expand('right').value
        name_col = column_lst.index('Customer')
        total_sum_col = column_lst.index('Gain/LossTotal')
        last_row=CTM_ws.range(f'{getColumnName(name_col+1)}' + str(CTM_ws.cells.last_cell.row)).end('up').row
        CTM_Customer = CTM_ws.range(f'{getColumnName(name_col+1)}2').expand('down').value
        CTM_LGTotal = CTM_ws.range(f'{getColumnName(total_sum_col+1)}2').expand('down').value

        num_row = ws1.range('A1').end('down').row
        ws1.range(f'A{num_row+1}').options(transpose = True).value = CTM_Customer
        ws1.range(f'G{num_row+1}').options(transpose = True).value = CTM_LGTotal
       
         
        ws1.autofit()
        num_row = ws1.range('A1').end('down').row
        num_col = ws1.range('A1').end('right').column
        
        retry=0
        while retry<15:
            try:
                pivot_sht = wb.sheets["Pivot"]
                time.sleep(5)
                # pivot_sht.select()
                pivot_sht.activate()
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry==15:
                    raise e
        # pivot_sht.api.Select()
        
        pivotCount = wb.api.ActiveSheet.PivotTables().Count
         # 'Data 02.21.2022'!$A$1:$G$4731
        for j in range(1, pivotCount+1):
            wb.api.ActiveSheet.PivotTables(j).PivotCache().SourceData = f"'Data {input_date}'!R1C1:R{num_row}C{num_col}" #Updateing data source
            wb.api.ActiveSheet.PivotTables(j).PivotCache().Refresh()  
        # find pivot table last row of column A and B
        A_lastRow = pivot_sht.range('A5').end('down').row
        B_lastRow = pivot_sht.range('B5').end('down').row
        if A_lastRow != B_lastRow:
            pivot_sht.range(f'A{A_lastRow-1}').copy()
            pivot_sht.range(f'A{A_lastRow+1}:A{B_lastRow-1}').paste()
            # pivot_sht.api.Range('A5:G350').Copy()
        
        # wb.save()

        # BB report
        #logger.info('Opening CPR BB report Workbook')
        pivot_sht.range(f'A5:H{B_lastRow-1}').copy()
        retry = 0
        while retry<10:
            try:
                BB_wb = xw.Book(input_cpr_copy,update_links=True)
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry==10:
                    raise e
        
        #logger.info('Opening Master sheet')
        while True:
            try:
                BB_ws = BB_wb.sheets['Master']
                break
            except:
                time.sleep(10)
        BB_ws.range('A9').paste()
        
        BB_ws.range('C5').value = cpr_file_date
        BB_CustomerROW = BB_ws.range('B9').end('down').row
        last_row=BB_ws.range(f'D' + str(CTM_ws.cells.last_cell.row)).end('up').row
        BB_ws.range(f"D7:I{last_row}").api.NumberFormat = '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'
        for i in range(9,BB_CustomerROW+1):
            # BB_ws.range(f'C{i}').formula = f'=+A{i}'
            BB_ws.range(f'C{i}').value =  BB_ws.range(f'A{i}').value
            if  BB_ws.range(f'A{i}').value == None:
                BB_ws.range(f'C{i}').value = '#N/A'
            # BB_ws.range(f'D{i}').api.NumberFormat= '_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'
            BB_ws.range(f'I{i}').formula = f'=H{i}+D{i}+F{i}-E{i}-G{i}'

        # BB Master +-25K report
        ####logger.info('Opening BB Master +-25K report sheet')
        BB_lastRow = BB_ws.range('A9').end('down').row
        BB_ws.range(f'B9:I{BB_lastRow}').copy()

        while True:
            try:
                BB_Master25ws = BB_wb.sheets['Master +- 25K']
                break
            except:
                time.sleep(10)
        BB_Master25ws.range('A7').paste()
        # BB_Master25ws.range('A9').end('down').row

        Total_lst = (BB_Master25ws.range('H7').expand('down').value)
        BB_Master25ws.range('H7').options(transpose = True).value = Total_lst

        # BB_wb.save()

        # delete sum of total value column ("J") Positive and negative value less then 25K
        ####logger.info('Delete rows with value between -25K to 25K')
        BB_Master25_Row = BB_Master25ws.range('H9').end('down').row
        # for i in range(7,BB_Master25_Row+1):
        i = 7
        while i<= BB_Master25_Row:
            if (type(BB_Master25ws.range(f'H{i}').value) == int) or (type(BB_Master25ws.range(f'H{i}').value) == float):
                if  (-25000 < BB_Master25ws.range(f'H{i}').value) and (BB_Master25ws.range(f'H{i}').value <25000):
                    # BB_Master25ws.range(f'{i}:{i}').api.Delete(DeleteShiftDirection.xlShiftUp)
                    BB_Master25ws.range(f'{i}:{i}').api.Delete()
                    i-=1
                else:
                    i+=1
                   
            else:
                i+=1
        ####logger.info('Refreshing all tab')  
        BB_wb.api.RefreshAll()
        print()
        wb.save(output_cpr)
        BB_wb.save(output_cpr_copy)

        return f"CPR Reports for {input_date} is generated"
    except Exception as e:
        # ####logger.exception(str(e))
        raise e
    finally:
        try:
            wb.app.quit()
        except:
            pass

def ctm(input_date, output_date):
    try:
        input_sheet = r'J:\WEST PLAINS\REPORT\CTM Combined report\Raw Files\CTM Combined _'+input_date+'.xlsx' 
        output_location = r'J:\WEST PLAINS\REPORT\CTM Combined report\Output files\CTM Combined _'+input_date+".xlsx"
        # input_cpr = r'J:\WEST PLAINS\REPORT\CPR reports\Raw Files\Counter Party Risk Consolidated '+cpr_file_date+'.xlsx'    
        if not os.path.exists(input_sheet):
            return(f"{input_sheet} Excel file not present for date {input_date}")

        prev_month = datetime.strftime(datetime.strptime(input_date, "%m.%d.%Y"), "%B")
        ###logger.info("Opening operating workbook instance of excel")
        retry=0
        while retry < 10:
            try:
                wb=xw.Book(input_sheet)
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry ==10:
                    raise e
        ###logger.info("Adding sheet to the same workbook")
        wb.sheets.add("Excl Macq & IC",after=wb.sheets[f"CTM Combined _{input_date}"]) 
        ws2=wb.sheets["Excl Macq & IC"]
        ###logger.info("Clearing its contents")
        ws2.cells.clear_contents()
        ###logger.info("Accessing Particular WorkBook[0]")
        ws1=wb.sheets[0]

        ###logger.info("Declaring Variables for columns and rows")
        last_row = ws1.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row
        column_list = ws1.range("A1").expand('right').value
        Customer_no_column=column_list.index('Customer')+1
        Customer_letter_column = num_to_col_letters(column_list.index('Customer')+1)
        Customer_data = ws1.range(f"{Customer_letter_column}1").expand('down').value
        Location_no_column=column_list.index('Location Id')+1
        Location_letter_column = num_to_col_letters(column_list.index('Location Id')+1)
        Location_data = ws1.range(f"{Location_letter_column}1").expand('down').value


        ###logger.info("Applying Filter to the same workbook")
        ws1.api.Range(f"{Customer_letter_column}1").AutoFilter(Field:=f'{Customer_no_column}', Criteria1:=["<>MACQUARIE COMMODITIES (USA) INC."], Operator:=1,Criteria2=["<>INTER-COMPANY PURCH/SALES"])
        ws1.api.Range(f"{Location_letter_column}1").AutoFilter(Field:=f'{Location_no_column}', Criteria1:=["<>WPMEXICO"], Operator:=1)
        ###logger.info("Copying and pasting Worksheet")
        ws1.api.AutoFilter.Range.Copy()
        ws2.api.Paste()
        ###logger.info("Applying Autofit")
        ws2.autofit()

        ###logger.info("Declaring Variables for columns and rows")
        column_list = ws1.range("A1").expand('right').value
        Customer_column = num_to_col_letters(column_list.index('Customer')+1)
        Customer_column_num = column_list.index('Customer')+1

        ###logger.info("Copying Inter Company Data from inp sheet  to Intercompany Sheet")
        try:
            ws1.api.AutoFilterMode=False
            ws1.api.Range(f"{Customer_column}1").AutoFilter(Feild:=Customer_column_num,Criteria1:="INTER-COMPANY PURCH/SALES") #Removing Intercompany
            intcomp_sht = wb.sheets.add("Intercompany", after=ws1)

            ws1.api.AutoFilter.Range.Copy()
            time.sleep(1)
            intcomp_sht.range("A1").api.Select()
            while True:
                try:
                    intcomp_sht.api.Paste()
                    break
                except:
                    time.sleep(1)
            wb.app.api.CutCopyMode=False
            time.sleep(1)
            ws1.api.AutoFilterMode=False
            # ###logger.info("Deleting intercompant data from original sheet after copying it in previous code")
            # ws1.api.AutoFilterMode=False
            # time.sleep(1)
            # ###logger.info("Declaring Variables for columns and rows")
            # last_row = ws1.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row
            # # last_row+=1
            # column_list = ws1.range("A1").expand('right').value
            # Customer_no_column=column_list.index('Customer')+1
            # Customer_letter_column = num_to_col_letters(column_list.index('Customer')+1)
            # ###logger.info("Applying loop for deleting INTER-COMPANY PURCH/SALES")
            # i = 2
            # while i <= last_row:
            # # for i in range(2,int(f'{last_row}')):
            #     if ws1.range(f"{Customer_letter_column}{i}").value=="INTER-COMPANY PURCH/SALES": 
            #         ws1.range(f"{i}:{i}").api.Delete(win32c.DeleteShiftDirection.xlShiftUp)
            #         print(i)
            #         i-=1               
            #     else:
            #         i+=1
            #         # continue
        except Exception as e:
            print("No (INTER-COMPANY PURCH/SALES) Present ")
            print(e)             
        last_row = ws2.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row
        column_list = ws2.range("A1").expand('right').value
        Contract_No_no_column=column_list.index('Contract No')+1
        Contract_No_letter_column = num_to_col_letters(column_list.index('Contract No')+1)
        ###logger.info("Adding Tier Coloumn and inserting value and dragging them")
        ws2.api.Range(f"{Contract_No_letter_column}1").EntireColumn.Insert()
        ws2.range(f"{Contract_No_letter_column}1").value = "Ship Tier"
        column_list = ws2.range("A1").expand('right').value
        Date_no_column=column_list.index('Delivery End Date')+1
        Date_letter_column = num_to_col_letters(column_list.index('Delivery End Date')+1)
        Ship_Tier_column=column_list.index('Ship Tier')+1
        Ship_Tier_column = num_to_col_letters(column_list.index('Ship Tier')+1)
        Date_data = ws2.range(f"{Date_letter_column}2").expand('down').value
        for index,values in enumerate(Date_data):
            if values == "Delinquent":
                index+=2
                ws2.range(f"{Ship_Tier_column}{index}").value = "W/n 12 Months"
            else:
                index+=2
                date2=datetime.strptime(values, "%b-%y")
                date1=datetime.strptime(input_date, "%m.%d.%Y")
                diff=(date2.year - date1.year) * 12 + (date2.month - date1.month)
                if diff <=12:
                        ws2.range(f"{Ship_Tier_column}{index}").value = "W/n 12 Months"
                else:
                    ws2.range(f"{Ship_Tier_column}{index}").value = ">12 Months"

        try:
            column_list = ws1.range("A1").expand('right').value
            Commodity_Id_column = num_to_col_letters(column_list.index('Commodity Id')+1)
            Commodity_Id_column_num = column_list.index('Commodity Id')+1
            ws1.api.AutoFilterMode=False
            ws1.api.Range(f"{Commodity_Id_column}1").AutoFilter(Feild:=Commodity_Id_column_num,Criteria1:="EQUIP") #Removing Intercompany
            intcomp_sht = wb.sheets.add("EXTRA", after=ws1)

            ws1.api.AutoFilter.Range.Copy()
            time.sleep(1)
            intcomp_sht.range("A1").api.Select()
            while True:
                try:
                    intcomp_sht.api.Paste()
                    break
                except:
                    time.sleep(1)
            wb.app.api.CutCopyMode=False
            time.sleep(1)
            ###logger.info("Deleting intercompant data from original sheet after copying it in previous code")
            ws1.api.AutoFilterMode=False
            time.sleep(1)
            ###logger.info("Declaring Variables for columns and rows")
            last_row = ws1.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row
            # last_row+=1
            column_list = ws1.range("A1").expand('right').value
            Commodity_Id_column=column_list.index('Commodity Id')+1
            Commodity_Id_letter_column = num_to_col_letters(column_list.index('Commodity Id')+1)
            ###logger.info("Applying loop for deleting INTER-COMPANY PURCH/SALES")
            i = 2
            while i <= last_row:
            # for i in range(2,int(f'{last_row}')):
                if ws2.range(f"{Commodity_Id_letter_column}{i}").value=="EQUIP": 
                    ws2.range(f"{i}:{i}").api.Delete(win32c.DeleteShiftDirection.xlShiftUp)
                    print(i)
                    i-=1               
                else:
                    i+=1
                    # continue
            last_row = intcomp_sht.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row
            last_row2 = ws2.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row
            last_row2+=10
            last_column = intcomp_sht.range('A1').end('right').last_cell.column
            last_column_letter=num_to_col_letters(intcomp_sht.range('A1').end('right').last_cell.column)
            x=last_row2+last_row
            intcomp_sht.range(f"A2:{last_column_letter}{last_row}").copy(ws2.range(f"A{last_row2}:{last_column_letter}{x}"))
            intcomp_sht.delete()
                    
        except Exception as e:
         print("No (INTER-COMPANY PURCH/SALES) Present ")
         print(e) 

        

        ###logger.info("Adding Worksheet for Pivot Table")
        wb.sheets.add("Pivot BB",after=wb.sheets["Excl Macq & IC"])
        ###logger.info("Clearing New Worksheet")
        wb.sheets["Pivot BB"].clear_contents()
        ws3=wb.sheets["Pivot BB"]
        ws3.range("A1").value="West Plains, LLC"
        ws3.range("A2").value="Net Unrealized Gains on Forward Contracts - Non MCUI"
        ws3.range("A2").api.Font.Bold = True
        ws3.range('A2').color ="#fffff"
        ws3.range("A3").value=input_date
        ###logger.info("Declaring Variables for columns and rows")
        last_column = ws2.range('A1').end('right').last_cell.column
        last_column_letter=num_to_col_letters(ws2.range('A1').end('right').last_cell.column)
        num_row = ws2.range('A1').end('down').row
        ###logger.info("Creating Pivot table")
        PivotCache=wb.api.PivotCaches().Create(SourceType=win32c.PivotTableSourceType.xlDatabase, SourceData=f"\'Excl Macq & IC\'!R1C1:R{num_row}C{last_column}", Version=win32c.PivotTableVersionList.xlPivotTableVersion14)
        PivotTable = PivotCache.CreatePivotTable(TableDestination="'Pivot BB'!R7C1", TableName="PivotTable1", DefaultVersion=win32c.PivotTableVersionList.xlPivotTableVersion14)
        ###logger.info("Adding particular Row Data in Pivot Table")
        PivotTable.PivotFields('Location Id').Orientation = win32c.PivotFieldOrientation.xlRowField
        PivotTable.PivotFields('Location Id').Position = 1
        # PivotTable.PivotFields('Tier').RepeatLabels=True
        PivotTable.PivotFields('Commodity Id').Orientation = win32c.PivotFieldOrientation.xlRowField
        ###logger.info("Adding particular Data Field in Pivot Table")
        PivotTable.PivotFields('Gain/LossTotal').Orientation = win32c.PivotFieldOrientation.xlDataField
        PivotTable.PivotFields('Sum of Gain/LossTotal').NumberFormat= '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
        ###logger.info("Adding particular Page Field in Pivot Table")
        PivotTable.PivotFields('Ship Tier').Orientation = win32c.PivotFieldOrientation.xlPageField
        ###logger.info("Applying filter in pagefield in Pivot Table")
        PivotTable.PivotFields('Ship Tier').CurrentPage = "W/n 12 Months"
        ###logger.info("Changing No Format in Pivot Table")
        ###logger.info("Changing Table layout")
        PivotTable.PivotFields('Location Id').Subtotals=(False, False, False, False, False, False, False, False, False, False, False, False)
        PivotTable.RowAxisLayout(1)
        ###logger.info("Changing Table Style")
        PivotTable.TableStyle2 = ""

        ###logger.info("Declaring Variables for columns and rows")
        last_column = ws2.range('A1').end('right').last_cell.column
        last_column_letter=num_to_col_letters(ws2.range('A1').end('right').last_cell.column)
        num_row = ws2.range('A1').end('down').row
        last_row2 = ws3.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row
        last_row2+=10
        ###logger.info("Creating Pivot table")
        PivotCache=wb.api.PivotCaches().Create(SourceType=win32c.PivotTableSourceType.xlDatabase, SourceData=f"\'Excl Macq & IC\'!R1C1:R{num_row}C{last_column}", Version=win32c.PivotTableVersionList.xlPivotTableVersion14)
        PivotTable = PivotCache.CreatePivotTable(TableDestination=f"'Pivot BB'!R{last_row2}C1", TableName="PivotTable2", DefaultVersion=win32c.PivotTableVersionList.xlPivotTableVersion14)
        ###logger.info("Adding particular Row Data in Pivot Table")
        PivotTable.PivotFields('Location Id').Orientation = win32c.PivotFieldOrientation.xlRowField
        PivotTable.PivotFields('Location Id').Position = 1
        # PivotTable.PivotFields('Tier').RepeatLabels=True
        PivotTable.PivotFields('Commodity Id').Orientation = win32c.PivotFieldOrientation.xlRowField
        PivotTable.PivotFields('Delivery End Date').Orientation = win32c.PivotFieldOrientation.xlRowField
        ###logger.info("Adding particular Data Field in Pivot Table")
        PivotTable.PivotFields('Gain/LossTotal').Orientation = win32c.PivotFieldOrientation.xlDataField
        PivotTable.PivotFields('Sum of Gain/LossTotal').NumberFormat= '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
        ###logger.info("Adding particular Page Field in Pivot Table")
        PivotTable.PivotFields('Ship Tier').Orientation = win32c.PivotFieldOrientation.xlPageField
        ###logger.info("Applying filter in pagefield in Pivot Table")
        PivotTable.PivotFields('Ship Tier').CurrentPage = ">12 months"
        ###logger.info("Changing No Format in Pivot Table")
        ###logger.info("Changing Table layout")
        PivotTable.PivotFields('Location Id').Subtotals=(False, False, False, False, False, False, False, False, False, False, False, False)
        PivotTable.PivotFields('Commodity Id').Subtotals=(False, False, False, False, False, False, False, False, False, False, False, False)
        PivotTable.RowAxisLayout(1)
        ###logger.info("Changing Table Style")
        PivotTable.TableStyle2 = ""

        last_column = ws3.range('A7').end('right').last_cell.column
        last_column+=3
        last_row = ws3.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row
        last_row+=5
        last_column_letter=num_to_col_letters(last_column)
        ws3.range(f"{last_column_letter}{last_row}").value=f'=GETPIVOTDATA("Gain/LossTotal",$A$7)+GETPIVOTDATA("Gain/LossTotal",$A${last_row2})'

        ws3.range(f"{last_column_letter}{last_row}").api.NumberFormat= '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
        # last_col_num = ws1.range('A1').expand('right').last_cell.column 
        # # last_col = num_to_col_letters(last_col_num) 
        # last_row = ws2.range(f'A'+ str(ws2.cells.last_cell.row)).end('up').row 
        # ######logger.info("Adding Worksheet for Pivot Table") 
        # wb.sheets.add("For allocation entry",before=ws1) 
        # ######logger.info("Creating Pivot table") 
        # PivotCache=wb.api.PivotCaches().Create(SourceType=win32c.PivotTableSourceType.xlDatabase, SourceData=f'\'{ws1.name}\'!R1C1:R{last_row}C{last_col_num}', Version=win32c.PivotTableVersionList.xlPivotTableVersion14) 
        # PivotTable = PivotCache.CreatePivotTable(TableDestination="'For allocation entry'!R3C1", TableName="PivotTable1", DefaultVersion=win32c.PivotTableVersionList.xlPivotTableVersion14)
        #  ######logger.info("Adding particular Row in Pivot Table") 
        # PivotTable.PivotFields('Location Name').Orientation = win32c.PivotFieldOrientation.xlRowField
        # PivotTable.PivotFields('Net').Orientation = win32c.PivotFieldOrientation.xlDataField
        #  # PivotTable.PivotFields('Sum of Net').NumberFormat= '0.00'
        wb.save(output_location)
        wb.app.quit()
        return f"CTM Combined Report Generated for date {input_date}"
    except Exception as e:
        raise e
     
    finally:
        try:
            wb.app.quit()
        except:
            pass
def freight_analysis(input_date, output_date):
    try:
        inp_formula_sht = r'J:\WEST PLAINS\REPORT\Freight analysis reports\Col_N_Formulas.xlsx'
        
        output_location = r'J:\WEST PLAINS\REPORT\Freight analysis reports\Output files'
        raw_input = r'J:\WEST PLAINS\REPORT\Freight analysis reports\Raw files'

        Input_Sheets = ['Inbound','Outbound', 'DS Outbound', 'DS Inbound']

        inp_month = datetime.strftime(datetime.strptime(input_date, "%m.%d.%Y"),"%B")
        inp_month_2 = datetime.strftime(datetime.strptime(input_date, "%m.%d.%Y"),"%b")
        inp_year = datetime.strftime(datetime.strptime(input_date, "%m.%d.%Y"),"%y")
        inp_year_2 = datetime.strftime(datetime.strptime(input_date, "%m.%d.%Y"),"%Y")
        
        
        for sheet in Input_Sheets:
            #####logger.info(f"Starting for {sheet}")
            # inbound_sheet = os.getcwd()+f"\\Raw files\\{input_sheet} {prev_month} {current_year}.xlsx"
            
            input_sheet = raw_input+f"\\{sheet} {inp_month} {inp_year}.xlsx"
            #####logger.info(f"path is {input_sheet}")
            #####logger.info(f"Path exists {os.path.exists(input_sheet)}")
            print(input_sheet)
            location_check = os.path.exists(input_sheet)
            if not location_check:
                return(f"{input_sheet} Excel file not present for month year:{inp_month} {inp_year}")
            retry=0
            while retry<10:
                try:
                    wb = xw.Book(input_sheet, update_links=True)
                    break
                except Exception as e:
                    time.sleep(5)
                    retry+=1
                    if retry==10:
                        raise e
            
            #####logger.info("Sheet Opened")
            # time.sleep(10)
            while True:
                try:
                    inp_sht = wb.sheets[0]
                    break
                except:
                    time.sleep(10)
            


            last_row = inp_sht.range(f'A'+ str(inp_sht.cells.last_cell.row)).end('up').row


            df = pd.read_excel(inp_formula_sht, sheet_name=sheet)

            # data_dict = df.set_index('Column').T.to_dict('list')
            #####logger.info("Fetching colnformula sheet into dict")
            columns = df.set_index(['Column'])["Name"].to_dict()

            formulas = df.set_index(['Column'])["Formula"].to_dict()
            for name in columns:
                inp_sht.range(f"{name}:{name}").insert()
                inp_sht.range(f"{name}1").value = columns[name]
                inp_sht.range(f"{name}1").color = "#FFFF00"
            for column in formulas:
                formulas[column] = formulas[column].replace("inp_year_2", inp_year_2)
                formulas[column] = formulas[column].replace("inp_month_2", inp_month_2)

                formulas[column] = formulas[column].replace("inp_month", inp_month)
                formulas[column] = formulas[column].replace("inp_year", inp_year)
                
                
                formulas[column] = formulas[column].replace("input_date", input_date)
                inp_sht.range(f"{column}2").value = formulas[column].replace('"','')
                inp_sht.api.Select()
                inp_sht.api.Range(f"{column}2").Copy()
                inp_sht.range(f"{column}3:{column}{last_row}").api.Select()
                inp_sht.api.Paste()
                wb.app.api.CutCopyMode=False

            #####logger.info("Splitting Unapplied Contracts")

            column_list = inp_sht.range("A1").expand('right').value
            contract_no= column_list.index('Contract No')+1
            contract_no_column = num_to_col_letters(column_list.index('Contract No')+1)
            contract_column = num_to_col_letters(column_list.index('Contract No'))

            for i in range(1,last_row+1):
                if inp_sht.range(f"{contract_column}{i}").value == "TS":
                    for name in columns:
                        if columns[name] == 'Market Zone' or columns[name] == 'Destination Market zone':
                            inp_sht.range(f"{name}{i}").value = 0
                        elif columns[name] == 'Freight term' or columns[name] == 'Destination Freight Term':
                            inp_sht.range(f"{name}{i}").value = "TS"



            contract_data = inp_sht.range(f"{contract_no_column}1").expand('down').value

            orignal_name = inp_sht.name

            inp_sht.name = inp_sht.name+" Org"
            m_sht = wb.sheets.add(orignal_name)
            inp_sht.api.AutoFilterMode=False
            inp_sht.api.Range(f"{contract_no_column}1").AutoFilter(Feild:=contract_no,Criteria1:="<>") #Non Blank Data
            inp_sht.api.AutoFilter.Range.Copy()

            m_sht.range("A1").api.Select()

            m_sht.api.Paste()
            wb.app.api.CutCopyMode=False

            
            
            
            inp_sht.api.AutoFilterMode=False
            #####logger.info(f"Cotract No column is {contract_no_column}")
            inp_sht.api.Range(f"{contract_no_column}1").AutoFilter(Feild:=contract_no,Criteria1:="=") #Blank Data
            inp_sht.api.AutoFilter.Range.Copy()
            unap_sht = wb.sheets.add("Unapplied tickets")
            unap_sht.range("A1").api.Select()
            unap_sht.api.Paste()
            wb.app.api.CutCopyMode=False
            #####logger.info("Deleting input sheet")
            inp_sht.delete()


            wb.save(f"{output_location}\\{sheet} {inp_month} {inp_year}.xlsx")
            wb.app.quit()
        return f"Freight Analysis reports Generated for {inp_month} {inp_year}"
    except Exception as e:
        raise e
    finally:
        try:
            wb.app.quit()
        except:
            pass
def inv_mtm(input_date, output_date):
    try:
        print(input_date)
        # print(output_date)

        input_xl = r'J:\WEST PLAINS\REPORT\MTM reports\Raw Files\Inventory MTM_'+input_date+".xlsx"
        if not os.path.exists(input_xl):
            return(f"{input_xl} Excel file not present for date {input_date}")
        
        pdf_loc = r'J:\WEST PLAINS\REPORT\MTM reports\Raw Files\Inventory Market Valuation _'+input_date+'.pdf'
        if not os.path.exists(pdf_loc):
            return(f"{pdf_loc} Pdf file not present for date {input_date}")

        hrw_pdf_loc = r'J:\WEST PLAINS\REPORT\MTM reports\Raw Files\HRW_'+input_date+'.pdf'
        if not os.path.exists(hrw_pdf_loc):
            return(f"{hrw_pdf_loc} Pdf file not present for date {input_date}")

        yc_pdf_loc = r'J:\WEST PLAINS\REPORT\MTM reports\Raw Files\YC_'+input_date+'.pdf'
        if not os.path.exists(yc_pdf_loc):
            return(f"{yc_pdf_loc} Pdf file not present for date {input_date}")

        loc_sheet = r'J:\WEST PLAINS\REPORT\MTM reports\Loc_Abbr.xlsx'
        if not os.path.exists(loc_sheet):
            return(f"{loc_sheet}Excel file not present for date {input_date}")

        loc_dict, hrw_fut, yc_fut = mtm_pdf_data_extractor(input_date,pdf_loc, hrw_pdf_loc, yc_pdf_loc)
        output_location = r'J:\WEST PLAINS\REPORT\MTM reports\Output files\Inventory MTM_'+input_date+".xlsx"
        mtm_excel(input_date, input_xl,loc_dict,loc_sheet, output_location, hrw_fut, yc_fut)

        print("Done")
        return f"MTM report Generated for {input_date}"
    except Exception as e:
        raise e
    

def open_ar(input_date, output_date):
    try:
        input_sheet = r'J:\WEST PLAINS\REPORT\Open AR\Raw files'+f'\\Open AR _{input_date} - Production.xlsx' 
        if not os.path.exists(input_sheet):
            return(f"{input_sheet} Excel file not present for date {input_date}")
        prev_output=r'J:\WEST PLAINS\REPORT\Open AR\Output files'+f'\\Open AR _{output_date} - Production.xlsx'
        if not os.path.exists(prev_output):
            return(f"{prev_output} Excel file not present for date {output_date}")  

        output_location = r'J:\WEST PLAINS\REPORT\Open AR\Output files'  
        prev_month = datetime.strftime(datetime.strptime(input_date, "%m.%d.%Y"), "%B")
        ##logger.info("Opening operating workbook instance of excel")
        retry=0
        while retry < 10:
            try:
                wb=xw.Book(input_sheet)
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry ==10:
                    raise e
        ##logger.info("Adding sheet to the same workbook")
        wb.sheets.add("Excl Macq & IC",after=wb.sheets[f"Open AR _{input_date} - Productio"]) 
        ws2=wb.sheets["Excl Macq & IC"]
        ##logger.info("Clearing its contents")
        ws2.cells.clear_contents()
        ##logger.info("Accessing Particular WorkBook[0]")
        ws1=wb.sheets[f"Open AR _{input_date} - Productio"]

        ##logger.info("Declaring Variables for columns and rows")
        last_row = ws1.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row
        column_list = ws1.range("A1").expand('right').value
        Customer_no_column=column_list.index('Customer')+1
        Customer_letter_column = num_to_col_letters(column_list.index('Customer')+1)
        Customer_data = ws1.range(f"{Customer_letter_column}1").expand('down').value
        Location_no_column=column_list.index('Location')+1
        Location_letter_column = num_to_col_letters(column_list.index('Location')+1)
        Location_data = ws1.range(f"{Location_letter_column}1").expand('down').value
        Total_AR_no_column=column_list.index('Total AR')+1
        Total_AR_letter_column = num_to_col_letters(column_list.index('Total AR')+1)
        Total_AR_data = ws1.range(f"{Total_AR_letter_column}1").expand('down').value


        ##logger.info("Applying Filter to the same workbook")
        ws1.api.Range(f"{Customer_letter_column}1").AutoFilter(Field:=f'{Customer_no_column}', Criteria1:=["<>MACQUARIE COMMODITIES (USA) INC."], Operator:=1,Criteria2=["<>INTER-COMPANY PURCH/SALES"])
        ws1.api.Range(f"{Location_letter_column}1").AutoFilter(Field:=f'{Location_no_column}', Criteria1:=["<>WPMEXICO"], Operator:=1)
        ws1.api.Range(f"{Total_AR_letter_column}1").AutoFilter(Field:=f'{Total_AR_no_column}', Criteria1:="<>0", Operator:=1)
        ##logger.info("Copying and pasting Worksheet")
        ws1.api.AutoFilter.Range.Copy()
        ws2.api.Paste()
        ##logger.info("Applying Autofit")
        ws2.autofit()

        ##logger.info("Declaring Variables for columns and rows")
        column_list = ws1.range("A1").expand('right').value
        Customer_column = num_to_col_letters(column_list.index('Customer')+1)
        Customer_column_num = column_list.index('Customer')+1

        ##logger.info("Copying Inter Company Data from inp sheet  to Intercompany Sheet")
        try:
            ws1.api.AutoFilterMode=False
            ws1.api.Range(f"{Customer_column}1").AutoFilter(Feild:=Customer_column_num,Criteria1:="INTER-COMPANY PURCH/SALES") #Removing Intercompany
            intcomp_sht = wb.sheets.add("Intercompany", after=ws1)

            ws1.api.AutoFilter.Range.Copy()
            time.sleep(1)
            intcomp_sht.range("A1").api.Select()
            while True:
                try:
                    intcomp_sht.api.Paste()
                    break
                except:
                    time.sleep(1)
            wb.app.api.CutCopyMode=False
            time.sleep(1)
            ws1.api.AutoFilterMode=False
            # ##logger.info("Deleting intercompant data from original sheet after copying it in previous code")
            # ws1.api.AutoFilterMode=False
            # time.sleep(1)
            # ##logger.info("Declaring Variables for columns and rows")
            # last_row = ws1.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row
            # # last_row+=1
            # column_list = ws1.range("A1").expand('right').value
            # Customer_no_column=column_list.index('Customer')+1
            # Customer_letter_column = num_to_col_letters(column_list.index('Customer')+1)
            # ##logger.info("Applying loop for deleting INTER-COMPANY PURCH/SALES")
            # i = 2
            # while i <= last_row:
            #     if ws1.range(f"{Customer_letter_column}{i}").value=="INTER-COMPANY PURCH/SALES": 
            #         ws1.range(f"{i}:{i}").api.Delete(win32c.DeleteShiftDirection.xlShiftUp)
            #         # print(i)
            #         i-=1                   
            #     else:
            #         i+=1
        except Exception as e:
            print("No (INTER-COMPANY PURCH/SALES) Present ")
            print(e)   
        ##logger.info("Copying tier column from previous output sheet")   
        retry=0
        while retry < 10:
            try:
                tier_wb = xw.Book(prev_output,update_links=True)
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry ==10:
                    raise e     
        # tier_wb = xw.Book(prev_output,update_links=True)
        tier_sht = tier_wb.sheets("Tier")
        ##logger.info("Copy tier sheet AFTER the intercompany sheet of input book.")
        tier_sht.api.Copy(None, After=ws2.api)
        tier_wb.close()
        ##logger.info("Declaring Variables for columns and rows")
        ws3=wb.sheets['Tier']
        last_row = ws2.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row
        column_list = ws2.range("A1").expand('right').value
        Customer_no_column=column_list.index('Customer')+1
        Customer_letter_column = num_to_col_letters(column_list.index('Customer')+1)
        Customer_data = ws2.range(f"{Customer_letter_column}2").expand('down').value
        mylist = list(dict.fromkeys(Customer_data))
        ##logger.info("Declaring Variables for columns and rows")
        last_row = ws3.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row
        column_list = ws3.range("A1").expand('right').value
        Customer_Name_no_column=column_list.index('Customer Name')+1
        Customer_Name_letter_column = num_to_col_letters(column_list.index('Customer Name')+1)
        Customer_Name_data = ws3.range(f"{Customer_Name_letter_column}2").expand('down').value
        ##logger.info("Declaring Variables for columns and rows")
        last_row_value=last_row+1
        Tier_letter_column = num_to_col_letters(column_list.index('Tier')+1)

        for names in mylist:
            if names in Customer_Name_data:
                pass
            else:
                ws3.range(f"{Customer_Name_letter_column}{last_row_value}").value = names
                ws3.range(f"{Customer_Name_letter_column}{last_row_value}").font.name = 'Calibri'
                ws3.range(f"{Tier_letter_column}{last_row_value}").value = "Tier II"
                ws3.range(f"{Tier_letter_column}{last_row_value}").font.name = 'Calibri'
                last_row_value+=1
        ##logger.info("Declaring Variables for columns and rows")
        last_row = ws2.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row
        column_list = ws2.range("A1").expand('right').value
        Credit_Limit_no_column=column_list.index('Credit Limit')+1
        Credit_Limit_letter_column = num_to_col_letters(column_list.index('Credit Limit')+1)
        ##logger.info("Adding Tier Coloumn and inserting value and dragging them")
        ws2.api.Range(f"{Credit_Limit_letter_column}1").EntireColumn.Insert()
        ws2.range(f"{Credit_Limit_letter_column}1").value = "Tier"
        ws2.range(f"{Credit_Limit_letter_column}2").value ="=VLOOKUP(H2,Tier!A:B,2,0)"
        ws2.range(f"{Credit_Limit_letter_column}2").copy(ws2.range(f"{Credit_Limit_letter_column}2:{Credit_Limit_letter_column}{last_row}"))

        ##logger.info("Adding Worksheet for Pivot Table")
        wb.sheets.add("Pivot Summary",after=wb.sheets["Tier"])
        ##logger.info("Clearing New Worksheet")
        wb.sheets["Pivot Summary"].clear_contents()
        # ws4=wb.sheets["Pivot Summary"]
        ##logger.info("Declaring Variables for columns and rows")
        last_row = ws2.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row
        last_column = ws2.range('A1').end('right').last_cell.column
        last_column_letter=num_to_col_letters(ws2.range('A1').end('right').last_cell.column)
        ##logger.info("Creating Pivot table")
        PivotCache=wb.api.PivotCaches().Create(SourceType=win32c.PivotTableSourceType.xlDatabase, SourceData=f"\'Excl Macq & IC\'!R1C1:R{last_row}C{last_column}", Version=win32c.PivotTableVersionList.xlPivotTableVersion14)
        PivotTable = PivotCache.CreatePivotTable(TableDestination="'Pivot Summary'!R1C1", TableName="PivotTable1", DefaultVersion=win32c.PivotTableVersionList.xlPivotTableVersion14)
        ##logger.info("Adding particular Row Data in Pivot Table")
        PivotTable.PivotFields('Tier').Orientation = win32c.PivotFieldOrientation.xlRowField
        PivotTable.PivotFields('Tier').Position = 1
        PivotTable.PivotFields('Tier').RepeatLabels=True
        PivotTable.PivotFields('Customer').Orientation = win32c.PivotFieldOrientation.xlRowField
        ##logger.info("Adding particular Data Field in Pivot Table")
        PivotTable.PivotFields('Current').Orientation = win32c.PivotFieldOrientation.xlDataField
        PivotTable.PivotFields('Sum of Current').NumberFormat= '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'
        PivotTable.PivotFields('1 - 10').Orientation = win32c.PivotFieldOrientation.xlDataField
        PivotTable.PivotFields('1 - 10').NumberFormat= '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'
        PivotTable.PivotFields('11 - 30').Orientation = win32c.PivotFieldOrientation.xlDataField
        PivotTable.PivotFields('11 - 30').NumberFormat= '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'
        PivotTable.PivotFields('31 - 60').Orientation = win32c.PivotFieldOrientation.xlDataField
        PivotTable.PivotFields('31 - 60').NumberFormat= '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'
        PivotTable.PivotFields('61 - 9999').Orientation = win32c.PivotFieldOrientation.xlDataField
        PivotTable.PivotFields('61 - 9999').NumberFormat= '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'
        PivotTable.PivotFields('Total AR').Orientation = win32c.PivotFieldOrientation.xlDataField
        PivotTable.PivotFields('Total AR').NumberFormat= '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'
        ##logger.info("Changing No Format in Pivot Table")
        ##logger.info("Changing Table layout")
        PivotTable.RowAxisLayout(1)
        ##logger.info("Changing Table Style")
        PivotTable.TableStyle2 = ""

        # PivotTable.TableStyle2 = ""
        ##logger.info("Removing subtotal from Tier")
        PivotTable.PivotFields('Tier').Subtotals=(False, False, False, False, False, False, False, False, False, False, False, False)
        ws4=wb.sheets["Pivot Summary"]
        ##logger.info("Adding Worksheet Eligible")
        wb.sheets.add("Eligible",after=wb.sheets["Pivot Summary"])
        ws5=wb.sheets["Eligible"]
        ##logger.info("Declaring Variables for columns and rows and sheets")
        ws4=wb.sheets["Pivot Summary"]
        last_row = ws4.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row
        final=last_row-1
        last_column = ws4.range('A1').end('right').last_cell.column
        last_column_letter=num_to_col_letters(ws4.range('A1').end('right').last_cell.column)
        ##logger.info("Copying and pasting sheet to Eligible Worksheet")
        ws4.api.Range(f'A1:{last_column_letter}{final}').Copy()
        ws5.api.Paste()
        ws5.autofit()
        ##logger.info("Changing names of columns in new sheet")
        column_list = ws5.range("A1").expand('right').value
        changed_column_list=['Tier', 'Customer Name', 'Current', ' 1 - 10', ' 11 - 30', ' 31 - 60', ' 61 - 9999', 'Balance']
        i=0
        for values in column_list:
            values_column_no=column_list.index(values)+1
            values_letter_column = num_to_col_letters(column_list.index(values)+1)
            ws5.range(f"{values_letter_column}1").value = changed_column_list[i]
            ws3.range(f"{Tier_letter_column}{last_row_value}").font.name = 'Calibri'
            i+=1
        ##logger.info("Inserting extra Culumns,adding their values and dragging them")
        list1=["Portion of Customer Account Greater than 10 Days Past Due","Eligiblity","total","Diff"]
        list2=['=IF(H2<=0,0,SUM(E2:G2)/H2)','=IF(I2>=0.5,"Ineligible","Eligible")','=+SUM(C2:G2)','=+H2-K2'] 
        last_column = ws5.range('A1').end('right').last_cell.column
        last_column+=1
        i=0
        last_row = ws5.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row
        for values in list1:
            last_column_letter=num_to_col_letters(last_column)
            ws5.range(f"{last_column_letter}1").value = values
            ws5.range(f"{last_column_letter}1").api.Font.Bold = True
            ws5.range(f"{last_column_letter}2").value = list2[i]
            time.sleep(1)
            ws5.range(f"{last_column_letter}2").copy(ws5.range(f"{last_column_letter}2:{last_column_letter}{last_row}"))
            i+=1
            last_column+=1
        ##logger.info("Applying same previous operation for extra hidden columns")
        list3=["c","c1","d","d1"]
        list4=['=SUM(D2:G2)',' ','=+SUM(E2:G2)',' '] 
        last_column = ws5.range('A1').end('right').last_cell.column
        last_column+=1
        i=0
        last_row = ws5.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row
        for values in list3:
            last_column_letter=num_to_col_letters(last_column)
            ws5.range(f"{last_column_letter}1").value = values
            ws5.range(f"{last_column_letter}2").value = list4[i]
            time.sleep(1)
            ws5.range(f"{last_column_letter}2").copy(ws5.range(f"{last_column_letter}2:{last_column_letter}{last_row}"))
            i+=1
            last_column+=1
        #CHECK FOR CASH CUSTOMER
        ##logger.info("CHECK FOR CASH CUSTOMER AND MAKING HIM INELIGIBLE")
        column_list = ws5.range("A1").expand('right').value
        Customer_Name_no_column=column_list.index('Customer Name')+1
        Customer_Name_letter_column = num_to_col_letters(column_list.index('Customer Name')+1)
        Customer_Name_data = ws5.range(f"{Customer_Name_letter_column}2").expand('down').value
        for values in Customer_Name_data:
            if 'CASH CUSTOMER' in values:
                values_row_no=Customer_Name_data.index('CASH CUSTOMER')+2
                Eligiblity_letter_column = num_to_col_letters(column_list.index('Eligiblity')+1)
                ws5.range(f"{Eligiblity_letter_column}{values_row_no}").value = 'Ineligible'
            else:
                pass


        #Paste Special Values For Values In c & d
        ##logger.info("Paste Special Values For Values In c & d")
        column_list = ws5.range("A1").expand('right').value
        last_row = ws5.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row
        list5=["c","d"]
        list6=["c1","d1"]
        i=0
        for values in list5:
            c_no_column=column_list.index(values)+1
            c_letter_column = num_to_col_letters(column_list.index(values)+1)    
            ws5.api.Range(f'{c_letter_column}2:{c_letter_column}{last_row}').Copy()
            c1_no_column=column_list.index(list6[i])+1
            c1_letter_column = num_to_col_letters(column_list.index(list6[i])+1)
            ws5.api.Range(f'{c1_letter_column}2:{c1_letter_column}{last_row}')._PasteSpecial(Paste=-4163)
            i+=1
        ##logger.info("Declaring Variables for columns and rows")
        last_row = ws5.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row
        last_column = ws5.range('A1').end('right').last_cell.column
        last_column_letter=num_to_col_letters(last_column)
        ws5.range(f'A1:{last_column_letter}{last_row}').api.NumberFormat= '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'

        # c_no_column=column_list.index("c")+1
        # c_letter_column = num_to_col_letters(column_list.index("c")+1)  
        # ws5.api.Range(f"{c_letter_column}1").AutoFilter(Field:=f'{c_no_column}', Criteria1:="<0", Operator:=1)

        # i=0
        last_row = ws5.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row
        last_row+=1
        ##logger.info("Starting loop for C column adjustment")
        for i in range(2,int(f'{last_row}')):
            
            if ws5.range(f"M{i}").value<0:
                print(i)
                ws5.range(f"D{i}").value=ws5.range(f"N{i}").value
                ws5.range(f"E{i}").value=0
                ws5.range(f"F{i}").value=0
                ws5.range(f"G{i}").value=0
                ws5.range(f"P{i}").value=0
            # elif ws5.range(f"O{i}").value<0:
            #     ws5.range(f"E{i}").value=ws5.range(f"P{i}").value
            #     ws5.range(f"F{i}").value=0
            #     ws5.range(f"G{i}").value=0
            else:
                pass   
        ##logger.info("Adding variables")    
        last_row = ws5.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row
        last_row+=1
        ##logger.info("Starting loop for D column adjustment")
        for i in range(2,int(f'{last_row}')):
            
            if ws5.range(f"O{i}").value<0:
                print(i)
                ws5.range(f"E{i}").value=ws5.range(f"P{i}").value
                ws5.range(f"F{i}").value=0
                ws5.range(f"G{i}").value=0
            else:
                pass    
        
        for i in range(2,int(f'{last_row}')):
            if ws5.range(f"G{i}").value<0:
                ws5.range(f"F{i}").value = ws5.range(f"F{i}").value + ws5.range(f"G{i}").value
                ws5.range(f"G{i}").value = 0
        for i in range(2,int(f'{last_row}')):       
                if ws5.range(f"F{i}").value<0:
                    ws5.range(f"E{i}").value = ws5.range(f"E{i}").value + ws5.range(f"F{i}").value
                    ws5.range(f"F{i}").value = 0
        for i in range(2,int(f'{last_row}')):            
                    if ws5.range(f"E{i}").value<0:
                        ws5.range(f"D{i}").value = ws5.range(f"D{i}").value + ws5.range(f"E{i}").value
                        ws5.range(f"E{i}").value = 0
        
        ##logger.info("Adding Worksheet for Pivot Table")
        wb.sheets.add("Pivot BB",after=wb.sheets["Eligible"])
        ##logger.info("Clearing contents for new sheet")
        wb.sheets["Pivot BB"].clear_contents()
        ws6=wb.sheets["Pivot BB"]
        ##logger.info("Declaring Variables for columns and rows")
        last_row = ws5.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row
        last_column = ws5.range('A1').end('right').last_cell.column
        last_column_letter=num_to_col_letters(ws5.range('A1').end('right').last_cell.column)
        ##logger.info("Creating Pivot Table")
        PivotCache=wb.api.PivotCaches().Create(SourceType=win32c.PivotTableSourceType.xlDatabase, SourceData=f"\'Eligible\'!R1C1:R{last_row}C{last_column}", Version=win32c.PivotTableVersionList.xlPivotTableVersion14)
        PivotTable = PivotCache.CreatePivotTable(TableDestination="'Pivot BB'!R3C1", TableName="PivotTable1", DefaultVersion=win32c.PivotTableVersionList.xlPivotTableVersion14)
        ##logger.info("Adding particular Row in Pivot Table")

        PivotTable.PivotFields('Tier').Orientation = win32c.PivotFieldOrientation.xlRowField
        PivotTable.PivotFields('Tier').Position = 1
        PivotTable.PivotFields('Customer Name').Orientation = win32c.PivotFieldOrientation.xlRowField
        ##logger.info("Adding particular Data Field in Pivot Table")
        PivotTable.PivotFields('Current').Orientation = win32c.PivotFieldOrientation.xlDataField
        PivotTable.PivotFields('Sum of Current').NumberFormat= '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'
        PivotTable.PivotFields(' 1 - 10').Orientation = win32c.PivotFieldOrientation.xlDataField
        PivotTable.PivotFields('Sum of  1 - 10').NumberFormat= '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'
        PivotTable.PivotFields(' 11 - 30').Orientation = win32c.PivotFieldOrientation.xlDataField
        PivotTable.PivotFields('Sum of  11 - 30').NumberFormat= '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'
        PivotTable.PivotFields(' 31 - 60').Orientation = win32c.PivotFieldOrientation.xlDataField
        PivotTable.PivotFields('Sum of  31 - 60').NumberFormat= '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'
        PivotTable.PivotFields(' 61 - 9999').Orientation = win32c.PivotFieldOrientation.xlDataField
        PivotTable.PivotFields('Sum of  61 - 9999').NumberFormat= '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'
        PivotTable.PivotFields('Balance').Orientation = win32c.PivotFieldOrientation.xlDataField
        PivotTable.PivotFields('Sum of Balance').NumberFormat= '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'
        ##logger.info("Adding particular Page Field in Pivot Table")
        PivotTable.PivotFields('Eligiblity').Orientation = win32c.PivotFieldOrientation.xlPageField
        ##logger.info("Applying filter in Data Field in Pivot Table")
        PivotTable.PivotFields('Eligiblity').CurrentPage = "Eligible"
        ##logger.info("Changing No Format in Pivot Table")
        # PivotTable.RowAxisLayout(1)
        ##logger.info("Changing Table Style in Pivot Table")
        PivotTable.TableStyle2 = ""
        ##logger.info("Changing Table Layout in Pivot Table")
        PivotTable.RowAxisLayout(1)
        ##logger.info("Declaring Variables for columns and rows")
        last_row = ws5.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row
        last_column = ws5.range('A1').end('right').last_cell.column
        last_column_letter=num_to_col_letters(ws5.range('A1').end('right').last_cell.column)
        last_row2 = ws6.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row
        last_row2+=10
        ##logger.info("Creating Pivot Table")
        PivotCache=wb.api.PivotCaches().Create(SourceType=win32c.PivotTableSourceType.xlDatabase, SourceData=f"\'Eligible\'!R1C1:R{last_row}C{last_column}", Version=win32c.PivotTableVersionList.xlPivotTableVersion14)
        PivotTable = PivotCache.CreatePivotTable(TableDestination=f"'Pivot BB'!R{last_row2}C1", TableName="PivotTable2", DefaultVersion=win32c.PivotTableVersionList.xlPivotTableVersion14)
        ##logger.info("Adding particular data in RowField in Pivot Table")

        PivotTable.PivotFields('Tier').Orientation = win32c.PivotFieldOrientation.xlRowField
        PivotTable.PivotFields('Tier').Position = 1
        PivotTable.PivotFields('Customer Name').Orientation = win32c.PivotFieldOrientation.xlRowField
        ##logger.info("Adding particular Data Field in Pivot Table")
        PivotTable.PivotFields('Current').Orientation = win32c.PivotFieldOrientation.xlDataField
        PivotTable.PivotFields('Sum of Current').NumberFormat= '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'
        PivotTable.PivotFields(' 1 - 10').Orientation = win32c.PivotFieldOrientation.xlDataField
        PivotTable.PivotFields('Sum of  1 - 10').NumberFormat= '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'
        PivotTable.PivotFields(' 11 - 30').Orientation = win32c.PivotFieldOrientation.xlDataField
        PivotTable.PivotFields('Sum of  11 - 30').NumberFormat= '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'
        PivotTable.PivotFields(' 31 - 60').Orientation = win32c.PivotFieldOrientation.xlDataField
        PivotTable.PivotFields('Sum of  31 - 60').NumberFormat= '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'
        PivotTable.PivotFields(' 61 - 9999').Orientation = win32c.PivotFieldOrientation.xlDataField
        PivotTable.PivotFields('Sum of  61 - 9999').NumberFormat= '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'
        PivotTable.PivotFields('Balance').Orientation = win32c.PivotFieldOrientation.xlDataField
        PivotTable.PivotFields('Sum of Balance').NumberFormat= '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'
        ##logger.info("Adding particular Data Field in Pivot Table")
        PivotTable.PivotFields('Eligiblity').Orientation = win32c.PivotFieldOrientation.xlPageField
        ##logger.info("Applying filter in pagefield in Pivot Table")
        PivotTable.PivotFields('Eligiblity').CurrentPage = "Ineligible"
        ##logger.info("Changing No Format in Pivot Table")
        # PivotTable.RowAxisLayout(1)
        ##logger.info("Changing Table Style in Pivot Table")
        PivotTable.TableStyle2 = ""
        ##logger.info("Changing Layout for Pivot Table")
        PivotTable.RowAxisLayout(1)
        ##logger.info("Doing final adjustments for Sheets")
        ws6.autofit()
        wb.app.api.CutCopyMode=False
        wb.app.api.Autofilter=False
        wb.app.api.AutofilterMode=False

        last_col_num = ws1.range('A1').expand('right').last_cell.column 
        # last_col = num_to_col_letters(last_col_num) 
        last_row = ws1.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row 
        #####logger.info("Adding Worksheet for Pivot Table") 
        wb.sheets.add("For allocation entry",before=ws1) 
        #####logger.info("Creating Pivot table") 
        PivotCache=wb.api.PivotCaches().Create(SourceType=win32c.PivotTableSourceType.xlDatabase, SourceData=f'\'{ws1.name}\'!R1C1:R{last_row}C{last_col_num}', Version=win32c.PivotTableVersionList.xlPivotTableVersion14) 
        PivotTable = PivotCache.CreatePivotTable(TableDestination="'For allocation entry'!R3C1", TableName="PivotTable1", DefaultVersion=win32c.PivotTableVersionList.xlPivotTableVersion14)
            #####logger.info("Adding particular Row in Pivot Table") 
        PivotTable.PivotFields('Location').Orientation = win32c.PivotFieldOrientation.xlRowField
        PivotTable.PivotFields('Total AR').Orientation = win32c.PivotFieldOrientation.xlDataField
            # PivotTable.PivotFields('Sum of Net').NumberFormat= '0.00'
        
        #logic for adding sum in Eligible sheet last
        last_row4 = ws5.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row
        last_row5=last_row4+5
        column_list = ws5.range("A1").expand('right').value
        Balance_no_column=column_list.index('Balance')+1
        Balance_letter_column = num_to_col_letters(column_list.index('Balance')+1)
        ws5.range(f"{Balance_letter_column}{last_row5}").value=f'=SUM({Balance_letter_column}2:{Balance_letter_column}{last_row4})'
        ws5.range(f"{Balance_letter_column}{last_row5}").api.NumberFormat= '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'
        ws5.range(f"{Balance_letter_column}{last_row5}").api.Font.Bold = True

        #logic for hidding extra columns
        list7=['Diff','c','c1','d','d1']
        for items in list7:
                letter_column = num_to_col_letters(column_list.index(f'{items}')+1)
                ws5.api.Range(f"{letter_column}1").EntireColumn.Hidden=True
        #logic for adding sum in PIVOT BB last
        last_row3 = ws6.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row 
        last_row3+=5
        last_column = ws6.range(f'A{last_row2}').end('right').last_cell.column
        last_column_letter=num_to_col_letters(last_column)
        ws6.range(f"{last_column_letter}{last_row3}").value=f'=GETPIVOTDATA("Sum of Balance",$A$3)+GETPIVOTDATA("Sum of Balance",$A${last_row2})'
        ws6.range(f"{last_column_letter}{last_row3}").api.NumberFormat= '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'
        ws6.range(f"{last_column_letter}{last_row3}").api.Font.Bold = True
        ws6.activate()
        time.sleep(2)
        wb.save(f"{output_location}\\Open AR _"+input_date+' - Production.xlsx')
        wb.app.quit()
        return f"Open AR Report generated for {input_date}"
    except Exception as e:
        raise e
    finally:
        try:
            wb.app.quit()
        except:
            pass 
    

def open_ap(input_date, output_date):
    try:
        input_sheet = r'J:\WEST PLAINS\REPORT\Open AP\Raw files'+f'\\Open AP _{input_date}.xlsx' 
        output_location = r'J:\WEST PLAINS\REPORT\Open AP\Output files'
        if not os.path.exists(input_sheet):
            return(f"{input_sheet} Excel file not present for date {input_date}")
        #logger.info("Opening operating workbook instance of excel")
        retry=0
        while retry < 10:
            try:
                wb=xw.Book(input_sheet)
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry ==10:
                    raise e
        #logger.info("Adding sheet to the same workbook")
        wb.sheets.add("EXCLUDING",after=wb.sheets[f"Open AP _{input_date}"]) 
        #logger.info("Accessing Particular WorkBook ")
        ws1=wb.sheets[f"Open AP _{input_date}"]
        ws2=wb.sheets["EXCLUDING"]
        #logger.info("Declaring Variables for columns and rows")
        last_row = ws1.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row
        column_list = ws1.range("A1").expand('right').value
        Vendor_no_column=column_list.index('Vendor')+1
        Vendor_letter_column = num_to_col_letters(column_list.index('Vendor')+1)
        Location_no_column=column_list.index('Location')+1
        Location_letter_column = num_to_col_letters(column_list.index('Location')+1)
        #logger.info("Applying Filter to the same workbook")
        ws1.api.Range(f"{Vendor_letter_column}1").AutoFilter(Field:=f'{Vendor_no_column}', Criteria1:=["<>MACQUARIE COMMODITIES (USA) INC."], Operator:=1,Criteria2=["<>INTER-COMPANY PURCH/SALES"])
        ws1.api.Range(f"{Location_letter_column}1").AutoFilter(Field:=f'{Location_no_column}', Criteria1:=["<>WPMEXICO"], Operator:=1)
        #logger.info("Copying and pasting Worksheet")
        ws1.api.AutoFilter.Range.Copy()
        ws2.api.Paste()
        #logger.info("Renaming the worksheet")
        ws2.name='Excl Macq & IC'
        #logger.info("Applying autofit")
        ws2.autofit()
        ws2=wb.sheets['Excl Macq & IC']

        #logger.info("Declaring Variables for columns and rows")
        column_list = ws1.range("A1").expand('right').value
        Vendor_no_column=column_list.index('Vendor')+1
        Vendor_letter_column = num_to_col_letters(column_list.index('Vendor')+1)

        #logger.info("Copying Inter Company Data from inp sheet  to Intercompany Sheet")
        try:
            ws1.api.AutoFilterMode=False
            ws1.api.Range(f"{Vendor_letter_column}1").AutoFilter(Feild:=Vendor_no_column,Criteria1:="INTER-COMPANY PURCH/SALES") #Removing Intercompany
            intcomp_sht = wb.sheets.add("Intercompany", after=ws1)

            ws1.api.AutoFilter.Range.Copy()
            time.sleep(1)
            intcomp_sht.range("A1").api.Select()
            while True:
                try:
                    intcomp_sht.api.Paste()
                    break
                except:
                    time.sleep(1)
            wb.app.api.CutCopyMode=False
            time.sleep(1)
            ws1.api.AutoFilterMode=False
        except Exception as e:
            print("No (INTER-COMPANY PURCH/SALES) Present ")
            print(e)
        #logger.info("Adding Worksheet for Pivot Table")
        wb.sheets.add("Pivot BB",after=ws2)
        ws3=wb.sheets["Pivot BB"]
        #logger.info("Creating Pivot table")
        #logger.info("Declaring Variables for columns and rows")
        last_row = ws2.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row
        last_column = ws2.range('A1').end('right').last_cell.column
        last_column_letter=num_to_col_letters(ws2.range('A1').end('right').last_cell.column)
        num_row=ws3.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row
        num_row+=2
        PivotCache=wb.api.PivotCaches().Create(SourceType=win32c.PivotTableSourceType.xlDatabase, SourceData=f"\'Excl Macq & IC\'!R1C1:R{last_row}C{last_column}", Version=win32c.PivotTableVersionList.xlPivotTableVersion14)
        PivotTable = PivotCache.CreatePivotTable(TableDestination=f"'Pivot BB'!R{num_row}C1", TableName="Pivot", DefaultVersion=win32c.PivotTableVersionList.xlPivotTableVersion14)
        #logger.info("Adding particular Row in Pivot Table")
        PivotTable.PivotFields('Location').Orientation = win32c.PivotFieldOrientation.xlRowField
        PivotTable.PivotFields('Location').Position = 1
        #logger.info("Adding particular Column in Pivot Table")
        PivotTable.PivotFields('Journal Source').Orientation = win32c.PivotFieldOrientation.xlColumnField
        PivotTable.PivotFields('Journal Source').Position = 1
        #logger.info("Adding particular Data Field in Pivot Table")
        PivotTable.PivotFields('Invoice Balance').Orientation = win32c.PivotFieldOrientation.xlDataField
        #logger.info("Changing No Format in Pivot Table")
        PivotTable.PivotFields('Sum of Invoice Balance').NumberFormat= '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'
        
        #logger.info("Adding Worksheet for Pivot Table")
        wb.sheets.add("Product Settlement",after=ws3)
        # wb.save()
        ws4=wb.sheets["Product Settlement"]
        #logger.info("Applying Filter to the same workbook")
        ws2.api.Range(f"{Location_letter_column}1").AutoFilter(Field:=f'{Location_no_column}', Criteria1:=["<>WP CORP"], Operator:=1)
        #logger.info("Copying and pasting Worksheet")
        ws2.api.AutoFilter.Range.Copy()
        ws4.api.Paste()   
        #logger.info("Creating Pivot table")
        num_row=ws3.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row
        num_row+=7
        PivotCache=wb.api.PivotCaches().Create(SourceType=win32c.PivotTableSourceType.xlDatabase, SourceData=f"\'Product Settlement\'!R1C1:R{last_row}C{last_column}", Version=win32c.PivotTableVersionList.xlPivotTableVersion14)
        PivotTable = PivotCache.CreatePivotTable(TableDestination=f"'Pivot BB'!R{num_row}C1", TableName="Product Settlement", DefaultVersion=win32c.PivotTableVersionList.xlPivotTableVersion14)
        #logger.info("Adding particular Row in Pivot Table")
        PivotTable.PivotFields('Location').Orientation = win32c.PivotFieldOrientation.xlRowField
        PivotTable.PivotFields('Location').Position = 1
        #logger.info("Adding particular Column in Pivot Table")
        PivotTable.PivotFields('Journal Source').Orientation = win32c.PivotFieldOrientation.xlColumnField
        PivotTable.PivotFields('Journal Source').Position = 1
        #logger.info("Adding particular Data Field in Pivot Table")
        PivotTable.PivotFields('Invoice Balance').Orientation = win32c.PivotFieldOrientation.xlDataField
        #logger.info("Changing No Format in Pivot Table")
        PivotTable.PivotFields('Sum of Invoice Balance').NumberFormat= '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'

        #logger.info("Applying filters in Pivot Table")
        try:
            PivotTable.PivotFields('Journal Source').PivotItems('Accrual Invoice Matching').Visible= False
        except Exception as e:
            pass  
        try:  
            PivotTable.PivotFields('Journal Source').PivotItems('Accrual Invoice Matching Reversal').Visible= False
        except Exception as e:
            pass
        try:    
            PivotTable.PivotFields('Journal Source').PivotItems('(blank)').Visible= False
        except Exception as e:
            pass
        #Heading for 2nd Pivot Table
        num_row2= num_row-2  
        ws3.range(f"A{num_row2}").value="Product/Grain Settlement"
        ws3.range(f"A{num_row2}").api.Font.Bold = True  

        wb.sheets.add("FREIGHT",after=ws4)
        ws5=wb.sheets["FREIGHT"]
        # wb.save()
        ws2.api.Range(f"{Location_letter_column}1").AutoFilter(Field:=f'{Location_no_column}', Criteria1:=["<>WP CORP"], Operator:=1)
        ws2.api.AutoFilter.Range.Copy()
        ws5.api.Paste()   
        #logger.info("Creating Pivot table")
        num_row=ws3.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row
        num_row+=7
        PivotCache=wb.api.PivotCaches().Create(SourceType=win32c.PivotTableSourceType.xlDatabase, SourceData=f"\'FREIGHT\'!R1C1:R{last_row}C{last_column}", Version=win32c.PivotTableVersionList.xlPivotTableVersion14)
        PivotTable = PivotCache.CreatePivotTable(TableDestination=f"'Pivot BB'!R{num_row}C1", TableName="FREIGHT", DefaultVersion=win32c.PivotTableVersionList.xlPivotTableVersion14)
        #logger.info("Adding particular Row in Pivot Table")
        PivotTable.PivotFields('Location').Orientation = win32c.PivotFieldOrientation.xlRowField
        PivotTable.PivotFields('Location').Position = 1
        #logger.info("Adding particular Column in Pivot Table")
        PivotTable.PivotFields('Journal Source').Orientation = win32c.PivotFieldOrientation.xlColumnField
        PivotTable.PivotFields('Journal Source').Position = 1
        #logger.info("Adding particular Data Field in Pivot Table")
        PivotTable.PivotFields('Invoice Balance').Orientation = win32c.PivotFieldOrientation.xlDataField
        PivotTable.PivotFields('Sum of Invoice Balance').NumberFormat= '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'

        #logger.info("Applying filters in Pivot Table")
        try:
            PivotTable.PivotFields('Journal Source').PivotItems('Credit Memos').Visible= False
        except Exception as e:
            pass
        try:
            PivotTable.PivotFields('Journal Source').PivotItems('Debit Memos').Visible= False
        except Exception as e:
            pass    
        try:
            PivotTable.PivotFields('Journal Source').PivotItems('Final Settlement - Purchase').Visible= False
        except Exception as e:
            pass    
        try:
            PivotTable.PivotFields('Journal Source').PivotItems('(blank)').Visible= False
        except Exception as e:
            pass    
        num_row2= num_row-2  
        ws3.range(f"A{num_row2}").value="Freight"
        ws3.range(f"A{num_row2}").api.Font.Bold = True  
        #logger.info("Applying autofit in Pivot_table sheet")
        ws3.autofit()
        #logger.info("Deleting unneccessary worksheets")
        ws5.delete()
        ws4.delete()
        #logger.info("Removing filters applied")
        ws1.api.AutoFilterMode=False
        ws2.api.AutoFilterMode=False

        last_col_num = ws1.range('A1').expand('right').last_cell.column 
        # last_col = num_to_col_letters(last_col_num) 
        last_row = ws1.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row 
        ####logger.info("Adding Worksheet for Pivot Table") 
        wb.sheets.add("For allocation entry",before=ws1) 
        ####logger.info("Creating Pivot table") 
        PivotCache=wb.api.PivotCaches().Create(SourceType=win32c.PivotTableSourceType.xlDatabase, SourceData=f'\'{ws1.name}\'!R1C1:R{last_row}C{last_col_num}', Version=win32c.PivotTableVersionList.xlPivotTableVersion14) 
        PivotTable = PivotCache.CreatePivotTable(TableDestination="'For allocation entry'!R3C1", TableName="PivotTable1", DefaultVersion=win32c.PivotTableVersionList.xlPivotTableVersion14)
         ####logger.info("Adding particular Row in Pivot Table") 
        PivotTable.PivotFields('Location').Orientation = win32c.PivotFieldOrientation.xlRowField
        PivotTable.PivotFields('Invoice Balance').Orientation = win32c.PivotFieldOrientation.xlDataField
         # PivotTable.PivotFields('Sum of Net').NumberFormat= '0.00'

        #logger.info("Saving current worksheet")

        ws3.activate()
        time.sleep(2)
        wb.save(f"{output_location}\\Open AP _{input_date}.xlsx")
        #logger.info("quiting the current instance of excel app")
        wb.app.quit()
        return f"Open AP Report Generated for {input_date}"
    except Exception as e:
        raise e
    finally:
        try:
            wb.app.quit()
        except:
            pass

def unsetteled_payables(input_date, output_date):
    try:
        input_xl = r'J:\WEST PLAINS\REPORT\Unsettled Payables\Raw Files\Unsettled Payables _'+input_date+".xlsx"
        output_location = r'J:\WEST PLAINS\REPORT\Unsettled Payables\Output files\Unsettled Payables _'+input_date+".xlsx"
        
        if not os.path.exists(input_xl):
            return(f"Excel file not present for date {input_date}")

        prev_month = datetime.strftime(datetime.strptime(input_date, "%m.%d.%Y"), "%B")
        
        
        retry=0
        while retry<10:
            try:
                
                wb = xw.Book(input_xl, update_links=True)
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry==10:
                    raise e
        
        #######logger.info("Sheet Opened")
        # time.sleep(10)
        while True:
            try:
                # inp_sht = wb.sheets[0]
                inp_sht = wb.sheets[f"Unsettled Payables _{input_date}"]
                break
            except:
                time.sleep(10)
        

        # inp_sht.range('AB2').formula = '=O2+Q2'
        

        column_list = inp_sht.range("A1").expand('right').value
        vendor_column = num_to_col_letters(column_list.index('Customer/Vendor Name')+1)
        vendor_column_num = column_list.index('Customer/Vendor Name')+1
        locationId_column = num_to_col_letters(column_list.index('Location Id')+1)
        locationId_column_num = column_list.index('Location Id')+1
        

        inp_sht.api.AutoFilterMode=False
        #######logger.info("Removing  MACQUARIE COMMODITIES (USA) INC. and all INTER-COMPANY PURCH/SALES vendor")
        inp_sht.api.Range(f"{vendor_column}1").AutoFilter(Feild:=vendor_column_num,Criteria1:="<>MACQUARIE COMMODITIES (USA) INC", Operator:=1, Criteria2:="<>INTER-COMPANY PURCH/SALES") #Removing macquarie and intercompany
        #######logger.info("Removing WPMEXICO Location ID")
        inp_sht.api.Range(f"{locationId_column}1").AutoFilter(Feild:=locationId_column_num,Criteria1:="<>WPMEXICO", Operator:=7) #Removing WPMEXICO
        #######logger.info("Creating Excl IC & Macq and pasting data")

        
        exc_sht = wb.sheets.add("Excl Macq & IC", after=inp_sht)
        inp_sht.api.AutoFilter.Range.Copy()
        time.sleep(1)
        exc_sht.api.Select()
        exc_sht.range("A1").api.Select()
        while True:
            try:
                exc_sht.api.Paste()
                break
            except:
                time.sleep(1)

        wb.app.api.CutCopyMode=False

        #######logger.info("Copying Inter Company Data from inp sheet  to Intercompany Sheet")
        inp_sht.api.AutoFilterMode=False
        inp_sht.api.Range(f"{vendor_column}1").AutoFilter(Feild:=vendor_column_num,Criteria1:="INTER-COMPANY PURCH/SALES") #Removing macquarie and intercompany
        
        intcomp_sht = wb.sheets.add("Intercompany", after=exc_sht)

        inp_sht.api.AutoFilter.Range.Copy()
        time.sleep(1)
        intcomp_sht.range("A1").api.Select()
        while True:
            try:
                intcomp_sht.api.Paste()
                break
            except:
                time.sleep(1)
        wb.app.api.CutCopyMode=False
        inp_sht.api.AutoFilterMode=False


        last_col = exc_sht.range("A1").expand("right").last_cell.column
        last_row = exc_sht.range(f'A'+ str(exc_sht.cells.last_cell.row)).end('up').row

        #######logger.info("Adding Worksheet for Pivot Table")
        wb.sheets.add("Pivot BB",after=wb.sheets['Intercompany'])
        #######logger.info("Creating Pivot table")
        PivotCache=wb.api.PivotCaches().Create(SourceType=win32c.PivotTableSourceType.xlDatabase, SourceData=f'\'Excl Macq & IC\'!R1C1:R{last_row}C{last_col}', Version=win32c.PivotTableVersionList.xlPivotTableVersion14)
        PivotTable = PivotCache.CreatePivotTable(TableDestination="'Pivot BB'!R3C1", TableName="PivotTable1", DefaultVersion=win32c.PivotTableVersionList.xlPivotTableVersion14)
        #######logger.info("Adding particular Row in Pivot Table")
        PivotTable.PivotFields('Location Id').Orientation = win32c.PivotFieldOrientation.xlRowField
        
        
        #######logger.info("Adding particular Column in Pivot Table")
        PivotTable.PivotFields('Settlement').Orientation = win32c.PivotFieldOrientation.xlDataField
        PivotTable.PivotFields('Sum of Settlement').NumberFormat= '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'
        PivotTable.PivotFields('Advance').Orientation = win32c.PivotFieldOrientation.xlDataField
        PivotTable.PivotFields('Sum of Advance').NumberFormat= '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'
        PivotTable.PivotFields('Net').Orientation = win32c.PivotFieldOrientation.xlDataField
        PivotTable.PivotFields('Sum of Net').NumberFormat= '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'
        

        #######logger.info("Changing Layout")
        PivotTable.RowAxisLayout(1)
        PivotTable.TableStyle2 = ""


        last_col_num = inp_sht.range('A1').expand('right').last_cell.column
        # last_col = num_to_col_letters(last_col_num)
        last_row = inp_sht.range(f'A'+ str(inp_sht.cells.last_cell.row)).end('up').row
        ######logger.info("Adding Worksheet for Pivot Table")
        wb.sheets.add("For allocation entry",before=inp_sht)
        ######logger.info("Creating Pivot table")
        PivotCache=wb.api.PivotCaches().Create(SourceType=win32c.PivotTableSourceType.xlDatabase, SourceData=f'\'{inp_sht.name}\'!R1C1:R{last_row}C{last_col_num}', Version=win32c.PivotTableVersionList.xlPivotTableVersion14)
        PivotTable = PivotCache.CreatePivotTable(TableDestination="'For allocation entry'!R3C1", TableName="PivotTable1", DefaultVersion=win32c.PivotTableVersionList.xlPivotTableVersion14)
        ######logger.info("Adding particular Row in Pivot Table")
        PivotTable.PivotFields('Location Name').Orientation = win32c.PivotFieldOrientation.xlRowField
        
        PivotTable.PivotFields('Net').Orientation = win32c.PivotFieldOrientation.xlDataField
        # PivotTable.PivotFields('Sum of Net').NumberFormat= '0.00'

        wb.save(output_location)
        wb.app.quit()
        return f"Unsettled Receivables report Generated for {input_date}"
    except Exception as e:
        raise e
    finally:
        try:
            wb.app.quit()
        except:
            pass
    pass

def unsetteled_receivables(input_date, output_date):
    try:
        

        input_xl = r'J:\WEST PLAINS\REPORT\Unsettled Receivables\Raw Files\Unsettled Receivables _'+input_date+".xlsx"
        prev_output_location = r'J:\WEST PLAINS\REPORT\Unsettled Receivables\Output files\Unsettled Receivables _'+output_date+".xlsx"
        output_location = r'J:\WEST PLAINS\REPORT\Unsettled Receivables\Output files\Unsettled Receivables _'+input_date+".xlsx"
        
        if not os.path.exists(input_xl):
            return(f"{input_xl} Excel file not present for date {input_date}")

        if not os.path.exists(prev_output_location):
            return(f"{prev_output_location} Excel file not present for date {input_date}")

        # if not os.path.exists(output_location):
        #     return(f"{output_location} Excel file not present for date {input_date}")
        
        
        prev_month = datetime.strftime(datetime.strptime(input_date, "%m.%d.%Y"), "%B")
        
        
        retry=0
        while retry<10:
            try:
                
                wb = xw.Book(input_xl, update_links=False)
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry==10:
                    raise e
        
        ######logger.info("Sheet Opened")
        # time.sleep(10)
        while True:
            try:
                inp_sht = wb.sheets[0] #wb.sheets[0].name in 'Unsettled Receivables _'+input_date
                break
            except Exception as e:
                time.sleep(5)
                # retry+=1
                # if retry==10:
                #     raise e
        

        # inp_sht.range('AB2').formula = '=O2+Q2'
        

        column_list = inp_sht.range("A1").expand('right').value
        vendor_column = num_to_col_letters(column_list.index('Customer/Vendor Name')+1)
        vendor_column_num = column_list.index('Customer/Vendor Name')+1
        locationId_column = num_to_col_letters(column_list.index('Location Id')+1)
        locationId_column_num = column_list.index('Location Id')+1
        

        
        inp_sht.api.AutoFilterMode=False
        ######logger.info("Removing  MACQUARIE COMMODITIES (USA) INC. and all INTER-COMPANY PURCH/SALES vendor")
        inp_sht.api.Range(f"{vendor_column}1").AutoFilter(Feild:=vendor_column_num,Criteria1:="<>MACQUARIE COMMODITIES (USA) INC", Operator:=1, Criteria2:="<>INTER-COMPANY PURCH/SALES") #Removing macquarie and intercompany
        ######logger.info("Removing WPMEXICO Location ID")
        inp_sht.api.Range(f"{locationId_column}1").AutoFilter(Feild:=locationId_column_num,Criteria1:="<>WPMEXICO", Operator:=7) #Removing WPMEXICO
        ######logger.info("Creating Excl IC & Macq and pasting data")

        
        exc_sht = wb.sheets.add("Excl Macq & IC", after=inp_sht)
        inp_sht.api.AutoFilter.Range.Copy()
        time.sleep(1)
        exc_sht.api.Select()
        exc_sht.range("A1").api.Select()
        while True:
            try:
                exc_sht.api.Paste()
                break
            except:
                time.sleep(1)

        wb.app.api.CutCopyMode=False

        ######logger.info("Copying Inter Company Data from inp sheet  to Intercompany Sheet")
        inp_sht.api.AutoFilterMode=False
        inp_sht.api.Range(f"{vendor_column}1").AutoFilter(Feild:=vendor_column_num,Criteria1:="INTER-COMPANY PURCH/SALES") #Removing macquarie and intercompany
        
        intcomp_sht = wb.sheets.add("Intercompany", after=exc_sht)

        inp_sht.api.AutoFilter.Range.Copy()
        time.sleep(1)
        intcomp_sht.range("A1").api.Select()
        while True:
            try:
                intcomp_sht.api.Paste()
                break
            except:
                time.sleep(1)
        wb.app.api.CutCopyMode=False
        inp_sht.api.AutoFilterMode=False
        
        ######logger.info("Copying Tier data from latest output files")

        # file_list = glob.glob(output_location+"\\*.xlsx")
        # file_list.sort()
        # latest_output = file_list[-1]
        # if input_date == output_date:
        #     ######logger.info(f"current selected output file date is date passed in argument")
        #     latest_output = file_list[-1]
        #     ######logger.info(f"Now file name is {output_date}")

        tier_wb = xw.Book(prev_output_location, update_links=False)

        tier_sht = tier_wb.sheets("Tier")
        # ######logger.info("Copy tier sheet AFTER the intercompany sheet of input book.")
        tier_sht.api.Copy(None, After=intcomp_sht.api)

        # ######logger.info("Copy paste new columns from prev output files")
        prev_exc_sht = tier_wb.sheets("Excl Macq & IC")

        n_col_list = prev_exc_sht.range("A1").expand('right').value
        start_col = num_to_col_letters(n_col_list.index('Tier')+1)
        end_col = num_to_col_letters(n_col_list.index('Payment Indicator'))
        n_col_list = n_col_list[n_col_list.index('Tier'):n_col_list.index('Payment Indicator')]
        # formula_list = prev_exc_sht.range(f"{start_col}2:{end_col}2").formulas

        next_col = num_to_col_letters(column_list.index('Net')+2)
        next_col_num = column_list.index('Net')+2
        last_col_num = column_list.index('Net')+1 #Changed from 2 to 1 for picking correct last column
        
        n_col_list[1] = datetime.strptime(input_date, "%m.%d.%Y").date()

        for col in n_col_list[::-1]: #For properly inserting data
            last_col_num+=1
            exc_sht.range(f"{next_col}:{next_col}").insert()
            exc_sht.range(f"{next_col}1").value = f"{col}"
            
        
        last_col = num_to_col_letters(last_col_num)
        last_row = exc_sht.range(f'A'+ str(exc_sht.cells.last_cell.row)).end('up').row

        # exc_sht.range(f"{next_col}2").options(transpose=True).formula = list(prev_exc_sht.range(f"{start_col}2:{end_col}2").formula)
        prev_exc_sht.range(f"{start_col}2:{end_col}2").copy(exc_sht.range(f"{next_col}2:{last_col}{last_row}"))
        ######logger.info("Handling tier2 formula")
        exc_sht.range(f"{next_col}2").formula = prev_exc_sht.range(f"{next_col}2").formula
        exc_sht.range(f"{next_col}2").copy(exc_sht.range(f"{next_col}2:{next_col}{last_row}"))

        exc_sht.api.AutoFilterMode=False
        exc_sht.api.Range(f"{next_col}1").AutoFilter(Feild:=next_col_num ,Criteria1:="#N/A")

        test_sht = wb.sheets.add("Test_Tier", after=exc_sht)

        exc_sht.api.AutoFilter.Range.Copy()
        time.sleep(1)
        test_sht.range("A1").api.Select()
        while True:
            try:
                test_sht.api.Paste()
                break
            except:
                time.sleep(1)
        wb.app.api.CutCopyMode=False

        if type(test_sht.range(f"{vendor_column}2").expand("down").value)==list:
            new_tiew_names = list(set(test_sht.range(f"{vendor_column}2").expand("down").value))
        else:
            new_tiew_names = test_sht.range(f"{vendor_column}2").expand("down").value
        this_tier_sht = wb.sheets("Tier")
        last_row_tier = this_tier_sht.range(f'A'+ str(this_tier_sht.cells.last_cell.row)).end('up').row
        this_tier_sht.range(f"A{last_row_tier+1}").options(transpose=True).value = new_tiew_names
        if type(test_sht.range(f"{vendor_column}2").expand("down").value)==list:
            this_tier_sht.range(f"B{last_row_tier}").copy(this_tier_sht.range(f"B{last_row_tier+1}:B{last_row_tier+len(new_tiew_names)}"))
        else:
            this_tier_sht.range(f"B{last_row_tier}").copy(this_tier_sht.range(f"B{last_row_tier+1}:B{last_row_tier+1}"))
        test_sht.delete()
        exc_sht.api.AutoFilterMode=False

        ######logger.info("Adding Worksheet for Pivot Table")
        wb.sheets.add("Pivot BB",after=wb.sheets['Tier'])
        ######logger.info("Creating Pivot table")
        PivotCache=wb.api.PivotCaches().Create(SourceType=win32c.PivotTableSourceType.xlDatabase, SourceData=f'\'Excl Macq & IC\'!R1C1:R{last_row}C{last_col_num}', Version=win32c.PivotTableVersionList.xlPivotTableVersion14)
        PivotTable = PivotCache.CreatePivotTable(TableDestination="'Pivot BB'!R3C1", TableName="PivotTable1", DefaultVersion=win32c.PivotTableVersionList.xlPivotTableVersion14)
        ######logger.info("Adding particular Row in Pivot Table")
        PivotTable.PivotFields('Tier').Orientation = win32c.PivotFieldOrientation.xlRowField
        # PivotTable.PivotFields('Tier').Position = 1
        PivotTable.PivotFields('Customer/Vendor Name').Orientation = win32c.PivotFieldOrientation.xlRowField
        # PivotTable.PivotFields('Customer/Vendor Name').Position = 1
        
        ######logger.info("Adding particular Column in Pivot Table")
        PivotTable.PivotFields('Net').Orientation = win32c.PivotFieldOrientation.xlDataField
        PivotTable.PivotFields('Sum of Net').NumberFormat= '0.00'
        PivotTable.PivotFields('Ticket Date <=30 Days').Orientation = win32c.PivotFieldOrientation.xlDataField
        PivotTable.PivotFields('Sum of Ticket Date <=30 Days').NumberFormat= '0.00'
        PivotTable.PivotFields('Ticket Date 31-60 Days').Orientation = win32c.PivotFieldOrientation.xlDataField
        PivotTable.PivotFields('Sum of Ticket Date 31-60 Days').NumberFormat= '0.00'
        PivotTable.PivotFields('Ticket Date +60 Days').Orientation = win32c.PivotFieldOrientation.xlDataField
        PivotTable.PivotFields('Sum of Ticket Date +60 Days').NumberFormat= '0.00'

        ######logger.info("Changing Layout")
        PivotTable.RowAxisLayout(1)
        PivotTable.TableStyle2 = ""


        last_col_num = inp_sht.range('A1').expand('right').last_cell.column
        # last_col = num_to_col_letters(last_col_num)
        last_row = inp_sht.range(f'A'+ str(inp_sht.cells.last_cell.row)).end('up').row
        ######logger.info("Adding Worksheet for Pivot Table")
        wb.sheets.add("For allocation entry",before=inp_sht)
        ######logger.info("Creating Pivot table")
        PivotCache=wb.api.PivotCaches().Create(SourceType=win32c.PivotTableSourceType.xlDatabase, SourceData=f'\'{inp_sht.name}\'!R1C1:R{last_row}C{last_col_num}', Version=win32c.PivotTableVersionList.xlPivotTableVersion14)
        PivotTable = PivotCache.CreatePivotTable(TableDestination="'For allocation entry'!R3C1", TableName="PivotTable1", DefaultVersion=win32c.PivotTableVersionList.xlPivotTableVersion14)
        ######logger.info("Adding particular Row in Pivot Table")
        PivotTable.PivotFields('Location Name').Orientation = win32c.PivotFieldOrientation.xlRowField
        
        PivotTable.PivotFields('Net').Orientation = win32c.PivotFieldOrientation.xlDataField
        # PivotTable.PivotFields('Sum of Net').NumberFormat= '0.00'
        
  
        print()

        wb.save(output_location)
        wb.app.quit()
        return f"Unsettled Receivables report Generated for {input_date}"
    except Exception as e:
        raise e
    finally:
        try:
            wb.app.quit()
        except:
            pass

def moc_interest_alloc(input_date, output_date):
    try:
        input_xl = r"J:\WEST PLAINS\REPORT\MOC Interest allocation\Raw files\Inventory MTM Excel Report " + input_date + ".xlsx"
        if not os.path.exists(input_xl):
                return(f"{input_xl} Excel file not present for date {input_date}")
        dt = datetime.strptime(input_date,"%m.%d.%Y")
        mtm_input_date = dt.strftime("%B %Y")

        mtm_file = r"J:\WEST PLAINS\REPORT\MOC Interest allocation\Raw files\Inventory MTM Excel Report " + mtm_input_date +'.xlsx'

        if not os.path.exists(mtm_file):
                return(f"{mtm_file} Excel file not present for date {input_date}")

        unsettled_recev_file = r'J:\WEST PLAINS\REPORT\Unsettled Receivables\Output files\Unsettled Receivables _'+input_date+'.xlsx'

        if not os.path.exists(unsettled_recev_file):
                return(f"{unsettled_recev_file} Excel file not present for date {input_date}")

        unsettled_pay_file = r'J:\WEST PLAINS\REPORT\Unsettled Payables\Output files\Unsettled Payables _'+input_date+'.xlsx'

        if not os.path.exists(unsettled_pay_file):
                return(f"{unsettled_pay_file} Excel file not present for date {input_date}")

        open_ar_file = r'J:\WEST PLAINS\REPORT\Open AR\Output files\Open AR _'+input_date+' - Production.xlsx'

        if not os.path.exists(open_ar_file):
                return(f"{open_ar_file} Excel file not present for date {input_date}")

        open_ap_file = r'J:\WEST PLAINS\REPORT\Open AP\Output files\Open AP _'+input_date+'.xlsx'

        if not os.path.exists(open_ap_file):
                return(f"{open_ap_file} Excel file not present for date {input_date}")

        


        output_dir = r"J:\WEST PLAINS\REPORT\MOC Interest allocation\Output Files"
        template_dir = r"J:\WEST PLAINS\REPORT\MOC Interest allocation\Raw files\template"


          
        main_df = moc_get_df_from_input_excel(input_xl, mtm_file, open_ap_file, open_ar_file,unsettled_pay_file, unsettled_recev_file)
        update_moc_excel(main_df, template_dir, output_dir, input_date)
        return "MOC Interest Allocation Report Generated"
    except Exception as e:
        raise e
    




def main():
    def on_closing():
        if messagebox.askokcancel("Quit", "Do you want to quit?"):
            
            root.destroy()
            sys.exit()
    def callback_2():
        if submit_text.get() != "Started" and 'Select' not in Rep_variable.get():
            submit_text.set("Started")
            text_box.delete(1.0, "end")
            text_box.tag_configure("center", justify='center')
            text_box.tag_add("center", 6.0, "end")
            text_box.insert("end", f"In Process", "center")
            root.update()
            
            print(inp_date.get())
            print(Rep_variable.get())
            input_date = inp_date.get()
            output_date = out_date.get()
            func_to_call = Rep_variable.get()
            msg = wp_job_ids[func_to_call](input_date, output_date)
            text_box.delete(1.0, "end")
            text_box.insert("end", f"\n{msg}", "center")
            submit_text.set("Submit")
            Rep_variable.set('Select')
            root.update()

            print()
        
        elif 'Select' in Rep_variable.get():
            text_box.insert("end", f"Please select job first", "center")

       
        root.update()
    def callback():
        threading.Thread(target=callback_2).start()
        
    def report_callback_exception(self, exc, val, tb):
        showerror("Error", message=str(exc + val +tb))

    Tk.report_callback_exception = report_callback_exception
    frame_title.grid(row=0, column=1,pady=(24,0),columnspan=3, padx=(30,0))
    logo = PhotoImage(file = path + '\\'+'wp_logo.png')


    title = Label(frame_title,image=logo, background='white')
    # title = Label(frame_title, text="Revelio Report Generator", font=("Algerian", 28), foreground='black', background="white")
    title.grid(column=1,row=0)

    root.protocol("WM_DELETE_WINDOW", on_closing)
    # input_date=None
    # output_date = None
    frame_options.grid(row=1,column=0, pady=30, padx=35, columnspan=2, rowspan=3)
    wp_job_ids = {'ABS':1,'BBR':bbr,'CPR Report':cpr, 'Freight analysis':freight_analysis, 'CTM combined':ctm,'Inventory MTM excel report summary':inv_mtm,
                    'MOC Interest Allocation':moc_interest_alloc,'Open AR':open_ar,'Open AP':open_ap, 'Unsettled Payable Report':unsetteled_payables,'Unsettled Receivable Report':unsetteled_receivables}
    # department_ids = {'Select \t\t\t\t\t\t\t\t\t': 9, 'Ethanol\t\t\t\t\t\t\t\t': 1, 'WestPlains': 8}
    Rep_variable = StringVar()
    doc_type_variable = StringVar()
    doc_type_variable.set('Select')
    folderPath = StringVar()
    macroPath = StringVar()
    # Dep_variable.trace('w', update_options_B)
    dep_label = Label(frame_options, text='Select Job:                  ', font=("Book Antiqua bold", 16), foreground="#ff8c00", background="white")
    dep_label.grid(row=0, column=0)
    Dep_option = OptionMenu(frame_options, Rep_variable, *wp_job_ids.keys())
    
    Dep_option["menu"].configure(background="white", font=("Arial", 12)) #, bg='#20bebe', fg='white')
    # Dep_option["menu"].config(width=19)
    Dep_option.grid(row=0, column=1)
    Rep_variable.set('Select \t\t\t\t\t\t\t\t\t')
    blank = Label(frame_options, text="                                ", font=("Helvetica", 16), foreground="blue", justify='left', background="white")
    blank.grid(row=1, column=0)
    blank2 = Label(frame_options, text="             ", font=("Helvetica", 16), foreground="green", justify='left', background="white")
    blank2.grid(row=1, column=1)
    # doc_label = Label(frame_options, text="Select Doc_Type:     ", font=("Book Antiqua bold", 16), foreground="#ff8c00", background="white")
    # doc_label.grid(row=2, column=0)
    # doc_type_option = OptionMenu(frame_options, doc_type_variable, '')
    # doc_type_option["menu"].configure(background="white", font=("Arial", 12))
    # doc_type_option.grid(row=2, column=1)

    blank3 = Label(frame_options, text="                                ", font=("Helvetica", 16), foreground="blue", justify='left', background="white")
    blank3.grid(row=3, column=0)
    folder_label = Label(frame_options, text="Select Input Date:     ", font=("Book Antiqua bold", 16), foreground="#ff8c00", background="white",justify='left')
    folder_label.grid(row=4, column=0)
    browse_text = StringVar()
    inp_date = MyDateEntry(master=frame_options, width=17, selectmode='day') #Button(frame_options, textvariable=browse_text, command=getFolderPath, font = ("Book Antiqua bold", 12), bg="#20bebe", fg="white", height=1, width=14, activebackground="#20bebb")
    browse_text.set("Browse")
    inp_date.grid(row=4, column=1)

    blank4 = Label(frame_options, text="                                ", font=("Helvetica", 16), foreground="blue", justify='left', background="white")
    blank4.grid(row=5, column=0)
    macro_label = Label(frame_options, text="Select Prev File Date:", font=("Book Antiqua bold", 16), foreground="#ff8c00", background="white",justify='left')
    macro_label.grid(row=6, column=0)
    browse_text2 = StringVar()
    out_date = MyDateEntry(master=frame_options, width=17, selectmode='day') #Button(frame_options, textvariable=browse_text2, command=getFilePath, font = ("Book Antiqua bold", 12), bg="#20bebe", fg="white", height=1, width=14, activebackground="#20bebb")
    browse_text2.set("Browse")
    out_date.grid(row=6, column=1)

    frame_folder.grid(row=2, column=2, padx=(28,0))
    

    frame_submit.grid(row=5, column=1,columnspan=3)
    submit_text = StringVar()
    submit = Button(frame_submit, textvariable=submit_text, font = ("Book Antiqua bold", 12), bg="#20bebe", fg="white", height=1, width=14, command=callback, activebackground="#20bebb")
    submit.grid(row=0, column=1, padx=(30,0))
    submit_text.set("Submit")
    
    # if doc_type_variable.get() == "Select \t\t\t\t\t\t\t\t\t":
    #     sel_Folder["state"] = "disabled"
    #     submit["state"] = "disabled"
        

    # text_box = Text(root, height=10, width=50, padx=15, pady=15)
    # text_box.insert(1.0, "Select Details, and click Select folder n Submit")
    # text_box.tag_configure("center", justify="center")
    # text_box.tag_add("center", 1.0), "end"
    # text_box.grid(column=1, row=6)
    blank3 = Label(frame_submit, text="             ", font=("Helvetica", 16), foreground="green", justify='left', background="white")
    blank3.grid(row=1, column=1)
    
    
    staus_text = StringVar()
    frame_msg.grid(row=7,column=1,columnspan=3) ##, padx=(180,0))
    text_box = Text(frame_msg, background="white",font=("Raleway", 10), width=88, height=10, borderwidth=0)

    # label_2 = Label(root, textvariable=label_2_text, background="white", justify='center',font=("Raleway", 12)) 
    text_box.grid(row=7, column=1,columnspan=3, padx=(14,0)) # column
    # label_2.grid(row=8, column=1,columnspan=2)
    # 
    # label_2_text.set("")

    root.mainloop()


if __name__ == '__main__':
    main()

