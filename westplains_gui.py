import re
import os
import PyPDF2
import shutil
import calendar
import requests
import traceback
import glob, time
import numpy as np
import pandas as pd
import xlwings as xw
import sys, traceback
from tkinter import Button
from tabula import read_pdf
from tkinter.ttk import Label
from tkinter import PhotoImage
from selenium import webdriver
from tkinter import messagebox
from tkcalendar import DateEntry
from tkinter import Label as label
import xlwings.constants as win32c
from tkinter.ttk import OptionMenu
from collections import defaultdict
from collections import OrderedDict
from tkinter.ttk import Frame, Style
from tkinter.messagebox import showerror,showwarning
from tkinter import N, Tk, StringVar, Text
from selenium.webdriver.common.by import By
from datetime import date, datetime, timedelta
from dateutil.relativedelta import relativedelta
from selenium.webdriver.firefox.options import Options
from webdriver_manager.firefox import GeckoDriverManager
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


# path = r'C:\Users\imam.khan\OneDrive - BioUrja Trading LLC\Documents\Revelio'

today = datetime.strftime(date.today(), format = "%d%m%Y")

# drive = r"I:\RPTWP_Swati"
# drive = r"J:\WEST PLAINS"
drive = r"\\biourja.local\biourja\India Sync\WEST PLAINS"

path = drive+r'\REPORT\Westplains_gui'


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

def set_borders(border_range):
    for border_id in range(7,13):
        border_range.api.Borders(border_id).LineStyle=1
        border_range.api.Borders(border_id).Weight=2


def xlOpner(inputFile):
    try:
        retry = 0
        while retry<10:
            try:
                input_wb = xw.Book(inputFile, update_links=False)
                return input_wb
            except Exception as e:
                time.sleep(2)
                retry+=1
                if retry==9:
                    raise e
    except Exception as e:
        print(f"Exception caught in xlOpner :{e}")
        raise e


def row_range_calc(filter_col:str, input_sht,wb):

    sp_lst_row = input_sht.range(f'{filter_col}'+ str(input_sht.cells.last_cell.row)).end('up').row
    sp_address= input_sht.api.Range(f"{filter_col}2:{filter_col}{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Address
    sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]        
    row_range = sorted([int(i) for i in list(set(re.findall("\d+",sp_address.replace("$",""))))])

    while row_range[-1]!=sp_lst_row:
        sp_lst_row = input_sht.range(f'{filter_col}'+ str(input_sht.cells.last_cell.row)).end('up').row
        sp_address = sp_address+','+(input_sht.api.Range(f"{filter_col}{row_range[-1]+1}:{filter_col}{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Address)
        # sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]        
        row_range.extend(sorted([int(i) for i in list(set(re.findall("\d+",sp_address.replace("$",""))))]))

    sp_address = sp_address.replace("$","").split(",")
    init_list= [list(range(int(i.split(":")[0]), int(i.split(":")[1])+1)) for i in sp_address]
    sublist = []
    flat_list = [item for sublist in init_list for item in sublist]
    return flat_list, sp_lst_row,sp_address


def insert_all_borders(cellrange:str,working_sheet,working_workbook):
        working_sheet.api.Range(cellrange).Select()
        working_workbook.app.selection.api.Borders(win32c.BordersIndex.xlDiagonalDown).LineStyle = win32c.Constants.xlNone
        working_workbook.app.selection.api.Borders(win32c.BordersIndex.xlDiagonalUp).LineStyle = win32c.Constants.xlNone
        linestylevalues=[win32c.BordersIndex.xlEdgeLeft,win32c.BordersIndex.xlEdgeTop,win32c.BordersIndex.xlEdgeBottom,win32c.BordersIndex.xlEdgeRight,win32c.BordersIndex.xlInsideVertical,win32c.BordersIndex.xlInsideHorizontal]
        for values in linestylevalues:
            a=working_workbook.app.selection.api.Borders(values)
            a.LineStyle = win32c.LineStyle.xlContinuous
            a.ColorIndex = 0
            a.TintAndShade = 0
            a.Weight = win32c.BorderWeight.xlThin


def freezepanes_for_tab(cellrange:str,working_sheet,working_workbook):
    try:
        working_sheet.activate()
        working_sheet.api.Rows(cellrange).Select()
        working_workbook.app.api.ActiveWindow.FreezePanes = True
    except Exception as e:
        raise e        

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

def interior_coloring_temp(colour_value,cellrange:str,working_sheet,working_workbook):
    try:
        working_sheet.activate()
        working_sheet.api.Range(cellrange).Select()
        a = working_workbook.app.selection.api.Interior
        a.Pattern = win32c.Constants.xlSolid
        a.PatternColorIndex = win32c.Constants.xlAutomatic
        a.Color = colour_value
        a.TintAndShade = 0
        a.PatternTintAndShade = 0        
    except Exception as e:
        raise e  

def interior_coloring(colour_value,cellrange:str,working_sheet,working_workbook):
    try:
        working_sheet.activate()
        if working_sheet.api.AutoFilterMode:
            working_sheet.api.Range(cellrange).SpecialCells(win32c.CellType.xlCellTypeVisible).Select()
        else:
            working_sheet.api.Range(cellrange).Select()
        a = working_workbook.app.selection.api.Interior
        a.Pattern = win32c.Constants.xlSolid
        a.PatternColorIndex = win32c.Constants.xlAutomatic
        a.Color = colour_value
        a.TintAndShade = 0
        a.PatternTintAndShade = 0        
    except Exception as e:
        raise e  


def interior_coloring_by_theme(pattern_tns,tintandshade,colour_value,cellrange:str,working_sheet,working_workbook):
    try:
        working_sheet.activate()
        if working_sheet.api.AutoFilterMode:
            working_sheet.api.Range(cellrange).SpecialCells(win32c.CellType.xlCellTypeVisible).Select()
        else:
            working_sheet.api.Range(cellrange).Select()
        a = working_workbook.app.selection.api.Interior
        # a.Pattern = win32c.Constants.xlSolid
        a.PatternColorIndex = win32c.Constants.xlAutomatic
        a.ThemeColor = colour_value
        a.TintAndShade = tintandshade
        a.PatternTintAndShade = pattern_tns    
    except Exception as e:
        raise e  


def mtd_new_trades(input_date, output_date):
    try:       
        job_name = 'MTD New Trades Automation'

        ##### Defining paths and checking availability of files ###########
        input_sheet= r'J:\WEST PLAINS\REPORT\MTD Trades\Input'+f'\\Sale Open Close_{input_date}.xlsx'
        input_sheet2 = r'J:\WEST PLAINS\REPORT\MTD Trades\Input'+f'\\Purchase Open Close_{input_date}.xlsx'
        ctm_book = r'J:\WEST PLAINS\REPORT\CTM Combined report\Output files'+f'\\CTM Combined _{input_date}.xlsx'
        template_sheet= r'J:\WEST PLAINS\REPORT\MTD Trades\Template'+f'\\MTD New Trades Templets.xlsx'
        mapping_workbook = r'J:\WEST PLAINS\REPORT\MTD Trades\Template'+f'\\Market Zone Mapping.xlsx'
        pricing_workbook = r'J:\WEST PLAINS\REPORT\MTD Trades\Active Pricing'+f'\\2023  Weekly and Monthly Values.xlsx'
        if not os.path.exists(input_sheet):
            return(f"{input_sheet} Excel file not present for date {input_date}")
        if not os.path.exists(input_sheet2):
            return(f"{input_sheet2} Excel file not present") 
        if not os.path.exists(template_sheet):
            return(f"{template_sheet} Excel file not present")  
        if not os.path.exists(ctm_book):
            return(f"{ctm_book} Excel file not present")           

        #####opening required workbooks ###########
        """
        1.Opening Template workbook
        2.Opening Sales workbook
        3.Opening Purchase workbook
        4.CTM workbook for same date
        """
        template_wb = xlOpner(template_sheet)

        sale_wb = xlOpner(input_sheet)

        purchase_wb = xlOpner(input_sheet2)

        ctm_wb = xlOpner(ctm_book)

        ###########Clearing GS QUERY SHEET ############
        template_wb.activate()
        gs_sheet=template_wb.sheets[f"GS QUERY"]
        gs_sheet.activate()
        last_rw = gs_sheet.api.UsedRange.Rows.Count
        gs_sheet.range(f"A2:BQ{last_rw}").clear()
        gs_sheet.range(f"BR3:BV{last_rw}").clear()

        ###########Copying FROM Sales sheet ==> GS QUERY SHEET ############
        sale_wb.activate()
        sale_sheet=sale_wb.sheets[0]
        sale_last_rw = sale_sheet.api.UsedRange.Rows.Count
        last_column_letter_s=num_to_col_letters(sale_sheet.range('A1').end('right').last_cell.column)
        sale_sheet.range(f"A2:{last_column_letter_s}{sale_last_rw}").copy(gs_sheet.range(f"A2"))

        ###########Copying FROM Purchase sheet ==> GS QUERY SHEET ############
        purchase_wb.activate()
        purchase_sheet=purchase_wb.sheets[0]
        purchase_last_rw = purchase_sheet.api.UsedRange.Rows.Count
        last_column_letter_p=num_to_col_letters(purchase_sheet.range('A1').end('right').last_cell.column)
        purchase_sheet.range(f"A2:{last_column_letter_p}{purchase_last_rw}").copy(gs_sheet.range(f"A{sale_last_rw+1}"))

        ########### GS QUERY SHEET Vlookup dragdown############
        gs_sheet.activate()
        gs_sheet.range(f'BR2:BV{sale_last_rw + purchase_last_rw - 1}').select()
        template_wb.app.selection.api.FillDown()
        template_wb.api.ActiveSheet.Outline.ShowLevels(RowLevels:=0,ColumnLevels:=1)

        ########### Copying data from CTM ==> MTM - CONTRACTS REPORT SHEET ############
        template_wb.activate()
        mtm_sheet=template_wb.sheets[f"MTM - CONTRACTS REPORT"]
        mtm_sheet.activate()
        lst_mtm = mtm_sheet.api.UsedRange.Rows.Count
        mtm_sheet.range(f"A3:AQ{lst_mtm}").clear()

        ctm_wb.activate()
        ctm_sheet=ctm_wb.sheets[0]
        ctm_sheet.activate()
        ctm_last = ctm_sheet.api.UsedRange.Rows.Count
        lst_clm = num_to_col_letters(ctm_sheet.api.UsedRange.Columns.Count)
        ctm_sheet.range(f"A2:{lst_clm}{ctm_last}").copy(mtm_sheet.range(f"B2"))

        mtm_sheet.activate()
        mtm_sheet.range(f'A2:A{ctm_last}').select()
        template_wb.app.selection.api.FillDown()
        template_wb.api.ActiveSheet.Outline.ShowLevels(RowLevels:=0,ColumnLevels:=1)        
        print(f"{mtm_sheet.name} processed successfully")
        
        ############### END DATE LOOKUP SHEET ############
        template_wb.activate()
        date_sheet=template_wb.sheets[f"END DATE LOOKUP"]
        date_sheet.activate()        

        column_list = gs_sheet.range("A1").expand('right').value
        edate_cl_no = column_list.index('End Date')+1
        pivotCount = template_wb.api.ActiveSheet.PivotTables().Count
         # 'INPUT DATA'!$A$3:$I$86
        for j in range(1, pivotCount+1):
            # if template_wb.api.ActiveSheet.PivotTables(j).PivotCache().SourceData != f"'INPUT DATA'!R3C1:R{num_row}C{num_col}": #Updateing data source
            template_wb.api.ActiveSheet.PivotTables(j).PivotCache().SourceData = f"'GS QUERY'!R1C{edate_cl_no}:R{sale_last_rw + purchase_last_rw - 1}C{edate_cl_no}" #Updateing data source
            template_wb.api.ActiveSheet.PivotTables(j).PivotCache().Refresh()
        date_sheet.range(f"B:B").clear()

        ########### date logic ######################
        dte_lst_rw = date_sheet.range(f'A'+ str(date_sheet.cells.last_cell.row)).end('up').row
        date_values = date_sheet.range(f"A2:A{dte_lst_rw -1}").value
        for index,dat_val in enumerate(date_values):
            if type(dat_val) == datetime:
                if dat_val.replace(day=1) < datetime.strptime(input_date,"%m.%d.%Y").replace(day=1):
                    print("found datetime date")
                    date_sheet.range(f"B{index+2}").value = f"Delinquent"
                else:
                    print("found datetime date")
                    mnt = datetime.strftime(dat_val,"%b")
                    yr = datetime.strftime(dat_val,"%y")
                    date_sheet.range(f"B{index+2}").value = f"'{mnt}-{yr}" 
            elif type(dat_val) == str:
                print(f"found str value format unknown :: row no {index+2}")
            else:
                print(f"unable to detect format check row no {index+2}")    

        ####################            #############################
        gs_sheet.activate()
        mpb_cl_no = column_list.index('Market Price Basis')+1
        mpb_letter_column = num_to_col_letters(mpb_cl_no)
        try:
            gs_sheet.api.Range(f"{mpb_letter_column}1").AutoFilter(Field:=f'{mpb_cl_no}', Criteria1:=["#N/A"], Operator:=7)
        except:
            print("no na values found")    

        mapping_wb = xlOpner(mapping_workbook)

        pricing_wb = xlOpner(pricing_workbook)
        pricing_sheet = pricing_wb.sheets[0]
        loc_sht=mapping_wb.sheets["Market Zone Mapping"]
        tlast_row = loc_sht.range(f'F'+ str(loc_sht.cells.last_cell.row)).end('up').row
        loc_dict=loc_sht.range(f'E1:F{tlast_row}').options(dict).value

        commodity_sht=mapping_wb.sheets["Commodity Mapping"]
        tlast_row2 = commodity_sht.range(f'A'+ str(commodity_sht.cells.last_cell.row)).end('up').row
        commodity_dict=commodity_sht.range(f'A1:B{tlast_row2}').options(dict).value

        option_mnt_sht=mapping_wb.sheets["Option Month mapping IT"]
        tlast_row3 = option_mnt_sht.range(f'A'+ str(option_mnt_sht.cells.last_cell.row)).end('up').row
        option_mnt_df=option_mnt_sht.range(f'A1:B{tlast_row3}').options(pd.DataFrame,header=False,index=False).value
        option_mnt_dict = {}
        for country in option_mnt_df[0].unique():
            option_mnt_dict[country] = option_mnt_df[option_mnt_df[0] == country][1].tolist()
                
        flat_list, sp_lst_row,sp_address = row_range_calc("A",gs_sheet,template_wb)
        str1 = ''
        for row in flat_list:
            gs_sheet.activate()
            if gs_sheet.range(f'AO{row}').value == 'IN STORE':
                gs_sheet.range(f'BU{row}').value = 0
                interior_coloring_temp(5287936,f'BU{row}',gs_sheet,template_wb)
                continue
            market_zone = loc_dict[gs_sheet.range(f'AO{row}').value]
            commodity = commodity_dict[gs_sheet.range(f'I{row}').value]
            if gs_sheet.range(f'Y{row}').value !=None:
                text = (re.findall(r'(\d+|\D+)', gs_sheet.range(f'Y{row}').value))
            else:
                str1+=f' - Option month not found for row:{row}\n'
                print(f"Option month value not found for row :: {row}")
                interior_coloring_temp(49407,f'BU{row}',gs_sheet,template_wb)
                continue   
            year = "2" + text[1]
            mnth = option_mnt_dict[text[0]]
            price_flag = False
            for months in mnth:
                ori_mnth = datetime.strftime(datetime.strptime(months,"%B"),"%b")
                matching_mnth = ori_mnth + f"-{year}"
                pricing_sheet.activate()
                pricing_sheet.api.Range("A1").Select()
                pricing_sheet.range('A:A').api.Cells.Find(What:=commodity, After:=pricing_sheet.api.Application.ActiveCell,LookIn:=win32c.FindLookIn.xlFormulas,
                                    LookAt:=win32c.LookAt.xlPart, SearchOrder:=win32c.SearchOrder.xlByRows, SearchDirection:=win32c.SearchDirection.xlNext).Activate()
                first_row = pricing_sheet.api.Application.ActiveCell.Row  
                end_row = pricing_sheet.range(f"A{int(row)}").expand('down').last_cell.row 
                pricing_sheet.api.Range(f"B{first_row}").Activate()

                pricing_sheet.api.Cells.Find(What:=market_zone, After:=pricing_sheet.api.Application.ActiveCell,LookIn:=win32c.FindLookIn.xlFormulas,
                                    LookAt:=win32c.LookAt.xlWhole, SearchOrder:=win32c.SearchOrder.xlByRows, SearchDirection:=win32c.SearchDirection.xlNext).Activate()                        

                price_row = pricing_sheet.api.Application.ActiveCell.Row

                pricing_sheet.api.Range(f"B{first_row}").Activate()
                pricing_sheet.api.Cells.Find(What:=matching_mnth, After:=pricing_sheet.api.Application.ActiveCell,LookIn:=win32c.FindLookIn.xlFormulas,
                                    LookAt:=win32c.LookAt.xlPart, SearchOrder:=win32c.SearchOrder.xlByColumns, SearchDirection:=win32c.SearchDirection.xlNext).Activate() 
                price_column = num_to_col_letters(pricing_sheet.api.Application.ActiveCell.Column)    
                if pricing_sheet.range(f'{price_column}{price_row}').value !=None:
                    actual_price = pricing_sheet.range(f'{price_column}{price_row}').value
                    gs_sheet.activate()
                    gs_sheet.range(f'BU{row}').value = actual_price
                    ######.Color = 15773696 ###########
                    interior_coloring_temp(15773696,f'BU{row}',gs_sheet,template_wb)
                    gs_sheet.api.Range(f"BU{row}").Font.Bold = True
                    price_flag = True
                    break
                else:
                    pass

            if price_flag == False:    
                    print(f"price not found for months ==> {mnth} ::commodity {commodity} ::{market_zone} :: reffer output row {row}")
                    str1+= f" - Price not found for months ==> {mnth} ::commodity => {commodity} :: market_zone => {market_zone} :: reffer row(GS Query) =>{row}\n"
                    interior_coloring_temp(255,f'BU{row}',gs_sheet,template_wb)
                    
        ################# REMOVING INSTORE FROM MAKET-ZONE  ##############
        gs_sheet.api.AutoFilterMode=False
        curr_col_list=gs_sheet.range("A1").expand('right').value
        market_zone_index=curr_col_list.index("Market Zone")
        market_zone_letter=num_to_col_letters(market_zone_index+1)
        commodity_index=curr_col_list.index("Commodity")
        commodity_letter=num_to_col_letters(commodity_index+1)
        last_rw = gs_sheet.range(f'A'+ str(gs_sheet.cells.last_cell.row)).end('up').row
        for row in range(last_rw, 1, -1):
            cell_value = gs_sheet[market_zone_letter + str(row)].value
            cell_value2 = gs_sheet[commodity_letter + str(row)].value
            if cell_value=="IN STORE":
                gs_sheet.range('A' + str(row)).api.EntireRow.Delete()
                print(f"value found for :: {row}, value :: {cell_value}")
                continue
            elif cell_value2=="EQUIP" or cell_value2=="FRT TON":
                gs_sheet.range('A' + str(row)).api.EntireRow.Delete()  
                print(f"value found for :: {row}, value :: {cell_value2}")
                continue              
            else:
                print(f"value not found for :: {row}")
        #####################################################################
        tablist={gs_sheet:255,mtm_sheet:49407,date_sheet:win32c.ThemeColor.xlThemeColorLight2}
        for tab,color in tablist.items():
            freezepanes_for_tab(cellrange="2:2",working_sheet=tab,working_workbook=template_wb) 
            try:
                tab.api.Tab.Color = color
            except:
                tab.api.Tab.ThemeColor =color
            tab.api.AutoFilterMode=False    

        output_location = r'J:\WEST PLAINS\REPORT\MTD Trades\Output'   
        template_wb.save(f"{output_location}\\MTD New Trades_"+input_date+'.xlsx')
        template_wb.app.quit()
        return[f"{job_name} Report for {input_date} generated succesfully", str1]

    except Exception as e:
        raise e
    finally:
        try:
            template_wb.app.quit()
        except:
            pass


def ar_exposure(input_date, output_date):
    try:
        # drive= r"C:\Users\amanullah.khan\OneDrive - BioUrja Trading LLC\p drive backup\projects\westplains_gui\AR Exposure Files"      
        job_name = 'ar_exposure_automation'
        input_sheet2 = drive+r'\REPORT\AR EXPOSURE\Output'+f'\\Open AR_{output_date}.xlsx'
        input_sheet= drive+r'\REPORT\AR EXPOSURE\Output'+f'\\Unsettled AR_{output_date}.xlsx'
        ap_ar_template = drive+r'\REPORT\AR EXPOSURE'+f'\\WPLLC - AP_AR_Template.xlsx'
        ar_exposure_template = drive+r'\REPORT\AR EXPOSURE'+f'\\WPLLC - AR Exposure_Template.xlsm'
        previous_sheet_ar_axposure = drive+r'\REPORT\AR EXPOSURE\Output'+f'\\AR Exposure {output_date.replace(".","")}.xlsm'
        output_location = drive+r'\REPORT\AR EXPOSURE\Ar_exposure_reports' 
        if not os.path.exists(input_sheet):
            return(f"{input_sheet} Excel file not present for date {output_date}")
        if not os.path.exists(input_sheet2):
            return(f"{input_sheet2} Excel file not present for date {output_date}")   
        if not os.path.exists(ap_ar_template):
            return(f"{ap_ar_template} Excel file not present")     
        if not os.path.exists(ar_exposure_template):
            return(f"{ar_exposure_template} Excel file not present")
        if not os.path.exists(previous_sheet_ar_axposure):
            return(f"{previous_sheet_ar_axposure} Excel file not present")                  
        retry=0
        while retry < 10:
            try:
                wb = xw.Book(input_sheet,update_links=False) 
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry ==10:
                    raise e 
        input_tab=wb.sheets[f"Unsettled AR_{output_date}"]
        # column_list = input_tab.range("A1").expand('right').value        

        retry=0
        while retry < 10:
            try:
                wb_open_ar = xw.Book(input_sheet2,update_links=False) 
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry ==10:
                    raise e 

        # # sheet1=wb_open_ar.sheets[0].name
        input_tab2=wb_open_ar.sheets[f"Open AR_{output_date}"]            
        retry=0
        while retry < 10:
            try:
                wb_ap_ar = xw.Book(ap_ar_template,update_links=False) 
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry ==10:
                    raise e 

        ws2p2=wb_ap_ar.sheets["Pivot AR summary new"] 
        ws2p2.activate()
        wlast_row = ws2p2.range(f'A'+ str(ws2p2.cells.last_cell.row)).end('up').row
        retry=0
        while retry < 10:
            try:
                ar_exposure_wb = xw.Book(ar_exposure_template,update_links=False) 
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry ==10:
                    raise e 
        credit_over_utilized_tab=ar_exposure_wb.sheets["Credit Over Utilized"]
        credit_over_utilized_tab.activate()
        coulast_row = credit_over_utilized_tab.range(f'A'+ str(credit_over_utilized_tab.cells.last_cell.row)).end('up').row
        credit_over_utilized_tab.api.Range(f"A3:H{coulast_row}").Delete()
        ws2p2.range(f"A2:E{wlast_row-1}").copy(credit_over_utilized_tab.range(f"A2"))
        ulast_row = credit_over_utilized_tab.range(f'A'+ str(credit_over_utilized_tab.cells.last_cell.row)).end('up').row
        credit_over_utilized_tab.api.Range(f"F2:H{ulast_row}").Select()
        ar_exposure_wb.app.api.Selection.FillDown()
        credit_over_utilized_tab.range(f"D2:E{ulast_row}").number_format='_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'
        exposure_ar_tab = ar_exposure_wb.sheets["Open AR"]
        exposure_ar_tab.clear()
        input_tab.activate()

        exposure_un_ar_tab = ar_exposure_wb.sheets["Unsettled AR"]
        exposure_un_ar_tab.clear()
        last_row_it = input_tab.range(f'A'+ str(input_tab.cells.last_cell.row)).end('up').row
        last_column_letter_it=num_to_col_letters(input_tab.range('A1').end('right').last_cell.column)
        # input_tab.range(f"A1:{last_column_letter_it}{last_row_it}").copy(exposure_ar_tab.range(f"A1"))
        input_tab.range(f"A1:{last_column_letter_it}{last_row_it}").copy(exposure_un_ar_tab.range(f"A1"))

        # exposure_un_ar_tab = ar_exposure_wb.sheets["Unsettled AR"]
        input_tab2.activate()
        last_row_it = input_tab2.range(f'A'+ str(input_tab2.cells.last_cell.row)).end('up').row
        last_column_letter_it=num_to_col_letters(input_tab2.range('A1').end('right').last_cell.column)
        # input_tab2.range(f"A1:{last_column_letter_it}{last_row_it}").copy(exposure_un_ar_tab.range(f"A1"))
        input_tab2.range(f"A1:{last_column_letter_it}{last_row_it}").copy(exposure_ar_tab.range(f"A1"))

        over_limit_tab = ar_exposure_wb.sheets["Over Limit"]
        over_limit_tab.activate()
        over_limit_tab.clear()
        credit_over_utilized_tab.activate()
        # last_row_cou = credit_over_utilized_tab.range(f'A'+ str(credit_over_utilized_tab.cells.last_cell.row)).end('up').row
        credit_over_utilized_tab.api.Range(f"B:H").Copy()
        over_limit_tab.api.Range(f"A1")._PasteSpecial(Paste=win32c.PasteType.xlPasteAllUsingSourceTheme,Operation=win32c.Constants.xlNone)
        over_limit_tab.api.Range(f"A1")._PasteSpecial(Paste=-4163,Operation=win32c.Constants.xlNone)
        over_limit_tab.range(f"C:G").number_format="0.00"
        last_row_ol = over_limit_tab.range(f'G'+ str(over_limit_tab.cells.last_cell.row)).end('up').row
        over_limit_tab.activate()
        over_limit_tab.range(f"A2:G{last_row_ol}").api.Sort(Key1=over_limit_tab.range(f"G2:G{last_row_ol}").api,Order1=win32c.SortOrder.xlAscending,DataOption1=win32c.SortDataOption.xlSortNormal,Orientation=1,SortMethod=1)
        over_limit_tab.range(f"C:G").number_format="_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)" 
        unapplied_credit_tab = ar_exposure_wb.sheets["Unapplied Credit"]
        unapplied_credit_tab.activate()
        unapplied_credit_tab.clear()
        
        unsettled_receivables_tab = ar_exposure_wb.sheets["Unsettled Receivables"]
        unsettled_receivables_tab.activate()
        unsettled_receivables_tab.clear()        
        
        credit_over_utilized_tab.activate()
        last_row_cou = credit_over_utilized_tab.range(f'A'+ str(credit_over_utilized_tab.cells.last_cell.row)).end('up').row
        credit_over_utilized_tab.api.Range(f"A1:H{last_row_cou}").AutoFilter(Field:=5, Criteria1:=["<0"], Operator:=1) 
        credit_over_utilized_tab.api.Range(f"A1:H{last_row_cou}").AutoFilter(Field:=6, Criteria1:=["<0"], Operator:=1) 

        sp_lst_row = credit_over_utilized_tab.range(f'A'+ str(credit_over_utilized_tab.cells.last_cell.row)).end('up').row
        sp_address= credit_over_utilized_tab.api.Range(f"A2:A{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Address
        sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]
    

        credit_over_utilized_tab.api.Range(f"A1:H{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Copy()
        unapplied_credit_tab.api.Range(f"A1")._PasteSpecial(Paste=win32c.PasteType.xlPasteAllUsingSourceTheme,Operation=win32c.Constants.xlNone)
        unapplied_credit_tab.autofit()
        unapplied_credit_tab.activate()

        slst_row = unapplied_credit_tab.range(f'A'+ str(unapplied_credit_tab.cells.last_cell.row)).end('up').row
        unapplied_credit_tab.api.Range(f"A1:H{slst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Copy()
        unsettled_receivables_tab.api.Range(f"A1")._PasteSpecial(Paste=win32c.PasteType.xlPasteAllUsingSourceTheme,Operation=win32c.Constants.xlNone)

        unsettled_receivables_tab.activate()
        last_row_urt = unsettled_receivables_tab.range(f'A'+ str(unsettled_receivables_tab.cells.last_cell.row)).end('up').row
        unsettled_receivables_tab.range(f"A2:H{last_row_urt}").api.Sort(Key1=unsettled_receivables_tab.range(f"D2:D{last_row_urt}").api,Order1=win32c.SortOrder.xlAscending,DataOption1=win32c.SortDataOption.xlSortNormal,Orientation=1,SortMethod=1)
        unsettled_receivables_tab.range(f"D:H").number_format="0.00"

        unsettled_receivables_tab.api.Range(f"A1:H{last_row_urt}").AutoFilter(Field:=4, Criteria1:=["0.00"]) 
        last_row_urtf = unsettled_receivables_tab.range(f'A'+ str(unsettled_receivables_tab.cells.last_cell.row)).end('up').row
        unsettled_receivables_tab.api.Range(f"A2:H{last_row_urtf}").SpecialCells(win32c.CellType.xlCellTypeVisible).Delete(Shift=win32c.Direction.xlUp)
        unsettled_receivables_tab.api.Range(f"A1:H{last_row_urt}").AutoFilter(Field:=4) 
        last_row_urt_u = unsettled_receivables_tab.range(f'A'+ str(unsettled_receivables_tab.cells.last_cell.row)).end('up').row
        unsettled_receivables_tab.range(f"A2:H{last_row_urt_u}").api.Sort(Key1=unsettled_receivables_tab.range(f"D2:D{last_row_urt_u}").api,Order1=win32c.SortOrder.xlDescending,DataOption1=win32c.SortDataOption.xlSortNormal,Orientation=1,SortMethod=1)
        unsettled_receivables_tab.range(f"D:H").number_format="_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"

        insert_all_borders(cellrange=f"A2:A{last_row_urt_u}",working_sheet=unsettled_receivables_tab,working_workbook=ar_exposure_wb)

        credit_over_utilized_tab.activate()
        credit_over_utilized_tab.api.AutoFilterMode=False

        watch_accounts_tab = ar_exposure_wb.sheets["Watch Accounts"]
        credit_over_utilized_tab.api.Range(f"A:H").Copy()
        watch_accounts_tab.api.Range(f"A1")._PasteSpecial(Paste=win32c.PasteType.xlPasteAllUsingSourceTheme,Operation=win32c.Constants.xlNone)
        watch_accounts_tab.api.AutoFilterMode=False
        watch_accounts_tab.activate()
        i_last_row_wa = watch_accounts_tab.range(f'I'+ str(watch_accounts_tab.cells.last_cell.row)).end('up').row
        watch_accounts_tab.api.Range(f"I3:I{i_last_row_wa}").Delete(Shift=win32c.Direction.xlUp)
        last_row_wa = watch_accounts_tab.range(f'A'+ str(watch_accounts_tab.cells.last_cell.row)).end('up').row

        watch_accounts_tab.api.Range(f"I2:I{last_row_wa}").Select()
        ar_exposure_wb.app.api.Selection.FillDown()

        watch_accounts_tab.range(f"A2:I{last_row_wa}").api.Sort(Key1=watch_accounts_tab.range(f"I2:I{last_row_wa}").api,Order1=win32c.SortOrder.xlAscending,DataOption1=win32c.SortDataOption.xlSortNormal,Orientation=1,SortMethod=1)
        watch_accounts_tab.api.Range(f"A1:I{last_row_wa}").AutoFilter(Field:=9, Criteria1:=["<=69%"],Operator:=1) 

        last_row_waf = watch_accounts_tab.range(f'A'+ str(watch_accounts_tab.cells.last_cell.row)).end('up').row
        watch_accounts_tab.api.Range(f"A2:I{last_row_waf}").SpecialCells(win32c.CellType.xlCellTypeVisible).Delete(Shift=win32c.Direction.xlUp)
        watch_accounts_tab.api.AutoFilterMode=False

        watch_accounts_tab.api.Range(f"A1:I{last_row_wa}").AutoFilter(Field:=9, Criteria1:=["#DIV/0!"],Operator:=1)
        last_row_waf = watch_accounts_tab.range(f'A'+ str(watch_accounts_tab.cells.last_cell.row)).end('up').row
        watch_accounts_tab.api.Range(f"A2:I{last_row_waf}").SpecialCells(win32c.CellType.xlCellTypeVisible).Delete(Shift=win32c.Direction.xlUp)

        watch_accounts_tab.api.AutoFilterMode=False
        watch_accounts_tab.range(f"A2:I{last_row_wa}").api.Sort(Key1=watch_accounts_tab.range(f"I2:I{last_row_wa}").api,Order1=win32c.SortOrder.xlDescending,DataOption1=win32c.SortDataOption.xlSortNormal,Orientation=1,SortMethod=1)
    
        last_row_waf = watch_accounts_tab.range(f'A'+ str(watch_accounts_tab.cells.last_cell.row)).end('up').row
        insert_all_borders(cellrange=f"A2:A{last_row_waf}",working_sheet=watch_accounts_tab,working_workbook=ar_exposure_wb)

        credit_over_utilized_tab.activate()
        exposure_ar_tab.api.Move(Before=ar_exposure_wb.api.Sheets(1))
        exposure_ar_tab.api.Range("1:1").Font.Bold = True

        exposure_un_ar_tab.api.Move(Before=ar_exposure_wb.api.Sheets(4))
        exposure_un_ar_tab.api.Range("1:1").Font.Bold = True

        unsettled_ar_loc_tab = ar_exposure_wb.sheets["Unsettled AR Loc"]
        unsettled_ar_loc_tab.api.Move(Before=ar_exposure_wb.api.Sheets(5))
        unsettled_ar_loc_tab.api.Range("1:1").Font.Bold = True

        unsettled_ar_loc_tab.api.Tab.ThemeColor = win32c.ThemeColor.xlThemeColorAccent2

        exposure_un_ar_tab.api.Tab.ThemeColor = win32c.ThemeColor.xlThemeColorAccent2

        unsettled_receivables_tab.api.Tab.ThemeColor = win32c.ThemeColor.xlThemeColorAccent2

        exposure_ar_tab.api.Tab.Color = 49407

        unsettled_receivables_tab.name = "Unsettled AR Summary"
        exposure_un_ar_tab.name = "Unsettled AR Detail"

        unapplied_credit_tab.api.Move(Before=ar_exposure_wb.api.Sheets(6))
        unapplied_credit_tab.api.Tab.ThemeColor = win32c.ThemeColor.xlThemeColorLight2
        unapplied_credit_tab.api.Range("1:1").Font.Bold = True

        unsettled_ar_ic = ar_exposure_wb.sheets["Unsettled AR IC"]
        unsettled_ar_ic.name = "Unsettled AR-Interco"
        unsettled_ar_ic.api.Range("1:1").Font.Bold = True
        unsettled_ar_ic.api.Tab.Color = 255


        tablist=[exposure_ar_tab,over_limit_tab,unsettled_receivables_tab,exposure_un_ar_tab,unapplied_credit_tab,watch_accounts_tab,unsettled_ar_ic,credit_over_utilized_tab]
        for tab in tablist:
            freezepanes_for_tab(cellrange="2:2",working_sheet=tab,working_workbook=ar_exposure_wb) 

        over_limit_tab.activate()
        over_limit_tab.range(f"A2:G{last_row_ol}").api.Sort(Key1=over_limit_tab.range(f"G2:G{last_row_ol}").api,Order1=win32c.SortOrder.xlDescending,DataOption1=win32c.SortDataOption.xlSortNormal,Orientation=1,SortMethod=1)
        
        interior_coloring(colour_value=65535,cellrange=f"G1:G{last_row_cou}",working_sheet=over_limit_tab,working_workbook=ar_exposure_wb)

        credit_over_utilized_tab.activate()
        credit_over_utilized_tab.range(f"A2:H{last_row_cou}").api.Sort(Key1=credit_over_utilized_tab.range(f"D2:D{last_row_cou}").api,Order1=win32c.SortOrder.xlDescending,DataOption1=win32c.SortDataOption.xlSortNormal,Orientation=1,SortMethod=1)
        
        unsettled_receivables_tab.activate()
        unsettled_receivables_tab.api.Range(f"A2:H{last_row_urt_u}").Delete()
        
        credit_over_utilized_tab.activate() 
        credit_over_utilized_tab.api.Range(f"A1:H{last_row_cou}").AutoFilter(Field:=4, Criteria1:=[">0"]) 

        sp_lst_row = credit_over_utilized_tab.range(f'A'+ str(credit_over_utilized_tab.cells.last_cell.row)).end('up').row
        sp_address= credit_over_utilized_tab.api.Range(f"A2:A{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Address
        sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]

        credit_over_utilized_tab.api.Range(f"A1:H{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Copy()
        unsettled_receivables_tab.api.Range(f"A1")._PasteSpecial(Paste=win32c.PasteType.xlPasteAllUsingSourceTheme,Operation=win32c.Constants.xlNone) 
        unsettled_receivables_tab.api.Range(f"A1")._PasteSpecial(Paste=-4163,Operation=win32c.Constants.xlNone)     

        credit_over_utilized_tab.api.AutoFilterMode=False
        ar_exposure_wb.app.api.CutCopyMode=False

        exposure_un_ar_tab.activate()
        exposure_un_ar_tab.api.Range(f"D:D").Delete(win32c.DeleteShiftDirection.xlShiftToLeft)
        exposure_un_ar_tab.api.Range(f"E:E").Delete(win32c.DeleteShiftDirection.xlShiftToLeft)
        exposure_un_ar_tab.api.Range(f"F:F").Delete(win32c.DeleteShiftDirection.xlShiftToLeft)

        last_row_un_ar_de = exposure_un_ar_tab.range(f'A'+ str(exposure_un_ar_tab.cells.last_cell.row)).end('up').row
        last_column_letter_un_ar_de=num_to_col_letters(exposure_un_ar_tab.range('A1').end('right').last_cell.column)
        exposure_un_ar_tab.range(f"A2:{last_column_letter_un_ar_de}{last_row_un_ar_de}").api.Sort(Key1=exposure_un_ar_tab.range(f"M2:M{last_row_un_ar_de}").api,Order1=win32c.SortOrder.xlAscending,DataOption1=win32c.SortDataOption.xlSortNormal,Orientation=1,SortMethod=1)
    

        unapplied_credit_tab.activate()
        last_row_un_Cr_t = unapplied_credit_tab.range(f'A'+ str(unapplied_credit_tab.cells.last_cell.row)).end('up').row
        unapplied_credit_tab.range(f"A2:H{last_row_un_Cr_t}").api.Sort(Key1=unapplied_credit_tab.range(f"E2:E{last_row_un_Cr_t}").api,Order1=win32c.SortOrder.xlAscending,DataOption1=win32c.SortDataOption.xlSortNormal,Orientation=1,SortMethod=1)
            
        exposure_ar_tab.activate()
        exposure_ar_tab.api.Range(f"C:C").Delete(win32c.DeleteShiftDirection.xlShiftToLeft)
        exposure_ar_tab.api.Range(f"I:I").Delete(win32c.DeleteShiftDirection.xlShiftToLeft)
        exposure_ar_tab.api.Range(f"J:J").Delete(win32c.DeleteShiftDirection.xlShiftToLeft)
        exposure_ar_tab.api.Range(f"Q:Q").Delete(win32c.DeleteShiftDirection.xlShiftToLeft)

        exposure_ar_tab.range(f"F:F").number_format="mm/dd/yy;@"
        exposure_ar_tab.range(f"I:O").number_format="#,##0.00"

        last_row_o_ar = exposure_ar_tab.range(f'A'+ str(exposure_ar_tab.cells.last_cell.row)).end('up').row
        last_column_letter_o_ar=num_to_col_letters(exposure_ar_tab.range('A1').end('right').last_cell.column)
        exposure_ar_tab.range(f"A2:{last_column_letter_o_ar}{last_row_o_ar}").api.Sort(Key1=exposure_ar_tab.range(f"O2:O{last_row_o_ar}").api,Order1=win32c.SortOrder.xlDescending,DataOption1=win32c.SortDataOption.xlSortNormal,Orientation=1,SortMethod=1)
        exposure_ar_tab.range(f"A2:{last_column_letter_o_ar}{last_row_o_ar}").api.Sort(Key1=exposure_ar_tab.range(f"N2:N{last_row_o_ar}").api,Order1=win32c.SortOrder.xlDescending,DataOption1=win32c.SortDataOption.xlSortNormal,Orientation=1,SortMethod=1)
        exposure_ar_tab.range(f"A2:{last_column_letter_o_ar}{last_row_o_ar}").api.Sort(Key1=exposure_ar_tab.range(f"M2:M{last_row_o_ar}").api,Order1=win32c.SortOrder.xlDescending,DataOption1=win32c.SortDataOption.xlSortNormal,Orientation=1,SortMethod=1)
        exposure_ar_tab.range(f"A2:{last_column_letter_o_ar}{last_row_o_ar}").api.Sort(Key1=exposure_ar_tab.range(f"L2:L{last_row_o_ar}").api,Order1=win32c.SortOrder.xlDescending,DataOption1=win32c.SortDataOption.xlSortNormal,Orientation=1,SortMethod=1)

        exposure_ar_tab.range(f"H:H").number_format="m/d/yyyy"

        exposure_ar_tab.api.Range(f"A1:{last_column_letter_o_ar}{last_row_o_ar}").AutoFilter(Field:=15, Criteria1:=[">0"],Operator:=1) 


        retry=0
        while retry < 10:
            try:
                pre_ar_exposure_wb = xw.Book(previous_sheet_ar_axposure,update_links=False) 
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry ==10:
                    raise e 
        pre_open_ar = pre_ar_exposure_wb.sheets["Open AR"]           
        pre_open_ar.activate()
        last_row_o_ar_pre = pre_open_ar.range(f'A'+ str(pre_open_ar.cells.last_cell.row)).end('up').row
        last_column_letter_o_ar_pre=num_to_col_letters(pre_open_ar.range('A1').end('right').last_cell.column)
        pre_open_ar.api.Range(f"A1:{last_column_letter_o_ar_pre}{last_row_o_ar_pre}").AutoFilter(Field:=15, Criteria1:=[">0"],Operator:=1)

        pre_open_ar.api.Range(f"O:P").SpecialCells(win32c.CellType.xlCellTypeVisible).Copy()
        exposure_ar_tab.api.Range(f"X:X")._PasteSpecial(Paste=-4163,Operation=win32c.Constants.xlNone) 
        exposure_ar_tab.activate()

        sp_lst_row = exposure_ar_tab.range(f'A'+ str(exposure_ar_tab.cells.last_cell.row)).end('up').row
        sp_address= exposure_ar_tab.api.Range(f"A2:P{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Address
        sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]
        exposure_ar_tab.api.Range(f"P{sp_initial_rw}").Value = f'=IFERROR(VLOOKUP(O{sp_initial_rw},X:Y,2,0),"New item")'
        
        exposure_ar_tab.api.Range(f"P{sp_initial_rw}:P{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Select()
        ar_exposure_wb.app.api.Selection.FillDown()

        exposure_ar_tab.api.Range(f"P{sp_initial_rw}:P{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Copy()
        exposure_ar_tab.api.Range(f"P{sp_initial_rw}:P{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible)._PasteSpecial(Paste=-4163,Operation=win32c.Constants.xlNone)
        exposure_ar_tab.api.Range(f"X:Y").Delete(win32c.DeleteShiftDirection.xlShiftToLeft)

        interior_coloring(colour_value=65535,cellrange=f"A{sp_initial_rw}:P{sp_lst_row}",working_sheet=exposure_ar_tab,working_workbook=ar_exposure_wb)
        exposure_ar_tab.api.AutoFilterMode=False
        ar_exposure_wb.app.api.CutCopyMode=False
        pre_open_ar.api.AutoFilterMode=False
        exposure_ar_tab.range(f"A2:{last_column_letter_o_ar}{last_row_o_ar}").api.Sort(Key1=exposure_ar_tab.range(f"N2:N{last_row_o_ar}").api,Order1=win32c.SortOrder.xlDescending,DataOption1=win32c.SortDataOption.xlSortNormal,Orientation=1,SortMethod=1)
    

        exposure_ar_tab.api.Range(f"A1:{last_column_letter_o_ar}{last_row_o_ar}").AutoFilter(Field:=14, Criteria1:=[">0"],Operator:=1)
        filter_check=exposure_ar_tab.range(f'A'+ str(exposure_ar_tab.cells.last_cell.row)).end('up').row
        if filter_check!=1:
            sp_lst_row = exposure_ar_tab.range(f'A'+ str(exposure_ar_tab.cells.last_cell.row)).end('up').row
            sp_address= exposure_ar_tab.api.Range(f"A2:P{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Address
            sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]
            interior_coloring(colour_value=65535,cellrange=f"A{sp_initial_rw}:P{sp_lst_row}",working_sheet=exposure_ar_tab,working_workbook=ar_exposure_wb)
        else:
            print("no data found")
        exposure_ar_tab.api.AutoFilterMode=False
        exposure_ar_tab.range(f"A2:{last_column_letter_o_ar}{last_row_o_ar}").api.Sort(Key1=exposure_ar_tab.range(f"M2:M{last_row_o_ar}").api,Order1=win32c.SortOrder.xlDescending,DataOption1=win32c.SortDataOption.xlSortNormal,Orientation=1,SortMethod=1)
    

        exposure_ar_tab.api.Range(f"A1:{last_column_letter_o_ar}{last_row_o_ar}").AutoFilter(Field:=13, Criteria1:=[">0"],Operator:=1)
        filter_check=exposure_ar_tab.range(f'A'+ str(exposure_ar_tab.cells.last_cell.row)).end('up').row
        if filter_check!=1:
            sp_lst_row = exposure_ar_tab.range(f'A'+ str(exposure_ar_tab.cells.last_cell.row)).end('up').row
            sp_address= exposure_ar_tab.api.Range(f"A2:P{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Address
            sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]
            interior_coloring(colour_value=5296274,cellrange=f"A{sp_initial_rw}:P{sp_lst_row}",working_sheet=exposure_ar_tab,working_workbook=ar_exposure_wb)
        else:
            print("no data found")
        exposure_ar_tab.api.AutoFilterMode=False
        exposure_ar_tab.range(f"A2:{last_column_letter_o_ar}{last_row_o_ar}").api.Sort(Key1=exposure_ar_tab.range(f"O2:O{last_row_o_ar}").api,Order1=win32c.SortOrder.xlDescending,DataOption1=win32c.SortDataOption.xlSortNormal,Orientation=1,SortMethod=1)
        exposure_ar_tab.range(f"A2:{last_column_letter_o_ar}{last_row_o_ar}").api.Sort(Key1=exposure_ar_tab.range(f"N2:N{last_row_o_ar}").api,Order1=win32c.SortOrder.xlDescending,DataOption1=win32c.SortDataOption.xlSortNormal,Orientation=1,SortMethod=1)       
        exposure_ar_tab.range(f"A2:{last_column_letter_o_ar}{last_row_o_ar}").api.Sort(Key1=exposure_ar_tab.range(f"M2:M{last_row_o_ar}").api,Order1=win32c.SortOrder.xlDescending,DataOption1=win32c.SortDataOption.xlSortNormal,Orientation=1,SortMethod=1)       
        
        insert_all_borders(cellrange=f"A1:{last_column_letter_o_ar}{last_row_o_ar}",working_sheet=exposure_ar_tab,working_workbook=ar_exposure_wb)

        over_limit_tab.activate()
        last_column_letter_ol=num_to_col_letters(over_limit_tab.range('A1').end('right').last_cell.column)
        over_limit_tab.api.Range(f"A1:{last_column_letter_ol}{last_row_ol}").AutoFilter(Field:=7, Criteria1:=[">=75000"],Operator:=1)
        
        filter_check=over_limit_tab.range(f'A'+ str(over_limit_tab.cells.last_cell.row)).end('up').row
        if filter_check!=1:
            sp_lst_row = over_limit_tab.range(f'A'+ str(over_limit_tab.cells.last_cell.row)).end('up').row
            sp_address= over_limit_tab.api.Range(f"A2:{last_column_letter_ol}{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Address
            sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]
            over_limit_tab.api.Range(f"A2:{last_column_letter_ol}{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Font.Bold=True
        else:
            print("no data found")
        over_limit_tab.api.AutoFilterMode=False        
        
        unsettled_receivables_tab.activate()
        Unarsum_lst_row = unsettled_receivables_tab.range(f'A'+ str(unsettled_receivables_tab.cells.last_cell.row)).end('up').row
        unsettled_receivables_tab.range(f"A2:H{Unarsum_lst_row}").api.Sort(Key1=unsettled_receivables_tab.range(f"D2:D{Unarsum_lst_row}").api,Order1=win32c.SortOrder.xlDescending,DataOption1=win32c.SortDataOption.xlSortNormal,Orientation=1,SortMethod=1)       
        
        unsettled_receivables_tab.api.Range(f"A1:H{Unarsum_lst_row}").AutoFilter(Field:=4, Criteria1:=[">=75000"],Operator:=1)
        
        filter_check=unsettled_receivables_tab.range(f'A'+ str(unsettled_receivables_tab.cells.last_cell.row)).end('up').row
        if filter_check!=1:
            sp_lst_row = unsettled_receivables_tab.range(f'A'+ str(unsettled_receivables_tab.cells.last_cell.row)).end('up').row
            sp_address= unsettled_receivables_tab.api.Range(f"A2:H{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Address
            sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]
            unsettled_receivables_tab.api.Range(f"A2:H{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Font.Bold=True
        else:
            print("no data found")
        unsettled_receivables_tab.api.AutoFilterMode=False   
        insert_all_borders(cellrange=f"A2:H{Unarsum_lst_row}",working_sheet=unsettled_receivables_tab,working_workbook=ar_exposure_wb)

        unapplied_credit_tab.activate()
        unapplied_credit_tab.range(f"A2:H{last_row_un_Cr_t}").api.Sort(Key1=unapplied_credit_tab.range(f"E2:E{last_row_un_Cr_t}").api,Order1=win32c.SortOrder.xlAscending,DataOption1=win32c.SortDataOption.xlSortNormal,Orientation=1,SortMethod=1)
        
        unapplied_credit_tab.api.Range(f"A1:H{last_row_un_Cr_t}").AutoFilter(Field:=5, Criteria1:=["<=-75000"],Operator:=1)
        filter_check=unapplied_credit_tab.range(f'A'+ str(unapplied_credit_tab.cells.last_cell.row)).end('up').row
        if filter_check!=1:
            sp_lst_row = unsettled_receivables_tab.range(f'A'+ str(unsettled_receivables_tab.cells.last_cell.row)).end('up').row
            sp_address= unsettled_receivables_tab.api.Range(f"A2:H{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Address
            sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]
            unapplied_credit_tab.api.Range(f"A2:H{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Font.Bold=True
        else:
            print("no data found")
        unapplied_credit_tab.api.AutoFilterMode=False  

        # unapplied_credit_tab.api.Range(f"E2:E{last_row_un_Cr_t}")
        interior_coloring_by_theme(pattern_tns=0,tintandshade=0.399975585192419,colour_value=win32c.ThemeColor.xlThemeColorAccent1,cellrange=f"E2:E{last_row_un_Cr_t}",working_sheet=unapplied_credit_tab,working_workbook=ar_exposure_wb)
        
        
        exposure_un_ar_tab.activate()
        insert_all_borders(cellrange=f"A1:{last_column_letter_un_ar_de}{last_row_un_ar_de}",working_sheet=exposure_un_ar_tab,working_workbook=ar_exposure_wb)

        unapplied_credit_tab.activate()
        insert_all_borders(cellrange=f"A1:H{last_row_un_Cr_t}",working_sheet=unapplied_credit_tab,working_workbook=ar_exposure_wb)

        over_limit_tab.activate()
        insert_all_borders(cellrange=f"A2:G{last_row_ol}",working_sheet=over_limit_tab,working_workbook=ar_exposure_wb)

        exposure_ar_tab.activate()
        exposure_ar_tab.autofit()
        # filter_dict = {"O":[15,65535],"N":[14,6684927],"M":[13,5296274]}
        filter_dict = {"O":[15,65535],"N":[14,14857357],"M":[13,5296274]}
        for key,value in filter_dict.items():
            exposure_ar_tab.api.Range(f"A1:{last_column_letter_o_ar}{last_row_o_ar}").AutoFilter(Field:=value[0], Criteria1:=[">0"],Operator:=1)
            filter_check=exposure_ar_tab.range(f'A'+ str(exposure_ar_tab.cells.last_cell.row)).end('up').row
            if filter_check!=1:
                sp_lst_row = exposure_ar_tab.range(f'A'+ str(exposure_ar_tab.cells.last_cell.row)).end('up').row
                sp_address= exposure_ar_tab.api.Range(f"A2:{last_column_letter_o_ar}{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Address
                sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]
                exposure_ar_tab.api.Range(f"R{sp_lst_row}").Value=f"=SUM({key}{sp_initial_rw}:{key}{sp_lst_row})"
                exposure_ar_tab.range(f"R{sp_lst_row}").number_format = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                exposure_ar_tab.api.Range(f"R{sp_lst_row}").Font.Bold=True
                exposure_ar_tab.api.Range(f"R{sp_lst_row-1}").Value=f"Today"
                pre_ar_exposure_wb.activate()
                pre_open_ar.activate()
                pre_open_ar.api.Range(f"A1:{last_column_letter_o_ar_pre}{last_row_o_ar_pre}").AutoFilter(Field:=value[0], Criteria1:=[">0"],Operator:=1)
                sp_lst_row_pre = pre_open_ar.range(f'A'+ str(pre_open_ar.cells.last_cell.row)).end('up').row
                pre_open_ar.api.Range(f"R{sp_lst_row_pre}").Copy()
                exposure_ar_tab.api.Range(f"S{sp_lst_row}")._PasteSpecial(Paste=win32c.PasteType.xlPasteValues)
                ar_exposure_wb.app.api.CutCopyMode=False
                exposure_ar_tab.api.Range(f"S{sp_lst_row}").Font.Bold=True
                exposure_ar_tab.range(f"S{sp_lst_row}").number_format = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                exposure_ar_tab.api.Range(f"S{sp_lst_row-1}").Value=f"Last Report"
                exposure_ar_tab.api.Range(f"T{sp_lst_row-1}").Value=f"Change"
                exposure_ar_tab.api.Range(f"T{sp_lst_row}").Value=f"=R{sp_lst_row}-S{sp_lst_row}"
                exposure_ar_tab.range(f"T{sp_lst_row}").number_format = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                interior_coloring(colour_value=value[1],cellrange=f"R{sp_lst_row-1}:T{sp_lst_row}",working_sheet=exposure_ar_tab,working_workbook=ar_exposure_wb)
            else:
                print(key,value)
                time.sleep
                print("no data found")
            exposure_ar_tab.api.AutoFilterMode=False 
            pre_open_ar.api.AutoFilterMode=False

        over_limit_tab.activate()
        over_limit_tab.api.Range(f"A1:G{last_row_ol}").AutoFilter(Field:=5, Criteria1:=["$-"],Operator:=1)
        filter_check=over_limit_tab.range(f'A'+ str(over_limit_tab.cells.last_cell.row)).end('up').row
        if filter_check!=1:
            sp_lst_row = over_limit_tab.range(f'A'+ str(over_limit_tab.cells.last_cell.row)).end('up').row
            sp_address= over_limit_tab.api.Range(f"A2:H{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Address
            sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]
            over_limit_tab.api.Range(f"A{sp_initial_rw}:G{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Delete(Shift=win32c.Direction.xlUp)
        else:
            print("no data found")
        over_limit_tab.api.AutoFilterMode=False 

        overlimit_lst_rw=over_limit_tab.range(f'A'+ str(over_limit_tab.cells.last_cell.row)).end('up').row
        over_limit_tab.api.Range(f"G{overlimit_lst_rw+1}").Value=f"=SUM(G2:G{overlimit_lst_rw})"
        over_limit_tab.api.Range(f"G{overlimit_lst_rw+1}").Font.Bold=True
        over_limit_tab.api.Range(f"G{overlimit_lst_rw+1}").Font.Size=14
        over_limit_tab.api.Range(f"H{overlimit_lst_rw+1}").Value=f"Today"
        pre_ar_exposure_wb.activate()
        pre_over_limit = pre_ar_exposure_wb.sheets["Over Limit"] 
        pre_over_limit.activate()
        pre_over_limit_lst_Rw = pre_over_limit.range(f'A'+ str(pre_over_limit.cells.last_cell.row)).end('up').row
        
        pre_over_limit.api.Range(f"G{pre_over_limit_lst_Rw+1}").Copy()
        over_limit_tab.api.Range(f"G{overlimit_lst_rw+2}")._PasteSpecial(Paste=win32c.PasteType.xlPasteValues)
        ar_exposure_wb.app.api.CutCopyMode=False
        over_limit_tab.activate()
        over_limit_tab.api.Range(f"H{overlimit_lst_rw+2}").Value=f"Prior Report"

        over_limit_tab.api.Range(f"G{overlimit_lst_rw+3}").Value=f"=G{overlimit_lst_rw+1}-G{overlimit_lst_rw+2}"
        over_limit_tab.api.Range(f"G{overlimit_lst_rw+2}").Font.Bold=True
        over_limit_tab.api.Range(f"G{overlimit_lst_rw+3}").Font.Bold=True
        over_limit_tab.api.Range(f"H{overlimit_lst_rw+3}").Value=f"Change"


        unsettled_receivables_tab.activate()

        unsettled_receivables_tab.api.AutoFilterMode=False 
        Unarsum_lst_row=unsettled_receivables_tab.range(f'A'+ str(unsettled_receivables_tab.cells.last_cell.row)).end('up').row
        unsettled_receivables_tab.api.Range(f"D{Unarsum_lst_row+1}").Value=f"=SUM(D2:D{Unarsum_lst_row})"

        interior_coloring(colour_value=10079232,cellrange=f"D2:D{Unarsum_lst_row+3}",working_sheet=unsettled_receivables_tab,working_workbook=ar_exposure_wb)

        unsettled_receivables_tab.api.Range(f"D{Unarsum_lst_row+1}").Font.Bold=True
        unsettled_receivables_tab.api.Range(f"D{Unarsum_lst_row+1}").Font.Size=14
        unsettled_receivables_tab.api.Range(f"E{Unarsum_lst_row+1}").Value=f"Today"
        pre_ar_exposure_wb.activate()
        pre_un_ar_sum = pre_ar_exposure_wb.sheets["Unsettled AR Summary"] 
        pre_un_ar_sum.activate()
        pre_over_limit_lst_Rw = pre_un_ar_sum.range(f'A'+ str(pre_un_ar_sum.cells.last_cell.row)).end('up').row
        
        pre_un_ar_sum.api.Range(f"D{pre_over_limit_lst_Rw+1}").Copy()
        unsettled_receivables_tab.api.Range(f"D{Unarsum_lst_row+2}")._PasteSpecial(Paste=win32c.PasteType.xlPasteValues)
        ar_exposure_wb.app.api.CutCopyMode=False
        unsettled_receivables_tab.activate()

        unsettled_receivables_tab.range(f"D{Unarsum_lst_row+2}").number_format = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"

        unsettled_receivables_tab.api.Range(f"E{Unarsum_lst_row+2}").Value=f"Prior Report"

        unsettled_receivables_tab.api.Range(f"D{Unarsum_lst_row+3}").Value=f"=D{Unarsum_lst_row+1}-D{Unarsum_lst_row+2}"
        unsettled_receivables_tab.api.Range(f"D{Unarsum_lst_row+2}").Font.Bold=True
        unsettled_receivables_tab.api.Range(f"D{Unarsum_lst_row+3}").Font.Bold=True
        unsettled_receivables_tab.api.Range(f"E{Unarsum_lst_row+3}").Value=f"Change"


        unapplied_credit_tab.activate()

        unapplied_credit_tab.api.AutoFilterMode=False 
        Uncr_lst_row=unapplied_credit_tab.range(f'A'+ str(unapplied_credit_tab.cells.last_cell.row)).end('up').row
        unapplied_credit_tab.api.Range(f"E{Uncr_lst_row+1}").Value=f"=SUM(E2:E{Uncr_lst_row})"

        interior_coloring(colour_value=16764006,cellrange=f"E2:E{Uncr_lst_row+3}",working_sheet=unapplied_credit_tab,working_workbook=ar_exposure_wb)

        unapplied_credit_tab.api.Range(f"E{Uncr_lst_row+1}").Font.Bold=True
        unapplied_credit_tab.api.Range(f"E{Uncr_lst_row+1}").Font.Size=14
        unapplied_credit_tab.api.Range(f"F{Uncr_lst_row+1}").Value=f"Today"
        pre_ar_exposure_wb.activate()
        pre_un_cre = pre_ar_exposure_wb.sheets["Unapplied Credit"] 
        pre_un_cre.activate()
        pre_over_limit_lst_Rw = pre_un_cre.range(f'A'+ str(pre_un_cre.cells.last_cell.row)).end('up').row
        
        pre_un_cre.api.Range(f"E{pre_over_limit_lst_Rw+1}").Copy()
        unapplied_credit_tab.api.Range(f"E{Uncr_lst_row+2}")._PasteSpecial(Paste=win32c.PasteType.xlPasteValues)
        ar_exposure_wb.app.api.CutCopyMode=False
        unapplied_credit_tab.activate()

        unapplied_credit_tab.range(f"E{Uncr_lst_row+2}").number_format = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"

        unapplied_credit_tab.api.Range(f"F{Uncr_lst_row+2}").Value=f"Prior Report"

        unapplied_credit_tab.api.Range(f"E{Uncr_lst_row+3}").Value=f"=E{Uncr_lst_row+1}-E{Uncr_lst_row+2}"
        unapplied_credit_tab.api.Range(f"E{Uncr_lst_row+2}").Font.Bold=True
        unapplied_credit_tab.api.Range(f"E{Uncr_lst_row+3}").Font.Bold=True
        unapplied_credit_tab.api.Range(f"F{Uncr_lst_row+3}").Value=f"Change"


        credit_over_utilized_tab.activate()
        credit_over_utilized_tab.api.Range(f"A1:H{last_row_cou}").AutoFilter(Field:=6, Criteria1:=["$-"],Operator:=1)
        filter_check=credit_over_utilized_tab.range(f'A'+ str(credit_over_utilized_tab.cells.last_cell.row)).end('up').row
        if filter_check!=1:
            sp_lst_row = credit_over_utilized_tab.range(f'A'+ str(credit_over_utilized_tab.cells.last_cell.row)).end('up').row
            sp_address= credit_over_utilized_tab.api.Range(f"A2:H{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Address
            sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]
            credit_over_utilized_tab.api.Range(f"A{sp_initial_rw}:H{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Delete(Shift=win32c.Direction.xlUp)
        else:
            print("no data found")
        credit_over_utilized_tab.api.AutoFilterMode=False 

        pre_ar_exposure_wb.activate()
        pre_ar_exposure_wb.api.Sheets(["Sheet1", "Dispute-Legal "]).Copy(Before=ar_exposure_wb.api.Sheets(1))

        dis_legal_sheet = ar_exposure_wb.sheets["Dispute-Legal "]
        dis_legal_sheet.api.Move(Before=ar_exposure_wb.api.Sheets(5))
        lst_rw_dis=dis_legal_sheet.range(f'B'+ str(dis_legal_sheet.cells.last_cell.row)).end('up').row

        over_limit_tab.activate()
        last_row_ol = over_limit_tab.range(f'A'+ str(over_limit_tab.cells.last_cell.row)).end('up').row
        over_limit_tab.api.Range(f"A1:G{last_row_ol}").AutoFilter(Field:=7, Criteria1:=["$-"],Operator:=1)
        filter_check=over_limit_tab.range(f'A'+ str(over_limit_tab.cells.last_cell.row)).end('up').row
        if filter_check!=1:
            sp_lst_row = over_limit_tab.range(f'A'+ str(over_limit_tab.cells.last_cell.row)).end('up').row
            sp_address= over_limit_tab.api.Range(f"A2:H{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Address
            sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]
            over_limit_tab.api.Range(f"A{sp_initial_rw}:G{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Delete(Shift=win32c.Direction.xlUp)
        else:
            print("no data found")
        over_limit_tab.api.AutoFilterMode=False 

        dis_legal_sheet.activate()
        disputed_cus = dis_legal_sheet.range(f"B2:B{lst_rw_dis}").value
        disputed_cus2 = dis_legal_sheet.range(f"A2:A{lst_rw_dis}").value
        over_limit_tab.activate()
        last_row_ol = over_limit_tab.range(f'A'+ str(over_limit_tab.cells.last_cell.row)).end('up').row
        i=2
        while i <=last_row_ol:
        # for i in range(2,int(f'{last_row_ol}')):
            if over_limit_tab.range(f"B{i}").value!=None:
                if (over_limit_tab.range(f"B{i}").value).upper() in disputed_cus:
                    print(f"Match found in {i}") 
                    over_limit_tab.api.Range(f"{i}:{i}").Delete(Shift=win32c.Direction.xlUp)
                    i-=1
                elif (over_limit_tab.range(f"A{i}").value).upper() in disputed_cus2:
                    print(f"Match found in {i}") 
                    over_limit_tab.api.Range(f"{i}:{i}").Delete(Shift=win32c.Direction.xlUp)
                    i-=1                
                else:
                    i+=1  
            else:
                i+=1                        

        last_row_ol = over_limit_tab.range(f'A'+ str(over_limit_tab.cells.last_cell.row)).end('up').row
        lisr_cus = over_limit_tab.range(f"A2:A{last_row_ol}").value
        
        duplicates = [number for number in lisr_cus if lisr_cus.count(number) > 1]
        unique_duplicates = list(set(duplicates))
        for filter in unique_duplicates:
            over_limit_tab.api.Range(f"A1:G{last_row_ol}").AutoFilter(Field:=1, Criteria1:=[filter])
            filter_check=over_limit_tab.range(f'A'+ str(over_limit_tab.cells.last_cell.row)).end('up').row
            if filter_check!=1:
                sp_lst_row = over_limit_tab.range(f'A'+ str(over_limit_tab.cells.last_cell.row)).end('up').row
                sp_address= over_limit_tab.api.Range(f"A2:H{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Address
                sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]
                over_limit_tab.api.Range(f"{sp_initial_rw}:{sp_initial_rw}").SpecialCells(win32c.CellType.xlCellTypeVisible).Insert(Shift=win32c.Direction.xlDown)
                over_limit_tab.api.Range(f"A{int(sp_initial_rw)+1}:B{int(sp_initial_rw)+1}").SpecialCells(win32c.CellType.xlCellTypeVisible).Copy()
                over_limit_tab.api.Range(f"A{sp_initial_rw}")._PasteSpecial(Paste=win32c.PasteType.xlPasteValues)
                ar_exposure_wb.app.api.CutCopyMode=False
                over_limit_tab.api.Range(f"F{int(sp_initial_rw)+1}").Copy()
                over_limit_tab.api.Range(f"F{sp_initial_rw}")._PasteSpecial(Paste=win32c.PasteType.xlPasteValues)
                ar_exposure_wb.app.api.CutCopyMode=False
                sp_lst_row2 = over_limit_tab.range(f'A'+ str(over_limit_tab.cells.last_cell.row)).end('up').row
                updated_rows_add=over_limit_tab.api.Range(f"A2:A{sp_lst_row2}").SpecialCells(win32c.CellType.xlCellTypeVisible).Rows.Address
                sum_Rows_C = "=C"+"+C".join(re.findall('\d+', updated_rows_add)[1:])
                sum_Rows_D = "=D"+"+D".join(re.findall('\d+', updated_rows_add)[1:])
                sum_Rows_E = "=E"+"+E".join(re.findall('\d+', updated_rows_add)[1:])
                over_limit_tab.api.Range(f"C{sp_initial_rw}").Value= sum_Rows_C
                over_limit_tab.api.Range(f"D{sp_initial_rw}").Value= sum_Rows_D
                over_limit_tab.api.Range(f"E{sp_initial_rw}").Value= sum_Rows_E
                over_limit_tab.api.Range(f"G{sp_initial_rw}").Value= f'=IF(E{sp_initial_rw}>F{sp_initial_rw},E{sp_initial_rw}-F{sp_initial_rw},0)'
                over_limit_tab.api.Range(f"{sp_initial_rw}:{sp_initial_rw}").Copy()
                over_limit_tab.api.Range(f"{sp_initial_rw}:{sp_initial_rw}")._PasteSpecial(Paste=win32c.PasteType.xlPasteValues)
                ar_exposure_wb.app.api.CutCopyMode=False
                over_limit_tab.api.Range(f"{int(sp_initial_rw)+1}:{sp_lst_row2}").SpecialCells(win32c.CellType.xlCellTypeVisible).Delete(Shift=win32c.Direction.xlUp)
            else:
                print("no data found")
            over_limit_tab.api.AutoFilterMode=False 

        
        ########################################## NEW COLOR LOGIC ################################################
        last_row_open_ar_sh = exposure_ar_tab.range(f'A'+ str(exposure_ar_tab.cells.last_cell.row)).end('up').row

        # Apply the filter to Column N (assuming it's the 14th column)
        exposure_ar_tab.api.Range(f"A1:P{last_row_open_ar_sh}").AutoFilter(Field=14, Criteria1=">0", Operator=2)
        interior_coloring(colour_value=14857357,cellrange=f"A2:P{last_row_open_ar_sh}",working_sheet=exposure_ar_tab,working_workbook=ar_exposure_wb)
        exposure_ar_tab.api.AutoFilterMode=False
        
        exposure_ar_tab.range(f"A2:T{last_row_o_ar}").api.Sort(Key1=exposure_ar_tab.range(f"L2:L{last_row_o_ar}").api,Order1=win32c.SortOrder.xlDescending,DataOption1=win32c.SortDataOption.xlSortNormal,Orientation=1,SortMethod=1)
        exposure_ar_tab.range(f"A2:T{last_row_o_ar}").api.Sort(Key1=exposure_ar_tab.range(f"M2:M{last_row_o_ar}").api,Order1=win32c.SortOrder.xlDescending,DataOption1=win32c.SortDataOption.xlSortNormal,Orientation=1,SortMethod=1)
        exposure_ar_tab.range(f"A2:T{last_row_o_ar}").api.Sort(Key1=exposure_ar_tab.range(f"N2:N{last_row_o_ar}").api,Order1=win32c.SortOrder.xlDescending,DataOption1=win32c.SortDataOption.xlSortNormal,Orientation=1,SortMethod=1)
        exposure_ar_tab.range(f"A2:T{last_row_o_ar}").api.Sort(Key1=exposure_ar_tab.range(f"O2:O{last_row_o_ar}").api,Order1=win32c.SortOrder.xlDescending,DataOption1=win32c.SortDataOption.xlSortNormal,Orientation=1,SortMethod=1)
        
        # Descending order in over limit column
        over_limit_tab.activate()
        last_row_a_over_limit = over_limit_tab.range(f'A'+ str(over_limit_tab.cells.last_cell.row)).end('up').row
        over_limit_tab.range(f"A2:G{last_row_a_over_limit}").api.Sort(Key1=over_limit_tab.range(f"G2:G{last_row_a_over_limit}").api,Order1=win32c.SortOrder.xlDescending,DataOption1=win32c.SortDataOption.xlSortNormal,Orientation=1,SortMethod=1)
        
        # Adding dollar sign in sheet1 column B        
        sh1_ws=ar_exposure_wb.sheets[0]
        sh1_ws.activate()
        last_row_b_sh1 = sh1_ws.range(f'B'+ str(sh1_ws.cells.last_cell.row)).end('up').row
        sh1_ws.range(f"B4:B{last_row_b_sh1}").api.NumberFormat = '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'
        ###########################################################################################################
        ar_exposure_wb.save(f"{output_location}\\AR Exposure "+input_date+'.xlsm')

        try:
            wb.app.quit()
        except:
            wb.app.quit()  
        return f"{job_name} Report for {input_date} generated succesfully"

    except Exception as e:
        wb.app.kill()
        raise e
    finally:
        try:
            wb.app.quit()
        except:
            pass



def ar_reports_exposure(input_date, output_date):
    try:       
        job_name = 'exposure_ar_reports'
        input_sheet2 = drive+r'\REPORT\Ar_Exposure(Open AR,Unsettled AR)\Input'+f'\\Open AR_{input_date}.xlsx'
        input_sheet= drive+r'\REPORT\Ar_Exposure(Open AR,Unsettled AR)\Input'+f'\\Unsettled AR_{input_date}.xlsx'
        previous_sheet_unsettled= drive+r'\REPORT\Ar_Exposure(Open AR,Unsettled AR)\Output'+f'\\Unsettled AR_{output_date}.xlsx'
        ticket_summary_sheet = drive+r'\REPORT\Ar_Exposure(Open AR,Unsettled AR)\Summary_sheet'+f'\\ticket query elevator 2015.xlsx'
        ap_ar_template = drive+r'\REPORT\Ar_Exposure(Open AR,Unsettled AR)'+f'\\WPLLC - AP_AR_Template.xlsx'
        output_location = drive+r'\REPORT\AR EXPOSURE' 
        output_location2 = drive+r'\REPORT\AR EXPOSURE\Output'
        if not os.path.exists(input_sheet):
            return(f"{input_sheet} Excel file not present for date {input_date}")
        if not os.path.exists(input_sheet2):
            return(f"{input_sheet2} Excel file not present for date {input_date}") 
        if not os.path.exists(previous_sheet_unsettled):
            return(f"{previous_sheet_unsettled} Excel file not present for date {output_date}") 
        if not os.path.exists(ticket_summary_sheet):
            return(f"{ticket_summary_sheet} Excel file not present")    
        if not os.path.exists(ap_ar_template):
            return(f"{ap_ar_template} Excel file not present")                     
        retry=0
        while retry < 10:
            try:
                wb2 = xw.Book(previous_sheet_unsettled) 
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry ==10:
                    raise e 
        retry=0
        while retry < 10:
            try:
                wb = xw.Book(input_sheet,update_links=False) 
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry ==10:
                    raise e 
        input_tab=wb.sheets[f"Unsettled AR_{input_date}"]
        column_list = input_tab.range("A1").expand('right').value

        Quantity_Name_column_no = column_list.index('Quantity')+1
        list1=["Lk Up","Lk Up","Ticket Add Date","Diff"]
        list2=[f"=C2&O2",f"=O2&I2",f"=VLOOKUP(S2,'[Unsettled AR_{output_date}.xlsx]MASTER'!$S:$T,2,0)",f"=T2-P2"] 
        Quantity_Name_column_no+=1
        i=0
        last_row = input_tab.range(f'A'+ str(input_tab.cells.last_cell.row)).end('up').row
        for values in list1:
            last_column_letter=num_to_col_letters(Quantity_Name_column_no)
            input_tab.api.Range(f"{last_column_letter}1").EntireColumn.Insert()
            input_tab.range(f"{last_column_letter}1").value = values
            input_tab.range(f"{last_column_letter}2").value = list2[i]
            time.sleep(1)
            input_tab.range(f"{last_column_letter}2").copy(input_tab.range(f"{last_column_letter}2:{last_column_letter}{last_row}"))
            if values == 'Diff':
                input_tab.range(f"{last_column_letter}2:{last_column_letter}{last_row}").number_format="General"
            if values == 'Ticket Add Date':
                input_tab.range(f"{last_column_letter}2:{last_column_letter}{last_row}").number_format="mm-dd-yyyy"
            i+=1
            Quantity_Name_column_no+=1

        column_list = input_tab.range("A1").expand('right').value
        Ticket_column_no = column_list.index('Ticket Add Date')+1
        Ticket_column_letter=num_to_col_letters(Ticket_column_no)
        input_tab.api.Range(f"{Ticket_column_letter}1").AutoFilter(Field:=f'{Ticket_column_no}', Criteria1:=["#N/A"]) 
        retry=0
        while retry < 10:
            try:
                wb_ex = xw.Book(ticket_summary_sheet,update_links=False) 
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry ==10:
                    raise e 
        input_tab.activate()
        last_row = input_tab.range(f'A'+ str(input_tab.cells.last_cell.row)).end('up').row
        last_column_letter=num_to_col_letters(input_tab.range('A1').end('right').last_cell.column)

        initial_row = re.findall("\d+",input_tab.api.Range(f"{Ticket_column_letter}2:{Ticket_column_letter}{last_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Address.split(',')[0].replace('$',""))[0]
        input_tab.range(f"{Ticket_column_letter}{initial_row}").number_format="mm-dd-yyyy"
        # input_tab.api.Range(f"{Ticket_column_letter}{initial_row}").Value = f"=VLOOKUP(S{initial_row},'[{wb2.name}]MASTER'!$S:$T,2,0)"
        input_tab.api.Range(f"{Ticket_column_letter}{initial_row}").Value = f"=VLOOKUP(S{initial_row},'[ticket query elevator 2015.xlsx]Sheet1'!$G:$Q,11,0)"
        input_tab.api.Range(f"{Ticket_column_letter}{initial_row}:{Ticket_column_letter}{last_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Select()
        wb.app.api.Selection.FillDown()
        wb.app.api.ActiveSheet.ShowAllData()

        last_row = input_tab.range(f'A'+ str(input_tab.cells.last_cell.row)).end('up').row
        input_tab.api.Range(f"{Ticket_column_letter}2:{Ticket_column_letter}{last_row}").Copy()
        input_tab.api.Range(f"{Ticket_column_letter}2")._PasteSpecial(Paste=win32c.PasteType.xlPasteValues)

        # a = input_tab.range(f"{Ticket_column_letter}2:{Ticket_column_letter}{last_row}").value
        # try:
        #     b = [datetime.strftime(no,"%d-%m-%Y").strip().split(" ")[0] if no!=None else no for no in a]
        # except:
        #     print("Please check dates once")    
        # input_tab.range(f"{Ticket_column_letter}2").options(transpose=True).value = b

        input_tab.api.Select()
        input_tab.api.Copy(None, After=input_tab.api)
        wb.sheets[-1].name = "MASTER"
        input_tab.activate()
        #Selection.FillDown
        # Sheets("Unsettled AR_11.18.2022").Select
        # Sheets("Unsettled AR_11.18.2022").Copy After:=Sheets(1)
        column_list = input_tab.range("A1").expand('right').value
        Customer_Name_column_no = column_list.index('Customer/Vendor Name')+1
        Customer_Name_column_letter=num_to_col_letters(Customer_Name_column_no)
        Location_column_no = column_list.index('Location Name')+1
        Location_column_letter=num_to_col_letters(Location_column_no)
        last_row = input_tab.range(f'A'+ str(input_tab.cells.last_cell.row)).end('up').row
        input_tab.range(f"U2:U{last_row}").number_format="General"
        last_column_letter=num_to_col_letters(input_tab.range('A1').end('right').last_cell.column)
        dict1={"MACQUARIE COMMODITIES (USA) INC.":[Customer_Name_column_no,Customer_Name_column_letter],"INTER-COMPANY PURCH/SALES":[Customer_Name_column_no,Customer_Name_column_letter],"WPMEXICO":[Location_column_no,Location_column_letter]}
        for key, value in dict1.items():
            try:
                input_tab.api.Range(f"{value[1]}1").AutoFilter(Field:=f'{value[0]}', Criteria1:=[key], Operator:=7)
                time.sleep(1)
                input_tab.api.Range(f"{value[1]}2:{last_column_letter}{last_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Select()
                time.sleep(1)
                wb.app.api.Selection.Delete(win32c.DeleteShiftDirection.xlShiftUp)
                time.sleep(1)
                wb.app.api.ActiveSheet.ShowAllData()
            except:
                wb.app.api.ActiveSheet.ShowAllData()
                pass    
        
        wb.sheets.add("Pivot",after=input_tab)
        ###logger.info("Clearing contents for new sheet")
        wb.sheets["Pivot"].clear_contents()
        ws2p=wb.sheets["Pivot"]
        ###logger.info("Declaring Variables for columns and rows")
        last_column = input_tab.range('A1').end('right').last_cell.column
        last_column_letter=num_to_col_letters(input_tab.range('A1').end('right').last_cell.column)
        last_row = input_tab.range(f'A'+ str(input_tab.cells.last_cell.row)).end('up').row
        ###logger.info("Creating Pivot Table")
        PivotCache=wb.api.PivotCaches().Create(SourceType=win32c.PivotTableSourceType.xlDatabase, SourceData=f"\'Unsettled AR_{input_date}\'!R1C1:R{last_row}C{last_column}", Version=win32c.PivotTableVersionList.xlPivotTableVersion14)
        PivotTable = PivotCache.CreatePivotTable(TableDestination=f"'Pivot'!R1C1", TableName="PivotTable1", DefaultVersion=win32c.PivotTableVersionList.xlPivotTableVersion14)        ###logger.info("Adding particular Row in Pivot Table")
        PivotTable.PivotFields('Location Name').Orientation = win32c.PivotFieldOrientation.xlRowField
        PivotTable.PivotFields('Location Name').Position = 1
        PivotTable.PivotFields('Location Name').Subtotals=(False, False, False, False, False, False, False, False, False, False, False, False)
        PivotTable.PivotFields('Customer/Vendor Id').Orientation = win32c.PivotFieldOrientation.xlRowField
        PivotTable.PivotFields('Customer/Vendor Id').Subtotals=(False, False, False, False, False, False, False, False, False, False, False, False)
        PivotTable.PivotFields('Customer/Vendor Name').Orientation = win32c.PivotFieldOrientation.xlRowField
        PivotTable.PivotFields('Customer/Vendor Name').Subtotals=(False, False, False, False, False, False, False, False, False, False, False, False)
        # PivotTable.PivotFields('City').Orientation = win32c.PivotFieldOrientation.xlRowField
        # PivotTable.PivotFields('City').Subtotals=(False, False, False, False, False, False, False, False, False, False, False, False)
        # PivotTable.PivotFields('State').Orientation = win32c.PivotFieldOrientation.xlRowField
        # PivotTable.PivotFields('State').Subtotals=(False, False, False, False, False, False, False, False, False, False, False, False)
        # PivotTable.PivotFields('Zip Code').Orientation = win32c.PivotFieldOrientation.xlRowField
        # PivotTable.PivotFields('Zip Code').Subtotals=(False, False, False, False, False, False, False, False, False, False, False, False)
        ###logger.info("Adding particular Data Field in Pivot Table")
        PivotTable.PivotFields('Net').Orientation = win32c.PivotFieldOrientation.xlDataField
        # PivotTable.PivotFields('Current').Orientation = win32c.PivotFieldOrientation.xlDataField
        # PivotTable.PivotFields('1 - 30').Orientation = win32c.PivotFieldOrientation.xlDataField
        # # PivotTable.PivotFields('Sum of  1 - 10').NumberFormat= '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'
        # PivotTable.PivotFields('31 - 60').Orientation = win32c.PivotFieldOrientation.xlDataField
        # # PivotTable.PivotFields('Sum of  31 - 60').NumberFormat= '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'
        # PivotTable.PivotFields('61 - 90').Orientation = win32c.PivotFieldOrientation.xlDataField
        # # PivotTable.PivotFields('Sum of  61 - 9999').NumberFormat= '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'
        # PivotTable.PivotFields('90+').Orientation = win32c.PivotFieldOrientation.xlDataField
        PivotTable.TableStyle2 = ""
        ###logger.info("Changing Table Layout in Pivot Table")
        PivotTable.RowAxisLayout(1)
        wb.api.ActiveSheet.PivotTables("PivotTable1").PivotFields('Location Name').RepeatLabels = True
        # wb.api.ActiveSheet.PivotTables("PivotTable1").InGridDropZones = True
        # wb.api.ActiveSheet.PivotTables("PivotTable1").DataPivotField.Caption = "Values"
        time.sleep(1)        

        retry=0
        while retry < 10:
            try:
                wb_open_ar = xw.Book(input_sheet2) 
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry ==10:
                    raise e 

        # sheet1=wb_open_ar.sheets[0].name
        input_tab2=wb_open_ar.sheets[f"Open AR_{input_date}"]            
        input_tab2.activate()            
        input_tab2.api.Select()
        input_tab2.api.Copy(None, After=input_tab2.api)
        wb_open_ar.sheets[-1].name = "MASTER"
        input_tab2.activate()

        column_list = input_tab2.range("A1").expand('right').value
        Customer_Name_column_no = column_list.index('Customer Name')+1
        Customer_Name_column_letter=num_to_col_letters(Customer_Name_column_no)
        Location_column_no = column_list.index('Location')+1
        Location_column_letter=num_to_col_letters(Location_column_no)
        last_row = input_tab2.range(f'A'+ str(input_tab2.cells.last_cell.row)).end('up').row
        last_column_letter=num_to_col_letters(input_tab2.range('A1').end('right').last_cell.column)
        dict1={"MACQUARIE COMMODITIES (USA) INC.":[Customer_Name_column_no,Customer_Name_column_letter],"INTER-COMPANY PURCH/SALES":[Customer_Name_column_no,Customer_Name_column_letter],"WEST PLAINS MEXICO":[Customer_Name_column_no,Customer_Name_column_letter],"WEST PLAINS MEXICO S. DE R.L. DE C.V.":[Customer_Name_column_no,Customer_Name_column_letter],"WPMEXICO":[Location_column_no,Location_column_letter]}
        for key, value in dict1.items():
            try:
                input_tab2.api.Range(f"{value[1]}1").AutoFilter(Field:=f'{value[0]}', Criteria1:=[key], Operator:=7)
                time.sleep(1)
                input_tab2.api.Range(f"{value[1]}2:{last_column_letter}{last_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Select()
                time.sleep(1)
                wb_open_ar.app.api.Selection.Delete(win32c.DeleteShiftDirection.xlShiftUp)
                time.sleep(1)
                wb_open_ar.app.api.ActiveSheet.ShowAllData()
            except:
                wb_open_ar.app.api.ActiveSheet.ShowAllData()
                pass 

        wb_open_ar.sheets.add("Pivot",after=input_tab2)
        ###logger.info("Clearing contents for new sheet")
        wb_open_ar.sheets["Pivot"].clear_contents()
        ws2=wb_open_ar.sheets["Pivot"]
        ###logger.info("Declaring Variables for columns and rows")
        last_column = input_tab2.range('A1').end('right').last_cell.column
        last_column_letter=num_to_col_letters(input_tab2.range('A1').end('right').last_cell.column)
        last_row = input_tab2.range(f'A'+ str(input_tab2.cells.last_cell.row)).end('up').row
        ###logger.info("Creating Pivot Table")
        PivotCache=wb_open_ar.api.PivotCaches().Create(SourceType=win32c.PivotTableSourceType.xlDatabase, SourceData=f"\'Open AR_{input_date}\'!R1C1:R{last_row}C{last_column}", Version=win32c.PivotTableVersionList.xlPivotTableVersion14)
        PivotTable = PivotCache.CreatePivotTable(TableDestination=f"'Pivot'!R1C1", TableName="PivotTable1", DefaultVersion=win32c.PivotTableVersionList.xlPivotTableVersion14)        ###logger.info("Adding particular Row in Pivot Table")
        PivotTable.PivotFields('Location').Orientation = win32c.PivotFieldOrientation.xlRowField
        PivotTable.PivotFields('Location').Position = 1
        PivotTable.PivotFields('Location').Subtotals=(False, False, False, False, False, False, False, False, False, False, False, False)
        PivotTable.PivotFields('Customer Id').Orientation = win32c.PivotFieldOrientation.xlRowField
        PivotTable.PivotFields('Customer Id').Subtotals=(False, False, False, False, False, False, False, False, False, False, False, False)
        PivotTable.PivotFields('Customer Name').Orientation = win32c.PivotFieldOrientation.xlRowField
        PivotTable.PivotFields('Customer Name').Subtotals=(False, False, False, False, False, False, False, False, False, False, False, False)
        # PivotTable.PivotFields('City').Orientation = win32c.PivotFieldOrientation.xlRowField
        # PivotTable.PivotFields('City').Subtotals=(False, False, False, False, False, False, False, False, False, False, False, False)
        # PivotTable.PivotFields('State').Orientation = win32c.PivotFieldOrientation.xlRowField
        # PivotTable.PivotFields('State').Subtotals=(False, False, False, False, False, False, False, False, False, False, False, False)
        # PivotTable.PivotFields('Zip Code').Orientation = win32c.PivotFieldOrientation.xlRowField
        # PivotTable.PivotFields('Zip Code').Subtotals=(False, False, False, False, False, False, False, False, False, False, False, False)
        ###logger.info("Adding particular Data Field in Pivot Table")
        PivotTable.PivotFields('Balance').Orientation = win32c.PivotFieldOrientation.xlDataField
        # PivotTable.PivotFields('Current').Orientation = win32c.PivotFieldOrientation.xlDataField
        # PivotTable.PivotFields('1 - 30').Orientation = win32c.PivotFieldOrientation.xlDataField
        # # PivotTable.PivotFields('Sum of  1 - 10').NumberFormat= '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'
        # PivotTable.PivotFields('31 - 60').Orientation = win32c.PivotFieldOrientation.xlDataField
        # # PivotTable.PivotFields('Sum of  31 - 60').NumberFormat= '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'
        # PivotTable.PivotFields('61 - 90').Orientation = win32c.PivotFieldOrientation.xlDataField
        # # PivotTable.PivotFields('Sum of  61 - 9999').NumberFormat= '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'
        # PivotTable.PivotFields('90+').Orientation = win32c.PivotFieldOrientation.xlDataField
        PivotTable.TableStyle2 = ""
        ###logger.info("Changing Table Layout in Pivot Table")
        PivotTable.RowAxisLayout(1)
        wb_open_ar.api.ActiveSheet.PivotTables("PivotTable1").PivotFields('Location').RepeatLabels = True
        # wb.api.ActiveSheet.PivotTables("PivotTable1").InGridDropZones = True
        # wb.api.ActiveSheet.PivotTables("PivotTable1").DataPivotField.Caption = "Values"
        time.sleep(1) 

        retry=0
        while retry < 10:
            try:
                wb_ap_ar = xw.Book(ap_ar_template) 
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry ==10:
                    raise e 

        ap_ar_tab3 = wb_ap_ar.sheets['Open & unsett AR']
        last_row3 = ap_ar_tab3.range(f'A'+ str(ap_ar_tab3.cells.last_cell.row)).end('up').row
        ap_ar_tab3.range(f"A2:F{last_row3}").delete()
        ws2.activate()
        last_rowp = ws2.range(f'A'+ str(ws2.cells.last_cell.row)).end('up').row
        ws2.range(f"A2:D{last_rowp-1}").copy(ap_ar_tab3.range(f"A2"))
        ap_ar_tab3.activate()
        last_row1 = ap_ar_tab3.range(f'A'+ str(ap_ar_tab3.cells.last_cell.row)).end('up').row
        ap_ar_tab3.range(f"D2:D{last_row1}").number_format='_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'
        # = VLOOKUP(A188,Location!G:H,2,0)(f column)
        # _("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)
        ws2p.activate()
        last_rowur = ws2p.range(f'A'+ str(ws2p.cells.last_cell.row)).end('up').row
        ws2p.range(f"A2:D{last_rowur-1}").copy(ap_ar_tab3.range(f"A{last_rowp}"))
        n_last_row = ap_ar_tab3.range(f'A'+ str(ap_ar_tab3.cells.last_cell.row)).end('up').row
        ap_ar_tab3.api.Range(f"D{last_rowp}:D{n_last_row}").Cut(ap_ar_tab3.api.Range(f"E{last_rowp}"))
        ap_ar_tab3.range(f"E{last_rowp}:E{n_last_row}").number_format='_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'
        ap_ar_tab3.api.Range(f"E2").Value = 0
        ap_ar_tab3.range(f"E2").number_format='_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'
        ap_ar_tab3.activate()
        time.sleep(1)
        ap_ar_tab3.api.Range(f"E2:E{last_row1}").Select()
        wb_ap_ar.app.api.Selection.FillDown()
        ap_ar_tab3.api.Range(f"D{last_row1+1}").Value = 0
        ap_ar_tab3.range(f"D{last_rowp}:D{n_last_row}").number_format='_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'
        time.sleep(1)
        ap_ar_tab3.api.Range(f"D{last_rowp}:D{n_last_row}").Select()
        wb_ap_ar.app.api.Selection.FillDown()
        ap_ar_tab3.api.Range(f"F{last_row1+1}").Value = f'= VLOOKUP(A{last_row1+1},Location!G:H,2,0)'
        time.sleep(1)
        ap_ar_tab3.api.Range(f"F{last_rowp}:F{n_last_row}").Select()
        wb_ap_ar.app.api.Selection.FillDown()        
        ap_ar_tab3.api.Range(f"F{last_rowp}:F{n_last_row}").Copy()
        ap_ar_tab3.api.Range(f"A{last_rowp}")._PasteSpecial(Paste=win32c.PasteType.xlPasteValues)
        wb_ap_ar.sheets.add("Pivot AR summary new",before=ap_ar_tab3)
        ###logger.info("Clearing contents for new sheet")
        wb_ap_ar.sheets["Pivot AR summary new"].clear_contents()
        ws2p2=wb_ap_ar.sheets["Pivot AR summary new"]
        ###logger.info("Declaring Variables for columns and rows")
        last_column = ap_ar_tab3.range('A1').end('right').last_cell.column
        last_column_letter=num_to_col_letters(ap_ar_tab3.range('A1').end('right').last_cell.column)
        last_row = ap_ar_tab3.range(f'A'+ str(ap_ar_tab3.cells.last_cell.row)).end('up').row
        ###logger.info("Creating Pivot Table")
        PivotCache=wb_ap_ar.api.PivotCaches().Create(SourceType=win32c.PivotTableSourceType.xlDatabase, SourceData=f"\'Open & unsett AR\'!R1C1:R{last_row}C{last_column}", Version=win32c.PivotTableVersionList.xlPivotTableVersion14)
        PivotTable = PivotCache.CreatePivotTable(TableDestination=f"'Pivot AR summary new'!R1C1", TableName="PivotTable2", DefaultVersion=win32c.PivotTableVersionList.xlPivotTableVersion14)        ###logger.info("Adding particular Row in Pivot Table") PivotTable.PivotFields('Location').Orientation = win32c.PivotFieldOrientation.xlRowField
        PivotTable.PivotFields('Location Name').Orientation = win32c.PivotFieldOrientation.xlRowField
        PivotTable.PivotFields('Location Name').Position = 1
        PivotTable.PivotFields('Location Name').Subtotals=(False, False, False, False, False, False, False, False, False, False, False, False)
        PivotTable.PivotFields('Customer Id').Orientation = win32c.PivotFieldOrientation.xlRowField
        PivotTable.PivotFields('Customer Id').Subtotals=(False, False, False, False, False, False, False, False, False, False, False, False)
        PivotTable.PivotFields('Customer Name').Orientation = win32c.PivotFieldOrientation.xlRowField
        PivotTable.PivotFields('Customer Name').Subtotals=(False, False, False, False, False, False, False, False, False, False, False, False)
        # PivotTable.PivotFields('City').Orientation = win32c.PivotFieldOrientation.xlRowField
        # PivotTable.PivotFields('City').Subtotals=(False, False, False, False, False, False, False, False, False, False, False, False)
        # PivotTable.PivotFields('State').Orientation = win32c.PivotFieldOrientation.xlRowField
        # PivotTable.PivotFields('State').Subtotals=(False, False, False, False, False, False, False, False, False, False, False, False)
        # PivotTable.PivotFields('Zip Code').Orientation = win32c.PivotFieldOrientation.xlRowField
        # PivotTable.PivotFields('Zip Code').Subtotals=(False, False, False, False, False, False, False, False, False, False, False, False)
        ###logger.info("Adding particular Data Field in Pivot Table")
        PivotTable.PivotFields('Unsettled AR').Orientation = win32c.PivotFieldOrientation.xlDataField
        PivotTable.PivotFields('Open AR').Orientation = win32c.PivotFieldOrientation.xlDataField
        # PivotTable.PivotFields('1 - 30').Orientation = win32c.PivotFieldOrientation.xlDataField
        # # PivotTable.PivotFields('Sum of  1 - 10').NumberFormat= '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'
        # PivotTable.PivotFields('31 - 60').Orientation = win32c.PivotFieldOrientation.xlDataField
        # # PivotTable.PivotFields('Sum of  31 - 60').NumberFormat= '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'
        # PivotTable.PivotFields('61 - 90').Orientation = win32c.PivotFieldOrientation.xlDataField
        # # PivotTable.PivotFields('Sum of  61 - 9999').NumberFormat= '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'
        # PivotTable.PivotFields('90+').Orientation = win32c.PivotFieldOrientation.xlDataField
        PivotTable.TableStyle2 = ""
        ###logger.info("Changing Table Layout in Pivot Table")
        PivotTable.RowAxisLayout(1)
        wb_ap_ar.api.ActiveSheet.PivotTables("PivotTable2").PivotFields('Location Name').RepeatLabels = True
        # wb_ap_ar.api.ActiveSheet.PivotTables("PivotTable1").InGridDropZones = True
        # wb_ap_ar.api.ActiveSheet.PivotTables("PivotTable2").DataPivotField.Caption = "Values"
        time.sleep(1) 
        ### date fromat changes ###
        input_tab.activate()
        input_tab.range(f"P:P").number_format="mm-dd-yyyy"
        input_tab2.activate()
        input_tab2.range(f"G:G").number_format="mm-dd-yyyy"
        input_tab2.range(f"I:I").number_format="mm-dd-yyyy"
        master_tab=wb.sheets[f"MASTER"]
        master_tab.activate()
        master_tab.range(f"P:P").number_format="mm-dd-yyyy"
        wb.save(f"{output_location2}\\Unsettled AR_"+input_date+'.xlsx') 

        wb_open_ar.save(f"{output_location2}\\Open AR_"+input_date+'.xlsx')

        wb_ap_ar.save(f"{output_location}\\WPLLC - AP_AR_Template.xlsx")
        try:
            wb.app.quit()
        except:
            wb.app.quit()  
        return f"{job_name} Report for {input_date} generated succesfully"

    except Exception as e:
        wb.app.kill()
        raise e
    finally:
        try:
            wb.app.quit()
        except:
            pass

def payroll_pdf_extractor(input_pdf, input_datetime, monthYear):
    try:
        main_dict = {}
        count = 0
        for loc in glob.glob(input_pdf):       #add month difference if ==2 then not consider that file
            file_date = loc.split()[-1].split(".pdf")[0].replace(".","-")
            file_datetime = datetime.strptime(loc.split()[-1].split(".pdf")[0],"%m.%d.%y")
            file_date = datetime.strftime(file_datetime, "%d-%m-%Y")
            diff = relativedelta(input_datetime.replace(day=1),file_datetime.replace(day=1))
            diff = diff.months*(diff.years+1)

            if diff == 0: # or diff == 1 or diff==-1:
                if count == 2:
                    raise Exception(f"3rd file found for input month {monthYear}")
                count+=1
                date_df = read_pdf(loc, pages = 1, guess = False, stream = True ,
                                    pandas_options={'header':0}, area = ["30,290,120,415"], columns=["320"])[0]
                dates = date_df.iloc[0,1].split("to")
                beg_date = datetime.strptime(dates[0].strip(), "%m/%d/%Y")
                end_date = datetime.strptime(dates[1].strip(), "%m/%d/%Y")
                monthYear1 = datetime.strftime(datetime.strptime(dates[0].strip(), "%m/%d/%Y"), "%b %y")
                monthYear2 = datetime.strftime(datetime.strptime(dates[1].strip(), "%m/%d/%Y"), "%b %y")
                # if monthYear1 == monthYear or monthYear2 == monthYear:
                pdfReader = PyPDF2.PdfFileReader(loc)
                
                
                for page in range(pdfReader.numPages - 1):
                    
                    pageObj = pdfReader.getPage(page)
                    a=pageObj.extractText()
                    
                    ada_group = int(a.split('Totals for Department: ')[1].split("-")[0].strip())
                    
                    
                    df = read_pdf(loc, pages = page+1, guess = False, stream = True ,
                                        pandas_options={'header':0}, area = ["150,5,560,850"], columns=["65,120,145,200,330,380,430,470,700,750"])[0]
                    # print(df)
                    gross_df = read_pdf(loc, pages = page+1, guess = False, stream = True ,
                                        pandas_options={'header':0}, area = ["60,300,190,400"])[0]
                    # print(gross_df)
                    gross_value = float(gross_df.iloc[-1,-1].replace(",",""))
                    state_fed_df = df.iloc[:,:4]
                    state_fed_df = state_fed_df[state_fed_df[state_fed_df.columns[0]].notna()].reset_index(drop=True)
                    state_taxable_df = df.iloc[:,4:8]
                    state_taxable_df = state_taxable_df[state_taxable_df[state_taxable_df.columns[0]].notna()].reset_index(drop=True)
                    deduc_ana_df = df.iloc[:,8:]
                    deduc_ana_df = deduc_ana_df[deduc_ana_df[deduc_ana_df.columns[0]].notna()].reset_index(drop=True)
                    deduc_ana_df = deduc_ana_df[deduc_ana_df[deduc_ana_df.columns[-1]].notna()].reset_index(drop=True)

                    medicare_ee = 0  #ER-Med	       Medicare -EE R
                    soc_sec_er = 0   #ER-SS            Social Security - ER
                    futa_nesui = 0   #FUTA             NESUI
                    suta_cosui = 0   #SUTA             COSUI
                    suta_wysui = 0   #SUTA             WYSUI
                    sui=0            #For all other sui
                    ffcra = 0        #FFCRA            Value Not Received Till Now ( Blank )
                    benefits = 0     #Benefits         Value Not Received Till Now ( Blank )
                    med_dent_vis = 0 # Med/Dent/Vis	   Total Value of Cafeteria 125 Deds
                    voluntary = 0 #Voluntary            Sum of All Misc. Expenses with no Parent Name (Deduction Analysis )
                    garnish_chldi = 0 #Garnishment     Deduction Analysis – CHLD1+GARN1
                    ee_401k = 0 #EE 401k               Deduction Analysis 401K
                    er_401k = 0 #ER401k	               Deduction Analysis  401L1   4ROTH
                    ee_roth = 0 #EE Roth 	           Deduction Analysis  4ROTH   Value Not Received Till Now ( Blank )
                    kln_401 = 0 #401KLN	               Deduction Analysis  401L2        401L1
                
                    for col in range(len(state_fed_df)):
                        if state_fed_df[state_fed_df.columns[0]][col] == "Medicare-ER":
                            if "("  in state_fed_df[state_fed_df.columns[-1]][col] and ")" in state_fed_df[state_fed_df.columns[-1]][col]:
                                medicare_ee = float(state_fed_df[state_fed_df.columns[-1]][col].replace(",","").replace("(","").replace(")",""))*-1
                            else:
                                medicare_ee = float(state_fed_df[state_fed_df.columns[-1]][col].replace(",",""))
                        elif state_fed_df[state_fed_df.columns[0]][col] == "Social Security-" and state_fed_df[state_fed_df.columns[0]][col+1] == "ER":
                            if "("  in state_fed_df[state_fed_df.columns[-1]][col] and ")" in state_fed_df[state_fed_df.columns[-1]][col]:
                                soc_sec_er = float(state_fed_df[state_fed_df.columns[-1]][col].replace(",","").replace("(","").replace(")",""))*-1
                            else:
                                soc_sec_er = float(state_fed_df[state_fed_df.columns[-1]][col].replace(",",""))
                    
                    for col in range(len(state_taxable_df)):
                        if state_taxable_df[state_taxable_df.columns[0]][col] == "NESUI":
                            if "("  in state_taxable_df[state_taxable_df.columns[-1]][col] and ")" in state_taxable_df[state_taxable_df.columns[-1]][col]:
                                futa_nesui = float(state_taxable_df[state_taxable_df.columns[-1]][col].replace(",","").replace("(","").replace(")",""))*-1
                            else:
                                futa_nesui = float(state_taxable_df[state_taxable_df.columns[-1]][col].replace(",",""))
                        elif state_taxable_df[state_taxable_df.columns[0]][col] == "COSUI":
                            if "("  in state_taxable_df[state_taxable_df.columns[-1]][col] and ")" in state_taxable_df[state_taxable_df.columns[-1]][col]:
                                suta_cosui = float(state_taxable_df[state_taxable_df.columns[-1]][col].replace(",","").replace("(","").replace(")",""))*-1
                            else:
                                suta_cosui = float(state_taxable_df[state_taxable_df.columns[-1]][col].replace(",",""))
                        elif state_taxable_df[state_taxable_df.columns[0]][col] == "WYSUI":
                            if "("  in state_taxable_df[state_taxable_df.columns[-1]][col] and ")" in state_taxable_df[state_taxable_df.columns[-1]][col]:
                                suta_wysui = float(state_taxable_df[state_taxable_df.columns[-1]][col].replace(",","").replace("(","").replace(")",""))*-1
                            else:
                                suta_wysui = float(state_taxable_df[state_taxable_df.columns[-1]][col].replace(",",""))

                        elif "SUI" in state_taxable_df[state_taxable_df.columns[0]][col].upper():
                            if "("  in state_taxable_df[state_taxable_df.columns[-1]][col] and ")" in state_taxable_df[state_taxable_df.columns[-1]][col]:
                                sui = float(state_taxable_df[state_taxable_df.columns[-1]][col].replace(",","").replace("(","").replace(")",""))*-1
                            else:
                                sui = float(state_taxable_df[state_taxable_df.columns[-1]][col].replace(",",""))
                
                    for col in range(len(deduc_ana_df)):
                        if deduc_ana_df[deduc_ana_df.columns[0]][col] == "Cafeteria 125":
                            if deduc_ana_df.iloc[-1,0]!="Cafeteria 125":
                                while deduc_ana_df[deduc_ana_df.columns[0]][col] !="Total":
                                    col+=1
                                if "("  in deduc_ana_df[deduc_ana_df.columns[-1]][col] and ")" in deduc_ana_df[deduc_ana_df.columns[-1]][col]:
                                    med_dent_vis = float(deduc_ana_df[deduc_ana_df.columns[-1]][col].replace(",","").replace("(","").replace(")",""))*-1
                                else:    
                                    med_dent_vis = float(deduc_ana_df[deduc_ana_df.columns[-1]][col].replace(",",""))*-1
                                break
                        
                        # elif deduc_ana_df[deduc_ana_df.columns[0]][col] == "CHLD1" or deduc_ana_df[deduc_ana_df.columns[0]][col] == "GARN1":
                        #     if "("  in deduc_ana_df[deduc_ana_df.columns[-1]][col] and ")" in deduc_ana_df[deduc_ana_df.columns[-1]][col]:
                        #         garnish_chldi += float(deduc_ana_df[deduc_ana_df.columns[-1]][col].replace(",","").replace("(","").replace(")",""))
                        #     else:
                        #         garnish_chldi += float(deduc_ana_df[deduc_ana_df.columns[-1]][col].replace(",",""))*-1
                        # elif deduc_ana_df[deduc_ana_df.columns[0]][col] == "401K":
                        #     if "("  in deduc_ana_df[deduc_ana_df.columns[-1]][col] and ")" in deduc_ana_df[deduc_ana_df.columns[-1]][col]:
                        #         ee_401k = float(deduc_ana_df[deduc_ana_df.columns[-1]][col].replace(",","").replace("(","").replace(")",""))
                        #     else:
                        #         ee_401k = float(deduc_ana_df[deduc_ana_df.columns[-1]][col].replace(",",""))*-1
                        # elif deduc_ana_df[deduc_ana_df.columns[0]][col] == "401L1":
                        #     if "("  in deduc_ana_df[deduc_ana_df.columns[-1]][col] and ")" in deduc_ana_df[deduc_ana_df.columns[-1]][col]:
                        #         er_401k = float(deduc_ana_df[deduc_ana_df.columns[-1]][col].replace(",","").replace("(","").replace(")",""))
                        #     else:
                        #         er_401k = float(deduc_ana_df[deduc_ana_df.columns[-1]][col].replace(",",""))*-1
                        # elif deduc_ana_df[deduc_ana_df.columns[0]][col] == "401L2":
                        #     if "("  in deduc_ana_df[deduc_ana_df.columns[-1]][col] and ")" in deduc_ana_df[deduc_ana_df.columns[-1]][col]:
                        #         kln_401 = float(deduc_ana_df[deduc_ana_df.columns[-1]][col].replace(",","").replace("(","").replace(")",""))
                        #     else:
                        #         kln_401 = float(deduc_ana_df[deduc_ana_df.columns[-1]][col].replace(",",""))*-1
                        # elif deduc_ana_df[deduc_ana_df.columns[0]][col] == "4ROTH":
                        #     if "("  in deduc_ana_df[deduc_ana_df.columns[-1]][col] and ")" in deduc_ana_df[deduc_ana_df.columns[-1]][col]:
                        #         ee_roth = float(deduc_ana_df[deduc_ana_df.columns[-1]][col].replace(",","").replace("(","").replace(")",""))
                        #     else:
                        #         ee_roth = float(deduc_ana_df[deduc_ana_df.columns[-1]][col].replace(",",""))*-1
                        else:
                            if deduc_ana_df[deduc_ana_df.columns[0]][col] != "Total":
                                if "("  in deduc_ana_df[deduc_ana_df.columns[-1]][col] and ")" in deduc_ana_df[deduc_ana_df.columns[-1]][col]:
                                    voluntary += float(deduc_ana_df[deduc_ana_df.columns[-1]][col].replace(",","").replace("(","").replace(")",""))*-1
                                else:
                                    voluntary += float(deduc_ana_df[deduc_ana_df.columns[-1]][col].replace(",",""))*-1
                    if file_date in main_dict.keys():  
                        
                        # main_dict[file_date][ada_group] = {"Gross":gross_value, "ER- SS":soc_sec_er, "ER - Med":medicare_ee, "FUTA":futa_nesui, "SUTA":suta_cosui+suta_wysui, "FFCRA": ffcra,
                        #     "Benefits":benefits, "Med/Dent/Vis":med_dent_vis, "Voluntary ":volutary, "Garnishment":garnish_chldi, "EE 401k ":ee_401k, "ER 401K":er_401k,
                        #     "EE Roth":ee_roth, "401KLN":kln_401}

                        main_dict[file_date][ada_group] = {"Gross Paid":gross_value, "SS - ER":soc_sec_er, "Medicare - ER":medicare_ee, "SUI":futa_nesui+suta_cosui+suta_wysui+sui,
                            "Cafeteria 125 Deductions":med_dent_vis, "Voluntary Deductions":voluntary, "Beggining Date":beg_date, "Ending Date":end_date}
                    else:  
                        main_dict[file_date] = {}
                        # main_dict[file_date][ada_group] = {"Gross":gross_value, "ER- SS":soc_sec_er, "ER - Med":medicare_ee, "FUTA":futa_nesui, "SUTA":suta_cosui+suta_wysui, "FFCRA": ffcra,
                        #         "Benefits":benefits, "Med/Dent/Vis":med_dent_vis, "Voluntary ":volutary, "Garnishment":garnish_chldi, "EE 401k ":ee_401k, "ER 401K":er_401k,
                        #         "EE Roth":ee_roth, "401KLN":kln_401}

                        main_dict[file_date][ada_group] = {"Gross Paid":gross_value, "SS - ER":soc_sec_er, "Medicare - ER":medicare_ee, "SUI":futa_nesui+suta_cosui+suta_wysui+sui,
                            "Cafeteria 125 Deductions":med_dent_vis, "Voluntary Deductions":voluntary, "Beggining Date":beg_date, "Ending Date":end_date}
                        
                        
            
        return main_dict
    except Exception as e:
        raise e


def other_loc_extractor(input_pdf):
    try:
        df = read_pdf(input_pdf, pages = 'all', guess = False, stream = True,
                                                pandas_options={'header':0}, area = ["50,200,580,740"], columns = ["290, 340, 490,590,640"])
        df = pd.concat(df, ignore_index=True)
        print(df)
        df = df[['Location','Product', 'Unit Cost']]
        df.set_index(['Location'])["Product"].to_dict()
        loc_dict = {}
        product=None
        for i in range(len(df)):
            
            
            if not pd.isnull(df.loc[:,'Location'][i]):
                location = df['Location'][i]
                # if location == "NGREEL":
                #     location = "NORTH GREELEY"
                if location == "OMA COMM":
                    location = "TERMINAL"
                # if location == "BROWNSVILL":
                #     location = "BROWNSVILLE"
            product = df['Product'][i]
            value = df['Unit Cost'][i]
            if product in loc_dict.keys():  
                    if location in loc_dict[product].keys():
                        loc_dict[product][location].append(value)
                    else:
                        loc_dict[product][location] = [value]
            else:  
                loc_dict[product] = {}
                loc_dict[product][location] = [value]

        print()
        return loc_dict
    except Exception as e:
        raise e

def mac_accr_pdf(input_pdf):
    try:
        acc_dict = {}
        
        acc_no = None
        pdfReader = PyPDF2.PdfFileReader(input_pdf)
        for page in range(pdfReader.numPages):
            
            pageObj = pdfReader.getPage(page)
            a=pageObj.extractText()

            if "MARKET REVALUATION" in a:
                acc_no = a[a.find('Account'):a.find('Account')+17]
                acc_no = acc_no.replace("Account: ","")
                #taking acc_no last 3 digits
                acc_no = acc_no[-3:]
                # print(f"account_num = {acc_no}, prev_acc = {prev_acc_no} and page is {page}")
                if acc_no == "":
                    continue
                # if prev_acc_no is None:
                #     prev_acc_no=acc_no #a[25:42]
                # elif prev_acc_no != acc_no:
                #     print(page-1)
                    
                #     print(acc_no)
                #     if str(prev_acc_no) in input_pdf:
                df2 = None
                # df = read_pdf(input_pdf, pages = page+1, guess = False, stream = True ,
                #             pandas_options={'header':0}, area = ["50,10,725,850"], columns=["195,280,430"])
                df = read_pdf(input_pdf, pages = page+1, guess = False, stream = True ,
                            pandas_options={'header':0}, area = ["50,10,740,850"], columns=["195,280,430"])
                df = pd.concat(df, ignore_index=True)
                print(df)

                i=0
                while df.iloc[i,0]!='MARKET REVALUATION':
                    i+=1
                j=0
                try:
                    while df.iloc[j,0]!='TOTALS':
                        j+=1
                except Exception as e:
                    df2=read_pdf(input_pdf, pages = page+2, guess = False, stream = True ,
                            area = ["50,10,725,850"], columns=["195,280,430"])
                    df2 = pd.concat(df2, ignore_index=True)
                    print(df2)
                    k=0
                    try:
                        while df2.iloc[k,0]!='TOTALS':
                            k+=1
                    except Exception as e:
                        raise e

                if df2 is None or df2.iloc[0,0] == "TOTALS":    
                    df = df.iloc[i+3:j-1,:]
                else:
                    df.columns = df2.columns
                    df = pd.concat([df.iloc[i+3:,:], df2.iloc[:k-1,:]], ignore_index=True)

                df = df.dropna(subset=[df.columns[-1]])
                for i in range(len(df)):
                    
                    if df.iloc[i,3] != "Profit/Loss":
                        commodity = df.iloc[i,0]
                        # price = df.iloc[i,2]
                        valuation = df.iloc[i,3]
                        if acc_no in acc_dict.keys():  
                            
                            acc_dict[acc_no][commodity]= float(valuation.replace(',',''))
                            
                        else:  
                            acc_dict[acc_no] = {}
                            acc_dict[acc_no][commodity]= float(valuation.replace(',',''))
                        



                # amount_dict[prev_acc_no] = float(df.iloc[-1,-1].replace(",","")) 
                        
                # print(prev_acc_no)
                print()
                # prev_acc_no = acc_no
            elif page == (pdfReader.numPages - 1):
                df = read_pdf(input_pdf, pages = page+1, guess = False, stream = True ,
                            pandas_options={'header':0}, area = ["70,10,725,850"], columns=["195,280,430"])
                df = pd.concat(df, ignore_index=True)
                print(df)
                net_liq = float(df.iloc[-1,2].replace(",",""))

        return acc_dict, net_liq
    except Exception as e:
        raise e

def inv_mtm_pdf_data_extractor(input_date, f, hrw_pdf_loc=None, yc_pdf_loc=None, mtm_report=False):
    try:
        hrw_fut = None
        yc_fut = None
        # reader = PyPDF2.PdfFileReader(open(f, mode='rb' ))
        # n = reader.getNumPages() 
        inp_month_year = datetime.strptime(input_date,"%m.%d.%Y").replace(day=1)
        # data_list = []
        if mtm_report:
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

        try:
            old_pdf=True
            com_loc = list(com_loc[0].str.split('Commodity: ',expand=True)[1])
        except:
            try:
                old_pdf=False
                com_loc = list(com_loc[1])
            except Exception as e:
                raise e
        # loc_dict = dict(zip(com_loc, [[]]*len(com_loc)))
        loc_dict = defaultdict(list)
            
        for page in range(1,len(com_loc)+1):
            if old_pdf:
                df = read_pdf(f, pages = page, guess = False, stream = True ,
                            pandas_options={'header':0}, area = ["75,10,580,850"], columns=["65,85, 180,225, 260, 280,300,360,400,430,480,525,570,620,665,720"])
                # df = read_pdf(f, pages = page, guess = True, stream = True ,
                #         pandas_options={'header':0}, area = ["75,10,580,850"], columns=["50,85, 180,225, 260, 280,300,360,400,430,480,525,570,620,665,720"])
            else:
                df = read_pdf(f, pages = page, guess = False, stream = True ,
                            pandas_options={'header':0}, area = ["75,10,580,850"], columns=["65,85, 180,225, 260, 275,300,360,400,430,480,525,570,620,665,713"])
                # df = read_pdf(f, pages = page, guess = True, stream = True ,
                #             pandas_options={'header':0}, area = ["75,10,580,850"], columns=["47,85, 180,225, 255,275,300,360,400,430,480,525,570,620,665,713"])
            df = pd.concat(df, ignore_index=True)
            ########logger.info("Filtering only required columns")
            df = df.iloc[:,[0,1,2,3,-2,-1]]
            # df = df[df['Offsite Name Cont. No.'].str.contains("Company Owned Risk:"),df['Offsite Name Cont. No.'].str.contains("Unpriced Sales:")]
            df = df[(df['Offsite Name Cont. No.'].str.contains("Company Owned Risk:")) | (df['Offsite Name Cont. No.'].str.contains("priced Sales:"))]
            # for i in df.loc[:,"Offsite Name Cont. No."]:

            df["Quantity.5"].fillna(0, inplace=True)
            try:
                df["Value.5"].fillna(0, inplace=True)
            except:
                try:
                    df = read_pdf(f, pages = page, guess = False, stream = True ,
                            pandas_options={'header':0}, area = ["75,10,580,850"], columns=["65,85, 180,225, 260, 275,300,360,400,430,480,525,570,620,665,713"])
                    # df = read_pdf(f, pages = page, guess = True, stream = True ,
                    #         pandas_options={'header':0}, area = ["75,10,580,850"], columns=["46,85, 180,225, 255,275,300,360,400,430,480,525,570,620,665,713"])
                    df = pd.concat(df, ignore_index=True)
                    df["Value.5"].fillna(0, inplace=True)
                    old_pdf=False
                except Exception as e:
                    raise e

            df["Quantity.5"] = df["Quantity.5"].astype(str).str.replace("(","-").str.replace(",","").str.replace(")","").astype(float)
            df["Value.5"] = df["Value.5"].astype(str).str.replace("(","-").str.replace(",","").str.replace(")","").astype(float)

            for i in range(len(df)):
                try:
                    print(df.iloc[i,2]) #2 for "Offsite Name Cont. No."
                except:
                    continue
                if "priced Sales" in df.iloc[i,2]:
                    print("Unprised Value found")
                    if df.iloc[-2,2] == 'Unpriced Sales:' and df.iloc[-2,-2]==0: #pd.isna(df.iloc[-2,-1]):
                        pass
                    else:
                        df.iloc[i+1,-2] = df.iloc[i+1,-2] - df.iloc[i,-2]
                        df.iloc[i+1,-1] = df.iloc[i+1,-1] - df.iloc[i,-1]

                if i>0 and (isinstance(df.iloc[i-1,0], str) and isinstance(df.iloc[i,0], str)):
                    if df.iloc[i-1,0]==df.iloc[i,0] or ("ALLIANCE" in df.iloc[i-1,0] and "ALLIANCE" in df.iloc[i,0]):
                        #Price Remains last one
                        #Adding Quantity and Value
                        df.iloc[i,4] = df.iloc[i,4]+df.iloc[i-1,4]
                        df.iloc[i,5] = df.iloc[i,5]+df.iloc[i-1,5]
                        #droping i-1 index row
                        df.drop([df.index[i-1]], inplace=True)
                        pass

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
        
        if mtm_report:
            return loc_dict, hrw_fut, yc_fut
        else:
            return loc_dict
    except Exception as e:
        raise e

def storage_qty(input_date,input_qty_pdf, input_qty_xl, monthYear2, qty_loc_dict):
    try:
        output_loc = drive+r'\REPORT\Storage Month End Report\Output Files' + f'\\STORAGE QTY {monthYear2}.xlsx'
        page_df = read_pdf(input_qty_pdf,pages = 1,guess = False,stream = True,
                        pandas_options={'header':0},area = ["65,630,600,735"],columns=["675"])[0]
        try:
            old_file = True
            page_num = int(page_df['e Types'][3][-4:])
        except:
            try:
                old_file = False
                page_df = read_pdf(input_qty_pdf,pages = 1,guess = False,stream = True,
                            pandas_options={'header':0},area = ["65,630,600,735"],columns=["670"])[0]
                page_num = int(page_df['e Types'][3][-4:])
            except Exception as e:
                raise e

        
        loc_dict = {}
        
            
            # df = read_pdf(input_qty_pdf,pages = i,guess = False,stream = True,
            #         pandas_options={'header':0},area = ["65,630,580,735"],columns=["675"])[0]
            
            # location_df = read_pdf(input_qty_pdf,pages = i,guess = False,stream = True,
            #                 pandas_options={'header':0},area = ["5,15,80,300"],columns=["60"])[0]
        if old_file:
            df = read_pdf(input_qty_pdf,pages = f"1-{page_num}",guess = False,stream = True,
                    pandas_options={'header':0},area = ["65,630,580,735"],columns=["680"])
            
            location_df = read_pdf(input_qty_pdf,pages = f"1-{page_num}",guess = False,stream = True,
                            pandas_options={'header':0},area = ["5,15,80,300"],columns=["60"])
        else:
            df = read_pdf(input_qty_pdf,pages = f"1-{page_num}",guess = False,stream = True,
                    pandas_options={'header':0},area = ["65,630,580,735"],columns=["672"])
            
            location_df = read_pdf(input_qty_pdf,pages = f"1-{page_num}",guess = False,stream = True,
                            pandas_options={'header':0},area = ["5,15,80,295"],columns=["53"])
        for i in range(page_num):

            # loc_lst.append(location_df['Daily Position R'][0])
            # commodity_lst.append(location_df['Daily Position R'][1])
            location = location_df[i]['Daily Position R'][0].split('-')[0].strip()
            if location == "ALLIANCETE":
                location = "ALLIANCE TERMINAL"
            if location == "HAYSPRING":
                location = "HAY SPRINGS"
            if location == "BROWNSVILL":
                location = "BROWNSVILLE"
            if location == "WESTPLAINS":
                location = "HAY SPRINGS"
            if location == "NGREELEY":
                location = "NORTH GREELEY"
            if location == "PLAWES":
                location = "TERMINAL"
            commodity = location_df[i]['Daily Position R'][1].split(' ')[1].strip()
            if commodity == 'SUNFLWR':
                commodity = 'SUNFL'
            # loc_dict[location_df['Daily Position R'][0]] = location_df['Daily Position R'][1]
            value = df[i][df[i].columns[len(df[i].columns)-1]].tail(1)
            value = list(value)[0]
            if value == '(14)' or type(value) ==float :
                continue
                
            else:
                if location in loc_dict.keys():  
                    if commodity in loc_dict[location].keys():
                        
                        loc_dict[location][commodity].append(value)
                    else:
                        loc_dict[location][commodity] = [value]

                else:  
                    loc_dict[location] = {}
                    loc_dict[location][commodity] = [value]
        

        # try:
        #     AT = loc_dict['ALLIANCETE']
        #     del loc_dict['ALLIANCETE']
        #     loc_dict['ALLIANCE TERMINAL'] = AT
        # except:
        #     pass
        # try:
        #     BW = loc_dict['BROWNSVILL']
        #     del loc_dict['BROWNSVILL']
        #     loc_dict['BROWNSVILLE'] = BW
        # except:
        #     pass
        # try:
        #     HS = loc_dict['WESTPLAINS']
        #     del loc_dict['WESTPLAINS']
        #     loc_dict['HAY SPRINGS'] = HS
        # except:
        #     pass
        # try:
        #     NG = loc_dict['NGREELEY']
        #     del loc_dict['NGREELEY']
        #     loc_dict['NORTH GREELEY'] = NG
        # except:
        #     pass

        # try:
        #     OMAHA = loc_dict['PLAWES']
        #     del loc_dict['PLAWES']
        #     loc_dict['TERMINAL'] = OMAHA
        # except:
        #     pass

        print()
        retry = 0
        while retry<10:
            try:
                wb = xw.Book(input_qty_xl, update_links=False)
                break
            except:
                time.sleep(5)
                retry+=1

        retry = 0
        while retry<10:
            try:
                ws1 = wb.sheets['Storage Accrual (2)']
                break
            except:
                time.sleep(5)
                retry+=1
        ws1.range('A3').value = input_date
        # xl_commodity = ws1.range('C5').expand('right').value
        last_row =  ws1.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row
        # col_lst = ws1.range("C5").expand('right').value

        for i in range(6,int(last_row)+1):
            if ws1.range(f'A{i}').value is not None:
                if  ws1.range(f'A{i}').value in loc_dict.keys():
                    for j in range(len(ws1.range("C5").expand('right'))):
                        if ws1.range(chr(ord("C")+j)+"5").value != 'TOTALS':
                            try:
                                ws1.range(chr(ord("C")+j)+f"{i}").value = qty_loc_dict[ws1.range(f"A{i}").value][ws1.range(chr(ord("C")+j)+"5").value]
                            except:
                                ws1.range(chr(ord("C")+j)+f"{i}").value = 0
                            try:
                                ws1.range(chr(ord("C")+j)+f"{i+1}").value = loc_dict[ws1.range(f"A{i}").value][ws1.range(chr(ord("C")+j)+"5").value]
                                    
                            except Exception as e:
                                ws1.range(chr(ord("C")+j)+f"{i+1}").value = 0
                        else:
                            pass
        wb.save(output_loc)
    except Exception as e:
        raise e
    finally:
        try:
            wb.app.quit()
        except:
            pass


def storage_accrual(input_date,strg_accr_inp_loc, monthYear, loc_dict):
    try:
        output_location = drive+r'\REPORT\Storage Month End Report\Output Files'+f"\\STORAGE ACCRUAL {monthYear}.xlsx"
        # output_location = r'C:\Users\imam.khan\OneDrive - BioUrja Trading LLC\Documents\WEST PLAINS\REPORT\Storage Month End Report\Output Files'+f"\\{monthYear}.xlsx"
        retry=0
        while retry < 10:
            try:
                wb=xw.Book(strg_accr_inp_loc)
                break
            except Exception as e:
                time.sleep(2)
                retry+=1
                if retry ==9:
                    raise e
        retry=0
        while retry < 10:
            try:
                accr_sht = wb.sheets[0]
                break
            except Exception as e:
                time.sleep(2)
                retry+=1
                if retry ==9:
                    raise e
        print()
        accr_sht.range("A5").value = f"Schedule of inventory held for third parties (open storage ticket report) as of {input_date}"
        last_row = accr_sht.range(f'A'+ str(accr_sht.cells.last_cell.row)).end('up').row

        for i in range(10,last_row):
            if accr_sht.range(f"A{i}").value is not None:
                if accr_sht.range(f"A{i}").value in loc_dict.keys():
                    print(accr_sht.range(f"A{i}").value)
                    for j in range(len(accr_sht.range("C9").expand('right'))):
                        try:
                            accr_sht.range(chr(ord("C")+j)+f"{i}").value = loc_dict[accr_sht.range(f"A{i}").value][accr_sht.range(chr(ord("C")+j)+"9").value]
                        except:
                            pass
        wb.save(output_location)
        
        print()
        return f"Storage Accrual Sheet Generated for {monthYear}"
    except Exception as e:
        raise e
    finally:
        try:
            wb.app.quit()
        except:
            pass

def storage_je(strg_je_inp_loc, input_date, loc_dict):
    try:
        xl_inp_date = datetime.strftime(datetime.strptime(input_date, "%m.%d.%Y"), "%m/%d/%Y")
        output_location = drive+r'\REPORT\Storage Month End Report\Output Files'+"\\STORAGE ACCRUAL JE_" +f"{input_date}.xlsx"
        # output_location = r'C:\Users\imam.khan\OneDrive - BioUrja Trading LLC\Documents\WEST PLAINS\REPORT\Storage Month End Report\Output Files'+"\\STORAGE ACCRUAL JE_" +f"{input_date}.xlsx"
        JE_dict = {'ALLIANCETE':'ALLIANCE TERMINAL','BATESLAND':'BATESLAND','BROWNSVILLE':'BROWNSVILL','CHADRON':'CHADRON','CLINTON':'CLINTON',
                    'CRAWFORD':'CRAWFORD','GERING':'GERING','HAYSPRG':'HAY SPRINGS','JTELEV':'JOHNSTOWN',
                    'LINGLE':'LINGLE','MITCHELL':'MITCHELL','NGREEL':'NORTH GREELEY','PLATNER':'PLATNER','TERMINAL':'TERMINAL','YUMA':'YUMA'}
                
        retry=0
        while retry < 10:
            try:
                wb=xw.Book(strg_je_inp_loc)
                break
            except Exception as e:
                time.sleep(2)
                retry+=1
                if retry ==9:
                    raise e
        retry=0
        while retry < 10:
            try:
                JE_sht = wb.sheets[0]
                break
            except Exception as e:
                time.sleep(2)
                retry+=1
                if retry ==9:
                    raise e
        print()
        JE_sht.range("A1").value = xl_inp_date
        last_row = JE_sht.range(f'A'+ str(JE_sht.cells.last_cell.row)).end('up').row

        for i in range(4,last_row+1):
            if JE_sht.range(f"B{i}").value is not None:
                try:
                    JE_sht.range(f"G{i}").value = loc_dict[JE_dict[JE_sht.range(f"B{i}").value]][JE_sht.range(f"E{i}").value]
                except:
                    JE_sht.range(f"G{i}").value = 0
        num_row = JE_sht.range('A3').end('down').row
        num_col = JE_sht.range('A3').end('right').column
       
        retry=0
        while retry<15:
            try:
                pivot_sht = wb.sheets["JE"]
                time.sleep(5)
                # pivot_sht.select()
                pivot_sht.activate()
                time.sleep(1)
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry==15:
                    raise e
        pivotCount = wb.api.ActiveSheet.PivotTables().Count
         # 'INPUT DATA'!$A$3:$I$86
        for j in range(1, pivotCount+1):
            if wb.api.ActiveSheet.PivotTables(j).PivotCache().SourceData != f"'INPUT DATA'!R3C1:R{num_row}C{num_col}": #Updateing data source
                wb.api.ActiveSheet.PivotTables(j).PivotCache().SourceData = f"'INPUT DATA'!R3C1:R{num_row}C{num_col}" #Updateing data source
            wb.api.ActiveSheet.PivotTables(j).PivotCache().Refresh()

        wb.save(output_location)
       
        print()
        
    except Exception as e:
        raise e
    finally:
        try:
            wb.app.quit()
        except:
            pass



def bbr_other_tabs(input_date, wb, input_ar, input_ctm):
    try:
        # input_date = "02.07.2022"
        # input_xl = drive+r"\REPORT\BBR Reports\Raw Files" +f"\\{input_date}_Borrowing Base Report.xlsx"
        # input_xl = r"C:\Users\Yashn.jain\Desktop\WEST PLAINS\REPORT\BBR Reports\Raw Files"+f"\\{input_date}_Borrowing Base Report.xlsx"
        # input_ar = drive+r"\REPORT\Open AR\Output files"+f"\\Open AR _{input_date} - Production.xlsx"
        # input_ar = r"C:\Users\Yashn.jain\Desktop\WEST PLAINS\REPORT\Open AR\Output files"+f"\\Open AR _{input_date} - Production.xlsx"
        # input_ctm = drive+r"\REPORT\CTM Combined report\Output files"+f"\\CTM Combined _{input_date}.xlsx"
        # input_ctm = r"C:\Users\Yashn.jain\Desktop\WEST PLAINS\REPORT\CTM Combined report\Output files"+f"\\CTM Combined _{input_date}.xlsx"
        # output_location=drive+r"\REPORT\BBR Reports\Output files"
        # output_location=r"C:\Users\Yashn.jain\Desktop\Sample_BBR"
        retry=0
        while retry < 10:
            try:
                ar_wb=xw.Book(input_ar)
                break
            except Exception as e:
                time.sleep(2)
                retry+=1
                if retry ==9:
                    raise e           
        wsar1=ar_wb.sheets["Eligible"]
        wsar1.activate()
        last_row = wsar1.range(f'A'+ str(wsar1.cells.last_cell.row)).end('up').row
        column_list = wsar1.range("A1").expand('right').value
        total_column=column_list.index('total')+1
        total_letter_column = num_to_col_letters(column_list.index('total')+1)
        # ar_wb.app.quit()
        ar_wb.close()

        # retry=0
        # while retry < 10:
        #     try:
        #         wb=xw.Book(input_xl)
        #         break
        #     except Exception as e:
        #         time.sleep(2)
        #         retry+=1
        #         if retry ==9:
        #             raise e


        ws1=wb.sheets["AR-Trade By Tier - Eligible"]
        ws1.select()
        ws1.clear_contents()
        ws1.activate()
        # pivotCount = wb.api.ActiveSheet.PivotTables().Count
        # #'\\Bio-India-FS\India sync$\WEST PLAINS\REPORT\BBR Reports\Raw Files\[Open AR _02.07.2022 - Production.xlsx]Eligible'!$A$1:$K$123
        # # 'Data 02.21.2022'!$A$1:$G$4731
        # #'\\Bio-India-FS\India sync$\WEST PLAINS\REPORT\BBR Reports\Raw Files\[Open AR _02.07.2022 - Production.xlsx]Eligible'!$A$1:$K$123
        # for j in range(1, pivotCount+1):
        #     wb.api.ActiveSheet.PivotTables("PivotTable1").PivotSelect("Tier[All]", win32c.PTSelectionMode.xlLabelOnly,True)
        #     # wb.api.ActiveSheet.PivotTables(j).PivotCache().SourceData = f"'J:\WEST PLAINS\REPORT\Open AR\Output files\[Open AR _{input_date} - Production]Eligible'!R1C1:R{last_row}C{total_column}"
        #     wb.api.ActiveSheet.PivotTables(j).PivotCache().Refresh()  

        ###logger.info("Adding Worksheet for Pivot Table")
        # wb.sheets.add("AR-Trade By Tier - Eligible2",after=wb.sheets["Account Receivable Summary"])
        ###logger.info("Clearing contents for new sheet")
        # wb.sheets["AR-Trade By Tier - Eligible2"].clear_contents()
        # ws2=wb.sheets["AR-Trade By Tier - Eligible2"]
        ###logger.info("Declaring Variables for columns and rows")
        # last_row = ws5.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row
        # last_column = ws5.range('A1').end('right').last_cell.column
        # last_column_letter=num_to_col_letters(ws5.range('A1').end('right').last_cell.column)
        ###logger.info("Creating Pivot Table")
        PivotCache=wb.api.PivotCaches().Create(SourceType=win32c.PivotTableSourceType.xlDatabase, SourceData=f"'J:\\WEST PLAINS\\REPORT\\Open AR\\Output files\\[Open AR _{input_date} - Production.xlsx]Eligible'!R1C1:R{last_row}C{total_column}", Version=win32c.PivotTableVersionList.xlPivotTableVersion14)
        PivotTable = PivotCache.CreatePivotTable(TableDestination=f"'AR-Trade By Tier - Eligible'!R7C1", TableName="PivotTable1", DefaultVersion=win32c.PivotTableVersionList.xlPivotTableVersion14)        ###logger.info("Adding particular Row in Pivot Table")
        PivotTable.PivotFields('Tier').Orientation = win32c.PivotFieldOrientation.xlRowField
        PivotTable.PivotFields('Tier').Position = 1
        PivotTable.PivotFields('Customer Name').Orientation = win32c.PivotFieldOrientation.xlRowField
        ###logger.info("Adding particular Data Field in Pivot Table")
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
        ###logger.info("Adding particular Page Field in Pivot Table")
        PivotTable.PivotFields('Eligiblity').Orientation = win32c.PivotFieldOrientation.xlPageField
        ###logger.info("Applying filter in Data Field in Pivot Table")
        PivotTable.PivotFields('Eligiblity').CurrentPage = "Eligible"
        ###logger.info("Changing No Format in Pivot Table")
        # PivotTable.RowAxisLayout(1)
        ###logger.info("Changing Table Style in Pivot Table")
        PivotTable.TableStyle2 = ""
        ###logger.info("Changing Table Layout in Pivot Table")
        PivotTable.RowAxisLayout(1)
        wb.api.ActiveSheet.PivotTables("PivotTable1").InGridDropZones = True
        wb.api.ActiveSheet.PivotTables("PivotTable1").DataPivotField.Caption = "Data"

        # ws1.api.Range("A1:A3").Copy()
        # ws2.api.Paste()
        # wb.app.api.CutCopyMode=False
        # ws1.delete()
        # ws2.name="AR-Trade By Tier - Eligible"
        ws1.range("A1").value = "West Plains, LLC"  
        ws1.range("A2").value = "Open Accounts Receivable -  by Tier"
        ws1.range("A3").formula = "='Cash Collateral'!A3"
        ws1.api.Range("A3").NumberFormat = 'mm/dd/yyyy'


        ws3=wb.sheets["AR-Trade By Tier - Ineligible"]
        ws3.select()
        ws3.clear_contents()
        # wb.sheets.add("AR-Trade By Tier - Ineligible2",after=wb.sheets["AR-Trade By Tier - Eligible"])
        # ###logger.info("Clearing contents for new sheet")
        # wb.sheets["AR-Trade By Tier - Ineligible2"].clear_contents()
        # ws4=wb.sheets["AR-Trade By Tier - Ineligible2"]
        ###logger.info("Declaring Variables for columns and rows")
        # last_row = ws5.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row
        # last_column = ws5.range('A1').end('right').last_cell.column
        # last_column_letter=num_to_col_letters(ws5.range('A1').end('right').last_cell.column)
        ###logger.info("Creating Pivot Table")
        PivotCache=wb.api.PivotCaches().Create(SourceType=win32c.PivotTableSourceType.xlDatabase, SourceData=f"'J:\\WEST PLAINS\\REPORT\\Open AR\\Output files\\[Open AR _{input_date} - Production.xlsx]Eligible'!R1C1:R{last_row}C{total_column}", Version=win32c.PivotTableVersionList.xlPivotTableVersion14)
        PivotTable = PivotCache.CreatePivotTable(TableDestination=f"'AR-Trade By Tier - Ineligible'!R7C1", TableName="PivotTable1", DefaultVersion=win32c.PivotTableVersionList.xlPivotTableVersion14)        ###logger.info("Adding particular Row in Pivot Table")
        PivotTable.PivotFields('Tier').Orientation = win32c.PivotFieldOrientation.xlRowField
        PivotTable.PivotFields('Tier').Position = 1
        PivotTable.PivotFields('Customer Name').Orientation = win32c.PivotFieldOrientation.xlRowField
        ###logger.info("Adding particular Data Field in Pivot Table")
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
        ###logger.info("Adding particular Page Field in Pivot Table")
        PivotTable.PivotFields('Eligiblity').Orientation = win32c.PivotFieldOrientation.xlPageField
        ###logger.info("Applying filter in Data Field in Pivot Table")
        PivotTable.PivotFields('Eligiblity').CurrentPage = "Ineligible"
        ###logger.info("Changing No Format in Pivot Table")
        # PivotTable.RowAxisLayout(1)
        ###logger.info("Changing Table Style in Pivot Table")
        PivotTable.TableStyle2 = ""
        ###logger.info("Changing Table Layout in Pivot Table")
        PivotTable.RowAxisLayout(1)
        wb.api.ActiveSheet.PivotTables("PivotTable1").InGridDropZones = True
        wb.api.ActiveSheet.PivotTables("PivotTable1").DataPivotField.Caption = "Data"

        ws3.range("A1").value = "West Plains, LLC"  
        ws3.range("A2").value = "Open Accounts Receivable -  by Tier"
        ws3.range("A3").formula = "='Cash Collateral'!A3"
        ws3.api.Range("A3").NumberFormat = 'mm/dd/yyyy'

        # ws3.api.Range("A1:A3").Copy()
        # ws4.api.Paste()
        # wb.app.api.CutCopyMode=False
        # ws3.delete()
        # ws4.name="AR-Trade By Tier - Ineligible"
        # ws5=wb.sheets['Detail CTM Non MCUI']
        ar_re_last_row = wb.sheets['AR-Re-Purchase Storage Rcbl'].range(f'I' + str(wb.sheets['AR-Re-Purchase Storage Rcbl'].cells.last_cell.row)).end('up').row
        wb.sheets["Account Receivable Summary"].range("C8").formula = '=+GETPIVOTDATA("Sum of  1 - 10",\'AR-Trade By Tier - Eligible\'!$A$7,"Tier","Tier I")'
        wb.sheets["Account Receivable Summary"].range("E8").formula = '=+GETPIVOTDATA("Sum of  1 - 10",\'AR-Trade By Tier - Eligible\'!$A$7,"Tier","Tier II")'
        wb.sheets["Account Receivable Summary"].formula = "='Cash Collateral'!A3"
        wb.sheets["Account Receivable Summary"].api.Range("A3").NumberFormat = 'mm/dd/yyyy'
        wb.sheets["Account Receivable Summary"].range("C11").formula = f'=\'AR-Re-Purchase Storage Rcbl\'!I{ar_re_last_row}'
        retry=0
        while retry < 10:
            try:
                wb1=xw.Book(input_ctm)
                break
            except Exception as e:
                time.sleep(2)
                retry+=1
                if retry ==9:
                    raise e 

        excl_sht = wb1.sheets("Excl Macq & IC")
        ##logger.info("Copy tier sheet AFTER the intercompany sheet of input book.")
        # num_row = excl_sht.range('A1').end('down').row
        num_row=excl_sht.range(f'A' + str(excl_sht.cells.last_cell.row)).end('up').row

        #Condtion for removing only header line coming at last
        if excl_sht.range(f'I{num_row}').value.upper()!="EQUIP":
            if excl_sht.range(f'I{num_row}').value.upper()=="Commodity Id".upper():
                excl_sht.range(f"{num_row}:{num_row}").delete()
        
        last_column = excl_sht.range('A1').end('right').last_cell.column
        last_column_letter=num_to_col_letters(last_column)
        # excl_sht.range(f'A1:{last_column_letter}{num_row}').copy()
        wb.activate()
        ws5 = wb.sheets['Detail CTM Non MBL']
        ws5.clear_contents()
        excl_sht.range(f'A1:{last_column_letter}{num_row}').copy()
        wb.activate()
        ws5.range('A1').paste()
        wb.app.api.CutCopyMode=False
        wb1.activate()
        wb1.close()
        wb.activate()
        ws6 = wb.sheets['Unrlz Gains Contracts Non MBL']
        ws6.select()
        ws6.clear_contents()

        #logger.info("Adding Worksheet for Pivot Table")
        # wb.sheets.add("Unrlz- Gains- Contracts Non MC2",after=wb.sheets["Inventory -Other"])
        #logger.info("Clearing New Worksheet")
        # wb.sheets["Unrlz- Gains- Contracts Non MC2"].clear_contents()
        # ws7=wb.sheets["Unrlz- Gains- Contracts Non MC2"]
        #logger.info("Declaring Variables for columns and rows")
        last_column = ws5.range('A1').end('right').last_cell.column
        last_column_letter=num_to_col_letters(ws5.range('A1').end('right').last_cell.column)
        num_row = ws5.range('A1').end('down').row
        #logger.info("Creating Pivot table")
        PivotCache=wb.api.PivotCaches().Create(SourceType=win32c.PivotTableSourceType.xlDatabase, SourceData=f"\'Detail CTM Non MBL\'!R1C1:R{num_row}C{last_column}", Version=win32c.PivotTableVersionList.xlPivotTableVersion14)
        PivotTable = PivotCache.CreatePivotTable(TableDestination="'Unrlz Gains Contracts Non MBL'!R7C1", TableName="PivotTable1", DefaultVersion=win32c.PivotTableVersionList.xlPivotTableVersion14)
        #logger.info("Adding particular Row Data in Pivot Table")
        PivotTable.PivotFields('Location Id').Orientation = win32c.PivotFieldOrientation.xlRowField
        PivotTable.PivotFields('Location Id').Position = 1
        # PivotTable.PivotFields('Tier').RepeatLabels=True
        PivotTable.PivotFields('Commodity Id').Orientation = win32c.PivotFieldOrientation.xlRowField
        #logger.info("Adding particular Data Field in Pivot Table")
        PivotTable.PivotFields('Gain/LossTotal').Orientation = win32c.PivotFieldOrientation.xlDataField
        PivotTable.PivotFields('Sum of Gain/LossTotal').NumberFormat= '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
        #logger.info("Adding particular Page Field in Pivot Table")
        PivotTable.PivotFields('Ship Tier').Orientation = win32c.PivotFieldOrientation.xlPageField
        #logger.info("Applying filter in pagefield in Pivot Table")
        PivotTable.PivotFields('Ship Tier').CurrentPage = "W/n 12 Months"
        #logger.info("Changing No Format in Pivot Table")
        #logger.info("Changing Table layout")
        PivotTable.PivotFields('Location Id').Subtotals=(False, False, False, False, False, False, False, False, False, False, False, False)
        PivotTable.RowAxisLayout(1)
        #logger.info("Changing Table Style")
        PivotTable.TableStyle2 = ""
        wb.api.ActiveSheet.PivotTables("PivotTable1").InGridDropZones = True

        #logger.info("Declaring Variables for columns and rows")
        last_column = ws5.range('A1').end('right').last_cell.column
        last_column_letter=num_to_col_letters(ws5.range('A1').end('right').last_cell.column)
        num_row = ws5.range('A1').end('down').row
        last_row2 = ws6.range(f'A'+ str(ws6.cells.last_cell.row)).end('up').row
        last_row2+=10
        #logger.info("Creating Pivot table")
        PivotCache=wb.api.PivotCaches().Create(SourceType=win32c.PivotTableSourceType.xlDatabase, SourceData=f"\'Detail CTM Non MBL\'!R1C1:R{num_row}C{last_column}", Version=win32c.PivotTableVersionList.xlPivotTableVersion14)
        PivotTable = PivotCache.CreatePivotTable(TableDestination=f"'Unrlz Gains Contracts Non MBL'!R{last_row2}C1", TableName="PivotTable2", DefaultVersion=win32c.PivotTableVersionList.xlPivotTableVersion14)
        #logger.info("Adding particular Row Data in Pivot Table")
        PivotTable.PivotFields('Location Id').Orientation = win32c.PivotFieldOrientation.xlRowField
        PivotTable.PivotFields('Location Id').Position = 1
        # PivotTable.PivotFields('Tier').RepeatLabels=True
        PivotTable.PivotFields('Commodity Id').Orientation = win32c.PivotFieldOrientation.xlRowField
        PivotTable.PivotFields('Delivery End Date').Orientation = win32c.PivotFieldOrientation.xlRowField
        #logger.info("Adding particular Data Field in Pivot Table")
        PivotTable.PivotFields('Gain/LossTotal').Orientation = win32c.PivotFieldOrientation.xlDataField
        PivotTable.PivotFields('Sum of Gain/LossTotal').NumberFormat= '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
        #logger.info("Adding particular Page Field in Pivot Table")
        PivotTable.PivotFields('Ship Tier').Orientation = win32c.PivotFieldOrientation.xlPageField
        #logger.info("Applying filter in pagefield in Pivot Table")

        Event = False
        try:
            PivotTable.PivotFields('Ship Tier').CurrentPage = ">12 months"
            ###logger.info("Changing No Format in Pivot Table")
            ###logger.info("Changing Table layout")
            PivotTable.PivotFields('Location Id').Subtotals=(False, False, False, False, False, False, False, False, False, False, False, False)
            PivotTable.PivotFields('Commodity Id').Subtotals=(False, False, False, False, False, False, False, False, False, False, False, False)
            PivotTable.RowAxisLayout(1)
            ###logger.info("Changing Table Style")
            PivotTable.TableStyle2 = ""
            wb.api.ActiveSheet.PivotTables("PivotTable2").InGridDropZones = True
        except:
            wb.api.ActiveSheet.PivotTables("PivotTable2").PivotSelect("",win32c.PTSelectionMode.xlDataAndLabel,True) 
            wb.app.selection.api.ClearContents()
            Event = True
        finally:
            last_row3 = ws6.range(f'A'+ str(ws6.cells.last_cell.row)).end('up').row 
            last_row3+=5
        if Event:    
            ws6.range(f"E58").value=f'=+GETPIVOTDATA("Gain/LossTotal",$A$7)'
        else:
            ws6.range(f"E58").value=f'=+GETPIVOTDATA("Gain/LossTotal",$A$7)+GETPIVOTDATA("Gain/LossTotal",$A${last_row2})'  

        # PivotTable.PivotFields('Ship Tier').CurrentPage = ">12 months"
        # #logger.info("Changing No Format in Pivot Table")
        # #logger.info("Changing Table layout")
        # PivotTable.PivotFields('Location Id').Subtotals=(False, False, False, False, False, False, False, False, False, False, False, False)
        # PivotTable.PivotFields('Commodity Id').Subtotals=(False, False, False, False, False, False, False, False, False, False, False, False)
        # PivotTable.RowAxisLayout(1)
        # #logger.info("Changing Table Style")
        # PivotTable.TableStyle2 = ""
        # wb.api.ActiveSheet.PivotTables("PivotTable2").InGridDropZones = True
        #logic for adding total
        # last_row3 = ws6.range(f'A'+ str(ws6.cells.last_cell.row)).end('up').row 
        # last_row3+=5
        # ws6.range(f"E58").value=f'=+GETPIVOTDATA("Gain/LossTotal",$A$7)+GETPIVOTDATA("Gain/LossTotal",$A${last_row2})'
        ws6.range(f"E58").api.NumberFormat= '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'

        ws6.range("A1").value = "West Plains, LLC"  
        ws6.range("A2").value = "Net Unrealized Gains on Forward Contracts - Non MBL"
        ws6.range("A3").formula = "='Cash Collateral'!A3"
        ws6.api.Range("A3").NumberFormat = 'mm/dd/yyyy'
        wb.sheets["BBR"].range(f'G32').formula = "=+'Unrlz Gains Contracts Non MBL'!E58"
        print()
        # ws6.api.Range("A1:A3").Copy()
        # ws7.api.Paste()
        # ws7.api.Columns("C:C").ColumnWidth = 17
        # wb.app.api.CutCopyMode=False
        # ws6.delete()
        # ws7.name="Unrlz- Gains- Contracts Non MC"
        # wb.save(f"{output_location}\\{input_date}_Borrowing Base Report_y.xlsx")
        # wb.app.quit()
    except Exception as e:
        raise e

def cash_colat(wb,bank_recons_loc, input_date_date):
    try:
        
        retry=0
        while retry < 10:
            try:
                bank_wb=xw.Book(bank_recons_loc)
                break
            except Exception as e:
                time.sleep(2)
                retry+=1
                if retry ==9:
                    raise e

        while True:
            try:
                cash_colat_sht = wb.sheets["Cash Collateral"] #wb.sheets[0].name in 'Unsettled Receivables _'+input_date
                break
            except Exception as e:
                time.sleep(2)

        while True:
            try:
                bank_colat_sht = bank_wb.sheets["BANK REC"] #wb.sheets[0].name in 'Unsettled Receivables _'+input_date
                break
            except Exception as e:
                time.sleep(2)
        cash_colat_sht.range("A3").value = input_date_date
        cash_colat_sht.api.Range("A3").NumberFormat = 'mm/dd/yyyy'
        # cash_colat_sht.range("B58").value = bank_colat_sht.range("B12").value
        # cash_colat_sht.range("E58").value = bank_colat_sht.range("B14").value
        cash_colat_sht.range("B12").value = bank_colat_sht.range("B58").value
        cash_colat_sht.range("B14").value = -1*bank_colat_sht.range("E58").value

        jp_morgan_amount = -1*bank_colat_sht.range("E27").value
        bank_wb.close()
        return jp_morgan_amount
    except Exception as e:
        raise e

def ar_unsettled_by_tier(wb, unset_rec_loc, input_date):
    try:
        retry=0
        while retry < 10:
            try:
                unset_rec_wb=xw.Book(unset_rec_loc)
                break
            except Exception as e:
                time.sleep(2)
                retry+=1
                if retry ==9:
                    raise e

        while True:
            try:
                xl_mac_n_ic = unset_rec_wb.sheets["Excl Macq & IC"] #wb.sheets[0].name in 'Unsettled Receivables _'+input_date
                break
            except Exception as e:
                time.sleep(2)
        last_row=xl_mac_n_ic.range(f'A' + str(xl_mac_n_ic.cells.last_cell.row)).end('up').row
        
        # 'J:\WEST PLAINS\REPORT\BBR Reports\Output\[Unsettled Receivables _02.14.2022.xlsx]Excl IC & Macq'!$A$1:$AJ$892
        unset_rec_wb.close()
        while True:
            try:
                ar_unsettled_by_tier_sht = wb.sheets["AR Unsettled ByTier"] #wb.sheets[0].name in 'Unsettled Receivables _'+input_date
                break
            except Exception as e:
                time.sleep(2)

        while True:
            try:
                ar_unsettled_by_tier_sht.select()
                break
            except Exception as e:
                time.sleep(2)
        
        # sht = wb.sheets["AR-Trade By Tier - Eligible"]
        wb.api.ActiveSheet.PivotTables(1).PivotCache().SourceData = f"'J:\\WEST PLAINS\\REPORT\\Unsettled Receivables\\Output Files\\[Unsettled Receivables _{input_date}.xlsx]Excl Macq & IC'!R1C1:R{last_row}C36"
        
          #f'Details!R1C1:R{len(new_rows)+1}C18' #Updateing data source
        wb.api.ActiveSheet.PivotTables(1).PivotCache().Refresh()
        for pf in wb.api.ActiveSheet.PivotTables(1).DataFields:
            pf.NumberFormat = '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'

        ar_unsettled_by_tier_sht.api.Range("A3").Formula = "='Cash Collateral'!A3"
        ar_unsettled_by_tier_sht.api.Range("A3").NumberFormat = 'mm/dd/yyyy'
        print("Refreshed")
        print()
        

        pass
    except Exception as e:
        raise e

def comm_acc_pdf_ext(account_lst, pdf_loc):
    try:
        amount_dict = {}
        prev_acc_no = None
        acc_no = None
        pdfReader = PyPDF2.PdfFileReader(pdf_loc)
        for page in range(pdfReader.numPages):
            
            pageObj = pdfReader.getPage(page)
            a=pageObj.extractText()
            # acc_no = a[a.find('Account'):a.find('Account')+17]
            # acc_no = acc_no.replace("Account: ","")
            # print(f"account_num = {acc_no}, prev_acc = {prev_acc_no} and page is {page}")
            # if acc_no == "":
            #     continue
            # if prev_acc_no is None:
            #     prev_acc_no=acc_no #a[25:42]
            # elif prev_acc_no != acc_no:
            #     print(page-1)
                
            #     print(acc_no)
            #     if str(prev_acc_no) in account_lst:
            acc_no = a[a.find('Account'):a.find('Account')+17]
            acc_no = acc_no.replace("Account: ","")
            if str(acc_no) in account_lst:
                amount_dict[acc_no] = 0
                if "Net Liquidating Value" in a:
                        df = read_pdf(pdf_loc, pages = page+1, guess = False, stream = True ,
                                    pandas_options={'header':0}, area = ["75,10,725,850"], columns=["180,280"])
                        df = pd.concat(df, ignore_index=True)
                        print(df)
                        amount_dict[acc_no] = float(df.iloc[-1,-1].replace(",","")) 
                    # try:
                    #     amount_dict[acc_no] = float(df['Unnamed: 1'][len(df)-1].replace(",",""))
                    
                    # except:
                    #     try:
                            
                    #         amount_dict[acc_no] = float(df['NET USD'][len(df)-1].replace(",",""))
                    #     except:
                    #         try:
                    #             df = read_pdf(pdf_loc, pages = page, guess = False, stream = True ,
                    #                 pandas_options={'header':0}, area = ["100,10,580,850"], columns=["180,280"])
                    #             df = pd.concat(df, ignore_index=True)
                    #             print(df)
                    #             amount_dict[acc_no] = float(df['Unnamed: 1'][len(df)-1].replace(",",""))
                    #         except Exception as e:
                    #             raise e
                    # print(prev_acc_no)
                    # print()
                # prev_acc_no = acc_no

        return amount_dict
    except Exception as e:
        raise e
    
def comm_acc_xl(wb,pdf_loc):
    try:
        while True:
            try:
                com_acc_sht = wb.sheets["Commodity Accounts (NLV)"] #wb.sheets[0].name in 'Unsettled Receivables _'+input_date
                break
            except Exception as e:
                time.sleep(2)
        cell = 8
        com_acc_sht.api.Range("A3").Formula = "='Cash Collateral'!A3"
        com_acc_sht.api.Range("A3").NumberFormat = 'mm/dd/yyyy'
        account_lst = com_acc_sht.range("B8").expand("down").value
        account_lst = [str(account).replace(".0","") for account in account_lst]
        amount_dict = comm_acc_pdf_ext(account_lst, pdf_loc)
        for account in account_lst:
            try:
                com_acc_sht.range(f"C{cell}").value = amount_dict[account]
            except:
                com_acc_sht.range(f"C{cell}").value = None
            cell+=1
        print()
    except Exception as e:
        raise e

def ar_open_storage_rcbl(wb, strg_accr_loc, input_date):
    try:
        # p_m_last_date = datetime.strftime((datetime.strptime(input_date, "%m.%d.%Y").replace(day=1)-timedelta(days=1)), "%m.%d.%Y")
        # txt = f"Schedule of inventory held for third parties (open storage ticket report) as of {p_m_last_date}"

        
        retry=0
        while retry < 10:
            try:
                strg_accr_wb=xw.Book(strg_accr_loc)
                break
            except Exception as e:
                time.sleep(2)
                retry+=1
                if retry ==9:
                    raise e

        while True:
            try:
                strg_acc_sht = strg_accr_wb.sheets[0] #wb.sheets[0].name in 'Unsettled Receivables _'+input_date
                break
            except Exception as e:
                time.sleep(2)
                retry+=1
                if retry ==9:
                    raise e

        while True:
            try:
                bbr_strg_acc_sht = wb.sheets["AR-Open Storage Rcbl"] #wb.sheets[0].name in 'Unsettled Receivables _'+input_date
                break
            except Exception as e:
                time.sleep(2)
                retry+=1
                if retry ==9:
                    raise e
        # strg_acc_sht.copy(bbr_strg_acc_sht)
        bbr_strg_acc_sht.range("A5").value = strg_acc_sht.range("A5").value
        last_row = strg_acc_sht.range(f'A'+ str(strg_acc_sht.cells.last_cell.row)).end('up').row
        nxt_last_row = strg_acc_sht.range(f'A{last_row}').end('up').row
        print(nxt_last_row)
        strg_acc_sht.range(f"A10:M{nxt_last_row}").copy(bbr_strg_acc_sht.range("A10"))
        strg_accr_wb.close()
        print()
    except Exception as e:
        raise e

def inv_whre_n_in_trans(wb, mtm_loc, input_date):
    try:
        retry=0
        while retry < 10:
            try:
                mtm_wb=xw.Book(mtm_loc)
                break
            except Exception as e:
                time.sleep(2)
                retry+=1
                if retry ==9:
                    raise e

        while True:
            try:
                m_sht = mtm_wb.sheets[0] #wb.sheets[0].name in 'Unsettled Receivables _'+input_date
                break
            except Exception as e:
                time.sleep(2)

        while True:
            try:
                whre_sht = wb.sheets["Inventory Whre & In-Trans"] #wb.sheets[0].name in 'Unsettled Receivables _'+input_date
                break
            except Exception as e:
                time.sleep(2)
        while True:
            try:
                inv_oth_sht = wb.sheets["Inventory -Other"] #wb.sheets[0].name in 'Unsettled Receivables _'+input_date
                break
            except Exception as e:
                time.sleep(2)


        last_row=m_sht.range(f'A' + str(m_sht.cells.last_cell.row)).end('up').row
        main_loc = m_sht.range(f"A1:A{last_row}").value
        hrw_value=0
        yc_value = 0
        whre_sht.range(f"A3").formula = "='Cash Collateral'!A3"
        whre_sht.api.Range("A3").NumberFormat = 'mm/dd/yyyy'
        inv_oth_sht.range(f"A3").formula = "='Cash Collateral'!A3"
        inv_oth_sht.api.Range("A3").NumberFormat = 'mm/dd/yyyy'
        for i in range(len(main_loc)):

            if main_loc[i]=="HRW" and hrw_value==0:
                hrw = f"{i+1}"
                hrw_value+=1
            elif main_loc[i]=="HRW" and hrw_value==1:
                hrw_2 = i+3
                hrw_value+=1
            elif main_loc[i]=="YC" and yc_value == 0:
                yc = f"{i+1}"
                yc_value+=1

            elif main_loc[i]=="Commodity":
                other_loc = f"{i+3}"
            elif main_loc[i] == "FW":
                other_loc_2 = f"{i+1}"
            elif main_loc[i] == "Sunflowers":
                sunflwr = f"{i+1}"


        # whre_sht.range(f"C{hrw}").options(transpose=True).value = [float(n[0]) if n[0]!= "" else n[0] for n in m_sht.range(f"C{hrw}").expand("down").formula] #m_sht.range(f"C{hrw}").expand("down").value
        # whre_sht.range(f"F{hrw}").options(transpose=True).value = [float(n[0]) if n[0]!= "" else n[0] for n in m_sht.range(f"F{hrw}").expand("down").formula] #m_sht.range(f"F{hrw}").expand("down").value
        # whre_sht.range(f"I{hrw}").options(transpose=True).value = [float(n[0]) if n[0]!= "" else n[0] for n in m_sht.range(f"I{hrw}:I{int(yc)-4}").formula] #m_sht.range(f"I{hrw}:I{int(yc)-4}").value
        
        

        # whre_sht.range(f"C{yc}").options(transpose=True).value = [float(n[0]) if n[0]!= "" else n[0] for n in m_sht.range(f"C{yc}:C{int(other_loc_2)-5}").formula] #m_sht.range(f"C{yc}").expand("down").value
        # whre_sht.range(f"F{yc}").options(transpose=True).value = [float(n[0]) if n[0]!= "" else n[0] for n in m_sht.range(f"F{yc}:F{int(other_loc_2)-5}").formula] #m_sht.range(f"F{yc}").expand("down").value
        # whre_sht.range(f"I{yc}").options(transpose=True).value = [float(n[0]) if n[0]!= "" else n[0] for n in m_sht.range(f"I{yc}:I{int(other_loc_2)-5}").formula] #m_sht.range(f"I{yc}:I{int(other_loc_2)-5}").value

        # whre_sht.range(f"C{other_loc_2}").options(transpose=True).value = [float(n[0]) if n[0]!= "" else n[0] for n in m_sht.range(f"C{other_loc_2}").expand("down").formula] #m_sht.range(f"C{other_loc_2}").expand("down").value
        # whre_sht.range(f"F{other_loc_2}").options(transpose=True).value = [float(n[0]) if n[0]!= "" else n[0] for n in m_sht.range(f"F{other_loc_2}").expand("down").formula] #m_sht.range(f"F{other_loc_2}").expand("down").value


        # inv_oth_sht.range(f"C{int(other_loc)-64}").options(transpose=True).value = [float(n[0]) if n[0]!= "" else n[0] for n in m_sht.range(f"C{other_loc}:C{int(sunflwr)-6}").formula]
        # inv_oth_sht.range(f"F{int(other_loc)-64}").options(transpose=True).value = [float(n[0]) if n[0]!= "" else n[0] for n in m_sht.range(f"F{other_loc}:F{int(sunflwr)-6}").formula] #m_sht.range(f"F{other_loc}:F{int(sunflwr)-6}").value
        
        # inv_oth_sht.range(f"C{int(sunflwr)-64}").options(transpose=True).value = float(m_sht.range(f"C{sunflwr}").value)
        # inv_oth_sht.range(f"F{int(sunflwr)-64}").options(transpose=True).value = float(m_sht.range(f"F{sunflwr}").value)
        
        m_sht.range(f"C{hrw}").expand("down").copy(whre_sht.range(f"C{hrw}").options(transpose=True))
        m_sht.range(f"F{hrw}").expand("down").copy(whre_sht.range(f"F{hrw}").options(transpose=True))
        m_sht.range(f"I{hrw}:I{int(yc)-4}").copy(whre_sht.range(f"I{hrw}").options(transpose=True))
        
        

        m_sht.range(f"C{yc}").expand("down").copy(whre_sht.range(f"C{yc}").options(transpose=True))
        m_sht.range(f"F{yc}").expand("down").copy(whre_sht.range(f"F{yc}").options(transpose=True))
        m_sht.range(f"I{yc}:I{int(other_loc_2)-5}").copy(whre_sht.range(f"I{yc}").options(transpose=True))

        m_sht.range(f"C{other_loc_2}").expand("down").copy(whre_sht.range(f"C{other_loc_2}").options(transpose=True))
        m_sht.range(f"F{other_loc_2}").expand("down").copy(whre_sht.range(f"F{other_loc_2}").options(transpose=True))

        #Logic for copy pasting new YSB vlaues in cell I58 and J58
        m_sht.range(f"I{other_loc_2}").end("down").expand('right').copy(whre_sht.range(f"I{other_loc_2}").end("down").options(transpose=True))


        m_sht.range(f"C{other_loc}:C{int(sunflwr)-6}").copy(inv_oth_sht.range(f"C{int(other_loc)-64}").options(transpose=True))
        m_sht.range(f"F{other_loc}:F{int(sunflwr)-6}").copy(inv_oth_sht.range(f"F{int(other_loc)-64}").options(transpose=True))
        
        m_sht.range(f"C{sunflwr}").copy(inv_oth_sht.range(f"C{int(sunflwr)-64}").options(transpose=True))
        m_sht.range(f"F{sunflwr}").copy(inv_oth_sht.range(f"F{int(sunflwr)-64}").options(transpose=True))


        mtm_wb.close()
        
        print()
    except Exception as e:
        raise e

def payables(input_date,wb, bbr_mapping_loc, open_ap_loc,unset_pay_loc,jp_morgan_amount):
    try:
        df = pd.read_excel(bbr_mapping_loc, usecols="A,B")   
        new_dict = dict(zip(df.iloc[:,0],df.iloc[:,1]))
        inv_dict = dict(zip(df.iloc[:,1],df.iloc[:,0]))
        payab_df = pd.read_excel(bbr_mapping_loc, usecols="D,E")
        payab_dict = dict(zip(payab_df.iloc[:,0],payab_df.iloc[:,1]))
        inv_payab_dict = dict(zip(payab_df.iloc[:,1],payab_df.iloc[:,0]))
        retry=0
        while retry < 10:
            try:
                open_ap_wb=xw.Book(open_ap_loc)
                break
            except Exception as e:
                time.sleep(2)
                retry+=1
                if retry ==9:
                    raise e

        while True:
            try:
                open_ap_sht = open_ap_wb.sheets["Pivot BB"] #wb.sheets[0].name in 'Unsettled Receivables _'+input_date
                break
            except Exception as e:
                time.sleep(2)
        retry=0
        while retry < 10:
            try:
                payab_wb=xw.Book(unset_pay_loc)
                break
            except Exception as e:
                time.sleep(2)
                retry+=1
                if retry ==9:
                    raise e

        while True:
            try:
                payab_sht = payab_wb.sheets["Pivot BB"] #wb.sheets[0].name in 'Unsettled Receivables _'+input_date
                break
            except Exception as e:
                time.sleep(2)
        while True:
            try:
                bbr_payab_sht = wb.sheets["Payables"] #wb.sheets[0].name in 'Unsettled Receivables _'+input_date
                break
            except Exception as e:
                time.sleep(2)

        
        f_last_row = open_ap_sht.range("A5").end('down').row
        open_ap_loc_lst = open_ap_sht.range(f"A5:A{int(f_last_row)-1}").value

        last_col_num = open_ap_sht.range("A4").expand("right").last_cell.column
        last_col = num_to_col_letters(last_col_num)
        total_col = open_ap_sht.range(f"{last_col}5:{last_col}{int(f_last_row)-1}").value

        dict_1 = dict(zip(open_ap_loc_lst,total_col))

        last_row=open_ap_sht.range(f'A' + str(open_ap_sht.cells.last_cell.row)).end('up').row
        last_pivot = int(open_ap_sht.range(f"A{last_row}").end('up').row)+2
        open_ap_loc_2_lst = open_ap_sht.range(f"A{last_pivot}:A{int(last_row)-1}").value

        last_col_num = open_ap_sht.range(f"A{last_pivot-1}").expand("right").last_cell.column
        last_col = num_to_col_letters(last_col_num)
        grnd_ttl = open_ap_sht.range(f"{last_col}{last_pivot}:{last_col}{int(last_row)-1}").value

        dict_2 = dict(zip(open_ap_loc_2_lst,grnd_ttl))

        bbr_loc = bbr_payab_sht.range("A10").expand("down").value
        bbr_last_row = bbr_payab_sht.range("A10").end("down").row
        #inserting row
        if len(bbr_loc) != len(open_ap_loc_lst):
            if len(open_ap_loc_lst) > len(bbr_loc):
                for i in range(len(open_ap_loc_lst) - len(bbr_loc)):
                    bbr_payab_sht.range(f"{bbr_last_row+i+1}:{bbr_last_row+i+1}").insert()
                    new_loc = f"{bbr_last_row+i+1}"
            # else:
            #     for i in range(len(bbr_loc) - len(open_ap_loc_lst)):
            #         bbr_payab_sht.range(f"{bbr_last_row-i}:{bbr_last_row-i}").delete()
        else:
            pass

        i=10
        for loc in open_ap_loc_lst:
            try:
                if inv_dict[loc] not in bbr_loc:
                    bbr_payab_sht.range(f"A{new_loc}").value = inv_dict[loc]
                    bbr_payab_sht.range(f"C{int(new_loc)+1}").formula = f"=+SUM(C10:C{new_loc})"
                    bbr_payab_sht.range(f"D{int(new_loc)+1}").formula = f"=+SUM(D10:D{new_loc})"
                    bbr_payab_sht.range(f"E{int(new_loc)+1}").formula = f"=+SUM(E10:E{new_loc})"
                    bbr_payab_sht.range(f"F{int(new_loc)+1}").formula = f"=+SUM(F10:F{new_loc})"
                    bbr_payab_sht.range(f"F{int(new_loc)}").formula = f"=C{int(new_loc)}-D{int(new_loc)}-E{int(new_loc)}"
            except:
                pass
        first_loc = bbr_payab_sht.range("A3").end("down").row+1
        bbr_o_ap_last_row = bbr_payab_sht.range(f"A{first_loc}").end("down").row
        for i in range(first_loc,bbr_o_ap_last_row+1):      
            # if new_dict[loc] == bbr_payab_sht.range(f"A{i}").value:
                # bbr_payab_sht.range(f"A{i}").value = new_dict[loc]
            try:
                bbr_payab_sht.range(f"C{i}").value = dict_1[new_dict[bbr_payab_sht.range(f"A{i}").value]]
            except:
                bbr_payab_sht.range(f"C{i}").value = 0
            try:
                bbr_payab_sht.range(f"E{i}").value = dict_2[new_dict[bbr_payab_sht.range(f"A{i}").value]]
            except:
                bbr_payab_sht.range(f"E{i}").value = 0
            # i+=1
            # elif bbr_payab_sht.range(f"A{i}").value is None:
            #     bbr_payab_sht.range(f"A{i}").value = new_dict[loc]
            #     try:
            #         bbr_payab_sht.range(f"C{i}").value = dict_1[loc]
            #     except:
            #         bbr_payab_sht.range(f"C{i}").value = 0
            #     try:
            #         bbr_payab_sht.range(f"E{i}").value = dict_2[loc]
            #     except:
            #         bbr_payab_sht.range(f"E{i}").value = 0
            #     i+=1

        # for loc in open_ap_loc_lst:
        #     bbr_payab_sht.range(f"A{i}").value = new_dict[loc]
        #     try:
        #         bbr_payab_sht.range(f"C{i}").value = dict_1[loc]
        #     except:
        #         bbr_payab_sht.range(f"C{i}").value = 0
        #     try:
        #         bbr_payab_sht.range(f"E{i}").value = dict_2[loc]
        #     except:
        #         bbr_payab_sht.range(f"E{i}").value = 0
        #     i+=1
        p_last_row = payab_sht.range("A4").end('down').row
        payab_loc_lst = payab_sht.range(f"A4:A{int(p_last_row)-1}").value
        total_col = payab_sht.range(f"D4:D{int(p_last_row)-1}").value

        

        dict_3 = dict(zip(payab_loc_lst,total_col))

        payb_loc = bbr_payab_sht.range(f"A{bbr_o_ap_last_row}").end("down").end("down").row
        payb_last_loc = bbr_payab_sht.range(f"A{bbr_o_ap_last_row}").end("down").end("down").end("down").row
        bbr_payb_loc = bbr_payab_sht.range(f"A{payb_loc}:A{int(payb_last_loc)-1}").expand('down').value

        bbr_payb_loc_lst = bbr_payab_sht.range(f"A{payb_loc}").expand("down").value
        bbr_payb_loc_lst = bbr_payb_loc_lst[:-1]
        #inserting row
        if len(bbr_payb_loc_lst) != len(payab_loc_lst):
            if len(payab_loc_lst) > len(bbr_payb_loc_lst):
                for i in range(len(payab_loc_lst) - len(bbr_payb_loc_lst)):
                    bbr_payab_sht.range(f"{payb_last_loc+i}:{payb_last_loc+i}").insert()
                    new_loc = f"{payb_last_loc+i}"
            # else:
            #     for i in range(len(bbr_loc) - len(open_ap_loc_lst)):
            #         bbr_payab_sht.range(f"{bbr_last_row-i}:{bbr_last_row-i}").delete()
        else:
            pass
        for loc in payab_loc_lst:
            try:
                if payab_dict[loc] not in bbr_payb_loc:
                    bbr_payab_sht.range(f"A{new_loc}").value = inv_dict[loc]
                    bbr_payab_sht.range(f"C{int(new_loc)+1}").formula = f"=+SUM(C10:C{new_loc})"
                    bbr_payab_sht.range(f"D{int(new_loc)+1}").formula = f"=+SUM(D10:D{new_loc})"
                    bbr_payab_sht.range(f"E{int(new_loc)+1}").formula = f"=+SUM(E10:E{new_loc})"
                    
                    bbr_payab_sht.range(f"F{int(new_loc)+1}").formula = f"=C{new_loc}-D{new_loc}-E{new_loc}"
            except:
                pass
        for i in range(payb_loc, payb_last_loc):
            # bbr_payab_sht.range(f"A{i}").value = payab_dict[loc]
            try:
                bbr_payab_sht.range(f"C{i}").value = dict_3[inv_payab_dict[bbr_payab_sht.range(f"A{i}").value]]
            except:
                bbr_payab_sht.range(f"C{i}").value = 0
            # payb_loc+=1
        bbr_payab_sht.range("A3").formula = "='Cash Collateral'!A3"
        bbr_payab_sht.api.Range("A3").NumberFormat = 'mm/dd/yyyy'
        bbr_payab_sht.api.Range("C:F").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        bbr_payab_sht.range("I14").value = jp_morgan_amount
        bbr_payab_sht.api.Range("I14").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        open_ap_wb.close()
        payab_wb.close()

        
        print()
    except Exception as e:
        traceback.print_exc()
        raise e


def moc_get_df_from_input_excel(mtm_file, open_ap_file, open_ar_file,unsettled_pay_file, unsettled_recev_file):
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
            ws_mtm = wb_mtm.sheets['JE']
            last_row = ws_mtm.range(f'A'+ str(ws_mtm.cells.last_cell.row)).end('up').row
            first_row  = ws_mtm.range(f"A{last_row}").end('up').last_cell.row
            req_index = first_row + 1
            df_mtm = pd.read_excel(mtm_file,sheet_name='JE', usecols="A,B", skiprows=req_index)   
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
            raise e
        finally:
            try:
                wb_mtm.app.quit()
            except Exception as e:
                pass
        
        """"This is the code for Open AP files"""
        try:
            inner_keys = ['Alliance/Hay Springs','Gering','Omaha','Johnstown','KC','West Coast','BROWNSVILL']
            inner_dict = {}.fromkeys(inner_keys)
            # df_ap = pd.read_excel(open_ap_file,sheet_name='For allocation entry',usecols="A,B", skiprows=2)
            df_ap = pd.read_excel(open_ap_file,sheet_name = 0, usecols="A,B", skiprows=2)

            new_dict = dict(zip(df_ap.iloc[:,0],df_ap.iloc[:,1]))
            inner_dict['Alliance/Hay Springs'] = new_dict['HAYSPRG']
            inner_dict['Gering'] = new_dict.get('GERING')
            inner_dict['Omaha'] = new_dict.get('TERMINAL')
            inner_dict['Johnstown'] = new_dict.get('OMA COMM') + new_dict.get('JTELEV')
            inner_dict['KC'] = new_dict.get('KANSAS CTY')
            try:
                inner_dict['West Coast'] = new_dict.get('WEST COAST')
            except:
                inner_dict["West Coast"] = None
            inner_dict['BROWNSVILL'] = new_dict.get('BROWNSVILL')
            req_dict['Accounts Payable'] = inner_dict
        except Exception as e:
            print(e)
            print("The format of input file is wrong for Open AP or the file does not exist. Please enter the correct format")
        
        """"This is the code for Open AR files"""
        try:
            inner_keys = ['Alliance/Hay Springs','Gering','Omaha','Johnstown','KC','West Coast','BROWNSVILL']
            inner_dict = {}.fromkeys(inner_keys)
            # df_ar = pd.read_excel(open_ar_file, sheet_name='For allocation entry',usecols="A,B", skiprows=2)
            df_ar = pd.read_excel(open_ar_file, sheet_name = 0, usecols="A,B", skiprows=2)
            new_dict = dict(zip(df_ar.iloc[:,0],df_ar.iloc[:,1]))
            inner_dict['Alliance/Hay Springs'] = new_dict['HAYSPRG']
            inner_dict['Gering'] = new_dict.get('GERING')
            inner_dict['Omaha'] = new_dict.get('TERMINAL')
            inner_dict['Johnstown'] = new_dict.get('OMA COMM') + new_dict.get('JTELEV')
            inner_dict['KC'] = new_dict.get('KANSAS CTY')
            try:
                inner_dict['West Coast'] = new_dict.get('WEST COAST')
            except:
                inner_dict["West Coast"] = None
            inner_dict['BROWNSVILL'] = new_dict.get('BROWNSVILL')
            req_dict['Open A/R'] = inner_dict
        except Exception as e:
            print(e)
            print("The format of input file is wrong for Open AR or the file does not exist. Please enter the correct format")
        
        """This is the code for Unsettled Payables files"""
        try:
            inner_keys = ['Alliance/Hay Springs','Gering','Omaha','Johnstown','KC','West Coast','BROWNSVILL']
            inner_dict = {}.fromkeys(inner_keys)
            # df_pay = pd.read_excel(unsettled_pay_file, sheet_name = 'For allocation entry', usecols="A,B", skiprows=2)
            df_pay = pd.read_excel(unsettled_pay_file, sheet_name = 0, usecols="A,B", skiprows=2)
            new_dict = dict(zip(df_pay.iloc[:,0],df_pay.iloc[:,1]))
            inner_dict['Alliance/Hay Springs'] = new_dict['HAY SPRINGS - WEST PLAINS, LLC']
            inner_dict['Gering'] = new_dict.get('GERING - WEST PLAINS, LLC')
            inner_dict['Omaha'] = new_dict.get('OMAHA TERMINAL ELEVATOR - WEST PLAINS, LLC')
            inner_dict['Johnstown'] = new_dict.get('OMAHA COMM - WEST PLAINS, LLC') + new_dict.get('JOHNSTOWN - WEST PLAINS, LLC')
            inner_dict['KC'] = new_dict.get('KANSAS CTY')
            try:
                inner_dict['West Coast'] = new_dict.get('WEST COAST')
            except:
                inner_dict["West Coast"] = None
            inner_dict['BROWNSVILL'] = new_dict.get('BROWNSVILLE - WEST PLAINS, LLC')
            req_dict['Unsettled A/P'] = inner_dict
        except Exception as e:
            print(e)
            print("The format of input file is wrong for Unsettled A/P or the file does not exist. Please enter the correct format")
            
        """This is the code for Unsettled Receivables"""
        try:
            inner_keys = ['Alliance/Hay Springs','Gering','Omaha','Johnstown','KC','West Coast','BROWNSVILL']
            inner_dict = {}.fromkeys(inner_keys)
            # df_recev = pd.read_excel(unsettled_recev_file, sheet_name = 'For allocation entry', usecols="A,B", skiprows=2)
            df_recev = pd.read_excel(unsettled_recev_file, sheet_name = 0, usecols="A,B", skiprows=2)
            new_dict = dict(zip(df_recev.iloc[:,0],df_recev.iloc[:,1]))
            inner_dict['Alliance/Hay Springs'] = new_dict['HAY SPRINGS - WEST PLAINS, LLC']
            inner_dict['Gering'] = new_dict.get('GERING - WEST PLAINS, LLC')
            inner_dict['Omaha'] = new_dict.get('OMAHA TERMINAL ELEVATOR - WEST PLAINS, LLC')
            inner_dict['Johnstown'] = new_dict.get('OMAHA COMM - WEST PLAINS, LLC') + new_dict.get('JOHNSTOWN - WEST PLAINS, LLC')
            inner_dict['KC'] = new_dict.get('KANSAS CTY')
            try:
                inner_dict['West Coast'] = new_dict.get('WEST COAST')
            except:
                inner_dict["West Coast"] = None
            inner_dict['BROWNSVILL'] = new_dict.get('BROWNSVILLE - WEST PLAINS, LLC')
            req_dict['Unsettled A/R'] = inner_dict
        except Exception as e:
            print(e)
            print("The format of input file is wrong for Unsettled A/R or the file does not exist. Please enter the correct format")
            
        
        main_df = pd.DataFrame(req_dict)
        print("Main dataframe created")
        return main_df
    except Exception as e:
        raise e
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
                ws_alloc.range('M9:M15').options(transpose=True).value = main_df.values[5]
                ws_alloc.range('P9:P15').options(transpose=True).value = main_df.values[6]

                ws_alloc.range('E9:P15').api.NumberFormat = '_("$"* #,##0_);_("$"* (#,##0);_("$"* "-"??_);_(@_)'

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
                        elif key == 'M17':
                            ws_alloc.range('M29:M35').options(transpose=True).value = main_df.values[5]
                        elif key == 'P17':
                            ws_alloc.range('P29:P35').options(transpose=True).value = main_df.values[6]
                else:        
                    ws_alloc.range('E29:E35').options(transpose=True).value = main_df.values[0]
                    ws_alloc.range('F29:F35').options(transpose=True).value = main_df.values[1]
                    ws_alloc.range('G29:G35').options(transpose=True).value = main_df.values[2]
                    ws_alloc.range('I29:I35').options(transpose=True).value = main_df.values[3]
                    ws_alloc.range('J29:J35').options(transpose=True).value = main_df.values[4]
                    ws_alloc.range('M29:M35').options(transpose=True).value = main_df.values[5]
                    ws_alloc.range('P29:P35').options(transpose=True).value = main_df.values[6]

                # ws_alloc.range('E37:p37').formula = '=+E29+E30+E31-E32-E33-E34-E35'
                # ws_alloc.range('E39:p39').formula = '=E37/$Q$37'
                # ws_alloc.range('E40:p40').formula = '=E39*$E$62'

                ws_alloc.range('E29:P35').api.NumberFormat = '_("$"* #,##0_);_("$"* (#,##0);_("$"* "-"??_);_(@_)'
                wb_alloc.save(output_dir + '\\' + file.replace(file.split('_')[1],input_date) + '.xls')                
                print(f"MOC Allocment file generated for {input_date}")
    except Exception as e:
        print("Template file was not found or some other issue occured")
        raise e
    finally:
        try:
            wb_alloc.app.quit()
        except Exception as e:
            pass



def mtm_pdf_data_extractor(input_date, f, hrw_pdf_loc=None, yc_pdf_loc=None ,ysb_pdf_loc=None, mtm_excel_summ=False):
    try:
        # reader = PyPDF2.PdfFileReader(open(f, mode='rb' ))
        # n = reader.getNumPages() 
        inp_month_year = datetime.strptime(input_date,"%m.%d.%Y").replace(day=1)
        # data_list = []
        option = Options()
        option.headless=True
        driver=webdriver.Firefox(executable_path=GeckoDriverManager().install(),options=option)
        for loc in [hrw_pdf_loc, yc_pdf_loc , ysb_pdf_loc]:
            if loc is not None:
                date_datetime = datetime.strptime(input_date,"%m.%d.%Y")
                dmonth = date_datetime.strftime("%m")
                dday = date_datetime.strftime("%d")
                dyear = date_datetime.strftime("%Y")
                if loc == hrw_pdf_loc:
                    driver.get(f"https://www.cmegroup.com/markets/agriculture/grains/kc-wheat.settlements.html#tradeDate={dmonth}%2F{dday}%2F{dyear}")
                elif loc == yc_pdf_loc:   
                    driver.get(f"https://www.cmegroup.com/markets/agriculture/grains/corn.settlements.html#tradeDate={dmonth}%2F{dday}%2F{dyear}")
                elif loc == ysb_pdf_loc:
                    driver.get(f"https://www.cmegroup.com/markets/agriculture/oilseeds/soybean.settlements.html#tradeDate={dmonth}%2F{dday}%2F{dyear}")                
                time.sleep(5)
                table_element=WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.CLASS_NAME, "table-wrapper")))
                df=pd.read_html(table_element.get_attribute('outerHTML'),header=0,skiprows=[0])[0]      
                # df = read_pdf(loc, pages = 1, guess = False, stream = True ,
                #                         pandas_options={'header':0}, area = ["700,70,1000,1200"], columns=['150','480','550','650', '700','800','900'])
                # df = pd.concat(df, ignore_index=True)
                if df.iloc[0,0]=="Month":
                    df.columns = df.iloc[0]
                    df = df[1:]
                    df = df.reset_index(drop=True)
                df = df[["Month","Settle"]]
                df.columns = ["MONTH","SETTLE"]
                form_dict = {"'6":"75", "'4":"50", "'2":"25", "'0":"00"}
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
                                elif 'YSB' in loc.upper():
                                    ysb_fut =  int(settle_price.replace(key,form_dict[key]))/10000                                
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
                                elif 'YSB' in loc.upper():
                                    ysb_fut =  int(settle_price.replace(key,form_dict[key]))/10000                                
                                break
                        break
                    
                    
        try:
            driver.quit() 
        except:
            print("could not close the driver object") 
        date_df = read_pdf(f, pages = 1, guess = False, stream = True ,
                        pandas_options={'header':None}, area = ["20,40,40,800"])
        print(date_df)
        # pdf_date = date_df[0][0][0].split()[-1]

        com_loc  = read_pdf(f, pages = 'all', guess = False, stream = True ,
                        pandas_options={'header':None}, area = ["30,15,50,120"])
        com_loc = pd.concat(com_loc, ignore_index=True)

        try:
            old_pdf=True
            com_loc = list(com_loc[0].str.split('Commodity: ',expand=True)[1])
        except:
            try:
                old_pdf=False
                com_loc = list(com_loc[1])
            except Exception as e:
                raise e
        # loc_dict = dict(zip(com_loc, [[]]*len(com_loc)))
        loc_dict = defaultdict(list)
        for page in range(1,len(com_loc)+1):
            if old_pdf:
                df = read_pdf(f, pages = page, guess = True, stream = True ,
                            pandas_options={'header':0}, area = ["75,10,580,850"], columns=["50,85, 180,225, 260, 280,300,360,400,430,480,525,570,620,665,720"])
            else:
                df = read_pdf(f, pages = page, guess = True, stream = True ,
                                pandas_options={'header':0}, area = ["75,10,580,850"], columns=["47,85, 180,225, 255,275,300,360,400,430,480,525,570,620,665,713"])
            df = pd.concat(df, ignore_index=True)
            ########logger.info("Filtering only required columns")
            df = df.iloc[:,[0,1,2,3,-2,-1]]
            # df = df[df['Offsite Name Cont. No.'].str.contains("Company Owned Risk:"),df['Offsite Name Cont. No.'].str.contains("Unpriced Sales:")]
            df = df[(df['Offsite Name Cont. No.'].str.contains("Company Owned Risk:")) | (df['Offsite Name Cont. No.'].str.contains("priced Sales:"))]
            # for i in df.loc[:,"Offsite Name Cont. No."]:

            df["Quantity.5"].fillna(0, inplace=True)
            try:
                df["Value.5"].fillna(0, inplace=True)
            except:
                try:
                    df = read_pdf(f, pages = page, guess = True, stream = True ,
                                pandas_options={'header':0}, area = ["75,10,580,850"], columns=["46,85, 180,225, 255,275,300,360,400,430,480,525,570,620,665,713"])
                    df = pd.concat(df, ignore_index=True)
                    df["Value.5"].fillna(0, inplace=True)
                    old_pdf=False
                except Exception as e:
                    raise e

            df["Quantity.5"] = df["Quantity.5"].astype(str).str.replace("(","-").str.replace(",","").str.replace(")","").astype(float)
            # df["Value.5"] = df["Value.5"].astype(str).str.replace("(","-").str.replace(",","").str.replace(")","").astype(float)
            df["Value.5"] = df["Value.5"].astype(str).str.replace("(","-").str.replace(",","").str.replace(")","").astype(float)
            
            for i in range(len(df)):
                try:
                    print(df.iloc[i,2]) #2 for "Offsite Name Cont. No."
                except:
                    continue
                if "priced Sales" in df.iloc[i,2]:
                    print("Unprised Value found")
                    if df.iloc[-2,2] == 'Unpriced Sales:' and df.iloc[-2,-2]==0: #pd.isna(df.iloc[-2,-1]):
                        pass
                    else:
                        df.iloc[i+1,-2] = df.iloc[i+1,-2] - df.iloc[i,-2]
                        df.iloc[i+1,-1] = df.iloc[i+1,-1] - df.iloc[i,-1]
                # if i>0 and df.iloc[i-1,0]==df.iloc[i,0]:
                if i>0 and (isinstance(df.iloc[i-1,0], str) and isinstance(df.iloc[i,0], str)):
                    if df.iloc[i-1,0]==df.iloc[i,0] or ("ALLIANCE" in df.iloc[i-1,0] and "ALLIANCE" in df.iloc[i,0]):
                        #Price Remains last one
                        #Adding Quantity and Value
                        df.iloc[i,4] = df.iloc[i,4]+df.iloc[i-1,4]
                        df.iloc[i,5] = df.iloc[i,5]+df.iloc[i-1,5]
                        #droping i-1 index row
                        df.drop([df.index[i-1]], inplace=True)
                        pass

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
        if mtm_excel_summ:
            return loc_dict
        else:
            try:

                return loc_dict, hrw_fut, yc_fut ,ysb_fut
            except Exception as e:
                if "hrw_fut" in e.args[0]:
                    e_value = "hrw_fut"
                    hrw_fut=None
                elif "yc_fut" in e.args[0]:
                    e_value = "yc_fut"
                    yc_fut=None
                elif "ysb_fut" in e.args[0]:
                    e_value = "ysb_fut"
                    ysb_fut=None
                else:
                    raise e
                messagebox.showinfo("Settle Price Not found",f"Settle price for {e_value} not found, click ok to proceed further")
                return loc_dict, hrw_fut, yc_fut ,ysb_fut
    except Exception as e:
        raise e

def mtm_excel(input_date,input_xl,loc_dict,loc_sheet, output_location, hrw_fut, yc_fut ,ysb_fut):
    try:
        monthYear = datetime.strftime(datetime.strptime(input_xl.split("_")[-1].split(".xlsx")[0],"%m.%d.%Y"), "%d-%b")
        
        retry = 0
        while retry<10:
            try:
                wb = xw.Book(input_xl, update_links=True)

                break
            except Exception as e:
                time.sleep(2)
                retry+=1
                if retry==9:
                    raise e
        retry = 0
        while retry<10:
            try:
                m_sht = wb.sheets[0]

                break
            except Exception as e:
                time.sleep(2)
                retry+=1
                if retry==9:
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
            elif main_loc[i] == "Palm":
                palm = f"{i+1}"


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
        equip_row = m_sht.range("L1").end('down').end('down').end('down').row #57
        m_sht.range(f"P{equip_row}").value = loc_dict["EQUIP"][0].iloc[-1,-1]/loc_dict["EQUIP"][0].iloc[-1,-2] #loc_dict["HRW"][0].loc[loc_abbr[location]]["Price"]
        m_sht.range(f"M{equip_row}").value = loc_dict["EQUIP"][0].iloc[-1,-2]
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
                    if m_sht.range(f"C{i}").value is None:
                        m_sht.range(f"C{i}").value = 0
                    if m_sht.range(f"F{i}").value is None:
                        m_sht.range(f"F{i}").value = 0
                    if (m_sht.range(f"F{i}").value is not None) and (m_sht.range(f"F{i}").value != 0):
                        m_sht.range(f"I{i}").value = hrw_fut
                        m_sht.range(f"J{i}").formula = f"=F{i}-I{i}"
                    else:
                        m_sht.range(f"I{i}").value = None
                        m_sht.range(f"J{i}").value = None
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
                time.sleep(2)
                retry+=1
                if retry==9:
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
                if m_sht.range(f"C{i}").value is None:
                    m_sht.range(f"C{i}").value = 0
                if m_sht.range(f"F{i}").value is None:
                    m_sht.range(f"F{i}").value = 0
                if (m_sht.range(f"F{i}").value is not None) and (m_sht.range(f"F{i}").value != 0):
                    m_sht.range(f"I{i}").value = yc_fut
                    m_sht.range(f"J{i}").formula = f"=F{i}-I{i}"
                else:
                    m_sht.range(f"I{i}").value = None
                    m_sht.range(f"J{i}").value = None
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
                time.sleep(2)
                retry+=1
                if retry==9:
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
                if location.upper() == 'ZEOLITE':
                    m_sht.range(f"P69").value = loc_dict[loc_abbr[location]][0].iloc[-1,-1]/loc_dict[loc_abbr[location]][0].iloc[-1,-2]

                    m_sht.range(f"M69").value = loc_dict[loc_abbr[location]][0].iloc[-1,-2] #Quantity

                    # m_sht.range(f"F{i}").value = loc_dict[loc_abbr[location]][0].iloc[-1,-2] #Price
                    m_sht.range(f"F{i}").value = 0
                    # m_sht.range(f"C{i}").value = loc_dict[loc_abbr[location]][0].iloc[-1,-1]
                    m_sht.range(f"C{i}").value = 0
                else:

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

        m_sht.range(f"F{palm}").value = loc_dict['PALM'][0].iloc[-1,-1]/loc_dict['PALM'][0].iloc[-1,-2]  #Price
        # m_sht.range(f"C113").value = loc_dict['SUNFLWR'][0].iloc[-1,-1]
        m_sht.range(f"C{palm}").value = loc_dict['PALM'][0].iloc[-1,-2]  #Quantity

        other_loc_2_lst = m_sht.range(other_loc_2).expand('down').value
        i=int(other_loc_2.replace("A", ""))
        for location in other_loc_2_lst:
            try:
               
                m_sht.range(f"F{i}").value = loc_dict[loc_abbr[location]][0].iloc[-1,-1]/loc_dict[loc_abbr[location]][0].iloc[-1,-2] #Price
                # m_sht.range(f"F{i}").value = loc_dict[loc_abbr[location]][0].iloc[-1,-2]
                m_sht.range(f"C{i}").value = loc_dict[loc_abbr[location]][0].iloc[-1,-2] #Quantity
                # m_sht.range(f"C{i}").value = loc_dict[loc_abbr[location]][0].iloc[-1,-1]
                
            except:
                m_sht.range(f"F{i}").value = 0
                m_sht.range(f"C{i}").value = 0
                
            if location == 'YSB' and (m_sht.range(f"F{i}").value != 0 and m_sht.range(f"C{i}").value != 0):
                    m_sht.range(f"I{i}").value = ysb_fut
                    m_sht.range(f"J{i}").formula = f"=F{i}-I{i}"
            i+=1
        print()
        wb.save(output_location)
    except Exception as e:
        raise e
    finally:
        try:
            wb.app.quit()
        except:
            pass
    

def bbr(input_date, output_date):
    try:
        prev_files_loc= drive+r'\REPORT\BBR Reports\Output files'
        file_list = glob.glob(prev_files_loc+"\\*.xlsx")
        file_list.sort()
        prev_bbr = file_list[-1]
        output_location = drive+r'\REPORT\BBR Reports\Output files'+f"\\{input_date}_Borrowing Base Report.xlsx"
        input_date_date = datetime.strptime(input_date, "%m.%d.%Y").date()
        # prev_date = datetime.strptime(file_list[-1].split("_")[0].split("\\")[-1], "%m.%d.%Y").date()
        prev_bbr = prev_files_loc+f"\\{output_date}_Borrowing Base Report.xlsx"

        if not os.path.exists(prev_bbr):
            return(f"In Output Files {prev_bbr} Excel file not present for date {output_date}")
        # i=2
        # while prev_date>=input_date_date:
        #     prev_date = datetime.strptime(file_list[-i].split("_")[0].split("\\")[-1], "%m.%d.%Y").date()
        #     prev_bbr = file_list[-i]
        #     i+=1


        prev_month_year = datetime.strftime((datetime.strptime(input_date, "%m.%d.%Y").replace(day=1)-timedelta(days=1)),"%b %Y").upper()
        print(prev_month_year)
        input_xl = drive+r"\REPORT\BBR Reports\Raw Files" +f"\\{input_date}_Borrowing Base Report.xlsx"
        if not os.path.exists(input_xl):
                return(f"{input_xl} Excel file not present for date {input_date}")
        # account_lst = ["52311940", "523WP771", "523WP774", "523WP775", "523WP777", "523WP779", "523WP780", "523WP781", "523WP782", "523WP783", "523WP784", "523WP785", "523WP786", "523WP787", "523WP788", "523WP789", "523WP790", "523WP791", "523WP792", "523WP793", "523WP794", "523WP795", "523WPHLD"]
        pdf_loc = drive+r"\REPORT\BBR Reports\Raw Files\Macquarie Statement_"+input_date+".pdf"
        if not os.path.exists(pdf_loc):
                return(f"{pdf_loc} Pdf file not present for date {input_date}")

        bank_recons_loc = drive+r"\REPORT\BBR Reports\Raw Files\BANK RECONS_"+input_date+".xls"
        # bank_recons_loc = drive+r"\REPORT\Bank Recons\Output Files\BANK RECONS_"+input_date+".xls"

        if not os.path.exists(bank_recons_loc):
                return(f"{bank_recons_loc} Excel file not present for date {input_date}")

        # strg_accr_loc = drive+r"\REPORT\BBR Reports\Raw Files\STORAGE ACCRUAL "+prev_month_year+".xlsx"
        strg_accr_loc = drive+r"\REPORT\Storage Month End Report\Output Files\STORAGE ACCRUAL "+prev_month_year+".xlsx"

        if not os.path.exists(strg_accr_loc):
                return(f"{strg_accr_loc} Excel file not present for date {input_date}")

        bbr_mapping_loc = drive+r"\REPORT\BBR Reports\bbr_payables_mapping.xlsx"

        if not os.path.exists(bbr_mapping_loc):
                return(f"{bbr_mapping_loc} Excel file not present for date {input_date}")

        input_ar = drive+r"\REPORT\Open AR\Output files"+f"\\Open AR _{input_date} - Production.xlsx"
        if not os.path.exists(input_ar):
            return(f"{input_ar} Excel file not present for date {input_date}")
        input_ctm = drive+r"\REPORT\CTM Combined report\Output files"+f"\\CTM Combined _{input_date}.xlsx"
        if not os.path.exists(input_ctm):
            return(f"{input_ctm} Excel file not present for date {input_date}")

        unset_rec_loc = drive+r'\REPORT\Unsettled Receivables\Output files\Unsettled Receivables _'+input_date+'.xlsx'
        if not os.path.exists(unset_rec_loc):
            return(f"{unset_rec_loc} Excel file not present for date {input_date}")
        mtm_loc = drive+r'\REPORT\MTM reports\Output files\Inventory MTM_'+input_date+".xlsx"
        if not os.path.exists(mtm_loc):
            return(f"{mtm_loc} Excel file not present for date {input_date}")

        open_ap_loc = drive+r'\REPORT\Open AP\Output files\Open AP _'+input_date+'.xlsx'

        if not os.path.exists(open_ap_loc):
            return(f"{open_ap_loc} Excel file not present for date {input_date}")

        unset_pay_loc = drive+r'\REPORT\Unsettled Payables\Output files\Unsettled Payables _'+input_date+'.xlsx'

        if not os.path.exists(unset_pay_loc):
            return(f"{unset_pay_loc} Excel file not present for date {input_date}")
        # amount_dict = comm_acc_pdf_ext(account_lst, pdf_loc)
        # print(amount_dict)
        # print()

        retry=0
        while retry < 10:
            try:
                wb=xw.Book(input_xl, update_links=False)
                break
            except Exception as e:
                time.sleep(2)
                retry+=1
                if retry ==9:
                    raise e
        retry=0
        while retry < 10:
            try:
                bbr_sht = wb.sheets["BBR"]
                break
            except Exception as e:
                time.sleep(2)
                retry+=1
                if retry ==9:
                    raise e
        if 4 <= input_date_date.day <= 20 or 24 <= input_date_date.day <= 30:
            suffix = "th"
        else:
            suffix = ["st", "nd", "rd"][input_date_date.day % 10 - 1]
        cur_date= datetime.strftime(input_date_date, f"%B %d{suffix}, %Y")
        bbr_sht.range("A4").value = f'As of {cur_date} (the "Determination Date")'
        #Replcaing sheets from prev file
        
        retry=0
        while retry < 10:
            try:
                p_wb=xw.Book(prev_bbr)
                break
            except Exception as e:
                time.sleep(2)
                retry+=1
                if retry ==9:
                    raise e
        try:
            # wb.sheets['AR-Re-Purchase Storage Rcbl'].name = "AR-Re-Purch Org"
            wb.sheets['AR-Re-Purchase Storage Rcbl'].clear_contents()
        except:
            try:
            #    wb.sheets['AR-Re-Purchase Storage Rcbl '].name = "AR-Re-Purch Org"
                wb.sheets['AR-Re-Purchase Storage Rcbl '].clear_contents()
            except Exception as e:
                raise e
        
        try:
            # wb.sheets['Unrld Gains-Contracts MCUI'].name = "Unrld Gains Org"
            # wb.sheets['Unrld Gains-Contracts MCUI'].clear_contents()
            wb.sheets['Unrld Gains-Contracts MBL'].clear_contents()
        except:
            try:
                # wb.sheets["Unrld Gains-Contracts MCUI "].name = "Unrld Gains Org"
                # wb.sheets["Unrld Gains-Contracts MCUI "].clear_contents()
                wb.sheets["Unrld Gains-Contracts MBL "].clear_contents()
            except Exception as e:
                raise e
        
        try:
            wb.app.api.CutCopyMode=False
            p_wb.app.api.CutCopyMode=False
            # p_wb.sheets['AR-Re-Purchase Storage Rcbl'].copy(before = wb.sheets["AR-Re-Purch Org"])
            p_wb.sheets['AR-Re-Purchase Storage Rcbl'].api.Range(p_wb.sheets['AR-Re-Purchase Storage Rcbl'].api.Cells.SpecialCells(12).Address).Copy()
            wb.sheets['AR-Re-Purchase Storage Rcbl'].api.Activate()
            wb.sheets['AR-Re-Purchase Storage Rcbl'].api.Range("A1").Select()
            wb.sheets['AR-Re-Purchase Storage Rcbl'].api.Paste()
           
            wb.sheets['AR-Re-Purchase Storage Rcbl'].api.Range("A3").Formula = "='Cash Collateral'!A3"
            wb.sheets['AR-Re-Purchase Storage Rcbl'].api.Range("A3").NumberFormat = 'mm/dd/yyyy'

            wb.app.api.CutCopyMode=False
            p_wb.app.api.CutCopyMode=False
            time.sleep(1)
            wb.sheets['AR-Re-Purchase Storage Rcbl'].activate()
            try:
                wb.api.ChangeLink(Name = wb.api.LinkSources()[0], NewName=wb.fullname, Type=1)
            except:
                pass

            pass


        except:
            try:
                # p_wb.sheets['AR-Re-Purchase Storage Rcbl'].copy(before = wb.sheets["AR-Re-Purch Org"])
                wb.app.api.CutCopyMode=False
                p_wb.app.api.CutCopyMode=False
                # p_wb.sheets['AR-Re-Purchase Storage Rcbl'].copy(before = wb.sheets["AR-Re-Purch Org"])
                p_wb.sheets['AR-Re-Purchase Storage Rcbl '].api.Range(p_wb.sheets['AR-Re-Purchase Storage Rcbl'].api.Cells.SpecialCells(12).Address).Copy()
                wb.sheets['AR-Re-Purchase Storage Rcbl '].api.Activate()
                wb.sheets['AR-Re-Purchase Storage Rcbl '].api.Range("A1").Select()
                wb.sheets['AR-Re-Purchase Storage Rcbl '].api.Paste()

                wb.sheets['AR-Re-Purchase Storage Rcbl '].api.Range("A3").Formula = "='Cash Collateral'!A3"
                wb.sheets['AR-Re-Purchase Storage Rcbl '].api.Range("A3").NumberFormat = 'mm/dd/yyyy'

                wb.app.api.CutCopyMode=False
                p_wb.app.api.CutCopyMode=False
                time.sleep(1)
                wb.sheets['AR-Re-Purchase Storage Rcbl'].activate()
                wb.api.ChangeLink(Name = wb.api.LinkSources()[0], NewName=wb.fullname, Type=1)
            except Exception as e:
                raise e
        try:
            wb.app.api.CutCopyMode=False
            p_wb.app.api.CutCopyMode=False
            # p_wb.sheets['AR-Re-Purchase Storage Rcbl'].copy(before = wb.sheets["AR-Re-Purch Org"])
            p_wb.sheets['Unrld Gains-Contracts MBL'].api.Range(p_wb.sheets['Unrld Gains-Contracts MBL'].api.Cells.SpecialCells(12).Address).Copy()
            wb.sheets['Unrld Gains-Contracts MBL'].api.Activate()
            wb.sheets['Unrld Gains-Contracts MBL'].api.Range("A1").Select()
            wb.sheets['Unrld Gains-Contracts MBL'].api.Paste()
            
            wb.sheets['Unrld Gains-Contracts MBL'].api.Range("A3").Formula = "='Cash Collateral'!A3"
            wb.sheets['Unrld Gains-Contracts MBL'].api.Range("A3").NumberFormat = 'mm/dd/yyyy'

            wb.app.api.CutCopyMode=False
            p_wb.app.api.CutCopyMode=False
            
        except:
            try:
                wb.app.api.CutCopyMode=False
                p_wb.app.api.CutCopyMode=False
                # p_wb.sheets['AR-Re-Purchase Storage Rcbl'].copy(before = wb.sheets["AR-Re-Purch Org"])
                p_wb.sheets['Unrld Gains-Contracts MBL '].api.Range(p_wb.sheets['Unrld Gains-Contracts MBL '].api.Cells.SpecialCells(12).Address).Copy()
                wb.sheets['Unrld Gains-Contracts MBL '].api.Activate()
                wb.sheets['Unrld Gains-Contracts MBL '].api.Range("A1").Select()
                wb.sheets['Unrld Gains-Contracts MBL '].api.Paste()

                wb.sheets['Unrld Gains-Contracts MBL '].api.Range("A3").Formula = "='Cash Collateral'!A3"
                wb.sheets['Unrld Gains-Contracts MBL '].api.Range("A3").NumberFormat = 'mm/dd/yyyy'


                wb.app.api.CutCopyMode=False
                p_wb.app.api.CutCopyMode=False
                
            except Exception as e:
                raise e
        # wb.sheets["AR-Re-Purch Org"].delete()
        # wb.sheets["Unrld Gains Org"].delete()
        p_wb.close()
        # bbr_other_tabs(input_date, wb, input_ar, input_ctm)
        # payables(input_date,wb, bbr_mapping_loc, open_ap_loc,unset_pay_loc)
        
        jp_morgan_amount = cash_colat(wb,bank_recons_loc, input_date_date)
        comm_acc_xl(wb, pdf_loc)
        inv_whre_n_in_trans(wb, mtm_loc, input_date)
        
        ar_unsettled_by_tier(wb, unset_rec_loc, input_date)
        ar_open_storage_rcbl(wb, strg_accr_loc, input_date)
        
        payables(input_date,wb, bbr_mapping_loc, open_ap_loc,unset_pay_loc,jp_morgan_amount)
        bbr_other_tabs(input_date, wb, input_ar, input_ctm)

        wb.sheets[0].activate()
        wb.save(output_location)
        print()
        return f"BBR report generated for {input_date}"
    except Exception as e:
        raise e
    finally:
        try:
            wb.app.quit()
        except:
            pass


def cpr(input_date, output_date):
    try:
        cpr_file_date = input_date.replace('.','-')
        output_cpr  = drive+r'\REPORT\CPR reports\Output files'+'\\Counter Party Risk Consolidated '+cpr_file_date+'.xlsx'
        output_cpr_copy  = drive+r'\REPORT\CPR reports\Output files'+'\\Counter Party Risk Consolidated '+cpr_file_date+' Report Copy.xlsx'
        
        input_cpr = drive+r'\REPORT\CPR reports\Raw Files\Counter Party Risk Consolidated '+cpr_file_date+'.xlsx'

        input_cpr_copy = drive+r'\REPORT\CPR reports\Raw Files\Counter Party Risk Consolidated '+cpr_file_date+' Report Copy.xlsx'

        UnsettledRec_book = drive+r'\REPORT\Unsettled Receivables\Output files\Unsettled Receivables _'+input_date+'.xlsx'

        UnsettledPay_book = drive+r'\REPORT\Unsettled Payables\Output files\Unsettled Payables _'+input_date+'.xlsx'

        Open_AR_book = drive+r'\REPORT\Open AR\Output files\Open AR _'+input_date+' - Production.xlsx'

        Open_AP_book = drive+r'\REPORT\Open AP\Output files\Open AP _'+input_date+'.xlsx'

        CTM_book = drive+r'\REPORT\CTM Combined report\Output files\CTM Combined _'+input_date+'.xlsx'

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


    
        
        # input_file = f'{book_name} {sheet_date}.xlsx'

        retry = 0
        while retry<10:
            try:
                wb = xw.Book(input_cpr, update_links=False)
                break
            except:
                time.sleep(2)
                retry+=1
        
        retry = 0
        while retry<10:
            try:
                
                ws1 = wb.sheets[f'Data {input_date}']
                ws1.api.AutoFilterMode=False
                break
            except:
                time.sleep(2)
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
                time.sleep(2)
                retry+=1
                if retry==9:
                    raise e
        retry=0
        while retry<10:
            try:
                UnsettledRec_ws = UnsettledRec_wb.sheets['Excl Macq & IC']
                break
            except Exception as e:
                time.sleep(2)
                retry+=1
                if retry==9:
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
                time.sleep(2)
                retry+=1
                if retry==9:
                    raise e
        retry=0
        while retry<10:
            try:
                
                UnsettledPay_ws = UnsettledPay_wb.sheets['Excl Macq & IC']
                break
            except Exception as e:
                time.sleep(2)
                retry+=1
                if retry==9:
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
                time.sleep(2)
                retry+=1
                if retry==9:
                    raise e
        retry=0
        while retry<10:
            try:
                OpenAR_ws = OpenAR_wb.sheets['Eligible']
                break
            except Exception as e:
                time.sleep(2)
                retry+=1
                if retry==9:
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
                time.sleep(2)
                retry+=1
                if retry==9:
                    raise e
        retry = 0
        while retry<10:
            try:
                
                OpenAP_ws = OpenAP_wb.sheets['Excl Macq & IC']
                break
            except Exception as e:
                time.sleep(2)
                retry+=1
                if retry==9:
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
                time.sleep(2)
                retry+=1
                if retry==9:
                    raise e
        retry = 0
        while retry<10:
            try:
                CTM_ws = CTM_wb.sheets['Excl Macq & IC']
                break
            except Exception as e:
                time.sleep(2)
                retry+=1
                if retry==9:
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
        while retry<10:
            try:
                pivot_sht = wb.sheets["Pivot"]
                time.sleep(2)
                # pivot_sht.select()
                pivot_sht.activate()
                break
            except Exception as e:
                time.sleep(2)
                retry+=1
                if retry==9:
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
                time.sleep(2)
                retry+=1
                if retry==9:
                    raise e
        
        #logger.info('Opening Master sheet')
        while True:
            try:
                BB_ws = BB_wb.sheets['Master']
                BB_ws.api.AutoFilterMode=False
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
            # if (type(BB_Master25ws.range(f'H{i}').value) == int) or (type(BB_Master25ws.range(f'H{i}').value) == float):
            if BB_Master25ws.range(f'H{i}').value is None:
                break
            if  (-25000 < float(BB_Master25ws.range(f'H{i}').value)) and (float(BB_Master25ws.range(f'H{i}').value) <25000):
                # BB_Master25ws.range(f'{i}:{i}').api.Delete(win32c.DeleteShiftDirection.xlShiftDown)
                BB_Master25ws.range(f"{i}:{i}").api.Delete(win32c.DeleteShiftDirection.xlShiftUp)
                # BB_Master25ws.range(f'{i}:{i}').delete()
                # i+=1
            else:
                i+=1
                
        # gt
        time.sleep(1)
        last_row = BB_Master25ws.range(f'H'+ str(BB_Master25ws.cells.last_cell.row)).end('up').row
        last_column = num_to_col_letters(BB_Master25ws.range("A6").end('right').column)
        BB_Master25ws.range(f"A6:{last_column}{last_row}").api.Sort(Key1=BB_Master25ws.range(f"H6:H{last_row}").api,Order1=2,DataOption1=0,Orientation=1)
        # BB_Master25ws.range(f'H9:H{BB_Master25_Row}').api.NumberFormat = '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'
        BB_Master25ws.range(f'C:H').api.NumberFormat = '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'
        ####logger.info('Refreshing all tab')  
        BB_wb.api.RefreshAll()
        BB_wb.activate()
        BB_wb.sheets[2].select()
        try:
            BB_wb.api.ActiveSheet.PivotTables("PivotTable2").PivotFields('Cust Type').CurrentPage = "E"
        except:
            try:
                BB_wb.api.ActiveSheet.PivotTables("PivotTable7").PivotFields('Cust Type').CurrentPage = "(blank)"
            except Exception as e:
                raise e
        BB_wb.activate()
        BB_wb.sheets[3].select()
        try:
            BB_wb.api.ActiveSheet.PivotTables("PivotTable4").PivotFields('Cust Type').CurrentPage = "F"
        except:
            try:
                BB_wb.api.ActiveSheet.PivotTables("PivotTable7").PivotFields('Cust Type').CurrentPage = "(blank)"
            except Exception as e:
                raise e
        BB_wb.activate()
        BB_wb.sheets[4].select()
        try:
            BB_wb.api.ActiveSheet.PivotTables("PivotTable5").PivotFields('Cust Type').CurrentPage = "R"
        except:
            try:
                BB_wb.api.ActiveSheet.PivotTables("PivotTable7").PivotFields('Cust Type').CurrentPage = "(blank)"
            except Exception as e:
                raise e
        BB_wb.activate()
        BB_wb.sheets[5].select()
        try:
            BB_wb.api.ActiveSheet.PivotTables("PivotTable6").PivotFields('Cust Type').CurrentPage = "P"
        except:
            try:
                BB_wb.api.ActiveSheet.PivotTables("PivotTable7").PivotFields('Cust Type').CurrentPage = "(blank)"
            except Exception as e:
                raise e
        BB_wb.activate()
        BB_wb.sheets[6].select()
        try:
            BB_wb.api.ActiveSheet.PivotTables("PivotTable7").PivotFields('Cust Type').CurrentPage = "T"
        except:
            try:
                BB_wb.api.ActiveSheet.PivotTables("PivotTable7").PivotFields('Cust Type').CurrentPage = "(blank)"
            except Exception as e:
                raise e
        print()
        BB_Master25ws.activate()
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
        input_sheet = drive+r'\REPORT\CTM Combined report\Raw Files\CTM Combined _'+input_date+'.xlsx' 
        output_location = drive+r'\REPORT\CTM Combined report\Output files\CTM Combined _'+input_date+".xlsx"
        # input_cpr = drive+r'\REPORT\CPR reports\Raw Files\Counter Party Risk Consolidated '+cpr_file_date+'.xlsx'    
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
                time.sleep(2)
                retry+=1
                if retry ==9:
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
            intcomp_sht.range(f"A2:{last_column_letter}{last_row}").copy(ws2.range(f"A{last_row2}"))#:{last_column_letter}{x}
            intcomp_sht.delete()

            excl_sht = wb.sheets("Excl Macq & IC")
            ##logger.info("Copy tier sheet AFTER the intercompany sheet of input book.")
            # num_row = excl_sht.range('A1').end('down').row
            num_row=excl_sht.range(f'A' + str(excl_sht.cells.last_cell.row)).end('up').row

            #Condtion for removing only header line coming at last
            if excl_sht.range(f'I{num_row}').value.upper()!="EQUIP":
                excl_sht.range(f"{num_row}:{num_row}").delete()                   
        except Exception as e:
         print("No (INTER-COMPANY PURCH/SALES) Present ")
         print(e) 

        

        ###logger.info("Adding Worksheet for Pivot Table")
        wb.sheets.add("Pivot BB",after=wb.sheets["Excl Macq & IC"])
        ###logger.info("Clearing New Worksheet")
        wb.sheets["Pivot BB"].clear_contents()
        ws3=wb.sheets["Pivot BB"]
        ws3.range("A1").value="West Plains, LLC"
        ws3.range("A2").value="Net Unrealized Gains on Forward Contracts - Non MBL"
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
        Event = False
        try:
            PivotTable.PivotFields('Ship Tier').CurrentPage = ">12 months"
            ###logger.info("Changing No Format in Pivot Table")
            ###logger.info("Changing Table layout")
            PivotTable.PivotFields('Location Id').Subtotals=(False, False, False, False, False, False, False, False, False, False, False, False)
            PivotTable.PivotFields('Commodity Id').Subtotals=(False, False, False, False, False, False, False, False, False, False, False, False)
            PivotTable.RowAxisLayout(1)
            ###logger.info("Changing Table Style")
            PivotTable.TableStyle2 = ""
        except:
            wb.api.ActiveSheet.PivotTables("PivotTable2").PivotSelect("",win32c.PTSelectionMode.xlDataAndLabel,True) 
            wb.app.selection.api.ClearContents()
            Event = True
        finally:
            last_column = ws3.range('A7').end('right').last_cell.column
            last_column+=3
            last_row = ws3.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row
            last_row+=5
            last_column_letter=num_to_col_letters(last_column)
        if Event:    
            ws3.range(f"{last_column_letter}{last_row}").value=f'=GETPIVOTDATA("Gain/LossTotal",$A$7)'
        else:
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
        inp_formula_sht = drive+r'\REPORT\Freight analysis reports\Col_N_Formulas.xlsx'
        
        output_location = drive+r'\REPORT\Freight analysis reports\Output files'
        raw_input = drive+r'\REPORT\Freight analysis reports\Raw files'

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
                    time.sleep(2)
                    retry+=1
                    if retry==9:
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
def mtm_report(input_date, output_date):
    try:
        print(input_date)
        # print(output_date)

        input_xl = drive+r'\REPORT\MTM reports\Raw Files\Inventory MTM_'+input_date+".xlsx"
        if not os.path.exists(input_xl):
            return(f"{input_xl} Excel file not present for date {input_date}")
        
        pdf_loc = drive+r'\REPORT\MTM reports\Raw Files\Inventory Market Valuation _'+input_date+'.pdf'
        if not os.path.exists(pdf_loc):
            return(f"{pdf_loc} Pdf file not present for date {input_date}")

        hrw_pdf_loc = 'HRW_loc'
        # hrw_pdf_loc = drive+r'\REPORT\MTM reports\Raw Files\HRW_'+input_date+'.pdf'
        # if not os.path.exists(hrw_pdf_loc):
        #     return(f"{hrw_pdf_loc} Pdf file not present for date {input_date}")
        yc_pdf_loc = 'YC_loc'
        # yc_pdf_loc = drive+r'\REPORT\MTM reports\Raw Files\YC_'+input_date+'.pdf'
        # if not os.path.exists(yc_pdf_loc):
        #     return(f"{yc_pdf_loc} Pdf file not present for date {input_date}")
        ysb_pdf_loc = 'YSB_loc'

        loc_sheet = drive+r'\REPORT\MTM reports\Loc_Abbr.xlsx'
        if not os.path.exists(loc_sheet):
            return(f"{loc_sheet}Excel file not present for date {input_date}")

        loc_dict, hrw_fut, yc_fut, ysb_fut = mtm_pdf_data_extractor(input_date,pdf_loc, hrw_pdf_loc, yc_pdf_loc ,ysb_pdf_loc)
        output_location = drive+r'\REPORT\MTM reports\Output files\Inventory MTM_'+input_date+".xlsx"
        mtm_excel(input_date, input_xl,loc_dict,loc_sheet, output_location, hrw_fut, yc_fut , ysb_fut)

        print("Done")
        return f"MTM report Generated for {input_date}"
    except Exception as e:
        raise e
    

def open_ar(input_date, output_date):
    try:
        input_sheet = drive+r'\REPORT\Open AR\Raw files'+f'\\Open AR _{input_date} - Production.xlsx' 
        if not os.path.exists(input_sheet):
            return(f"{input_sheet} Excel file not present for date {input_date}")
        prev_output=drive+r'\REPORT\Open AR\Output files'+f'\\Open AR _{output_date} - Production.xlsx'
        if not os.path.exists(prev_output):
            return(f"{prev_output} Excel file not present for date {output_date}")  

        output_location = drive+r'\REPORT\Open AR\Output files'  
        prev_month = datetime.strftime(datetime.strptime(input_date, "%m.%d.%Y"), "%B")
        ##logger.info("Opening operating workbook instance of excel")
        retry=0
        while retry < 10:
            try:
                wb=xw.Book(input_sheet)
                break
            except Exception as e:
                time.sleep(2)
                retry+=1
                if retry ==9:
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
        ######################New Renewble customer removal logic$##############################
        try:
            ws2.api.AutoFilterMode=False
            ws2.api.Range(f"{Customer_column}1").AutoFilter(Feild:=Customer_column_num,Criteria1:="BIOURJA RENEWABLES, LLC") #Removing Intercompany
            # renewable_sht = wb.sheets.add("Renewables", after=ws1)

            # ws2.api.AutoFilter.Range.Copy()
            # time.sleep(1)
            # renewable_sht.range("A1").api.Select()
            # while True:
            #     try:
            #         renewable_sht.api.Paste()
            #         break
            #     except:
            #         time.sleep(1)
            # wb.app.api.CutCopyMode=False
            # time.sleep(1)
            

            filter_lst_row = ws2.range(f'A'+ str(ws2.cells.last_cell.row)).end('up').row
            if filter_lst_row!=1:
                filter_add = ws2.api.Range(f"A2:A{filter_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Address
                ws2.range(f"{filter_add}").delete()
            ws2.api.AutoFilterMode=False

        except Exception as e:
            print("BIOURJA RENEWABLES, LLC")
            print(e)

        ####################################################################################
        ##logger.info("Copying tier column from previous output sheet")   
        retry=0
        while retry < 10:
            try:
                tier_wb = xw.Book(prev_output,update_links=True)
                break
            except Exception as e:
                time.sleep(2)
                retry+=1
                if retry ==9:
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
        for i in range(2,int(f'{last_row}')):            
                    if ws5.range(f"D{i}").value<0:
                        if ws5.range(f"G{i}").value>0:
                            if ws5.range(f"G{i}").value>abs(ws5.range(f"D{i}").value):
                                ws5.range(f"G{i}").value = ws5.range(f"D{i}").value + ws5.range(f"G{i}").value
                                ws5.range(f"D{i}").value = 0 
                            else:
                                ws5.range(f"D{i}").value = ws5.range(f"D{i}").value + ws5.range(f"G{i}").value
                                ws5.range(f"G{i}").value = 0 
        for i in range(2,int(f'{last_row}')):            
                    if ws5.range(f"D{i}").value<0:
                        if ws5.range(f"F{i}").value>0:
                            if ws5.range(f"F{i}").value>abs(ws5.range(f"D{i}").value):
                                ws5.range(f"F{i}").value = ws5.range(f"D{i}").value + ws5.range(f"F{i}").value
                                ws5.range(f"D{i}").value = 0 
                            else:
                                ws5.range(f"D{i}").value = ws5.range(f"D{i}").value + ws5.range(f"F{i}").value
                                ws5.range(f"F{i}").value = 0                                
        for i in range(2,int(f'{last_row}')):            
                    if ws5.range(f"D{i}").value<0:
                        if ws5.range(f"E{i}").value>0:
                            if ws5.range(f"E{i}").value>abs(ws5.range(f"D{i}").value):
                                ws5.range(f"E{i}").value = ws5.range(f"D{i}").value + ws5.range(f"E{i}").value
                                ws5.range(f"D{i}").value = 0 
                            else:
                                ws5.range(f"D{i}").value = ws5.range(f"D{i}").value + ws5.range(f"E{i}").value
                                ws5.range(f"E{i}").value = 0 
        ws5.api.Range(f"A1:{last_column_letter}{last_row}").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        
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
        input_sheet = drive+r'\REPORT\Open AP\Raw files'+f'\\Open AP _{input_date}.xlsx' 
        output_location = drive+r'\REPORT\Open AP\Output files'
        if not os.path.exists(input_sheet):
            return(f"{input_sheet} Excel file not present for date {input_date}")
        #logger.info("Opening operating workbook instance of excel")
        retry=0
        while retry < 10:
            try:
                wb=xw.Book(input_sheet)
                break
            except Exception as e:
                time.sleep(2)
                retry+=1
                if retry ==9:
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
            PivotTable.PivotFields('Journal Source').PivotItems("A/P Invoice Entry").Visible = False
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
        input_xl = drive+r'\REPORT\Unsettled Payables\Raw Files\Unsettled Payables _'+input_date+".xlsx"
        output_location = drive+r'\REPORT\Unsettled Payables\Output files\Unsettled Payables _'+input_date+".xlsx"
        
        if not os.path.exists(input_xl):
            return(f"Excel file not present for date {input_date}")

        prev_month = datetime.strftime(datetime.strptime(input_date, "%m.%d.%Y"), "%B")
        
        
        retry=0
        while retry<10:
            try:
                
                wb = xw.Book(input_xl, update_links=True)
                break
            except Exception as e:
                time.sleep(2)
                retry+=1
                if retry==9:
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
        return f"Unsettled Payables report Generated for {input_date}"
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
        

        input_xl = drive+r'\REPORT\Unsettled Receivables\Raw Files\Unsettled Receivables _'+input_date+".xlsx"
        prev_output_location = drive+r'\REPORT\Unsettled Receivables\Output files\Unsettled Receivables _'+output_date+".xlsx"
        output_location = drive+r'\REPORT\Unsettled Receivables\Output files\Unsettled Receivables _'+input_date+".xlsx"
        
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
                time.sleep(2)
                retry+=1
                if retry==9:
                    raise e
        
        ######logger.info("Sheet Opened")
        # time.sleep(10)
        while True:
            try:
                inp_sht = wb.sheets[0] #wb.sheets[0].name in 'Unsettled Receivables _'+input_date
                
                break
            except Exception as e:
                time.sleep(2)
                # retry+=1
                # if retry==9:
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
        PivotTable.PivotFields('Sum of Net').NumberFormat= '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'
        PivotTable.PivotFields('Ticket Date <=30 Days').Orientation = win32c.PivotFieldOrientation.xlDataField
        PivotTable.PivotFields('Sum of Ticket Date <=30 Days').NumberFormat= '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'
        PivotTable.PivotFields('Ticket Date 31-60 Days').Orientation = win32c.PivotFieldOrientation.xlDataField
        PivotTable.PivotFields('Sum of Ticket Date 31-60 Days').NumberFormat= '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'
        PivotTable.PivotFields('Ticket Date +60 Days').Orientation = win32c.PivotFieldOrientation.xlDataField
        PivotTable.PivotFields('Sum of Ticket Date +60 Days').NumberFormat= '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'

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
        # input_xl = drive+r"\REPORT\MOC Interest allocation\Raw files\Inventory MTM Excel Report " + input_date + ".xlsx"
        # if not os.path.exists(input_xl):
        #         return(f"{input_xl} Excel file not present for date {input_date}")
        dt = datetime.strptime(input_date,"%m.%d.%Y")
        mtm_input_date = dt.strftime("%B %Y")

        # mtm_file = drive+r"\REPORT\MOC Interest allocation\Raw files\Inventory MTM Excel Report " + mtm_input_date +'.xlsx'
        mtm_file = drive+r"\REPORT\FIFO reports\Output files\Inventory MTM Excel Report " + mtm_input_date +'.xlsx'
        # mtm_file = drive+r"\REPORT\FIFO reports\Output files\Inventory MTM Excel Report_" + input_date +'.xlsx'
        if not os.path.exists(mtm_file):
                return(f"{mtm_file} Excel file not present for date {input_date}")

        unsettled_recev_file = drive+r'\REPORT\Unsettled Receivables\Output files\Unsettled Receivables _'+input_date+'.xlsx'

        if not os.path.exists(unsettled_recev_file):
                return(f"{unsettled_recev_file} Excel file not present for date {input_date}")

        unsettled_pay_file = drive+r'\REPORT\Unsettled Payables\Output files\Unsettled Payables _'+input_date+'.xlsx'

        if not os.path.exists(unsettled_pay_file):
                return(f"{unsettled_pay_file} Excel file not present for date {input_date}")

        open_ar_file = drive+r'\REPORT\Open AR\Output files\Open AR _'+input_date+' - Production.xlsx'

        if not os.path.exists(open_ar_file):
                return(f"{open_ar_file} Excel file not present for date {input_date}")

        open_ap_file = drive+r'\REPORT\Open AP\Output files\Open AP _'+input_date+'.xlsx'

        if not os.path.exists(open_ap_file):
                return(f"{open_ap_file} Excel file not present for date {input_date}")

        


        output_dir = drive+r"\REPORT\MOC Interest allocation\Output Files"
        template_dir = drive+r"\REPORT\MOC Interest allocation\Raw files\template"


          
        main_df = moc_get_df_from_input_excel(mtm_file, open_ap_file, open_ar_file,unsettled_pay_file, unsettled_recev_file)
        update_moc_excel(main_df, template_dir, output_dir, input_date)
        return f"MOC Interest Allocation Report Generated for {input_date}"
    except Exception as e:
        raise e
    
def bbr_monthEnd(input_date, output_date):
    try:
        monthYear = datetime.strftime(datetime.strptime(input_date, "%m.%d.%Y"), "%b %Y").upper()
        input_bbr = drive+r"\REPORT\BBR Reports\Output files" +f"\\{input_date}_Borrowing Base Report.xlsx"
        output_loc = drive+r"\REPORT\BBR Reports\Output files\Month_End" +f"\\{input_date}_Borrowing Base Report.xlsx"
        if not os.path.exists(input_bbr):
            return(f"{input_bbr} Excel file not present for date {input_date}")

        strg_accr = drive+r'\REPORT\Storage Month End Report\Output Files'+f"\\STORAGE ACCRUAL {monthYear}.xlsx" #f"\\{monthYear}.xlsx"
        if not os.path.exists(strg_accr):
            return(f"{strg_accr} Excel file not present for date {input_date}")

        retry=0
        while retry < 10:
            try:
                bbr_wb=xw.Book(input_bbr)
                break
            except Exception as e:
                time.sleep(2)
                retry+=1
                if retry ==9:
                    raise e

        retry=0
        while retry < 10:
            try:
                accr_wb=xw.Book(strg_accr)
                break
            except Exception as e:
                time.sleep(2)
                retry+=1
                if retry ==9:
                    raise e
        
        try:
            bbr_wb.sheets['AR-Open Storage Rcbl'].name = "AR-Open Storage Rcbl Org"
        except:
            try:
                bbr_wb.sheets['AR-Open Storage Rcbl '].name = "AR-Open Storage Rcbl Org"
            except Exception as e:
                raise e

        accr_wb.sheets[0].copy(before = bbr_wb.sheets["AR-Open Storage Rcbl Org"])
        bbr_wb.sheets["AR-Open Storage Rcbl Org"].delete()
        bbr_wb.sheets["Storage Accrual"].name = 'AR-Open Storage Rcbl'
        bbr_wb.save(output_loc)
        return f"BBR motnh End report generated for {monthYear}"
    except Exception as e:
        raise e
    finally:
        try:
            bbr_wb.app.quit()
        except:
            pass

def inv_mtm_excel_summ(input_date, output_date):
    try:
        monthYear = datetime.strftime(datetime.strptime(input_date, "%m.%d.%Y"), "%B %Y")
        pdf_loc = drive+r'\REPORT\MTM reports\Raw Files\Inventory Market Valuation _'+input_date+'.pdf'
        # pdf_loc = r'C:\Users\imam.khan\OneDrive - BioUrja Trading LLC\Documents\WEST PLAINS\REPORT\MTM reports\Raw Files\Inventory Market Valuation _'+input_date+'.pdf'
        if not os.path.exists(pdf_loc):
            return(f"{pdf_loc} Pdf file not present for date {input_date}")
        input_xl = drive+r'\REPORT\Inv_MTM_Excel_Report_Summ\Raw Files\Inventory_MTMExcel_SummTemplate.xlsx'
        # input_xl = r'C:\Users\imam.khan\OneDrive - BioUrja Trading LLC\Documents\WEST PLAINS\REPORT\Inv_MTM_Excel_Report_Summ\Raw Files\Inventory_MTMExcel_SummTemplate.xlsx'
        if not os.path.exists(input_xl):
            return(f"{input_xl} Excel file not present for date {input_date}")

        # output_loc = r"C:\Users\imam.khan\OneDrive - BioUrja Trading LLC\Documents\WEST PLAINS\REPORT\FIFO reports\Output files" +f"\\Inventory MTM Excel Report {monthYear}.xlsx"
        output_loc = drive+r'\REPORT\Inv_MTM_Excel_Report_Summ\Output files' +f"\\Inventory MTM Excel Report_{input_date}.xlsx"



        loc_dict = inv_mtm_pdf_data_extractor(input_date,pdf_loc)
        # loc_dict = mtm_pdf_data_extractor(input_date,pdf_loc, mtm_excel_summ=True)

        retry=0
        while retry < 10:
            try:
                mtm_wb=xw.Book(input_xl)
                break
            except Exception as e:
                time.sleep(2)
                retry+=1
                if retry ==9:
                    raise e
        retry=0
        while retry < 10:
            try:
                mtm_sht = mtm_wb.sheets["INPUT DATA"]
                break
            except Exception as e:
                time.sleep(2)
                retry+=1
                if retry ==9:
                    raise e
        mtm_sht.api.AutoFilterMode=False
        mtm_sht.range("A1").value = input_date.replace(".","-")
        mtm_last_row = mtm_sht.range(f'A'+ str(mtm_sht.cells.last_cell.row)).end('up').row
        for loc in ["HRW", "YC"]:
            mtm_sht.activate()
            mtm_sht.api.AutoFilterMode=False
            mtm_sht.api.Range(f"D3").AutoFilter(Field:=4,Criteria1:=loc, Operator:=7)

            #Updating HRW and YC Quantity and Values
            loc_dict[loc][0].set_index('Location Zone', inplace=True)
            loc_dict[loc][0].rename(index={'ALLIANCETE': 'ALLIANCE'}, inplace=True)
            loc_dict[loc][0].rename(index={'LISCO - WE': 'LISCO'}, inplace=True)
            mtm_sht.api.Range(f"G4:G{mtm_last_row}").SpecialCells(12).Select()
            for rng in mtm_wb.app.selection.address.split(','):
                for i in range(int(rng.split(":")[0].split("$")[-1]),int(rng.split(":")[-1].split("$")[-1])+1):
                    try:
                        mtm_sht.range(f"G{i}").value = float(loc_dict[loc][0].loc[mtm_sht.range(f"B{i}").value]["Quantity.5"])
                        mtm_sht.range(f"K{i}").value = float(loc_dict[loc][0].loc[mtm_sht.range(f"B{i}").value]["Value.5"]) #/float(loc_dict[loc][0].loc[mtm_sht.range(f"B{i}").value]["Quantity.5"])
                    except:
                        mtm_sht.range(f"G{i}").value = 0
                        mtm_sht.range(f"K{i}").value = 0
            
        mtm_sht.api.AutoFilterMode=False
        mtm_sht.api.Range(f"D3").AutoFilter(Field:=4,Criteria1:='<>HRW', Operator:=1, Criteria2:='<>YC')
        mtm_sht.api.Range(f"G4:G{mtm_last_row}").SpecialCells(12).Select()
        for rng in mtm_wb.app.selection.address.split(','):
            for i in range(int(rng.split(":")[0].split("$")[-1]),int(rng.split(":")[-1].split("$")[-1])+1):
                
                if mtm_sht.range(f"B{i}").value == "BROWNSVILL" and mtm_sht.range(f"D{i}").value == "MILO":
                    try:
                        loc_dict["SORGHUM"][0].set_index('Location Zone', inplace=True)
                    except:
                        pass
                    mtm_sht.range(f"G{i}").value = float(loc_dict["SORGHUM"][0].loc[mtm_sht.range(f"B{i}").value]["Quantity.5"])
                    mtm_sht.range(f"K{i}").value = float(loc_dict["SORGHUM"][0].loc[mtm_sht.range(f"B{i}").value]["Value.5"]) #/float(loc_dict["SORGHUM"][0].loc[mtm_sht.range(f"B{i}").value]["Quantity.5"])
                else:
                    try:
                        loc_dict[mtm_sht.range(f"D{i}").value][0].set_index('Location Zone', inplace=True)
                        try:
                            # loc_dict[mtm_sht.range(f"D{i}").value][0].rename(index={'OMA COMM': 'TERMINAL'}, inplace=True)
                            loc_dict[mtm_sht.range(f"D{i}").value][0].rename(index={'OMA COMM': 'OMCOM'}, inplace=True)
                        except:
                            pass
                    except:
                        pass
                    try:
                        mtm_sht.range(f"G{i}").value = float(loc_dict[mtm_sht.range(f"D{i}").value][0].loc[mtm_sht.range(f"B{i}").value]["Quantity.5"]) #Quantity
                    except:
                        pass
                    try:
                        if float(loc_dict[mtm_sht.range(f"D{i}").value][0].loc[mtm_sht.range(f"B{i}").value]["Quantity.5"]) == 0 and float(loc_dict[mtm_sht.range(f"D{i}").value][0].loc[mtm_sht.range(f"B{i}").value]["Value.5"]) == 0:
                            mtm_sht.range(f"K{i}").value = 0
                        else:
                            mtm_sht.range(f"K{i}").value = float(loc_dict[mtm_sht.range(f"D{i}").value][0].loc[mtm_sht.range(f"B{i}").value]["Value.5"]) #/float(loc_dict[mtm_sht.range(f"D{i}").value][0].loc[mtm_sht.range(f"B{i}").value]["Quantity.5"])
                    except:
                        pass
        mtm_sht.api.AutoFilterMode=False                   
        mtm_wb.save(output_loc)

        return f"MTM report Generated for {input_date}"
    except Exception as e:
        raise e
    finally:
        try:
            mtm_wb.app.quit()
        except:
            pass

def fifo(input_date, output_date):
    try:
        location = ["HRW", "YC"]
        inp_date = datetime.strftime(datetime.strptime(input_date, "%m.%d.%Y"), "%m.%d.%y")
        monthYear = datetime.strftime(datetime.strptime(input_date, "%m.%d.%Y"), "%B %Y")
        for loc in location:
            input_xl = drive+r"\REPORT\FIFO reports\Raw Files" +f"\\Inventory on site {loc}_{inp_date}.xlsx"
            # input_xl = r"C:\Users\imam.khan\OneDrive - BioUrja Trading LLC\Documents\WEST PLAINS\REPORT\FIFO reports\Raw Files" +f"\\Inventory on site {loc}_{inp_date}.xlsx"
            if not os.path.exists(input_xl):
                    return(f"{input_xl} Excel file not present for date {input_date}")
            
            # input_mtm = drive+r"\REPORT\MOC Interest allocation\Raw Files" +f"\\Inventory MTM Excel Report {monthYear}.xlsx"
            # input_mtm = drive+r'\REPORT\Inv_MTM_Excel_Report_Summ\Output files' +f"\\Inventory MTM Excel Report {monthYear}.xlsx"
            input_mtm = drive+r'\REPORT\Inv_MTM_Excel_Report_Summ\Output files' +f"\\Inventory MTM Excel Report_{input_date}.xlsx"
            # input_mtm = r"C:\Users\imam.khan\OneDrive - BioUrja Trading LLC\Documents\WEST PLAINS\REPORT\FIFO reports\Output files" +f"\\Inventory MTM Excel Report {monthYear}.xlsx"
            if not os.path.exists(input_mtm):
                    return(f"{input_mtm} Excel file not present for date {input_date}")

            input_mapping = drive+r"\REPORT\FIFO reports" +f"\\Sub_Loc_Mapping.xlsx"
            # input_mapping = r"C:\Users\imam.khan\OneDrive - BioUrja Trading LLC\Documents\WEST PLAINS\REPORT\FIFO reports" +f"\\Sub_Loc_Mapping.xlsx"
            if not os.path.exists(input_mapping):
                    return(f"{input_mapping} Excel file not present for date")
            
            input_pdf = drive+r"\REPORT\FIFO reports\Raw Files" +f"\\Inventory Trial Balance_{inp_date}.pdf"
            # input_pdf = r"C:\Users\imam.khan\OneDrive - BioUrja Trading LLC\Documents\WEST PLAINS\REPORT\FIFO reports\Raw Files" +f"\\Inventory Trial Balance_{inp_date}.pdf"
            if not os.path.exists(input_pdf):
                    return(f"{input_pdf} Excel file not present for date {input_date}")

            # input_yc = drive+r"\REPORT\FIFO reports\Raw Files" +f"\\Inventory on site YC_{inp_date}.xlsx"
            # if not os.path.exists(input_yc):
            #         return(f"{input_yc} Excel file not present for date {input_date}")

            output_loc = drive+r"\REPORT\FIFO reports\Output files" +f"\\Inventory on site {loc}_{inp_date}.xlsx"
            
            # output_loc = r"C:\Users\imam.khan\OneDrive - BioUrja Trading LLC\Documents\WEST PLAINS\REPORT\FIFO reports\Output files" +f"\\Inventory on site {loc}_{inp_date}.xlsx"
            # ouput_yc = drive+r"\REPORT\FIFO reports\Output files" +f"\\Inventory on site YC_{inp_date}.xlsx"
            mtm_ouput_loc = drive+r"\REPORT\FIFO reports\Output files" +f"\\Inventory MTM Excel Report {monthYear}.xlsx"
            # mtm_ouput_loc = r"C:\Users\imam.khan\OneDrive - BioUrja Trading LLC\Documents\WEST PLAINS\REPORT\FIFO reports\Output files" +f"\\Inventory MTM Excel Report {monthYear}.xlsx"

            

            retry=0
            while retry < 10:
                try:
                    wb=xw.Book(input_xl)
                    break
                except Exception as e:
                    time.sleep(2)
                    retry+=1
                    if retry ==9:
                        raise e

            retry=0
            while retry < 10:
                try:
                    inp_sht = wb.sheets[0]
                    break
                except Exception as e:
                    time.sleep(2)
                    retry+=1
                    if retry ==9:
                        raise e
            
            inp_sht.copy(wb.sheets[0],name="Master file")
            inp_sht.range("1:1").insert(shift = "down")
            inp_sht.api.AutoFilterMode=False

            last_column = num_to_col_letters(inp_sht.range("A2").end('right').column)
            if inp_sht.range("A2").value == None:
                inp_sht.range("2:2").delete()
            col_headers = inp_sht.range("A2").expand("right").value
            for col in range(len(col_headers)):
                if col_headers[col] == "Trans  Date":
                    transDate = f"{num_to_col_letters(col+1)}"
                elif col_headers[col] == "Quantity":
                    quantityCol = f"{num_to_col_letters(col+1)}"
                elif col_headers[col] == "Inventory Value":
                    invValCol = f"{num_to_col_letters(col+1)}"
                    qty_col = num_to_col_letters(col+2)
                    value_col = num_to_col_letters(col+3)
                    price_col = num_to_col_letters(col+4)
                elif col_headers[col] == "Unit Cost":
                    unit_cost_col_num = col+1
                    unit_cost_col = num_to_col_letters(col+1)
                elif col_headers[col] == "Name":
                    cust_name_col_num = col+1+3 #extra 3 for 3 columns being inserted
                    cust_name_col = f"{num_to_col_letters(col+1+3)}"
            
            last_row = inp_sht.range(f'{transDate}'+ str(inp_sht.cells.last_cell.row)).end('up').row
            inp_sht.range(f"A2:{last_column}{last_row}").api.Sort(Key1=inp_sht.range(f"{transDate}2:{transDate}{last_row}").api,Order1=2,DataOption1=0,Orientation=1)
            #inserting  columns after INVENTORY VALUE
            inp_sht.range(f"{qty_col}:{qty_col}").insert(shift = "down")
            inp_sht.range(f"{qty_col}2").value = "Qty"
            inp_sht.range(f"{value_col}:{value_col}").insert(shift = "down")
            inp_sht.range(f"{value_col}2").value = "Value"
            inp_sht.range(f"{price_col}:{price_col}").insert(shift = "down")
            inp_sht.range(f"{price_col}2").value = "Price"

            #Filtering Inter-Company and putting their quantiy to 0
            inp_sht.select()
            inp_sht.api.AutoFilterMode=False
            inp_sht.api.Range(f"{cust_name_col}2").AutoFilter(Field:=cust_name_col_num,Criteria1:="INTER-COMPANY PURCH/SALES")
            l_row = inp_sht.range(f"{quantityCol}2").end('down').row
            
            # inp_sht.api.Range(f"{quantityCol}3:{quantityCol}{l_row}").SpecialCells(12).Select()
            inp_sht.api.Range(f"3:{l_row}").SpecialCells(12).Select()
            # wb.app.selection.value = 0
            wb.app.selection.delete()
            inp_sht.api.AutoFilterMode=False
           
            if loc == "HRW":
                inp_sht.api.Range(f"{unit_cost_col}2").AutoFilter(Field:=unit_cost_col_num,Criteria1:="<=1.7")
            else:
                inp_sht.api.Range(f"{unit_cost_col}2").AutoFilter(Field:=unit_cost_col_num,Criteria1:="<=1")
            l_row = inp_sht.range(f"{quantityCol}2").end('down').row
            inp_sht.api.Range(f"{quantityCol}3:{quantityCol}{l_row}").SpecialCells(12).Select()
            wb.app.selection.value = 0
        
            
            inp_sht.api.AutoFilterMode=False

            if loc == "HRW":
                loc_dict = other_loc_extractor(input_pdf)
                retry=0
                while retry < 10:
                    try:
                        mtm_wb=xw.Book(input_mtm)
                        break
                    except Exception as e:
                        time.sleep(2)
                        retry+=1
                        if retry ==9:
                            raise e

                retry=0
                while retry < 10:
                    try:
                        mtm_sht = mtm_wb.sheets["INPUT DATA"]
                        break
                    except Exception as e:
                        time.sleep(2)
                        retry+=1
                        if retry ==9:
                            raise e
                retry=0
                while retry < 10:
                    try:
                        je_sht = mtm_wb.sheets["JE"]
                        break
                    except Exception as e:
                        time.sleep(2)
                        retry+=1
                        if retry ==9:
                            raise e

            df = pd.read_excel(input_mapping, sheet_name=loc)

                
            columns_1 = df.set_index(['Tab Name'])["SubLocation"].to_dict()
            columns_2 = df.set_index(['Tab'])["Pivot"].to_dict()

            for key in columns_1.keys():
                inp_sht.api.Range(f"A2").AutoFilter(Field:=1,Criteria1:=columns_1[key].split(','), Operator:=7)
                # inp_sht.api.Select()
                # inp_sht.api.AutoFilter.Range.SpecialCells(12).Select()
                # wb.app.selection.copy()
                new_sht = wb.sheets.add(key, after=wb.sheets[-1])
                inp_sht.api.AutoFilter.Range.Copy()
                new_sht.api.Range("A2").Select()
                new_sht.api.Paste()
                new_sht.range("N1").value = "MTM Qty"
                mtm_sht.api.AutoFilterMode=False
                mtm_last_row = mtm_sht.range(f'A'+ str(mtm_sht.cells.last_cell.row)).end('up').row
                #Freezing top 2 columns
                new_sht.api.Range("A3").Select()
                wb.app.api.ActiveWindow.FreezePanes = True
                # if loc  == "HRW":
                mtm_sht.activate()
                mtm_sht.api.AutoFilterMode=False #mtm_wb.app.selection
                mtm_sht.api.Range(f"D3").AutoFilter(Field:=4,Criteria1:=loc, Operator:=7)
                time.sleep(1)
                if key == 'HaySprings':
                    columns_1[key] = columns_1[key].replace("ALLIANCETE", "ALLIANCE")
                    columns_1[key] = columns_1[key].replace("LISCO - WE", "LISCO")
                # mtm_sht.api.Range(f"B3").AutoFilter(Field:=2,Criteria1:=columns_1[key].split(','), Operator:=7)
                mtm_sht.api.Range(f"C3").AutoFilter(Field:=3,Criteria1:=columns_2[key], Operator:=7)
                mtm_sht.api.Range(f"G4:G{mtm_last_row}").SpecialCells(12).Select()
                qty_sum=0
                price_sum = 0
                # je_sht.range(f'A'+ str(je_sht.cells.last_cell.row)).end('up').end('up').row
                for rng in mtm_wb.app.selection.address.split(','):
                    # if rng != '$G$3':
                    if type(mtm_sht.range(rng).value) is list:
                        qty_sum+=float(sum(mtm_sht.range(rng).value))
                        price_sum+=float(sum(mtm_sht.range(rng.replace("G","M")).value))
                    else:
                        qty_sum+=float(mtm_sht.range(rng).value)
                        price_sum+=float(mtm_sht.range(rng.replace("G","M")).value)



                            
                
                new_sht.range("O1").value = qty_sum
                new_sht.range("P1").value = price_sum
                
                
                new_sht.range("Q1").value = "MTM Price"
                new_sht.range("R1").formula = "=P1/O1"

                summ=0
                summ2=0
                i=2
                while summ<=new_sht.range("O1").value:
                    i+=1
                    # print(i)
                    # print(new_sht.range(f"M{i}").value)
                    summ+=new_sht.range(f"M{i}").value
                    summ2+=new_sht.range(f"O{i}").value
                
                new_sht.range(f"P{i}").value = summ
                new_sht.range(f"P{i}").color = "#FFFF00"
                new_sht.range(f"Q{i}").value = summ2
                new_sht.range(f"Q{i}").color = "#FFFF00"
                new_sht.range(f"R{i}").color = "#FFFF00"
                # new_sht.range(f"R{i}").value = summ2
                
                
                new_sht.range(f"R{i}").formula = f"=Q{i}/P{i}"
                print()


                mtm_sht.activate()
                mtm_sht.api.AutoFilterMode=False
                mtm_sht.api.Range(f"D3").AutoFilter(Field:=4,Criteria1:=loc, Operator:=7)
                time.sleep(1)
                mtm_sht.api.Range(f"B3").AutoFilter(Field:=2,Criteria1:=columns_1[key].split(','), Operator:=7)
                time.sleep(1)
                mtm_sht.api.Range(f"G3").AutoFilter(Field:=7,Criteria1:='<>0', Operator:=1, Criteria2:='<>')
                mtm_wb.activate()
                mtm_sht.activate()
                try:
                    mtm_sht.api.Range(f"O4:O{mtm_last_row}").SpecialCells(12).Select()
                    mtm_wb.app.selection.value = new_sht.range(f"R{i}").value
                except:
                    messagebox.showinfo("Info", f"Zero rows found for location {loc}-{columns_1[key].split(',')[0]}, check and continue")

            if loc == "HRW":
                mtm_sht.api.AutoFilterMode=False
                mtm_sht.api.Range(f"D3").AutoFilter(Field:=4,Criteria1:='<>HRW', Operator:=1, Criteria2:='<>YC')
                
                try:
                    rng_lst = mtm_sht.api.Range(f"D4:D{mtm_last_row}").SpecialCells(12).Address.split(",")
                except:
                    rng_lst = list(mtm_sht.api.Range(f"D4:D{mtm_last_row}").SpecialCells(12).Address)
                # 
                for rng in rng_lst:
                    rng.split("$")[2].replace(':','')
                    for i in range(int(rng.split("$")[2].replace(':','')), int(rng.split("$")[-1])+1):
                        # if i == 6:
                        #     continue
                        if mtm_sht.range(f"G{i}").value is not None and mtm_sht.range(f"G{i}").value != 0: #If quantity present
                            print(i)
                            try:
                                if mtm_sht.range(f"B{i}").value == "BROWNSVILL" and mtm_sht.range(f"D{i}").value == "MILO":
                                    mtm_sht.range(f"O{i}").value = loc_dict["SORGHUM"][mtm_sht.range(f"B{i}").value]
                                
                                else:
                                    mtm_sht.range(f"O{i}").value = loc_dict[mtm_sht.range(f"D{i}").value][mtm_sht.range(f"B{i}").value]
                            except:
                                mtm_sht.range(f"O{i}").value=0
                                pass

            wb.save(output_loc)
            wb.close()
        retry=0
        while retry < 10:
            try:
                mtm_je_sht = mtm_wb.sheets["JE"]
                break
            except Exception as e:
                time.sleep(2)
                retry+=1
                if retry ==9:
                    raise e
        mtm_sht.api.AutoFilterMode=False
        mtm_je_sht.activate()
        mtm_je_sht.api.AutoFilterMode=False
        pivotCount = mtm_wb.api.ActiveSheet.PivotTables().Count
        for j in range(1, pivotCount+1):
            mtm_je_sht.activate()
            PivotCache = mtm_wb.api.PivotCaches().Create(SourceType=win32c.PivotTableSourceType.xlDatabase, SourceData=f"\'INPUT DATA\'!R3C1:R{mtm_last_row}C19", Version=win32c.PivotTableVersionList.xlPivotTableVersion14)
            mtm_wb.api.ActiveSheet.PivotTables(j).ChangePivotCache(PivotCache)
            
            mtm_wb.api.ActiveSheet.PivotTables(j).PivotCache().Refresh()   
        # mtm_wb.api.ActiveSheet.PivotTables(2).PivotCache().Refresh() 
        mtm_sht.activate()
        mtm_sht.api.AutoFilterMode=False
        mtm_wb.save(mtm_ouput_loc)
        # mtm_wb.app.quit()
        
        return f"Fifo reports Genrated for {input_date}"
    except Exception as e:
        raise e
    finally:
        try:
            mtm_wb.app.quit()
        except:
            pass
    
def bank_recons_rep(input_date,output_date):
    try:
        input_sheet = drive+r'\REPORT\Bank Recons\Raw Files\Raw Template'+f'\\template.xls'
        if not os.path.exists(input_sheet):
            return(f"{input_sheet} Excel file not present for date {input_date}")
        pdf_input=drive+r'\REPORT\Bank Recons\Raw Files'+f'\\Outstanding Checks Report_{input_date}.pdf'
        if not os.path.exists(pdf_input):
                return(f"{pdf_input} Excel file not present for date {input_date}")
        pdf_input2=drive+r'\REPORT\Bank Recons\Raw Files'+f'\\BOA 4003_{input_date}.pdf'
        if not os.path.exists(pdf_input2):
                return(f"{pdf_input2} Excel file not present for date {input_date}")
        # job_name = "BANK_RECONS_Automation"
        output_location = drive+r'\REPORT\Bank Recons\Output Files'
        with open(pdf_input, 'rb') as f:
                    pdf = PyPDF2.PdfFileReader(f)
                    number_of_pages = pdf.getNumPages()
                    print(number_of_pages) 
        i=1 
        date_area=["8.798,105.876,47.048,508.266"]
        df=read_pdf(pdf_input,stream=True, multiple_tables=True,pages=i,area=date_area,silent=True,guess=False)
        text_value=df[0].columns[0]
        Required_date=text_value[text_value.find("To"):].split()[1]

        dictBOA={}
        dictJP={}
        for i in range(i,number_of_pages+1):
            test_area=["35.573,23.256,66.173,297.891"]
            df=read_pdf(pdf_input,stream=True, multiple_tables=True,pages=i,area=test_area,silent=True,guess=False)
            Extracted_value=df[0].columns[1]
            # Extracted_value=[item.replace(':', '') for item in Extracted_value]
            column_seperator=["408,500"]
            df=read_pdf(pdf_input,stream=True, multiple_tables=True,columns=column_seperator,pages=i,silent=True,guess=False)
            df[0].drop(len(df[0])-1,inplace=True)
            if Extracted_value in str(df[0].loc[len(df[0])-1]):
                total=df[0].iloc[-1,:][1]
                 
            if "BANK OF AMERICA" in Extracted_value.upper():
                if '-' in Extracted_value:
                    Extracted_value=Extracted_value.split('-')[0].strip()
                else:
                    Extracted_value=Extracted_value.split()[0].strip()  
                dictBOA[Extracted_value]=total

            if "JP MORGAN CD" in Extracted_value.upper():
                if '-' in Extracted_value:
                    Extracted_value=Extracted_value.split('-')[0].strip()
                else:
                    Extracted_value=Extracted_value.split()[0].strip()  
                dictJP[Extracted_value]=total    
               
        print("extraction done") 

        with open(pdf_input2, 'rb') as f:
                    pdf = PyPDF2.PdfFileReader(f)
                    page_object=pdf.getPage(0)
                    page_text=page_object.extractText() 
                    final_value=page_text[page_text.find("Closing Ledger Balance (015)"):].split()[4].split("€")[0]

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

        ws1=wb.sheets[0]       
        ws1.range("B6").value=Required_date
        ws1.range("B40").value=final_value
        jp_inserting_row=ws1.range('D1').end('down').end('down').end('down').row+1
        jp_inserting_end_row=ws1.range('D1').end('down').end('down').end('down').end('down').row
        for i in range(jp_inserting_row,jp_inserting_end_row+1):            
            try:
                ws1.range(f"E{i}").value = float(dictJP[((ws1.range(f'D{i}').value).split("CDA")[0].strip()).upper()].replace(',',''))
            except:
                ws1.range(f"E{i}").value = 0
        BOA_inserting_row=ws1.range(f'D'+ str(ws1.cells.last_cell.row)).end('up').end('up').end('up').row+1
        BOA_inserting_end_row=last_row =ws1.range(f'D'+ str(ws1.cells.last_cell.row)).end('up').end('up').row  
        for i in range(BOA_inserting_row,BOA_inserting_end_row+1):
            try:
                ws1.range(f"E{i}").value = float(dictBOA[((ws1.range(f'D{i}').value).split("CDA")[0].strip()).upper()].replace(',',''))
            except:
                ws1.range(f"E{i}").value = 0
        ws1.api.Range("B40").NumberFormat = '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'
        save_date=datetime.strptime(Required_date,"%m/%d/%Y")
        save_date=datetime.strftime(save_date,"%m.%d.%Y")       
        wb.save(f"{output_location}\\BANK RECONS_{save_date}.xls")
        return f"Bank Recons Report Generated for {save_date}"
    except Exception as e:
        raise e
    finally:
        try:
            wb.app.quit()
        except:
            pass


def strg_month_end_report(input_date, output_date):
    try:
        monthYear = datetime.strftime(datetime.strptime(input_date, "%m.%d.%Y"), "%b%Y").upper()
        monthYear2 = datetime.strftime(datetime.strptime(input_date, "%m.%d.%Y"), "%b %Y").upper()
        
        pdf_loc = drive+r'\REPORT\Storage Month End Report\Raw Files'+f"\\{monthYear}\\PDF"
        # pdf_loc = drive+r'\REPORT\Storage Month End Report\Raw Files'+f"\\{input_date}\\PDF"
        # pdf_loc = r'C:\Users\imam.khan\OneDrive - BioUrja Trading LLC\Documents\WEST PLAINS\REPORT\Storage Month End Report\Raw Files'+f"\\{monthYear}\\PDF"
        if not os.path.exists(pdf_loc):
            return(f"{pdf_loc} Excel file not present for date {input_date}")
        strg_accr_inp_loc = drive+r'\REPORT\Storage Month End Report\Raw Files\STORAGE ACCRUAL.xlsx'
        # strg_accr_inp_loc = r'C:\Users\imam.khan\OneDrive - BioUrja Trading LLC\Documents\WEST PLAINS\REPORT\Storage Month End Report\Raw Files\STORAGE ACCRUAL.xlsx'
        if not os.path.exists(strg_accr_inp_loc):
            return(f"{strg_accr_inp_loc} Excel file not present for date {input_date}")
        strg_je_inp_loc = drive+r'\REPORT\Storage Month End Report\Raw Files\STORAGE ACCRUAL JE.xlsx'
        # strg_je_inp_loc = r'C:\Users\imam.khan\OneDrive - BioUrja Trading LLC\Documents\WEST PLAINS\REPORT\Storage Month End Report\Raw Files\STORAGE ACCRUAL JE.xlsx'
        if not os.path.exists(strg_je_inp_loc):
            return(f"{strg_je_inp_loc} Excel file not present for date {input_date}")

        input_qty_xl = drive+r'\REPORT\Storage Month End Report\Raw Files\STORAGE QTY.xlsx'
        # input_qty_xl = r'C:\Users\imam.khan\OneDrive - BioUrja Trading LLC\Documents\WEST PLAINS\REPORT\Storage Month End Report\Raw Files\STORAGE QTY.xlsx'

        input_qty_pdf = drive+r'\REPORT\Storage Month End Report\Raw Files'f'\\{monthYear}\\DailyPositionRecordForm2539A.pdf'
        # input_qty_pdf = r'C:\Users\imam.khan\OneDrive - BioUrja Trading LLC\Documents\WEST PLAINS\REPORT\Storage Month End Report\Raw Files'f'\\{monthYear}\\DailyPositionRecordForm2539A.pdf'
        
        loc_dict = {}
        qty_loc_dict = {}

        for loc in glob.glob(pdf_loc+"\\*.pdf"):
            # loc =  drive+r'\REPORT\Storage Month End Report\Raw Files\FEB2022\DailyPositionRecordForm2539A.pdf'
            # df = read_pdf(loc, pages = 'all', guess = False, stream = True,
            #                                 pandas_options={'header':0}, area = ["65,630,590,735"], columns=["680"])
            df = read_pdf(loc, pages = 'all', guess = False, stream = True,
                                            pandas_options={'header':0}, area = ["65,320,590,735"], columns=["450,520,680"])

            df = pd.concat(df, ignore_index=True)
            location = loc.split("\\")[-1].split(".")[0]
            if location == "ALLIANCET":
                location = "ALLIANCE TERMINAL"
            if location == "BROWNSVILLE":
                location = "BROWNSVILL"                              
            if location == "HAYSPRING":
                location = "HAY SPRINGS"
            
            commodity = loc.split("\\")[-1].split(".")[1]
            value = float(df.iloc[-1][-1].replace(",",""))
            qty_value = float(df.iloc[-1][0].replace(",",""))

            # location_lst.append(loc.split("\\")[-1].split(".")[0])
            # commodity_lst.append(loc.split("\\")[-1].split(".")[1])
            # values_lst.append(float(df.iloc[-1][-1].replace(",","")))

            if location in loc_dict.keys():  
                if commodity in loc_dict[location].keys():
                    loc_dict[location][commodity].append(value)
                else:
                    loc_dict[location][commodity] = [value]
            else:  
                loc_dict[location] = {}
                loc_dict[location][commodity] = [value]


            if location in qty_loc_dict.keys():  
                if commodity in qty_loc_dict[location].keys():
                    qty_loc_dict[location][commodity].append(qty_value)
                else:
                    qty_loc_dict[location][commodity] = [qty_value]
            else:
                if location =="BROWNSVILL":
                    qty_loc_dict["BROWNSVILLE"] = {}
                    qty_loc_dict["BROWNSVILLE"][commodity] = [qty_value]
                else:
                    qty_loc_dict[location] = {}
                    qty_loc_dict[location][commodity] = [qty_value]


        storage_accrual(input_date,strg_accr_inp_loc, monthYear, loc_dict)
        storage_je(strg_je_inp_loc, input_date, loc_dict)
        storage_qty(input_date,input_qty_pdf, input_qty_xl, monthYear2, qty_loc_dict)

        return f"Storage Month End Report Generated for {input_date}"
    except Exception as e:
        raise e

def payables_gl_entry_monthly(input_date, output_date):
    try:
        job_name = "Payables_GL_Entry_Monthly"
        input_sheet = drive+r'\REPORT\Unsettled Payables\Output files'+f'\\Unsettled Payables _{input_date}.xlsx' 
        if not os.path.exists(input_sheet):
            return(f"{input_sheet} Excel file not present for date {input_date}")
        template_file=drive+r'\REPORT\Month End GL Entries\Raw Files\template'+f'\\template2.xlsx'
        if not os.path.exists(template_file):
            return(f"{template_file} Excel file not present for date {input_date}")
        
        output_location = drive+r'\REPORT\Month End GL Entries\Output Files'
        xw.App.display_alerts = False
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
        ws4=wb.sheets["Pivot BB"]            
        print('sbb') 
        retry=0
        while retry < 10:
            try:
                # app = xw.App(add_book=False)
                # app.display_alerts = False
                # previous_wb = app.books.api.Open(template_file, UpdateLinks=False)
                previous_wb = xw.Book(template_file,update_links=False)
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry ==10:
                    raise e     
        # previous_wb = xw.Book(template_file,update_links=True)
        CTM_JE_sht = previous_wb.sheets("JE")
        CTM_JE_sht.api.Copy(None, After=ws4.api)
        time.sleep(2)
        GL_Look_up_sht = previous_wb.sheets("GL Lookup Table")
        GL_Look_up_sht.api.Copy(None, After=ws4.api)       
        previous_wb.close()
        ws1=wb.sheets[f"Unsettled Payables _{input_date}"]
        ws1.activate()        
        column_list = ws1.range("A1").expand('right').value
        Location_letter_column = num_to_col_letters(column_list.index('Location Name')+1)
        Commodity_Id_letter_column = num_to_col_letters(column_list.index('Commodity Id')+1)    
        last_column = ws1.range('A1').end('right').last_cell.column
        concatenate_cl1=num_to_col_letters(last_column+1)
        concatenate_cl2=num_to_col_letters(last_column+2)
        list1=["Loc","Comm","GL COGS Entry Debit Acct","GL BS Credit Entry"]
        list2=[f"=VLOOKUP({Location_letter_column}2,'GL Lookup Table'!A:B,2,0)",f"=VLOOKUP({Commodity_Id_letter_column}2,'GL Lookup Table'!A:B,2,0)",f'=CONCATENATE({concatenate_cl1}2,"-",{concatenate_cl2}2,"-",1000)',"0010000-2260-1000"] 
        last_column+=1
        i=0
        last_row = ws1.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row
        for values in list1:
            last_column_letter=num_to_col_letters(last_column)
            ws1.range(f"{last_column_letter}1").value = values
            ws1.range(f"{last_column_letter}1").api.Font.Bold = True
            ws1.range(f"{last_column_letter}2").value = list2[i]
            time.sleep(1)
            ws1.range(f"{last_column_letter}2").copy(ws1.range(f"{last_column_letter}2:{last_column_letter}{last_row}"))
            i+=1
            last_column+=1
        last_column = ws1.range('A1').end('right').last_cell.column  
        last_row = ws1.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row  
        ws6=wb.sheets["JE"]
        ws6.activate()
        wb.api.ActiveSheet.PivotTables(1).PivotCache().SourceData =f"Unsettled Payables _{input_date}!R1C1:R{last_row}C{last_column}"           #f'Details!R1C1:R{len(new_rows)+1}C18' #Updateing data source
        wb.api.ActiveSheet.PivotTables(1).PivotCache().Refresh() 
        # try:
        #     wb.api.ActiveSheet.PivotTables(1).PivotFields("GL COGS Entry Debit Acct").Orientation = win32c.PivotFieldOrientation.xlHidden
        #     wb.api.ActiveSheet.PivotTables(1).PivotFields("GL BS Credit Entry").Orientation = win32c.PivotFieldOrientation.xlHidden
        # except Exception as e:
        #     print("Columns not found from previous pivot")
        #     pass
        # wb.api.ActiveSheet.PivotTables(1).PivotFields("GL CTM BS Entry").Orientation = win32c.PivotFieldOrientation.xlRowField
        # wb.api.ActiveSheet.PivotTables(1).PivotFields("GL CTM BS Entry").Position = 1
        # wb.api.ActiveSheet.PivotTables(1).PivotFields("GL CTM COGS Entry").Orientation = win32c.PivotFieldOrientation.xlRowField
        # wb.api.ActiveSheet.PivotTables(1).PivotFields("GL CTM COGS Entry").Position = 2
        wb.api.ActiveSheet.PivotTables(1).PivotFields("Customer/Vendor Name").CurrentPage = "(All)"
        wb.api.ActiveSheet.PivotTables(1).PivotFields("Customer/Vendor Name").EnableMultiplePageItems = True
        try:
            wb.api.ActiveSheet.PivotTables(1).PivotFields("Customer/Vendor Name").PivotItems("INTER-COMPANY PURCH/SALES").Visible = False
        except Exception as e:
            pass   
        try:
            wb.api.ActiveSheet.PivotTables(1).PivotFields("Customer/Vendor Name").PivotItems("MACQUARIE COMMODITIES (USA) INC.").Visible = False
        except Exception as e:
            pass 
        wb.api.ActiveSheet.PivotTables(1).PivotFields("Location Name").EnableMultiplePageItems = True
        try:
            wb.api.ActiveSheet.PivotTables(1).PivotFields("Location Name").PivotItems("WEST PLAINS MEXICO").Visible = False
        except Exception as e:
            pass
        last_row = ws6.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row
        E_last_row = ws6.range(f'E'+ str(ws1.cells.last_cell.row)).end('up').row
        if last_row!=E_last_row:
            if E_last_row<last_row:
                ws6.range(f"E{E_last_row}:S{E_last_row}").copy(ws6.range(f"E{E_last_row}:S{last_row}"))
            else:
                last_row+=1
                ws6.range(f'E{last_row}:S{E_last_row}').api.Delete(win32c.DeleteShiftDirection.xlShiftUp)                       
        else:
            pass
        last_row = ws6.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row
        E_last_row = ws6.range(f'E'+ str(ws1.cells.last_cell.row)).end('up').row                     
        Acc_Date=datetime.strptime(input_date,"%m.%d.%Y")
        Acc_Date_input=datetime.strftime(Acc_Date,"%Y%m%d") 
        Rev_Date=Acc_Date + timedelta(days=1)
        Rev_Date_input=datetime.strftime(Rev_Date,"%Y%m%d")
        ws6.range("C1").value=Acc_Date_input
        ws6.range("C2").value=Rev_Date_input
        wb.save(f"{output_location}\\Unsettled Payables _{input_date} JE.xlsx")
        #CTM Combined _03.31.2022 JE
        wb.app.quit()
        return f"{job_name} Report for {input_date} generated succesfully"
    except Exception as e:
        raise e
    finally:
        try:
            wb.app.quit()
        except:
            pass

def receivables_gl_entry_monthly(input_date, output_date):
    job_name = "Receivables_GL_Entry_Monthly"
    try:    
        input_sheet = drive+r'\REPORT\Unsettled Receivables\Output files'+f'\\Unsettled Receivables _{input_date}.xlsx' 
        if not os.path.exists(input_sheet):
            return(f"{input_sheet} Excel file not present for date {input_date}")
        template_file=drive+r'\REPORT\Month End GL Entries\Raw Files\template'+f'\\template3.xlsx'
        if not os.path.exists(template_file):
            return(f"{template_file} Excel file not present for date {input_date}")
        output_location = drive+r'\REPORT\Month End GL Entries\Output Files'
        xw.App.display_alerts = False
        retry=0
        while retry < 10:
            try:
                wb=xw.Book(input_sheet,update_links=False)
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry ==10:
                    raise e 
        ws4=wb.sheets["Pivot BB"]            
        print('sbb') 
        retry=0
        while retry < 10:
            try:
                # app = xw.App(add_book=False)
                # app.display_alerts = False
                # previous_wb = app.books.api.Open(template_file, UpdateLinks=False)
                previous_wb = xw.Book(template_file,update_links=False)
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry ==10:
                    raise e     
        # previous_wb = xw.Book(template_file,update_links=True)
        JE_sht = previous_wb.sheets("JE")
        JE_sht.api.Copy(None, After=ws4.api)
        time.sleep(2)
        GL_Look_up_sht = previous_wb.sheets("GL Lookup Table")
        GL_Look_up_sht.api.Copy(None, After=ws4.api)       
        previous_wb.close()
        sheet_name=f"Unsettled Receivables _{input_date}"
        ws1=wb.sheets[sheet_name[:31]]
        ws1.activate()        
        column_list = ws1.range("A1").expand('right').value
        Location_letter_column = num_to_col_letters(column_list.index('Location Name')+1)
        Commodity_Id_letter_column = num_to_col_letters(column_list.index('Commodity Id')+1)    
        last_column = ws1.range('A1').end('right').last_cell.column
        concatenate_cl1=num_to_col_letters(last_column+1)
        concatenate_cl2=num_to_col_letters(last_column+2)
        list1=["Loc","Comm","Debit GL Balance Sheet Entry (Single Posting)","Credit Entries Income Stmt Accounts"]
        list2=[f"=VLOOKUP({Location_letter_column}2,'GL Lookup Table'!A:B,2,0)",f"=VLOOKUP({Commodity_Id_letter_column}2,'GL Lookup Table'!A:B,2,0)","0010000-1120-1000",f'=CONCATENATE({concatenate_cl1}2,"-",{concatenate_cl2}2,"-",1000)'] 
        last_column+=1
        i=0
        last_row = ws1.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row
        for values in list1:
            last_column_letter=num_to_col_letters(last_column)
            ws1.range(f"{last_column_letter}1").value = values
            ws1.range(f"{last_column_letter}1").api.Font.Bold = True
            ws1.range(f"{last_column_letter}2").value = list2[i]
            time.sleep(1)
            ws1.range(f"{last_column_letter}2").copy(ws1.range(f"{last_column_letter}2:{last_column_letter}{last_row}"))
            i+=1
            last_column+=1
        last_column = ws1.range('A1').end('right').last_cell.column  
        last_row = ws1.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row  
        ws6=wb.sheets["JE"]
        ws6.activate()
        wb.api.ActiveSheet.PivotTables(1).PivotCache().SourceData =f"{sheet_name[:31]}!R1C1:R{last_row}C{last_column}"           #f'Details!R1C1:R{len(new_rows)+1}C18' #Updateing data source
        wb.api.ActiveSheet.PivotTables(1).PivotCache().Refresh() 
        wb.api.ActiveSheet.PivotTables(1).PivotFields("Customer/Vendor Name").CurrentPage = "(All)"
        wb.api.ActiveSheet.PivotTables(1).PivotFields("Customer/Vendor Name").EnableMultiplePageItems = True
        try:
            wb.api.ActiveSheet.PivotTables(1).PivotFields("Customer/Vendor Name").PivotItems("INTER-COMPANY PURCH/SALES").Visible = False
        except Exception as e:
            pass   
        try:
            wb.api.ActiveSheet.PivotTables(1).PivotFields("Customer/Vendor Name").PivotItems("MACQUARIE COMMODITIES (USA) INC.").Visible = False
        except Exception as e:
            pass 
        wb.api.ActiveSheet.PivotTables(1).PivotFields("Location Name").EnableMultiplePageItems = True
        try:
            wb.api.ActiveSheet.PivotTables(1).PivotFields("Location Name").PivotItems("WEST PLAINS MEXICO").Visible = False
        except Exception as e:
            pass
        last_row = ws6.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row
        E_last_row = ws6.range(f'E'+ str(ws1.cells.last_cell.row)).end('up').row
        if last_row!=E_last_row:
            if E_last_row<last_row:
                ws6.range(f"E{E_last_row}:S{E_last_row}").copy(ws6.range(f"E{E_last_row}:S{last_row}"))
            else:
                last_row+=1
                ws6.range(f'E{last_row}:S{E_last_row}').api.Delete(win32c.DeleteShiftDirection.xlShiftUp)                       
        else:
            pass
        last_row = ws6.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row
        E_last_row = ws6.range(f'E'+ str(ws1.cells.last_cell.row)).end('up').row                      
        Acc_Date=datetime.strptime(input_date,"%m.%d.%Y")
        Acc_Date_input=datetime.strftime(Acc_Date,"%Y%m%d") 
        Rev_Date=Acc_Date + timedelta(days=1)
        Rev_Date_input=datetime.strftime(Rev_Date,"%Y%m%d")
        ws6.range("C1").value=Acc_Date_input
        ws6.range("C2").value=Rev_Date_input
        wb.save(f"{output_location}\\Unsettled Receivables _{input_date} JE.xlsx")
        #Unsettled Receivables _03.31.2022 JE
        wb.app.quit()
        return f"{job_name} Report for {input_date} generated succesfully"
    except Exception as e:
        raise e
    finally:
        try:
            wb.app.quit()
        except:
            pass


def ctm_gl_entry_monthly(input_date, output_date):
    try:    
        job_name = "CTM_GL_Entry_Monthly"
        input_sheet = drive+r'\REPORT\CTM Combined report\Output files'+f'\\CTM Combined _{input_date}.xlsx' 
        if not os.path.exists(input_sheet):
            return(f"{input_sheet} Excel file not present for date {input_date}")
        template_file=drive+r'\REPORT\Month End GL Entries\Raw Files\template'+f'\\template.xlsx'
        if not os.path.exists(template_file):
            return(f"{template_file} Excel file not present for date {input_date}")
        output_location = drive+r'\REPORT\Month End GL Entries\Output Files'
        xw.App.display_alerts = False
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
        ws4=wb.sheets["Pivot BB"]            
        print('sbb') 
        retry=0
        while retry < 10:
            try:
                # app = xw.App(add_book=False)
                # app.display_alerts = False
                # previous_wb = app.books.api.Open(template_file, UpdateLinks=False)
                previous_wb = xw.Book(template_file,update_links=False)
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry ==10:
                    raise e     
        # previous_wb = xw.Book(template_file,update_links=True)
        CTM_JE_sht = previous_wb.sheets("CTM JE")
        CTM_JE_sht.api.Copy(None, After=ws4.api)
        time.sleep(2)
        LOOKUP_sht = previous_wb.sheets("LOOKUP")
        LOOKUP_sht.api.Copy(None, After=ws4.api)
        time.sleep(2)
        GL_Look_up_sht = previous_wb.sheets("GL Look up Table CTM")
        GL_Look_up_sht.api.Copy(None, After=ws4.api)       
        previous_wb.close()
        ws1=wb.sheets[f"CTM Combined _{input_date}"]
        ws1.activate()      
        column_list = ws1.range("A1").expand('right').value
        Location_Id_letter_column = num_to_col_letters(column_list.index('Location Id')+1)
        Location_no_column=column_list.index('Location')+1
        Location_letter_column = num_to_col_letters(column_list.index('Location')+1)
        MTM_Commodity_letter_column = num_to_col_letters(column_list.index('MTM Commodity')+1)
        try:
            last_row = ws1.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row
            last_row+=1
            for i in range(2,int(f'{last_row}')):               
                if ws1.range(f"{Location_letter_column}{i}").value=="WEST PLAINS, LLC":
                    print(i)
                    if  ws1.range(f"{Location_Id_letter_column}{i}").value=="WEST COAST":
                        ws1.range(f"{Location_letter_column}{i}").value='WEST COAST- WEST PLAINS, LLC'
                    elif ws1.range(f"{Location_Id_letter_column}{i}").value=="WPGKC":
                        ws1.range(f"{Location_letter_column}{i}").value='KANSAS CITY - WEST PLAINS, LLC'
                    else:
                        print("new location found")  
        except Exception as e:
                print("failed in changing location name")
                raise e                  
        last_column = ws1.range('A1').end('right').last_cell.column
        concatenate_cl1=num_to_col_letters(last_column+1)
        concatenate_cl2=num_to_col_letters(last_column+2)
        list1=["Loc","Comm","GL CTM COGS Entry","GL CTM BS Entry"]
        list2=[f"=VLOOKUP({Location_letter_column}2,'GL Look up Table CTM'!A:B,2,0)",f"=VLOOKUP({MTM_Commodity_letter_column}2,'GL Look up Table CTM'!A:B,2,0)",f'=CONCATENATE({concatenate_cl1}2,"-",{concatenate_cl2}2,"-",3000)',f"=VLOOKUP({MTM_Commodity_letter_column}2,'GL Look up Table CTM'!A:C,3,0)"] 
        last_column+=1
        i=0
        last_row = ws1.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row
        for values in list1:
            last_column_letter=num_to_col_letters(last_column)
            ws1.range(f"{last_column_letter}1").value = values
            ws1.range(f"{last_column_letter}1").api.Font.Bold = True
            ws1.range(f"{last_column_letter}2").value = list2[i]
            time.sleep(1)
            ws1.range(f"{last_column_letter}2").copy(ws1.range(f"{last_column_letter}2:{last_column_letter}{last_row}"))
            i+=1
            last_column+=1
        last_column = ws1.range('A1').end('right').last_cell.column  
        last_row = ws1.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row  
        ws6=wb.sheets["CTM JE"]
        ws6.activate()
        wb.api.ActiveSheet.PivotTables(1).PivotCache().SourceData =f"CTM Combined _{input_date}!R1C1:R{last_row}C{last_column}"           #f'Details!R1C1:R{len(new_rows)+1}C18' #Updateing data source
        wb.api.ActiveSheet.PivotTables(1).PivotCache().Refresh() 
        wb.api.ActiveSheet.PivotTables(1).PivotFields("Customer").CurrentPage = "(All)"
        try:
            wb.api.ActiveSheet.PivotTables(1).PivotFields("Customer").PivotItems("INTER-COMPANY PURCH/SALES").Visible = False
        except Exception as e:
            pass   
        try:
            wb.api.ActiveSheet.PivotTables(1).PivotFields("Customer").PivotItems("MACQUARIE COMMODITIES (USA) INC.").Visible = False
        except Exception as e:
            pass 
        try:
            wb.api.ActiveSheet.PivotTables(1).PivotFields("Location").PivotItems("WPMEXICO").Visible = False
        except Exception as e:
            pass
        last_row = ws6.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row
        E_last_row = ws6.range(f'E'+ str(ws1.cells.last_cell.row)).end('up').row
        if last_row!=E_last_row:
            if E_last_row<last_row:
                ws6.range(f"E{E_last_row}:S{E_last_row}").copy(ws6.range(f"E{E_last_row}:S{last_row}"))
            else:
                ws6.range(f'{last_row}:{E_last_row}').api.Delete(win32c.DeleteShiftDirection.xlShiftUp)     
        else:
            pass
        last_row = ws6.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row
        E_last_row = ws6.range(f'E'+ str(ws1.cells.last_cell.row)).end('up').row  
        if last_row==E_last_row:
            ws6.range(f'E{last_row}:S{E_last_row}').api.Delete(win32c.DeleteShiftDirection.xlShiftUp) 
        Acc_Date=datetime.strptime(input_date,"%m.%d.%Y")
        Acc_Date_input=datetime.strftime(Acc_Date,"%Y%m%d") 
        Rev_Date=Acc_Date + timedelta(days=1)
        Rev_Date_input=datetime.strftime(Rev_Date,"%Y%m%d")
        ws6.range("C1").value=Acc_Date_input
        ws6.range("C2").value=Rev_Date_input
        wb.save(f"{output_location}\\CTM Combined _{input_date} JE.xlsx")
        #CTM Combined _03.31.2022 JE
        wb.app.quit()
        return f"{job_name} Report for {input_date} generated succesfully"
    except Exception as e:
        raise e
    finally:
        try:
            wb.app.quit()
        except:
            pass

def macq_accr_entry(input_date, output_date):
    try:
        xl_date = datetime.strftime(datetime.strptime(input_date, "%m.%d.%Y"), "%Y%m%d")
        next_date =datetime.strftime((datetime.strptime(input_date, "%m.%d.%Y")+timedelta(days=1)), "%Y%m%d")
        output_loc =  drive+r"\REPORT\Macquaire Accrual Entry\Output Files" +f"\\Macq Statement_{input_date}.xlsx"
        input_pdf = drive+r"\REPORT\Macquaire Accrual Entry\Raw Files" +f"\\Macq Statement_{input_date}.pdf"
        # input_pdf = r"C:\Users\imam.khan\OneDrive - BioUrja Trading LLC\Documents\WEST PLAINS\REPORT\Macquaire Accrual Entry\Raw Files" +f"\\Macq Statement_{input_date}.pdf"
        if not os.path.exists(input_pdf):
                return(f"{input_pdf} PDF file not present for date {input_date}")
        input_xl = drive+r"\REPORT\Macquaire Accrual Entry\Raw Files" +f"\\Macq Accrual_{input_date}.xlsx"
        # input_xl = r"C:\Users\imam.khan\OneDrive - BioUrja Trading LLC\Documents\WEST PLAINS\REPORT\Macquaire Accrual Entry\Raw Files" +f"\\Macq Accrual_{input_date}.xlsx"
        if not os.path.exists(input_xl):
                return(f"{input_xl} Excel file not present for date {input_date}")

        mapping_loc = drive+r"\REPORT\Macquaire Accrual Entry\Mapping.xlsx"

        df = pd.read_excel(mapping_loc)
        # map_dict = {k : g["Mapping( Open Trade Equity)"].to_dict() for k, g in df.set_index('Location as per Macquarie').groupby('GL')}
        lookup_dict = {k : g["Mapping( Open Trade Equity)"].to_dict() for k, g in df.set_index('GL').groupby('Location as per Macquarie')}
        acc_dict, net_liq = mac_accr_pdf(input_pdf)


        retry=0
        while retry < 10:
            try:
                wb=xw.Book(input_xl)
                break
            except Exception as e:
                time.sleep(2)
                retry+=1
                if retry ==9:
                    raise e

        retry=0
        while retry < 10:
            try:
                inp_sht = wb.sheets["Market revaluation"]
                break
            except Exception as e:
                time.sleep(2)
                retry+=1
                if retry ==9:
                    raise e

        # lookup_dict = {}
        inp_sht.range("A1").value = xl_date
        inp_sht.range("A2").value = next_date
        comm_last_row = inp_sht.range(f'B'+ str(inp_sht.cells.last_cell.row)).end('up').row
        net_liq_loc = 0
        while inp_sht.range(f"B{comm_last_row-net_liq_loc}").value != "Net liquidity value":
            net_liq_loc+=1
        #updating net liquidating value
        inp_sht.range(f"C{comm_last_row-net_liq_loc}").value = net_liq

        last_acc_row = inp_sht.range(f'A'+ str(inp_sht.cells.last_cell.row)).end('up').row
        for i in range(inp_sht.range("A2").end("down").row, last_acc_row+1):
            if inp_sht.range(f"A{i}").value is not None:
                try:
                    if int(inp_sht.range(f"A{i}").value)==775 and inp_sht.range(f"B{i}").value == '0010000-5005-9100':
                        try:
                            wheat_value =    -1 * (acc_dict[str(int(inp_sht.range(f"A{i}").value))]['CBT Wheat Futures'])
                        except:
                            wheat_value = 0
                        try:
                            wheat_future_value = -1 * (acc_dict[str(int(inp_sht.range(f"A{i}").value))]['KANSAS HARD RED WINTER WHEAT Futures'])
                        except:
                            wheat_future_value = 0
                        inp_sht.range(f"C{i}").value = wheat_value + wheat_future_value
                    else:
                        inp_sht.range(f"C{i}").value = -1 * (acc_dict[str(int(inp_sht.range(f"A{i}").value))][lookup_dict[int(inp_sht.range(f"A{i}").value)][inp_sht.range(f"B{i}").value]])
                except:
                    inp_sht.range(f"C{i}").value = 0

        inp_sht.activate()
        wb.save(output_loc)
        return f"Macquarie Accrual Report for {input_date} generated succesfully"
    except Exception as e:
        raise e
    finally:
        try:
            wb.app.quit()
        except:
            pass


def tkt_n_settlement_summ(input_date, output_date):
    try:
        monthYear = datetime.strftime(datetime.strptime(input_date, "%m.%d.%Y"), "%b %Y")
        Year = datetime.strftime(datetime.strptime(input_date, "%m.%d.%Y"), "%Y")
        input_datetime = datetime.strptime(input_date, "%m.%d.%Y")
        end_date = datetime.strftime(input_datetime+timedelta(days=1), "%m-%d-%Y")#
        st_date = datetime.strftime(input_datetime.replace(day=1), "%m-%d-%Y")
        tkt_query_xl = drive+r"\REPORT\Ticket And Settlement Summary\Raw Files" +f"\\Ticket Query {Year}.xlsx"
        # input_xl = r"C:\Users\imam.khan\OneDrive - BioUrja Trading LLC\Documents\WEST PLAINS\REPORT\Macquaire Accrual Entry\Raw Files" +f"\\Macq Accrual_{input_date}.xlsx"
        if not os.path.exists(tkt_query_xl):
            return(f"{tkt_query_xl} Excel file not present for year {Year}")

        settlement_xl = drive+r"\REPORT\Ticket And Settlement Summary\Raw Files\SETTLEMENT MAKER.xlsx"
        if not os.path.exists(settlement_xl):
            return(f"{settlement_xl} Excel file not present")


        template_xl = drive+r"\REPORT\Ticket And Settlement Summary\Raw Files\Ticket Query monYearTemplate.xlsx"
        # template_xl = drive+r"\REPORT\Ticket And Settlement Summary\Raw Files\Test.xlsx"
        if not os.path.exists(template_xl):
            return(f"{template_xl} Excel file not present")

        final_input = drive+r"\REPORT\Ticket And Settlement Summary\Output Files\Tickets and Settlement.xlsx"
        # template_xl = drive+r"\REPORT\Ticket And Settlement Summary\Raw Files\Test.xlsx"
        if not os.path.exists(final_input):
            return(f"{final_input} Excel file not present")

        output_file =  drive+r"\REPORT\Ticket And Settlement Summary\Output Files"+f"\\Tickets and Settlement.xlsx"
        det_output_file = drive+r"\REPORT\Ticket And Settlement Summary\Output Files"+f"\\Ticket Query {monthYear} Details.xlsx"


        #getting data from ticket query file till M column
        #query generated via dn conn in business i drive
        retry=0
        while retry < 10:
            try:
                tkt_wb=xw.Book(tkt_query_xl)
                break
            except Exception as e:
                time.sleep(2)
                retry+=1
                if retry ==9:
                    raise e

        retry=0
        while retry < 10:
            try:
                tkt_sht = tkt_wb.sheets[Year]
                break
            except Exception as e:
                time.sleep(2)
                retry+=1
                if retry ==9:
                    raise e

        retry=0
        while retry < 10:
            try:
                wb=xw.Book(template_xl)
                break
            except Exception as e:
                time.sleep(2)
                retry+=1
                if retry ==9:
                    raise e

        retry=0
        while retry < 10:
            try:
                tkt_ent_sht = wb.sheets["Tickets Entry"]
                break
            except Exception as e:
                time.sleep(2)
                retry+=1
                if retry ==9:
                    raise e
        last_row = tkt_sht.range(f'A'+ str(tkt_sht.cells.last_cell.row)).end('up').row
        tkt_ent_sht.cells.clear_contents()

        tkt_wb.activate()
        tkt_sht.activate()
        tkt_wb.api.ActiveSheet.ListObjects(1).ShowAutoFilter=False
        tkt_sht.api.AutoFilterMode=False
        # tkt_sht.api.Range(f"L1").AutoFilter(Field:=12,Criteria1:=f">={st_date}", Operator:=1, Criteria2:=f"<={end_date}")
        tkt_sht.api.Range(f"L1").AutoFilter(Field:=12,Criteria1:=f">={datetime.strptime(st_date, '%m-%d-%Y')}", Operator:=1, Criteria2:=f"<={datetime.strptime(end_date, '%m-%d-%Y')}")
        
        # tkt_sht.api.Range(f"L1").AutoFilter(Field:=12, Operator:=7, Criteria2:=[1,"2/28/2022"])
        tkt_wb.activate()
        time.sleep(2)
        tkt_sht.activate()
        time.sleep(2)
        tkt_sht.api.Range(f"A1:M{last_row-1}").SpecialCells(12).Select()
        # messagebox.showinfo(title="Info",message="Data is filtered and selected in ticket query 2022 sheet")
        tkt_wb.app.selection.copy()
        # tkt_sht.range(f"A1:M{last_row-1}").copy()
        tkt_ent_sht.range("A1").paste(paste="values_and_number_formats") #pasting only values
        tkt_last_row = tkt_ent_sht.range(f'A'+ str(tkt_ent_sht.cells.last_cell.row)).end('up').row
        
        # #adding Add by column by copy pasting add_by column already present in column K
        # i=0
        # while tkt_ent_sht.range(chr(ord("M")-i)+"1").value != "add_by":
        #     i+=1
        # add_by_col = chr(ord("M")-i)
        # country_col = chr(ord("M")-i+1)
        # tkt_ent_sht.range(f"{add_by_col}1").value = "Add By"
        # tkt_ent_sht.range(f"{add_by_col}2").expand("down").copy(tkt_ent_sht.range("N2"))


        tkt_ent_sht.range("N1").value = "Add By"
        tkt_ent_sht.range("N2").formula = "=VLOOKUP(K2,'Name (2)'!A:B,2,0)"
        tkt_ent_sht.range("N2").copy(tkt_ent_sht.range(f"N3:N{tkt_last_row}"))

        tkt_ent_sht.range("O1").value = "Team"
        tkt_ent_sht.range("O2").formula = "=VLOOKUP(K2,'Name (2)'!A:C,3,0)"
        tkt_ent_sht.range("O2").copy(tkt_ent_sht.range(f"O3:O{tkt_last_row}"))
        tkt_wb.save()
        
        # messagebox.showinfo(title="Info",message="Data is pasted in ticket entry sheet")
        #Now getting settlemt data same as above
        retry=0
        while retry < 10:
            try:
                set_wb=xw.Book(settlement_xl)
                break
            except Exception as e:
                time.sleep(2)
                retry+=1
                if retry ==9:
                    raise e

        retry=0
        while retry < 10:
            try:
                # set_sht = set_wb.sheets["Sheet1"]
                set_sht = set_wb.sheets[0]
                break
            except Exception as e:
                time.sleep(2)
                retry+=1
                if retry ==9:
                    raise e
        retry=0
        while retry < 10:
            try:
                inp_set_sht = wb.sheets["Settlement"]
                break
            except Exception as e:
                time.sleep(2)
                retry+=1
                if retry ==9:
                    raise e
        inp_set_sht.cells.clear_contents()
        
        set_last_row = set_sht.range(f'A'+ str(set_sht.cells.last_cell.row)).end('up').row
        set_wb.activate()          
        set_sht.activate()
        # set_sht.api.AutoFilterMode=False
        # set_sht.api.Range(f"E1").AutoFilter(Field:=5,Criteria1:=[first_date, last_date], Operator:=7)
        set_sht.api.Range(f"A1:M{set_last_row}").SpecialCells(12).Select()
        set_wb.app.selection.copy()
        time.sleep(1)
        inp_set_sht.range("A1").paste(paste="values_and_number_formats")
        time.sleep(1)
        inp_set_last_row = inp_set_sht.range(f'A'+ str(inp_set_sht.cells.last_cell.row)).end('up').row
        #adding Add by column by copy pasting add_by column already present in column K
        # i=0
        # while inp_set_sht.range(chr(ord("M")-i)+"1").value != "add_by":
        #     i+=1
        # add_by_col = chr(ord("M")-i)
        inp_set_sht.range("L1").value = "Add"
        inp_set_sht.range("L2").formula = "=+VLOOKUP(@H:H,'Name (2)'!A:B,2,FALSE)"
        inp_set_sht.range("L2").copy(inp_set_sht.range(f"L3:L{inp_set_last_row}"))

        inp_set_sht.range("M1").value = "Team"
        inp_set_sht.range("M2").formula = "=VLOOKUP(H2,'Name (2)'!A:C,3,0)"
        inp_set_sht.range("M2").copy(inp_set_sht.range(f"M3:M{inp_set_last_row}"))

        #Refreshing Pivots
        while retry < 20:
            try:
                tkt_p_sht = wb.sheets["Ticket Summary (2)"]
                break
            except Exception as e:
                time.sleep(2)
                retry+=1
                if retry ==19:
                    raise e
        tkt_p_sht.activate()
        # time.sleep(5)
        tkt_p_sht.range("A:E").clear()
        # time.sleep(5)
        tkt_p_sht.range("A1").select()
        # messagebox.showinfo(title="Info",message=f"Currently source data is '{tkt_ent_sht.name}'!R1C1:R{tkt_last_row}C15")
        # time.sleep(150)
        # messagebox.showinfo(title="Info",message=f"Currently source data is '{tkt_ent_sht.name}'!R1C1:R{tkt_last_row}C15")
        #First pivot
        PivotCache=wb.api.PivotCaches().Create(SourceType=win32c.PivotTableSourceType.xlDatabase, SourceData=f"'{tkt_ent_sht.name}'!R1C1:R{tkt_last_row}C15", Version=win32c.PivotTableVersionList.xlPivotTableVersion14)
        # time.sleep(5)
        PivotTable = PivotCache.CreatePivotTable(TableDestination=f"'Ticket Summary (2)'!R1C1", TableName="PivotTable1", DefaultVersion=win32c.PivotTableVersionList.xlPivotTableVersion14)        ###logger.info("Adding particular Row in Pivot Table")
        # time.sleep(5)
        PivotTable.PivotFields('Team').Orientation = win32c.PivotFieldOrientation.xlRowField
        PivotTable.PivotFields('Team').Position = 1
        PivotTable.PivotFields('Add By').Orientation = win32c.PivotFieldOrientation.xlRowField
        PivotTable.PivotFields('Commodity').Orientation = win32c.PivotFieldOrientation.xlDataField
        PivotTable.TableStyle2 = "PivotStyleMedium13"
        PivotTable.RowAxisLayout(1)
        wb.api.ActiveSheet.PivotTables("PivotTable1").InGridDropZones = True
        last_row = tkt_p_sht.range(f'A'+ str(set_sht.cells.last_cell.row)).end('up').row

        tkt_p_sht.range(f"A{last_row+5}").select()
        #Second pivot
        PivotCache=wb.api.PivotCaches().Create(SourceType=win32c.PivotTableSourceType.xlDatabase, SourceData=f"'{tkt_ent_sht.name}'!R1C1:R{tkt_last_row}C15", Version=win32c.PivotTableVersionList.xlPivotTableVersion14)
        PivotTable = PivotCache.CreatePivotTable(TableDestination=f"'Ticket Summary (2)'!R{last_row+11}C1", TableName="PivotTable2", DefaultVersion=win32c.PivotTableVersionList.xlPivotTableVersion14)        ###logger.info("Adding particular Row in Pivot Table")
        PivotTable.PivotFields('Mode').Orientation = win32c.PivotFieldOrientation.xlRowField
        PivotTable.PivotFields('Mode').Position = 1
        PivotTable.PivotFields('Team').Orientation = win32c.PivotFieldOrientation.xlRowField
        PivotTable.PivotFields('Add By').Orientation = win32c.PivotFieldOrientation.xlRowField
        PivotTable.PivotFields('Commodity').Orientation = win32c.PivotFieldOrientation.xlDataField
        PivotTable.TableStyle2 = "PivotStyleMedium13"
        PivotTable.RowAxisLayout(1)
        wb.api.ActiveSheet.PivotTables("PivotTable2").InGridDropZones = True

        #Updating Railtickets
        tkt_p_sht.range("G7").formula = f'=+GETPIVOTDATA("Commodity",$A${last_row+11},"Mode","R","Team","USA")+GETPIVOTDATA("Commodity",$A${last_row+11},"Mode","R","Team","VERTICAL")'

        # pivotCount = wb.api.ActiveSheet.PivotTables().Count
            
        # for j in range(1, pivotCount+1):
        #     wb.api.ActiveSheet.PivotTables(j).PivotCache().SourceData = f"'{tkt_ent_sht.name}'!R1C1:R{tkt_last_row}C15" #15 for O col
        #     wb.api.ActiveSheet.PivotTables(j).PivotCache().Refresh()

        tkt_wb.close()

        #Refreshing Pivots
        while retry < 10:
            try:
                set_p_sht = wb.sheets["Settlement Summary (2)"]
                break
            except Exception as e:
                time.sleep(2)
                retry+=1
                if retry ==9:
                    raise e
        set_p_sht.activate()

        pivotCount = wb.api.ActiveSheet.PivotTables().Count
            
        for j in range(1, pivotCount+1):
            wb.api.ActiveSheet.PivotTables(j).PivotCache().SourceData = f"'{inp_set_sht.name}'!R1C1:R{inp_set_last_row}C13" #13 for M col
            wb.api.ActiveSheet.PivotTables(j).PivotCache().Refresh()

        #Combining data for summary tab
        left_df = tkt_p_sht.range(f"B2:C{tkt_p_sht.range('C2').end('down').row}").options(pd.DataFrame, 
                                header=1,
                                index=False 
                                ).value[:-1]
        left_df.columns = ["Row Labels", "Tickets"]
        left_df.replace(to_replace='None', value=np.nan).dropna(inplace=True)
        left_df.dropna(inplace=True)
        right_df = set_p_sht.range(f"B1:C{set_p_sht.range('C2').end('down').row}").options(pd.DataFrame, 
                                header=1,
                                index=False
                                ).value[:-1]
        right_df.columns = ["Row Labels", "Settlements"]
        right_df.replace(to_replace='None', value=np.nan).dropna(inplace=True)
        right_df.dropna(inplace=True)

        merged_df = left_df.merge(right_df, on='Row Labels', how='outer')

        #inserting merged data in sheet 1
        wb.sheets["Sheet1"].clear()
        wb.sheets["Sheet1"].range("A1").options(pd.DataFrame, header=1, index=False, expand='table').value = merged_df


        #Refreshing Pivots
        while retry < 10:
            try:
                summ_p_sht = wb.sheets["Summary"]
                break
            except Exception as e:
                time.sleep(2)
                retry+=1
                if retry ==9:
                    raise e
        summ_p_sht.activate()

        pivotCount = wb.api.ActiveSheet.PivotTables().Count
            
        for j in range(1, pivotCount+1):
            wb.api.ActiveSheet.PivotTables(j).PivotCache().SourceData = f"'Sheet1'!R1C1:R{len(merged_df)+1}C3" #3 for C col
            wb.api.ActiveSheet.PivotTables(j).PivotCache().Refresh()

    
        

        new_wb = xw.Book(final_input)
        try: 
            current_month_sht = new_wb.sheets[monthYear]
            current_month_sht.clear()
        except:
            new_wb.sheets.add(monthYear,after=new_wb.sheets[-1])
            current_month_sht = new_wb.sheets[monthYear]
        time.sleep(1)
        # new_wb.sheets[0].name = monthYear
        tkt_p_sht.activate()
        # tkt_p_sht.api.Range(tkt_p_sht.api.Cells.SpecialCells(12).Address).Copy()
        # new_wb.activate()
        # current_month_sht.activate()
        # current_month_sht.api.Range("A1").Select()
        # current_month_sht.range("A1").api.Range("A1").PasteSpecial(Paste=	-4163)    #xlPasteValues
        # current_month_sht.autofit(axis="columns")

        #Generating data for final file
        #Ticket Summary Data sheets
        tkt_df1 = tkt_p_sht.range(f"A2:C{tkt_p_sht.range('C2').end('down').row}").options(pd.DataFrame, 
                                header=1,
                                index=False 
                                ).value
        # tkt_df1.columns = ["Add By", "Total"]
        # tkt_df1.replace(to_replace='None', value=np.nan).dropna(inplace=True)

        # next_table_row = tkt_p_sht.range("A2").end("down").end("down").row+1
        next_table_row = tkt_p_sht.range('C2').end('down').end('down').row
        last_row = tkt_p_sht.range(f'D'+ str(tkt_p_sht.cells.last_cell.row)).end('up').row
        tkt_df2 = tkt_p_sht.range(f"A{next_table_row}:D{last_row}").options(pd.DataFrame, 
                                header=1,
                                index=False
                                ).value
        # tkt_df2.columns = ["Add By", "Settlements"]
        # tkt_df2.replace(to_replace='None', value=np.nan).dropna(inplace=True)
        table3_col = num_to_col_letters(tkt_p_sht.range("C2").end("right").column)
        last_row = tkt_p_sht.range(f'F'+ str(tkt_p_sht.cells.last_cell.row)).end('up').row
        tkt_df3 = tkt_p_sht.range(f"{table3_col}1:{table3_col}{last_row}").options(pd.DataFrame, 
                                header=1,
                                index=False, 
                                expand='right').value
        
        #pasting data in final workbook Ticket Summary Sheet
        new_wb.activate()
        current_month_sht.activate()
        current_month_sht.range("A1").value = "Ticket Summary"
        current_month_sht.range("A1").api.Font.Bold = True
        current_month_sht.range("A1").color = "#FFFF00"
        current_month_sht.range("A1:B1").merge()
        current_month_sht.range("A1:B1").api.HorizontalAlignment = win32c.HAlign.xlHAlignCenter
        current_month_sht.range("A3").options(pd.DataFrame, header=1, index=False, expand='table').value = tkt_df1
        current_month_sht.range("A3").expand('right').api.Font.Bold = True
        # new_wb.app.selection.api.Font.Bold = True
        current_month_sht.range(f"A{len(tkt_df1)+8}").options(pd.DataFrame, header=1, index=False, expand='table').value = tkt_df2
        current_month_sht.range(f"A{len(tkt_df1)+8}").expand('right').select()
        new_wb.app.selection.api.Font.Bold = True
        current_month_sht.range(f"{table3_col}3").options(pd.DataFrame, header=1, index=False, expand='table').value = tkt_df3
        current_month_sht.range(f"{table3_col}3").expand('right').select()
        new_wb.app.selection.api.Font.Bold = True
        table3_last_col = num_to_col_letters(current_month_sht.range(f"{table3_col}3").end("right").column)
        table3_last_row = current_month_sht.range(table3_last_col+ str(current_month_sht.cells.last_cell.row)).end('up').row
        current_month_sht.range(f"{table3_last_col}2:{table3_last_col}{table3_last_row}").api.NumberFormat = "0%"
        current_month_sht.autofit(axis="columns")

        #setting Borders for 1st table
        border_range = current_month_sht.range(f"A3:A{len(tkt_df1)+3}").expand("right")
        set_borders(border_range)

        #setting Borders for 2nd table
        border_range = current_month_sht.range(f"A{len(tkt_df1)+8}:A{len(tkt_df1)+8+len(tkt_df2)}").expand("right")
        set_borders(border_range)

        #setting Borders for 3rd table
        table3_last_row = current_month_sht.range(f'F'+ str(current_month_sht.cells.last_cell.row)).end('up').row
        border_range = current_month_sht.range(f"{table3_col}3:{table3_col}{table3_last_row}").expand("right")
        set_borders(border_range)




        #Settlement Summary Data sheets
        settlement_row =  current_month_sht.range(f'D'+ str(current_month_sht.cells.last_cell.row)).end('up').row+5
        # 
        current_month_sht.range(f"A{settlement_row}").value = "Settlement Summary"
        current_month_sht.range(f"A{settlement_row}").api.Font.Bold = True
        current_month_sht.range(f"A{settlement_row}").color = "#FFFF00"
        current_month_sht.range(f"A{settlement_row}:B{settlement_row}").merge()
        current_month_sht.range(f"A{settlement_row}:B{settlement_row}").api.HorizontalAlignment = win32c.HAlign.xlHAlignCenter


        set_df1 = set_p_sht.range(f"A1:C{set_p_sht.range('C2').end('down').row}").options(pd.DataFrame, 
                                header=1,
                                index=False
                                ).value
        

        
        table2_col = num_to_col_letters(set_p_sht.range("C2").end("right").column)
        last_row = set_p_sht.range(chr(ord(table2_col)+1)+ str(set_p_sht.cells.last_cell.row)).end('up').row
        set_df2 = set_p_sht.range(f"{table2_col}2:{table2_col}{last_row}").options(pd.DataFrame, 
                                header=False,
                                index=False, 
                                expand='right').value

        set_df2.columns = [["Profit Center","Count","Percentage"]]
        
        #pasting data in final workbook Settlement Summary Sheet
        
        current_month_sht.range(f"A{settlement_row+3}").options(pd.DataFrame, header=1, index=False, expand='table').value = set_df1
        current_month_sht.range(f"A{settlement_row+3}").expand('right').select()
        new_wb.app.selection.api.Font.Bold = True
        # current_month_sht.range(f"A{settlement_row+3}").expand("right").api.Font.Bold = True
        current_month_sht.range(f"A{settlement_row+3+len(set_df1)+5}").options(pd.DataFrame, header=1, index=False, expand='table').value = set_df2
        current_month_sht.range(f"A{settlement_row+3+len(set_df1)+5}").expand('right').select()
        new_wb.app.selection.api.Font.Bold = True
        # current_month_sht.range(f"A{settlement_row+3+len(set_df1)+5}").expand("right").api.Font.Bold = True

        table2_last_col = num_to_col_letters(current_month_sht.range(f"A{settlement_row+3+len(set_df1)+5}").end("right").column)
        table2_last_row = current_month_sht.range(table2_last_col+ str(current_month_sht.cells.last_cell.row)).end('up').row
        current_month_sht.range(f"{table2_last_col}{settlement_row+3+len(set_df1)+5}:{table2_last_col}{table2_last_row}").api.NumberFormat = "0%"
        current_month_sht.autofit(axis="columns")
        
        
        # new_wb.sheets['Settlement Summary'].autofit(axis="columns")

        #setting Borders for 1st table
        border_range = current_month_sht.range(f"A{settlement_row+3}:A{settlement_row+3+len(set_df1)}").expand("right")
        set_borders(border_range)

        #setting Borders for 2nd table
        border_range = current_month_sht.range(f"A{settlement_row+3+len(set_df1)+5}:A{settlement_row+3+len(set_df1)+5+len(set_df2)}").expand("right")
        set_borders(border_range)
        

        #Settlement Summary Data sheets
        # new_wb.sheets.add("Consolidated Summary",after=new_wb.sheets[f"Settlement Summary"]) 
        summ_row =  current_month_sht.range(f'B'+ str(current_month_sht.cells.last_cell.row)).end('up').row +5
        # new_wb.sheets.add("Settlement Summary",after=new_wb.sheets[f"Ticket Summary"])
        current_month_sht.range(f"A{summ_row}").value = "Consolidated Summary"
        current_month_sht.range(f"A{summ_row}").api.Font.Bold = True
        current_month_sht.range(f"A{summ_row}").color = "#FFFF00"
        current_month_sht.range(f"A{summ_row}:B{summ_row}").merge()
        current_month_sht.range(f"A{summ_row}:B{summ_row}").api.HorizontalAlignment = win32c.HAlign.xlHAlignCenter



        summ_df = summ_p_sht.range('A3').options(pd.DataFrame, 
                                header=1,
                                index=False,
                                expand='table').value
        summ_df.columns = ['User', 'Sum of Tickets', 'Sum of Settlements']

        current_month_sht.range(f"A{summ_row+3}").options(pd.DataFrame, header=1, index=False, expand='table').value = summ_df
        current_month_sht.range(f"A{summ_row+3}").expand('right').select()
        new_wb.app.selection.api.Font.Bold = True
        current_month_sht.autofit(axis="columns")
        #setting Borders
        border_range = current_month_sht.range(f"A{summ_row+3}").expand("table")
        set_borders(border_range)

        #Saving workbooks
        wb.save(det_output_file)
        new_wb.save(output_file)

        wb.app.quit()

        return f"Ticket an Settlement Summary Generated for {input_date} Successfully"
    except Exception as e:
        raise e
    finally:
        try:
            wb.app.quit()
        except:
            pass


def credit_card_entry(input_date, output_date):
    try:
        job_name = 'Credit_Card_Entry'
        datetime_input=datetime.strptime(input_date,"%m.%d.%Y")
        input_month=datetime.strftime(datetime_input,"%B")
        input_year=datetime.strftime(datetime_input,"%Y")
        input_month_small=datetime.strftime(datetime_input,"%b").upper()
        input_year_small=datetime.strftime(datetime_input,"%y")
        input_month_no=datetime.strftime(datetime_input,"%m")
        date=datetime_input.replace(day=1)-timedelta(1)
        previous_month=datetime.strftime(date,"%B")
        previous_year=datetime.strftime(date,"%Y")
        input_csv = drive+r'\REPORT\Credit Card Entry\Raw Files'+f'\\Credit_Card_{input_month_no}.{input_year}.csv' 
        input_sheet = drive+r'\REPORT\Credit Card Entry\Output files'+f'\\{previous_month} {previous_year} Credit Card expense.xlsx' 

        working_sheet = f'{input_month} {input_year}'           # current month sheet name
        output_location = drive+r'\REPORT\Credit Card Entry\Output files'               
        
        
        cardName_df = pd.read_excel(input_sheet,sheet_name='Card List', usecols="C,D",index_col = 0)       #Data Frame of Card List
        # required dictionary with card_num as KEY and Name as Value
        req_dict = cardName_df.to_dict()['Name']
        # logging.info('Opening Workbook')
        retry=0
        while retry < 10:
            try:
                wb = xw.Book(input_sheet, update_links=False) 
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry ==10:
                    raise e  
        # logging.info('Adding Current month Sheet')
        try:
            wb.sheets.add(working_sheet,after=wb.sheets[-1].name)
        except:
            pass

        # logging.info('opening current month sheet')
        ws1 = wb.sheets[working_sheet]          # opening current month sheet
        ws1.clear()

        # opening raw data workbook
        # logging.info(' Opening raw data workbook')
        # raw_wb = xw.Book('Chase4284_Activity20211229_20220128_20220203.csv',update_links=True)
        raw_df = pd.read_csv(f'{input_csv}')
        # logging.info('Changing columns order as required')
        raw_df = raw_df[['Card','Type','Transaction Date','Post Date','Description','Amount','Category']]      #change columns order
        # logging.info('Adding data from from raw file')
        ws1.range('B1').options(index = False).value = raw_df          #adding raw file data
        # wb = xw.Book(wb_path).save()

        card_lst = list(raw_df['Card'])
        amount_lst = list(raw_df['Amount'])

        # logging.info('Changing  credit card number format')
        for i in range(len(card_lst)):
            if len(str(card_lst[i]))==3:
                card_lst[i] = f'XX-0{card_lst[i]}'
            else:
                card_lst[i] = f'XX-{card_lst[i]}'
            
        # print(card_lst) 
        # logging.info('Adding credit card number')
        ws1.range('B2').options(transpose = True).value = card_lst    #adding card number
        ws1.range('A1').value = 'Name'
        ws1.range("B1").value = 'Credit Card No'

        
        name_lst = []
        for i in card_lst:
            name_lst.append(req_dict[i])
        # print(name_lst)

        # logging.info('Remove negative sign from amount column')
        for i in range(len(amount_lst)):
            if amount_lst[i] < 0:
                amount_lst[i]*= (-1)
        
        ws1.range('G2').options(transpose = True).value = amount_lst   
        # logging.info('Adding emloyee name')
        ws1.range('A2').options(transpose = True).value = name_lst      #adding employee name
        # wb = xw.Book(wb_path).save()
        # final_df = pd.read_excel(wb_path,sheet_name=working_sheet)
        final_df = pd.DataFrame(ws1.range('A1').expand('table').value, columns=ws1.range('A1').expand('right').value)
        # logging.info('Sorting data in ascending order')
        final_df = final_df.sort_values(by='Name')     #sorting data frame
        ws1.range('A1').options(index = False).value = final_df

        ws1.range('1:1').api.Font.Bold = True
        ws1['1:1'].font.size = 12
        ws1.autofit()
        num_row = ws1.range('A1').end('down').row

        ws1.range(f'G2:G{num_row}').api.NumberFormat= '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'
        # wb = xw.Book(wb_path).save()
        name_lst.sort()
        # logging.info('Put Arpita Bhandari on top of the sheet')

        last_column = ws1.range('A1').end('right').last_cell.column
        last_column_letter=num_to_col_letters(last_column)
        last_row = ws1.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row
        ws1.range(f"A2:{last_column_letter}{last_row}").api.Sort(Key1=ws1.range(f"A2:A{last_row}").api,Order1=1,DataOption1=0,Orientation=1)
        ws1.api.Sort.SortFields.Clear()
        top_row = 2
        for i in range(2,num_row*2):
            if ws1.range(f'A{i}').value == 'Arpita Bhandari':
                ws1.range(f"{top_row}:{top_row}").api.Insert(win32c.InsertShiftDirection.xlShiftDown,win32c.InsertFormatOrigin.xlFormatFromRightOrBelow)
                i+=1
                ws1.range(f'{top_row}:{top_row}').value = ws1.range(f'{i}:{i}').value
                ws1.range(f'{i}:{i}').api.Delete(win32c.DeleteShiftDirection.xlShiftUp) 
            else:
                continue
            top_row+=1

        top_row = 2
        for i in range(2,last_row):
            if (ws1.range(f'C{i}').value == 'Payment') or (ws1.range(f'C{i}').value == 'Return'):
                ws1.range(f'G{i}').value = ws1.range(f'G{i}').value*(-1) 
            else:
                pass
            top_row+=1

        prev = ws1.range('A2').value
        prev_row = 2
        # logging.info('Insert blank row after every employee and adding total amount ')
        for i in range(2,num_row*2):
            curr = ws1.range(f'A{i}').value
            if prev ==curr:
                # ws1.range(f'A{i}').color = (0,128,0)
                continue
            else:
               
                ws1.range(f"{i}:{i}").api.Insert(win32c.InsertShiftDirection.xlShiftToRight)
                ws1.range(f'A{i}').value = ws1.range(f'A{i-1}').value
                ws1.range(f'B{i}').value = ws1.range(f'B{i-1}').value
                ws1.range(f'F{i}').value = 'Chase Corp Card Clearing Acct'
                ws1.range(f'F{i}').api.Font.Bold = True
                ws1.range(f'G{i}').formula = f'=-SUM(G{prev_row}:G{i-1})'
                ws1.range(f'G{i}').api.Font.Bold = True
                # ws1.range(f"{i}:{i}").api.Font.Bold = True
            prev = curr
            prev_row = i+1

        # wb = xw.Book(wb_path).save()
        ws1.autofit()
        last_row = ws1.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row
        column_list = ws1.range("A1").expand('right').value
        Name_no_column=column_list.index('Name')+1
        Name_letter_column = num_to_col_letters(Name_no_column)
        i = 2
        while i <= last_row:
            if ws1.range(f"{Name_letter_column}{i}").value == "Name": 
                ws1.range(f"{i}:{i}").api.Delete(win32c.DeleteShiftDirection.xlShiftUp)
                # print(i)
                i-=1                   
            else:
                i+=1
        column_list = ws1.range("A1").expand('right').value
        list1=["G/L ID Center","ACCT Part","Sub Acct","Location","GL","Back up details"]
        list2=["","",'','=IFERROR(INDEX(Code!$B:$B,MATCH(H2,Code!A:A,0)),"")','=IFERROR(INDEX(Code!$F:$F,MATCH(I2,Code!$E:$E,0)),"")',''] 
        last_column = ws1.range('A1').end('right').last_cell.column
        last_column+=1
        i=0
        last_row = ws1.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row
        for values in list1:
            last_column_letter2=num_to_col_letters(last_column-1)
            ws1.api.Range(f"{last_column_letter2}1").EntireColumn.Insert()
            ws1.range(f"{last_column_letter2}1").value = values
            ws1.range(f"{last_column_letter2}1").api.Font.Bold = True
            ws1.range(f"{last_column_letter2}2").value = list2[i]
            time.sleep(1)
            ws1.range(f"{last_column_letter2}2").copy(ws1.range(f"{last_column_letter2}2:{last_column_letter2}{last_row}"))
            i+=1
            last_column+=1
        ws1.autofit()
        last_column = ws1.range('A1').end('right').last_cell.column
        last_column_letter = num_to_col_letters(last_column)
        last_row = ws1.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row
        ws1.activate()
        insert_all_borders(cellrange=f"A1:{last_column_letter}{last_row}",working_sheet=ws1,working_workbook=wb)
        dt=datetime_input.replace(day=1)-timedelta(1)
        previous_month=datetime.strftime(dt,"%B")
        previous_year=datetime.strftime(dt,"%Y")
        previous_sheet=wb.sheets[f"{previous_month} {previous_year}"]
        previous_sheet.activate()
        last_row_num1 = previous_sheet.range('F1').end('down').end('down').row
        last_row_num2=last_row_num1+9
        last_row = ws1.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row
        last_row+=5
        wb.app.display_alerts=False
        previous_sheet.range(f"F{last_row_num1}:H{last_row_num2}").copy(ws1.range(f"F{last_row}"))
        wb.app.display_alerts=True
        ws1.range(f"G{last_row}").value=f"='{previous_month} {previous_year}'!G{last_row_num2}"
        directories_created=[f"Excel_Files_{input_month} {input_year}"]
        for directory in directories_created:
            path3 = os.path.join(output_location,directory)  
            try:
                os.makedirs(path3, exist_ok = True)
                print("Directory '%s' created successfully" % directory)
            except OSError as error:
                print("Directory '%s' can not be created" % directory)
        path3 = os.path.join(output_location,directory)         
        def remove_existing_files(path3):
            """_summary_

            Args:
                path3 (_type_): _description_

            Raises:
                e: _description_
            """           
        try:
            files = os.listdir(path3)
            if len(files) > 0:
                for file in files:
                    os.remove(path3 + "\\" + file) 
            else:
                print("No existing files available to reomve")
            print("Pause")
        except Exception as e:
            raise e
        remove_existing_files(path3)   
           
        row_list = ws1.range("A2").expand('down').value
        row_list_n = list(OrderedDict.fromkeys(row_list))
        for values in row_list_n:
            # logging.info('Opening Workbook')
            retry=0
            while retry < 10:
                try:
                    wb2 = xw.Book() 
                    break
                except Exception as e:
                    time.sleep(5)
                    retry+=1
                    if retry ==10:
                        raise e  
            wss1=wb2.sheets[0]
            time.sleep(1)
            ws1.range(f"A1:G1").copy(wss1.range("A1"))
            time.sleep(1)
            wb.activate()
            ws1.activate()
            ws1.api.Range(f"A1").AutoFilter(Field:=1, Criteria1:=values, Operator:=7)
            last_row = ws1.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row
            ws1.api.Range(f"A2:G{last_row}").SpecialCells(12).Select()
            time.sleep(1)
            wb.app.selection.copy(wss1.range("A2"))
            time.sleep(1)
            wb2.activate()
            wss1.activate()
            wss1.api.Range(f"H1").Value="Location"
            wss1.range(f"H1").api.Font.Size = "12"
            last_row = wss1.range(f'A'+ str(wss1.cells.last_cell.row)).end('up').row
            insert_all_borders(cellrange=f"H1:H{last_row}",working_sheet=wss1,working_workbook=wb2)
            wss1.autofit()       
            wb2.save(f"{output_location}\\Excel_Files_{input_month} {input_year}\\{values}.xlsx")
            wb2.close()      

        ws1.api.AutoFilterMode=False
        ws1.api.Rows("2:2").Select()
        wb.app.api.ActiveWindow.FreezePanes = True
        ws1.api.Range("H:H").NumberFormat="General"
        ws1.api.Range("I:I").NumberFormat="General"
        ws1.api.Range("J:J").NumberFormat="General"
        
        # save_month = (datetime_input+relativedelta(months=+1)).strftime("%B")
        ws1.autofit()
        file_name = f'{input_month} {input_year} Credit Card expense'
        wb.save(f"{output_location}\\{file_name}.xlsx")
        # raw_wb.close()
        return f"{job_name} Report for {input_date} generated succesfully"
    except Exception as e:
        raise e
        # logging.exception(str(e))
    finally:
        wb.app.quit()


def payroll_summ(input_date, output_date):
    try:
        input_datetime = datetime.strptime(input_date,"%m.%d.%Y")
        monthYear = datetime.strftime(datetime.strptime(input_date, "%m.%d.%Y"), "%b %y")
        input_pdf = drive+r"\REPORT\Payroll summary accounting report\Raw Files" +f"\\Payroll Summary By Cost Center *.pdf"
        # input_pdf = r"C:\Users\imam.khan\OneDrive - BioUrja Trading LLC\Documents\WEST PLAINS\REPORT\Macquaire Accrual Entry\Raw Files" +f"\\Macq Statement_{input_date}.pdf"
        # if not os.path.exists(input_pdf):
        #         return(f"{input_pdf} PDF file not present for date {input_date}")
        # input_xl = drive+r"\REPORT\Payroll summary accounting report\Raw Files" +f"\\Payroll by Dept - {monthYear}.xlsx"
        # input_xl = r"C:\Users\imam.khan\OneDrive - BioUrja Trading LLC\Documents\WEST PLAINS\REPORT\Macquaire Accrual Entry\Raw Files" +f"\\Macq Accrual_{input_date}.xlsx"
        # if not osE.path.exists(input_xl):
        #         return(f"{input_xl} Excel file not present for date {input_date}")
        template_xl = drive+r"\REPORT\Payroll summary accounting report\Raw Files" +f"\\Template.xlsx"
        # input_xl = r"C:\Users\imam.khan\OneDrive - BioUrja Trading LLC\Documents\WEST PLAINS\REPORT\Macquaire Accrual Entry\Raw Files" +f"\\Macq Accrual_{input_date}.xlsx"
        if not os.path.exists(template_xl):
                return(f"{template_xl} Excel file not present")

        gl_map_xl = drive+r"\REPORT\Payroll summary accounting report\Raw Files" +f"\\PayrollMapping.xlsx"
        # input_xl = r"C:\Users\imam.khan\OneDrive - BioUrja Trading LLC\Documents\WEST PLAINS\REPORT\Macquaire Accrual Entry\Raw Files" +f"\\Macq Accrual_{input_date}.xlsx"
        if not os.path.exists(gl_map_xl):
            return(f"{gl_map_xl} Excel file not present")
        output_location = drive+r'\REPORT\Payroll summary accounting report\Output Files'+f"\\Payroll by Dept - {monthYear}.xlsx"

        data = payroll_pdf_extractor(input_pdf, input_datetime, monthYear)

        # retry=0
        # while retry < 10:
        #     try:
        #         wb=xw.Book(input_xl)
        #         break
        #     except Exception as e:
        #         time.sleep(2)
        #         retry+=1
        #         if retry ==9:
        #             raise e

        retry=0
        # while retry < 10:
        #     try:
        #         # inp_sht = wb.sheets["Detail"]
        #         inp_sht = wb.sheets["PAYLOCITY"]
        #         break
        #     except Exception as e:
        #         time.sleep(2)
        #         retry+=1
        #         if retry ==9:
        #             raise e
        while retry < 10:
            try:
                wb=xw.Book(template_xl)
                break
            except Exception as e:
                time.sleep(2)
                retry+=1
                if retry ==9:
                    raise e

        retry=0
        while retry < 10:
            try:
                # t_sht = t_wb.sheets["Detail"]
                t_sht = wb.sheets["PAYLOCITY"]
                break
            except Exception as e:
                time.sleep(2)
                retry+=1
                if retry ==9:
                    raise e
        # inp_sht.range("F4:T4").expand("down").expand("down").delete()
        # inp_sht.range("A4:T4").expand("down").expand("down").delete()

        

        first_row = t_sht.range(f"A2").end("down").row+1


        column_list = t_sht.range(f"A{first_row-1}").expand('right').value
        beg_date_col = num_to_col_letters(column_list.index('Pay Period Beginning Date')+1)
        end_date_col = num_to_col_letters(column_list.index('Pay Period Ending Date')+1)
        ledger_date_col = num_to_col_letters(column_list.index('Check/Ledger Date')+1)
        gl_code_col = num_to_col_letters(column_list.index('GL')+1)
        gross_col = num_to_col_letters(column_list.index('Gross Paid')+1)
        ss_er_col = num_to_col_letters(column_list.index('SS - ER')+1)
        med_er_col = num_to_col_letters(column_list.index('Medicare - ER')+1)
        sui_col = num_to_col_letters(column_list.index('SUI')+1)
        c_ded_col = num_to_col_letters(column_list.index('Cafeteria 125 Deductions')+1)
        v_ded_col = num_to_col_letters(column_list.index('Voluntary Deductions')+1)
        last_col = len(column_list)

        col_data_list = [gross_col, ss_er_col, med_er_col, sui_col, c_ded_col, v_ded_col]

        last_column = t_sht.range(f'{beg_date_col}{first_row}').end('right').address.split("$")[1]
        last_row = t_sht.range(f'{beg_date_col}'+ str(t_sht.cells.last_cell.row)).end('up').row

        curr_gl_code_list = t_sht.range(f"{gl_code_col}{first_row}:{gl_code_col}{last_row}").value
        curr_gl_code_list = list(map(lambda x:int(x.split('-')[0]),curr_gl_code_list))
        
        

        #Reading gl code mapping excel file
        gl_map_df = pd.read_excel(gl_map_xl)
        
        gl_map_dict = gl_map_df.set_index(['GL-CODE'])["PDF CODE"].to_dict()

        gl_map_keys = list(gl_map_dict.keys())
        gl_map_values = list(gl_map_dict.values())

        
        new_key1 = set(data[list(data.keys())[0]].keys()).difference(curr_gl_code_list).difference(gl_map_values)
        new_key2 = set(data[list(data.keys())[1]].keys()).difference(curr_gl_code_list).difference(gl_map_values)
        if len(new_key1):
            wb.app.quit()

            return f"""Payroll Summary Report for {input_date} not generated, new key found in file {list(data.keys())[0]}, 
            Please review, update new key: {new_key1} in template file and rerun the job"""
        
        elif len(new_key2):
            wb.app.quit()

            return f"""Payroll Summary Report for {input_date} not generated, new key found in file {list(data.keys())[1]}, 
            Please review, update new key: {new_key2} in template file and rerun the job"""

        

        if len(data)>1:
            i=0
            for key in sorted(data.keys()):
                last_row = t_sht.range(f'{beg_date_col}'+ str(t_sht.cells.last_cell.row)).end('up').row
                first_key = list(data[key].keys())[0]
                if i==0:
                    t_sht.range(f'{beg_date_col}{first_row}').value = data[key][first_key]["Beggining Date"]
                    t_sht.range(f'{end_date_col}{first_row}').value = data[key][first_key]["Ending Date"]
                    i+=1
                else:
                    t_sht.range(f'{beg_date_col}{first_row}:{last_column}{last_row}').copy(t_sht.range(f'A{last_row+1}'))
                    t_sht.range(f'{beg_date_col}{last_row+1}').value = data[key][first_key]["Beggining Date"]
                    t_sht.range(f'{end_date_col}{last_row+1}').value = data[key][first_key]["Ending Date"]
                

                

        else:
            pass
        
       
        last_row = t_sht.range(f'{beg_date_col}'+ str(t_sht.cells.last_cell.row)).end('up').row
        for row in range(first_row, last_row+1):
            gl_code = int(t_sht.range(f"{gl_code_col}{row}").value.split('-')[0])
            full_gl_code = t_sht.range(f"{gl_code_col}{row}").value
            ledger_date = datetime.strftime(t_sht.range(f'{ledger_date_col}{row}').value, "%d-%m-%Y")
            if full_gl_code in gl_map_keys:
                gl_code = gl_map_dict[full_gl_code]
            # for col in col_data_list:
            #     try:
            #         t_sht.range(f"{col}{row}").value = data[ledger_date][gl_map_dict[full_gl_code]][t_sht.range(f"{col}{first_row-1}").value]
            #     except:
            #         t_sht.range(f"{col}{row}").value = 0
            # else:
            try:
                for col in col_data_list:
                    t_sht.range(f"{col}{row}").value = data[ledger_date][gl_code][t_sht.range(f"{col}{first_row-1}").value]
            except:
                t_sht.range(f"{col}{row}").value = 0

        print("Done")

        

        # inp_sht.range(f"A4:T{last_row}").delete()
        # last_row=4


















        
        # init_chr = "F"
        # last_column = inp_sht.range("F3").end('right').column-6 #considering f as intial col
        # for pdf_data in range(len(data)-1,-1,-1):
        #     t_sht.range(f"A4:T4").expand("down").copy(inp_sht.range(f"A{first_row}"))#copying data from template
        #     last_row = inp_sht.range(f'A'+ str(inp_sht.cells.last_cell.row)).end('up').row
        #     #inserting dates
        #     inp_sht.range(f"C{first_row}:C{last_row}").value = list(data.keys())[pdf_data]
            
        #     for row in range(first_row,last_row+1):
        #         for col in range(last_column):
        #             try:
        #                 inp_sht.range(f"{chr(ord(init_chr)+col)}{row}").value = data[list(data.keys())[pdf_data]][inp_sht.range(f"A{row}").value][inp_sht.range(f"{chr(ord(init_chr)+col)}3").value]
        #             except:
        #                 inp_sht.range(f"{chr(ord(init_chr)+col)}{row}").value = 0
        #             if row == last_row and pdf_data == 0:
        #                 inp_sht.range(f"{chr(ord(init_chr)+col)}{row+2}").formula = f'=SUM({chr(ord(init_chr)+col)}4:{chr(ord(init_chr)+col)}{row})'
        #         #updating first row as last row
        #     first_row = last_row+1
        # inp_sht.range(f"F4:{chr(ord(init_chr)+last_column)}{row+2}").api.NumberFormat = '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)' #Standard format
        #Pivot part starts
        retry=0
        while retry < 10:
            try:
                # p_sht = wb.sheets["Pivot"]
                p_sht = wb.sheets["JE"]
                break
            except Exception as e:
                time.sleep(2)
                retry+=1
                if retry ==9:
                    raise e


        #Now update prior month end date and current month end date
        #getting current file month last date as file is parked based on rueup month
        nxt_mnth = input_datetime.replace(day=28) + timedelta(days=4)
        curr_month_last_date = (nxt_mnth.replace(day=1) -timedelta(days=1)).date()
        prev_month_last_date = (input_datetime.replace(day=1) -timedelta(days=1)).date()

        p_sht.range(f"A1").value = prev_month_last_date
        p_sht.range(f"A2").value = curr_month_last_date

        
        p_sht.activate()
        # wb.api.ActiveSheet.PivotTables(1).PivotCache().SourceData = f"Detail!R3C1:R{last_row}C{last_column+6}" #Updateing data source, Removing initialization for f
        wb.api.ActiveSheet.PivotTables(1).PivotCache().SourceData = f"PAYLOCITY!R3C1:R{last_row}C{len(column_list)}" #Updateing data source, Removing initialization for f
        wb.api.ActiveSheet.PivotTables(1).PivotCache().Refresh()  


        #getting gl entry first row AND COLUMN
        pivot_header_row = p_sht.range(f"A2").end('down').end('down').row
        data_col = num_to_col_letters(p_sht.range(f"A{pivot_header_row}").end('right').end('right').end('right').end('right').column) #Ignoring **RECLASS P&L FROM CORPORATE TO LOCATIONS present in firsdt row
        data_last_col = num_to_col_letters(p_sht.range(f"{data_col}{pivot_header_row}").end('right').column)
        data_last_row = p_sht.range(f'{data_col}'+ str(p_sht.cells.last_cell.row)).end('up').row

       #Pivot part starts
        retry=0
        while retry < 10:
            try:
                # p_sht = wb.sheets["Pivot"]
                entry_sht = wb.sheets["ENTRY"]
                break
            except Exception as e:
                time.sleep(2)
                retry+=1
                if retry ==9:
                    raise e
        
        entry_sht.clear()
        p_sht.range(f"{data_col}{pivot_header_row}:{data_last_col}{data_last_row}").api.Copy()
        entry_sht.range("A1").paste(paste="values_and_number_formats")


        # p_last_row = p_sht.range(f'A'+ str(inp_sht.cells.last_cell.row)).end('up').row -1
        # p_sht.range(f"C5:D{p_last_row}").copy(p_sht.range("G5"))

        # #Updating Dates
        # p_sht.range("K1").value = datetime.strftime(input_datetime.replace(day=1)-timedelta(days=1), "%m/%d/%Y") #Last Monthend
        # p_sht.range("M1").value = datetime.strftime(input_datetime, "%m/%d/%Y") #Monthend

        # p_sht.range("L4").formula = "=C4" #take first date from pivot
        # p_sht.range("M4").formula = "=D4" #take second date from pivot

        # p_sht.range("U5").expand("down").value = f"PAYROLL RECLASSIFICATION {monthYear}"
        # p_sht.range("V5").expand("down").value = f"PAYROLL {monthYear}"
        
        wb.save(output_location)
        print()
        wb.app.quit()

        return f"Payroll Summary Report for {input_date} generated succesfully"
    except Exception as e:
        raise e
    
    finally:
        try:
            wb.app.quit()
        except:
            pass

def credit_card_gl(input_date, output_date):
    try:
        job_name = "Credit_Card_GL_Monthly"
        datetime_input=datetime.strptime(input_date,"%m.%d.%Y")
        input_month=datetime.strftime(datetime_input,"%B")
        input_year=datetime.strftime(datetime_input,"%Y")
        input_month_small=datetime.strftime(datetime_input,"%b").upper()
        input_year_small=datetime.strftime(datetime_input,"%y")
        lastday=calendar.monthrange(datetime_input.year,datetime_input.month)[1]
        input_month_no=datetime.strftime(datetime_input,"%m")
        insertdate=f'{input_year}{input_month_no}{lastday}'

        template_sheet=drive+r'\REPORT\Credit_Card_GL\Raw Files\template'+f'\\template.xlsx'
        retry=0
        if not os.path.exists(template_sheet):
            return(f"{template_sheet} Excel file not present in template folder") 
        while retry < 10:
            try:              
                template_wb=xw.Book(template_sheet,update_links=False)
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry ==10:
                    raise e    
        sht=template_wb.sheets[0]
        tlast_row = sht.range(f'A'+ str(sht.cells.last_cell.row)).end('up').row
        dict1=sht.range(f'A1:B{tlast_row}').options(dict).value
        template_wb.close()
        input_sheet = drive+r'\REPORT\Credit_Card_GL\Raw Files'+f'\\{input_month} {input_year} Credit Card expense.xlsx' 
        if not os.path.exists(input_sheet):
            return(f"{input_sheet} Excel file not present for date {input_date}")           
        output_location = drive+r'\REPORT\Credit_Card_GL\Output files'
        output_location_file=f'{output_location}'+f'\\{input_month} {input_year} Credit Card expense.xlsx'
        if os.path.exists(output_location_file):
            input_sheet=output_location_file
        xw.App.display_alerts = False
        retry=0
        while retry < 10:
            try:
                wb=xw.Book(input_sheet,update_links=False)
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry ==10:
                    raise e                       
        input_tab=wb.sheets[f"{input_month} {input_year}"]
        entry_tab=0
        try:
            entry_tab=wb.sheets[f"GS entry {input_month_small} {input_year_small}"]
            entry_tab.clear() 
        except:
            wb.sheets.add(f"GS entry {input_month_small} {input_year_small}",after=input_tab)        
        entry_tab=wb.sheets[f"GS entry {input_month_small} {input_year_small}"]  
        input_tab.activate()
        column_list = input_tab.range("A1").expand('right').value
        gl_letter_column = num_to_col_letters(column_list.index('GL')+1)
        last_row = input_tab.range(f'A'+ str(input_tab.cells.last_cell.row)).end('up').row
        entry_tab.activate()
        entry_tab.api.Range("A1").Select()
        input_tab.api.Range(f"A1:{gl_letter_column}{last_row}").Copy()
        time.sleep(5)
        entry_tab.api.Paste()
        wb.app.api.CutCopyMode=False
        entry_tab.autofit()
        last_row = entry_tab.range(f'A'+ str(entry_tab.cells.last_cell.row)).end('up').row
        column_list = entry_tab.range("A1").expand('right').value
        Description_no_column=column_list.index('Description')+1
        Description_letter_column = num_to_col_letters(Description_no_column)
        # i = 2
        # while i <= last_row:
        #     color_hex="ffc000"
        #     rgb_value=tuple(int(color_hex[i:i+2], 16) for i in (0, 2, 4))
        #     if entry_tab.range(f"{Description_letter_column}{i}").color == rgb_value: 
        #         entry_tab.range(f"{i}:{i}").api.Delete(win32c.DeleteShiftDirection.xlShiftUp)
        #         # print(i)
        #         i-=1                   
        #     else:
        #         i+=1
        for key, value in dict1.items():
            try:
                entry_tab.api.Range(f"A1").AutoFilter(Field:=1, Criteria1:=value, Operator:=7)
                last_row = entry_tab.range(f'A'+ str(entry_tab.cells.last_cell.row)).end('up').row
                last_column_letter=num_to_col_letters(entry_tab.range('A1').end('right').end('right').last_cell.column)
                cell_range=entry_tab.api.Range(f"A2:{last_column_letter}{last_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Address
                starting_index=int(cell_range.split(':')[0].replace("$",""))
                ending_index=int(cell_range.split(':')[1].replace("$",""))
                i = starting_index
                while i <= ending_index:
                    if key in entry_tab.range(f"{Description_letter_column}{i}").value: 
                        entry_tab.range(f"{i}:{i}").api.Delete(win32c.DeleteShiftDirection.xlShiftUp)
                        # print(i)                 
                    else:
                        i+=1
                last_row = entry_tab.range(f'A'+ str(entry_tab.cells.last_cell.row)).end('up').row
                last_column_letter=num_to_col_letters(entry_tab.range('A1').end('right').end('right').last_cell.column)
                cell_range=entry_tab.api.Range(f"A2:{last_column_letter}{last_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Address
                starting_index=int(cell_range.split(':')[0].replace("$",""))
                ending_index=int(cell_range.split(':')[1].replace("$","")) 
                Amount_no_column=column_list.index('Amount')+1
                Amount_letter_column = num_to_col_letters(Amount_no_column)      
                if entry_tab.range(f"{Amount_letter_column}{ending_index}").value==None:
                    entry_tab.range(f"{ending_index}:{ending_index}").api.Delete(win32c.DeleteShiftDirection.xlShiftUp)
 
                wb.app.api.ActiveSheet.ShowAllData()
            except:
                wb.app.api.ActiveSheet.ShowAllData()
                pass

        Type_no_column=column_list.index('Type')+1
        Type_letter_column = num_to_col_letters(Type_no_column)
        i = 2
        while i <= last_row:
            if entry_tab.range(f"{Type_letter_column}{i}").value == "Payment": 
                entry_tab.range(f"{i}:{i}").api.Delete(win32c.DeleteShiftDirection.xlShiftUp)
                row_list = entry_tab.range(f"{Description_letter_column}1").expand('down').value
                Chase_Corp=row_list.index('Chase Corp Card Clearing Acct')+1
                Amount_no_column=column_list.index('Amount')+1
                Amount_letter_column = num_to_col_letters(Amount_no_column)
                if entry_tab.range(f"{Amount_letter_column}{Chase_Corp}").value == None:
                    entry_tab.range(f"{Chase_Corp}:{Chase_Corp}").api.Delete(win32c.DeleteShiftDirection.xlShiftUp)
                i-=1                   
            else:
                i+=1        

        column_list = entry_tab.range("A1").expand('right').value
        list1=["MONTH"," ","description"," ","fiscal_year","gl_acct_nbr","journal_source_code","transaction_date","description","refrence_id","amount"]
        list2=[f"{input_month_small}'{input_year_small}","'00",'=+F2&" "&B2&" "&M2',"","C",'=+N2&H2&"-"&I2&"-"&J2',"GLM",f"{insertdate}","=TRIM(O2)","=+VLOOKUP(@A:A,Code!I:J,2,FALSE)","=+ROUND(G2,2)"] 
        last_column = entry_tab.range('A1').end('right').last_cell.column
        last_column+=1
        i=0
        last_row = entry_tab.range(f'A'+ str(entry_tab.cells.last_cell.row)).end('up').row
        for values in list1:
            last_column_letter=num_to_col_letters(last_column)
            entry_tab.range(f"{last_column_letter}1").value = values
            entry_tab.range(f"{last_column_letter}1").api.Font.Bold = True
            entry_tab.range(f"{last_column_letter}2").value = list2[i]
            time.sleep(1)
            entry_tab.range(f"{last_column_letter}2").copy(entry_tab.range(f"{last_column_letter}2:{last_column_letter}{last_row}"))
            i+=1
            last_column+=1
        entry_tab.autofit()
        entry_tab.activate() 
        last_row = entry_tab.range(f'A'+ str(entry_tab.cells.last_cell.row)).end('up').row  
        column_list = entry_tab.range("A1").expand('right').value        
        amount_column=num_to_col_letters(column_list.index('amount')+1)
        MONTH_letter_column = num_to_col_letters(column_list.index('MONTH')+1)
        insert_all_borders(cellrange=f"{MONTH_letter_column}1:{amount_column}{last_row}",working_sheet=entry_tab,working_workbook=wb)                    
        save_month = (datetime_input).strftime("%B")
        wb.save(f"{output_location}\\{save_month} {input_year} Credit Card expense.xlsx")
        wb.app.quit()
        return f"{job_name} Report for {input_date} generated succesfully"
    except Exception as e:
        raise e
    finally:
        try:
            wb.app.quit()
        except:
            pass


def unsettled_ar_by_location_part1(input_date, output_date):
    try:       
        job_name = 'Unsettled AR By Location - Part 1'
        output_raw_date=datetime.strptime(output_date,"%m.%d.%Y")
        output_date=datetime.strftime(output_raw_date,"%m.%d.%y")   
        input_raw_date=datetime.strptime(input_date,"%m.%d.%Y")
        input_date_short=datetime.strftime(input_raw_date,"%m.%d.%y") 
        input_sheet = drive+r'\REPORT\Unsettled AR By Location - Part 1\Raw Files'+f'\\Unsettled AR_{input_date}.xlsx'
        previous_output= drive+r'\REPORT\Unsettled AR By Location - Part 1\Output_Files'+f'\\Unsettled AR {output_date}_with reason.xlsx'
        if not os.path.exists(input_sheet):
            return(f"{input_sheet} Excel file not present for date {input_date}")
        if not os.path.exists(previous_output):
            return(f"{previous_output} Excel file not present") 

        source_folder = drive+r"\REPORT\Unsettled AR By Location - Part 1\Output_Files"
        destination_folder = drive+r"\REPORT\Unsettled AR By Location - Part 1\Output_Files"
        file_name=f"Unsettled AR {output_date}_with reason.xlsx"
        file_name2=f"Unsettled AR {input_date_short}_with reason.xlsx"
        source = source_folder + "\\"+ file_name
        destination = destination_folder +"\\"+ file_name2
        if os.path.isfile(source):
                shutil.copy(source, destination)
                print('copied', file_name)
        else:
            print(f"{file_name} not present in the folder please recheck")           

        retry=0
        while retry < 10:
            try:
                wb = xw.Book(input_sheet) 
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry ==10:
                    raise e 
        ws1=wb.sheets[0]
        ws1.activate
        column_list = ws1.range("A1").expand('right').value
        list1=["Lk Up","Lk Up","Diff"]
        for values in list1:
             delete_column_no = column_list.index(values)+1
             delete_column_letter=num_to_col_letters(delete_column_no)
             ws1.api.Range(f"{delete_column_letter}1").EntireColumn.Delete()
             column_list = ws1.range("A1").expand('right').value
        last_row = ws1.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row
        last_column_letter=num_to_col_letters(ws1.range('A2').end('right').last_cell.column)     
   
        retry=0
        while retry < 10:
            try:
                wb2 = xw.Book(destination) 
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry ==10:
                    raise e 
        wss1=wb2.sheets[0]
        wss1.activate 
        column_list2 = wss1.range("A1").expand('right').value
        PurchaseSale_column_no = column_list2.index("Purchase/Sale")+1
        PurchaseSale_column_letter=num_to_col_letters(PurchaseSale_column_no)
        last_row2 = wss1.range(f'A'+ str(wss1.cells.last_cell.row)).end('up').row
        last_column_letter2=num_to_col_letters(wss1.range('A1').end('right').end('right').last_cell.column)
        time.sleep(1)
        wb.app.DisplayAlerts = False
        wss1.api.AutoFilterMode=False
        cnt1 = ws1.api.UsedRange.Rows.Count
        cnt2 = wss1.api.UsedRange.Rows.Count
        ws1.api.Range(f"A2:{last_column_letter}{last_row}").Copy(wss1.api.Range(f"{PurchaseSale_column_letter}2"))  
        wb.app.DisplayAlerts = True
        time.sleep(1)
        wss1.activate()
        try:
            Reason_column_no = column_list2.index("Reason")+1
            Reason_column_letter=num_to_col_letters(Reason_column_no)
        except:
            Reason_column_no = column_list2.index("Reasons")+1
            Reason_column_letter=num_to_col_letters(Reason_column_no)   
        last_row22 = wss1.range(f'F'+ str(ws1.cells.last_cell.row)).end('up').row  
        if cnt1>cnt2:
            wss1.range(f"A{last_row2}:{Reason_column_letter}{last_row2}").copy(wss1.range(f"A{last_row2}:{Reason_column_letter}{last_row22}")) #old change
            # wss1.range(f"F{last_row2+1}:AK{last_row}").delete()  #new change
        else:
            wss1.range(f"A{cnt1+1}").expand('table').delete()
        retry=0
        while retry < 10:
            try:
                wb3 = xw.Book(source) 
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry ==10:
                    raise e
        wb2.activate()
        wss1.activate()
        wss1.api.Range(f"{Reason_column_letter}2").Value=f"=VLOOKUP(A2,'[Unsettled AR {output_date}_with reason.xlsx]Sheet1'!$A:$E,5,0)"
        wss1.range(f"{Reason_column_letter}2").copy(wss1.range(f"{Reason_column_letter}2:{Reason_column_letter}{last_row22}"))
        time.sleep(1)
        wss1.range(f"{Reason_column_letter}:{Reason_column_letter}").copy()
        time.sleep(1)
        wss1.range(f"{Reason_column_letter}:{Reason_column_letter}").paste(paste="values_and_number_formats")
        wb2.app.api.CutCopyMode=False
        try:
            wss1.api.Range(f"{Reason_column_letter}1").AutoFilter(Field:=f'{Reason_column_no}', Criteria1:=['#N/A'], Operator:=7)
            time.sleep(1)
            wss1.api.Range(f"{Reason_column_letter}2:{Reason_column_letter}{last_row22}").SpecialCells(win32c.CellType.xlCellTypeVisible).Select()
            time.sleep(1)
            wb2.app.api.Selection.Value=None
            time.sleep(1)
            wb2.app.api.ActiveSheet.ShowAllData()
        except:
            wb2.app.api.ActiveSheet.ShowAllData()
            pass
        try:
            CVC_column_no = column_list2.index("Customer/Vendor Name")+1
            CVC_column_letter=num_to_col_letters(CVC_column_no)
            wss1.api.Range(f"{CVC_column_letter}1").AutoFilter(Field:=f'{CVC_column_no}', Criteria1:=['INTER-COMPANY PURCH/SALES'], Operator:=7)
            time.sleep(1)
            last_column_letter21=num_to_col_letters(wss1.range('A1').end('right').end('right').last_cell.column)
            wss1.api.Range(f"A2:{last_column_letter21}{last_row22}").SpecialCells(win32c.CellType.xlCellTypeVisible).Select()
            time.sleep(1)
            wb2.app.api.Selection.Delete(win32c.DeleteShiftDirection.xlShiftUp)
            time.sleep(1)
            wb2.app.api.ActiveSheet.ShowAllData()
        except:
            wb2.app.api.ActiveSheet.ShowAllData()
            pass
   
        wb2.save()   
       
        return f"{job_name} Report for {input_date} generated succesfully"

    except Exception as e:
        raise e
    finally:
        try:
            wb.app.quit()
        except:
            pass



def unsettled_ar_by_location_part2(input_date, output_date):
    try:       
        job_name = 'unsettled_ar_by_location_part2'  
        input_raw_date=datetime.strptime(input_date,"%m.%d.%Y")
        input_date_short=datetime.strftime(input_raw_date,"%m.%d.%y") 
        input_sheet = drive+r'\REPORT\Unsettled AR By Location - Part 2\Raw Files'+f'\\Unsettled AR {input_date_short}_with reason.xlsx'
        template_sheet= drive+r'\REPORT\Unsettled AR By Location - Part 2\Raw Files\Template'+f'\\Template.xlsx'
        output_location=drive+r"\REPORT\Unsettled AR By Location - Part 2\Output_Files"        
        if not os.path.exists(input_sheet):
            return(f"{input_sheet} Excel file not present for date {input_date}")
        if not os.path.exists(template_sheet):
            return(f"{template_sheet} Excel file not present") 

        retry=0
        while retry < 10:
            try:
                wb = xw.Book(template_sheet) 
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry ==10:
                    raise e 

        retry=0
        while retry < 10:
            try:
                wb2 = xw.Book(input_sheet) 
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry ==10:
                    raise e 

        wss1=wb2.sheets[0]
        wss1.activate()
        column_list = wss1.range("A1").expand('right').value
        Location_Name_column_no = column_list.index("Location Name")+1
        Location_Name_column_letter=num_to_col_letters(Location_Name_column_no)
        dict1={'Unsettled Gering':["GERING - WEST PLAINS, LLC"],'Unsettled AR HS':['HAY SPRINGS - WEST PLAINS, LLC'],"Unsettled Brownsville":["BROWNSVILLE - WEST PLAINS, LLC"],
        "Unsettled JT":['JOHNSTOWN - WEST PLAINS, LLC','OMAHA COMM - WEST PLAINS, LLC'],'Unsettled Omaha':['OMAHA TERMINAL ELEVATOR - WEST PLAINS, LLC'],'All Location':"<>"}      
        for key, value in dict1.items():
            try:
                wss1.api.Range(f"{Location_Name_column_letter}1").AutoFilter(Field:=f'{Location_Name_column_no}', Criteria1:=value, Operator:=7)
                last_row = wss1.range(f'A'+ str(wss1.cells.last_cell.row)).end('up').row
                last_column_letter=num_to_col_letters(wss1.range('A1').end('right').end('right').last_cell.column)
                working_worksheet=wb.sheets[key]
                wss1.api.Range(f"A2:{last_column_letter}{last_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Select()
                if str(wb2.app.api.Selection.Address).split()[0].replace('$','')=='A1:AK1':
                    wb2.app.api.ActiveSheet.ShowAllData()
                else:
                    wb2.app.api.Selection.Copy(working_worksheet.api.Range("A2"))
                    wss1.activate()  
                    wb2.app.api.ActiveSheet.ShowAllData()
            except:
                wb2.app.api.ActiveSheet.ShowAllData()
                pass

        ws1=wb.sheets[0]
        ws1.activate() 
        refresh_sheet=wb.sheets["All Location"]
        last_row = refresh_sheet.range(f'A'+ str(refresh_sheet.cells.last_cell.row)).end('up').row
        last_column=refresh_sheet.range('A1').end('right').last_cell.column 
        wb.api.ActiveSheet.PivotTables(1).PivotCache().SourceData=f"'All Location'!R1C1:R{last_row}C{last_column}"
        wb.api.ActiveSheet.PivotTables(1).PivotCache().Refresh()      
        wb.save(f"{output_location}\\Unsettled AR_{input_date_short} with reasons.xlsx")   
       
        return f"{job_name} Report for {input_date} generated succesfully"

    except Exception as e:
        raise e
    finally:
        try:
            wb.app.quit()
        except:
            pass


def open_ar_monthly(input_date, output_date):
    try:       
        job_name = 'open_ar_v2'
        input_sheet = drive+r'\REPORT\Open AR New\Raw Files'+f'\\Open AR_{input_date}.xlsx'
        input_sheet2= drive+r'\REPORT\Open AR New\Raw Files'+f'\\Profile.xls'
        if not os.path.exists(input_sheet):
            return(f"{input_sheet} Excel file not present for date {input_date}")
        if not os.path.exists(input_sheet2):
            return(f"{input_sheet2} Excel file not present") 
        retry=0
        while retry < 10:
            try:
                wb = xw.Book(input_sheet) 
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry ==10:
                    raise e 
        input_tab=wb.sheets[f"Open AR_{input_date}"]
        column_list = input_tab.range("A1").expand('right').value
        Customer_Name_column_no = column_list.index('Customer Name')+1
        Customer_Name_column_letter=num_to_col_letters(Customer_Name_column_no)
        Location_column_no = column_list.index('Location')+1
        Location_column_letter=num_to_col_letters(Location_column_no)
        last_row = input_tab.range(f'A'+ str(input_tab.cells.last_cell.row)).end('up').row
        last_column_letter=num_to_col_letters(input_tab.range('A1').end('right').last_cell.column)
        dict1={"MACQUARIE COMMODITIES (USA) INC.":[Customer_Name_column_no,Customer_Name_column_letter],"INTER-COMPANY PURCH/SALES":[Customer_Name_column_no,Customer_Name_column_letter],"WPMEXICO":[Location_column_no,Location_column_letter]}
        for key, value in dict1.items():
            try:
                input_tab.api.Range(f"{value[1]}1").AutoFilter(Field:=f'{value[0]}', Criteria1:=[key], Operator:=7)
                time.sleep(1)
                input_tab.api.Range(f"A2:{last_column_letter}{last_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Select()
                time.sleep(1)
                wb.app.api.Selection.Delete(win32c.DeleteShiftDirection.xlShiftUp)
                time.sleep(1)
                wb.app.api.ActiveSheet.ShowAllData()
            except:
                wb.app.api.ActiveSheet.ShowAllData()
                pass    
        retry=0
        while retry < 10:
            try:
                wb2 = xw.Book(input_sheet2) 
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry ==10:
                    raise e 

        column_list = input_tab.range("A1").expand('right').value
        Customer_Name_column_no = column_list.index('Customer Name')+1
        list1=["Address","City","State","Zip Code"]
        list2=[f"=VLOOKUP(A2,Profile.xls!$A$1:$I$29312,9,0)",f"=VLOOKUP(A2,Profile.xls!$A$1:$I$29312,4,0)",f"=VLOOKUP(A2,Profile.xls!$A$1:$I$29312,5,0)",f"=VLOOKUP(A2,Profile.xls!$A$1:$K$29312,11,0)"] 
        Customer_Name_column_no+=1
        i=0
        last_row = input_tab.range(f'A'+ str(input_tab.cells.last_cell.row)).end('up').row
        for values in list1:
            last_column_letter=num_to_col_letters(Customer_Name_column_no)
            input_tab.api.Range(f"{last_column_letter}1").EntireColumn.Insert()
            input_tab.range(f"{last_column_letter}1").value = values
            input_tab.range(f"{last_column_letter}2").value = list2[i]
            time.sleep(1)
            input_tab.range(f"{last_column_letter}2").copy(input_tab.range(f"{last_column_letter}2:{last_column_letter}{last_row}"))
            i+=1
            Customer_Name_column_no+=1
        #combining 1-10 & 11-30 column
        wb2.close()
        column_list = input_tab.range("A1").expand('right').value
        insert_column_no = column_list.index('11 - 30')+1
        insert_column_no+=1
        insert_column_letter=num_to_col_letters(insert_column_no)
        input_tab.api.Range(f"{insert_column_letter}1").EntireColumn.Insert()  
        input_tab.range(f"{insert_column_letter}1").number_format="@"
        input_tab.api.Range(f"{insert_column_letter}1").Value = '1 - 30'
        prevous_column1=num_to_col_letters(insert_column_no-1)
        prevous_column2=num_to_col_letters(insert_column_no-2)
        input_tab.range(f"{insert_column_letter}2").value = f'={prevous_column1}2+{prevous_column2}2'
        time.sleep(1)
        input_tab.range(f"{insert_column_letter}2").copy(input_tab.range(f"{insert_column_letter}2:{insert_column_letter}{last_row}"))
        time.sleep(1)
        input_tab.range(f"{insert_column_letter}:{insert_column_letter}").copy()
        time.sleep(1)
        input_tab.range(f"{insert_column_letter}:{insert_column_letter}").paste(paste="values_and_number_formats")
        time.sleep(1)
        wb.app.api.CutCopyMode=False
        input_tab.api.Range(f"{prevous_column2}1").EntireColumn.Delete()
        time.sleep(1)
        input_tab.api.Range(f"{prevous_column2}1").EntireColumn.Delete()
        #creating date and as of date columns
        column_list = input_tab.range("A1").expand('right').value
        Due_Date_column_no = column_list.index('Due Date')+1
        Due_Date_column_no+=1
        insert_column_letter=num_to_col_letters(Due_Date_column_no)
        input_tab.api.Range(f"{insert_column_letter}1").EntireColumn.Insert()
        As_of_date_CN=Due_Date_column_no+1
        As_of_date_letter=num_to_col_letters(As_of_date_CN)
        input_tab.api.Range(f"{As_of_date_letter}1").EntireColumn.Insert()
        prevous_column1=num_to_col_letters(Due_Date_column_no-1)
        input_tab.range(f"{prevous_column1}:{prevous_column1}").copy(input_tab.range(f"{insert_column_letter}:{insert_column_letter}"))
        time.sleep(1)
        Due_Date_letter=num_to_col_letters(Due_Date_column_no-1)
        input_tab.range(f"{Due_Date_letter}{1}:{Due_Date_letter}{last_row}").number_format='dd-mm-yyyy'
        next_letter=num_to_col_letters(Due_Date_column_no)
        input_tab.range(f"{next_letter}{1}:{next_letter}{last_row}").number_format='dd-mm-yyyy'
        input_tab.range(f"{insert_column_letter}1").value = 'Date'
        x=datetime.strptime(input_date,"%m.%d.%Y")
        # x=datetime.strftime(x,"%d-%m-%Y")
        input_tab.range(f"{As_of_date_letter}{1}").number_format='dd-mm-yyyy'
        input_tab.range(f"{As_of_date_letter}1").options(dates=datetime.date).value = x      
        input_tab.range(f"{As_of_date_letter}{1}").number_format='dd-mm-yyyy'
        # messagebox.showinfo("Info", 'Changing formating')
        
        for i in range(2,int(f'{last_row+1}')):
            
            if input_tab.range(f"N{i}").value==None:
                print(i)
                input_tab.range(f"N{i}").value=input_tab.range(f"K{i}").value
        input_tab.range(f"{As_of_date_letter}2").value = f'=+${As_of_date_letter}$1-{insert_column_letter}2'  
        input_tab.range(f"{As_of_date_letter}2").number_format="0.00"
        input_tab.range(f"{As_of_date_letter}2").copy(input_tab.range(f"{As_of_date_letter}2:{As_of_date_letter}{last_row}"))
        column_list = input_tab.range("A1").expand('right').value
        insert_column_no = column_list.index('61 - 9999')+1
        prevous_column1=num_to_col_letters(insert_column_no)
        insert_column_no+=1
        insert_column_letter=num_to_col_letters(insert_column_no)
        input_tab.api.Range(f"{insert_column_letter}1").EntireColumn.Insert() 
        input_tab.range(f"{insert_column_letter}1").number_format="General"
        input_tab.api.Range(f"{insert_column_letter}1").Value = '90+'
        input_tab.api.Range(f"{prevous_column1}1").Value = '61 - 90'
        input_tab.api.Range(f"{prevous_column1}2:{prevous_column1}{last_row}").Copy(input_tab.api.Range(f"{insert_column_letter}2:{insert_column_letter}{last_row}"))
        # messagebox.showinfo("Info", 'Checking Error')
        input_tab.api.Range(f"{As_of_date_letter}1").AutoFilter(Field:=f'{As_of_date_CN}', Criteria1:=[">90"], Operator:=1) 
        time.sleep(1) 
        input_tab.api.Range(f"{prevous_column1}2:{prevous_column1}{last_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Select()
        time.sleep(1)
        wb.app.api.Selection.Value=0
        time.sleep(1)
        wb.app.api.ActiveSheet.ShowAllData()
        input_tab.api.Range(f"{As_of_date_letter}1").AutoFilter(Field:=f'{As_of_date_CN}', Criteria1:=["<=90"], Operator:=1)  
        input_tab.api.Range(f"{insert_column_letter}2:{insert_column_letter}{last_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).Select()
        wb.app.api.Selection.Value=0
        time.sleep(1)
        wb.app.api.ActiveSheet.ShowAllData()
        print("ddd")        

        wb.sheets.add("Pivot",after=input_tab)
        ###logger.info("Clearing contents for new sheet")
        wb.sheets["Pivot"].clear_contents()
        ws2=wb.sheets["Pivot"]
        ###logger.info("Declaring Variables for columns and rows")
        last_column = input_tab.range('A1').end('right').last_cell.column
        last_column_letter=num_to_col_letters(input_tab.range('A1').end('right').last_cell.column)
        ###logger.info("Creating Pivot Table")
        PivotCache=wb.api.PivotCaches().Create(SourceType=win32c.PivotTableSourceType.xlDatabase, SourceData=f"\'Open AR_{input_date}\'!R1C1:R{last_row}C{last_column}", Version=win32c.PivotTableVersionList.xlPivotTableVersion14)
        PivotTable = PivotCache.CreatePivotTable(TableDestination=f"'Pivot'!R1C1", TableName="PivotTable1", DefaultVersion=win32c.PivotTableVersionList.xlPivotTableVersion14)        ###logger.info("Adding particular Row in Pivot Table")
        PivotTable.PivotFields('Customer Id').Orientation = win32c.PivotFieldOrientation.xlRowField
        PivotTable.PivotFields('Customer Id').Position = 1
        PivotTable.PivotFields('Customer Id').Subtotals=(False, False, False, False, False, False, False, False, False, False, False, False)
        PivotTable.PivotFields('Customer Name').Orientation = win32c.PivotFieldOrientation.xlRowField
        PivotTable.PivotFields('Customer Name').Subtotals=(False, False, False, False, False, False, False, False, False, False, False, False)
        PivotTable.PivotFields('Address').Orientation = win32c.PivotFieldOrientation.xlRowField
        PivotTable.PivotFields('Address').Subtotals=(False, False, False, False, False, False, False, False, False, False, False, False)
        PivotTable.PivotFields('City').Orientation = win32c.PivotFieldOrientation.xlRowField
        PivotTable.PivotFields('City').Subtotals=(False, False, False, False, False, False, False, False, False, False, False, False)
        PivotTable.PivotFields('State').Orientation = win32c.PivotFieldOrientation.xlRowField
        PivotTable.PivotFields('State').Subtotals=(False, False, False, False, False, False, False, False, False, False, False, False)
        PivotTable.PivotFields('Zip Code').Orientation = win32c.PivotFieldOrientation.xlRowField
        PivotTable.PivotFields('Zip Code').Subtotals=(False, False, False, False, False, False, False, False, False, False, False, False)
        ###logger.info("Adding particular Data Field in Pivot Table")
        PivotTable.PivotFields('Balance').Orientation = win32c.PivotFieldOrientation.xlDataField
        PivotTable.PivotFields('Current').Orientation = win32c.PivotFieldOrientation.xlDataField
        PivotTable.PivotFields('1 - 30').Orientation = win32c.PivotFieldOrientation.xlDataField
        # PivotTable.PivotFields('Sum of  1 - 10').NumberFormat= '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'
        PivotTable.PivotFields('31 - 60').Orientation = win32c.PivotFieldOrientation.xlDataField
        # PivotTable.PivotFields('Sum of  31 - 60').NumberFormat= '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'
        PivotTable.PivotFields('61 - 90').Orientation = win32c.PivotFieldOrientation.xlDataField
        # PivotTable.PivotFields('Sum of  61 - 9999').NumberFormat= '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'
        PivotTable.PivotFields('90+').Orientation = win32c.PivotFieldOrientation.xlDataField
        PivotTable.TableStyle2 = ""
        ###logger.info("Changing Table Layout in Pivot Table")
        PivotTable.RowAxisLayout(1)
        wb.api.ActiveSheet.PivotTables("PivotTable1").InGridDropZones = True
        wb.api.ActiveSheet.PivotTables("PivotTable1").DataPivotField.Caption = "Values"
        time.sleep(1)
        last_row_2 = ws2.range(f'A'+ str(ws2.cells.last_cell.row)).end('up').row
        wb.sheets.add("Summary",after=ws2)
        ###logger.info("Clearing contents for new sheet")
        wb.sheets["Summary"].clear_contents()
        ws3=wb.sheets["Summary"]
        last_column2 = ws2.range('A2').end('right').last_cell.column
        last_column_letter2=num_to_col_letters(last_column2)
        ws2.api.Range(f"A2:{last_column_letter2}{last_row_2}").Copy(ws3.api.Range(f"A1"))
        last_row_3 = ws3.range(f'A'+ str(ws2.cells.last_cell.row)).end('up').row
        last_column3 = ws3.range('A1').end('right').last_cell.column
        last_column_letter3=num_to_col_letters(last_column3)
        insert_all_borders(cellrange=f"A1:{last_column_letter3}{last_row_3}",working_sheet=ws3,working_workbook=wb)
        ws3.autofit()
        ws3.api.Range("1:1").Font.Bold = True
        ws2.activate()
        wb.app.api.ActiveSheet.PivotTables("PivotTable1").TableStyle2 = "PivotStyleLight16"
        ws3.api.Range(f"{last_row_3}:{last_row_3}").Font.Bold = True
        output_location = drive+r'\REPORT\Open AR New\Output Files'   
        wb.save(f"{output_location}\\Open AR_"+input_date+' updated.xlsx')
        wb.app.quit()
        return f"{job_name} Report for {input_date} generated succesfully"

    except Exception as e:
        raise e
    finally:
        try:
            wb.app.quit()
        except:
            pass




def weekly_estimate(input_date, output_date):
    try:
        job_name = "Weekly Estimate"

        input_datetime = datetime.strptime(input_date, "%m.%d.%Y")
        input_datetime2 = datetime.strptime(output_date, "%m.%d.%Y")
        input_date2 = datetime.strftime(input_datetime, "%m%d%Y")
        input_month_date = datetime.strftime(input_datetime, "%d%m")
        # inp_sht_date = datetime.strftime(input_datetime, "%d-%m-%Y")
        cur_year = datetime.strftime(input_datetime, "%Y")
        monthYear = datetime.strftime(input_datetime, "%b-%y")
        prev_datetime = input_datetime.replace(day=1)-timedelta(days=1)
        prev_date = datetime.strftime(prev_datetime,"%m.%d.%Y").upper()
        input_date3 = datetime.strftime(input_datetime2, "%m/%d/%Y")#05/22/2023
        # input_date_save = datetime.strftime(input_datetime, "%m-%d-%Y")#05/22/2023
        prior_date = input_datetime - timedelta(days=7)
        prior_date_str = datetime.strftime(prior_date, "%m-%d-%Y")
        curr_month_last_date = (input_datetime.replace(day=28) + timedelta(days=4)).replace(day=1) + timedelta(days=-1)
        output_loc = drive+f'\\REPORT\\Weekly_Estimate\\Output Files\\Weekly Estimate_{input_date}.xlsx'
        mac_output_loc = drive+f'\\REPORT\\Weekly_Estimate\\Output Files\\Open Macquire Repurchase Tracking Report_{input_date}.xlsx'
        
        ################Picking Latest Output File Location###############################################
        prior_wb_loc = drive+f'\\REPORT\\Weekly_Estimate\\Output Files\\Weekly Estimate_*.xlsx'
        prior_wb_list = glob.glob(prior_wb_loc)
        prior_wb_list.sort(key=os.path.getmtime)
        prior_wb_loc = prior_wb_list[-1]
        if not os.path.exists(prior_wb_loc):
            return(f"{prior_wb_loc} Excel file not present")
        input_xl = drive+r'\REPORT\Weekly_Estimate'+f'\\Weekly_Estimate_Template.xlsx'
        if not os.path.exists(input_xl):
            return(f"{input_xl} Excel Template file not present")
        inp_gl_query = drive +f'\\REPORT\\Weekly_Estimate\\Raw Files\\GL QUERY_{input_date}.xlsx'
        if not os.path.exists(inp_gl_query):
            return(f"{inp_gl_query} Excel file not present.")
        inp_open_futures = drive + f'\\REPORT\\Macquarie P & L\\FromEmail\\{input_date2}\\{input_month_date}F\\OC{input_month_date}F.csv'
        if not os.path.exists(inp_open_futures):
            return(f"{inp_open_futures} CSV file not present for date {input_date}")
        # inp_payables = drive +f'\\REPORT\\Unsettled Payables\\Output files\\Unsettled Payables _'+input_date+'.xlsx'
        inp_payables = drive+r'\REPORT\Unsettled Payables\Raw Files\Unsettled Payables _'+input_date+".xlsx"        
        if not os.path.exists(input_xl):
            return(f"Excel file not present for date {input_date}")
        if not os.path.exists(inp_payables):
            return(f"{inp_payables} Excel file not present for date {input_date}")
        # inp_receivables = drive +f'\\REPORT\\Unsettled Receivables\\Output files\\Unsettled Receivables _'+input_date+'.xlsx'
        inp_receivables = drive+r'\REPORT\Unsettled Receivables\Raw Files\Unsettled Receivables _'+input_date+".xlsx"
        if not os.path.exists(inp_receivables):
            return(f"{inp_receivables} Excel file not present for date {input_date}")
        # inp_ctm = drive +f'\\REPORT\\CTM Combined report\\Output files\\CTM Combined _{input_date}.xlsx'
        inp_ctm = drive+r'\REPORT\CTM Combined report\Raw Files\CTM Combined _'+input_date+'.xlsx' 
        if not os.path.exists(inp_ctm):
            return(f"{inp_ctm} Excel file not present for date {input_date}")
        inp_inventory = drive +f'\\REPORT\\Inv_MTM_Excel_Report_Summ\\Output Files\\Inventory MTM Excel Report_'+input_date+".xlsx"
        if not os.path.exists(inp_inventory):
            return(f"{inp_inventory} Excel file not present for date {input_date}")
        if curr_month_last_date == input_datetime:
            prev_date = input_date
        inp_strg_accr = drive +f'\\REPORT\\Storage Month End Report\\Output Files\\STORAGE ACCRUAL JE_'+prev_date+".xlsx"
        if not os.path.exists(inp_strg_accr):
            return(f"{inp_strg_accr} Excel file not present for date {prev_date}")
        inp_freight_accr = drive +f'\\REPORT\Weekly_Estimate\Raw Files\\Freight Accrual_'+input_date+".xlsx"
        if not os.path.exists(inp_freight_accr):
            return(f"{inp_freight_accr} Excel file not present for date {input_date}")
        # macq_file_path = drive+f'\\REPORT\Weekly_Estimate\\Raw Files\\Open Macquire Repurchase Tracking Report- {cur_year}.xlsx'
        macq_file_path = drive+f'\\REPORT\Weekly_Estimate\\Raw Files\\Open Macquire Repurchase Tracking Report_{input_date}.xlsx'
        if not os.path.exists(macq_file_path):
            return(f"{macq_file_path} Excel file not present for year {cur_year}")
        # inp_basis_path = drive+f'\\REPORT\Weekly_Estimate\\Raw Files\\{cur_year}  Weekly and Monthly Values.xlsx'
        inp_basis_path = drive+f'\\REPORT\Weekly_Estimate\\Raw Files\\Weekly and Monthly Values_{input_date}.xlsx'
        if not os.path.exists(inp_basis_path):
            return(f"{inp_basis_path} Excel file not present for year {cur_year}")
        mapping_loc = drive+r'\REPORT\Weekly_Estimate\Mapping.xlsx'
        if not os.path.exists(mapping_loc):
            return(f"{mapping_loc} Excel file not present for date {input_date}")
        
        retry=0
        while retry < 10:
            try:
                wb=xw.Book(input_xl, update_links=False)
                break
            except Exception as e:
                time.sleep(2)
                retry+=1
                if retry ==9:
                    raise e
    
        gl_query_sht = wb.sheets("GL QUERY")
        estimate_sht = wb.sheets("West Plains Estimate")
        open_futures_sht = wb.sheets("OPEN FUTURES")
        payab_sht = wb.sheets("UNSETTLED PAYABLES")
        receivables_sht = wb.sheets("UNSETTLED RECEIVABLES")
        ctm_sht = wb.sheets("CTM")
        inventory_sht = wb.sheets("INVENTORY")
        repos_sht = wb.sheets("REPOS")
        freight_sht = wb.sheets("FREIGHT ACCRUAL")
        storage_sht = wb.sheets("STORAGE ACCRUAL")

        ###############Updating dat in first sheet########################
        estimate_sht.range(f"B3").value = input_datetime

        ##################GL QUERY PART#####################################
        retry=0
        while retry < 10:
            try:
                gl_wb=xw.Book(inp_gl_query, update_links=False)
                break
            except Exception as e:
                time.sleep(2)
                retry+=1
                if retry ==9:
                    raise e
        inp_gl_sht = gl_wb.sheets("GL QUERY")
        inp_gl_last_row = inp_gl_sht.range(f"A{inp_gl_sht.cells.last_cell.row}").end('up').row
        inp_gl_sht.range(f"A2:Q{inp_gl_last_row}").copy(gl_query_sht.range("A2"))
        #Now calculating last row based on pasted data
        last_row_gl_query = gl_query_sht.range(f'A'+ str(gl_query_sht.cells.last_cell.row)).end('up').row
        wb.activate()
        gl_query_sht.activate()
        gl_query_sht.range(f'R2:T{last_row_gl_query}').select()
        wb.app.selection.api.FillDown()

        gl_wb.close()

        ##################getting data in dataframes#############
        open_futures_df = pd.read_csv(inp_open_futures)
        #Replacing data in Open Futures Sheet
        last_row_open_futures = open_futures_sht.range(f'AX'+ str(open_futures_sht.cells.last_cell.row)).end('up').row
        open_futures_sht.range(f'A2:AW{last_row_open_futures}').clear_contents()
        open_futures_sht.range(f'A1').options(pd.DataFrame, header=True, index=False).value = open_futures_df
        #Now calculating last row based on pasted data
        last_row_open_futures = open_futures_sht.range(f'A'+ str(open_futures_sht.cells.last_cell.row)).end('up').row
        wb.activate()
        open_futures_sht.activate()
        open_futures_sht.range(f'AX2:AY{last_row_open_futures}').select()
        wb.app.selection.api.FillDown()
        ############Unsettled payable part##################################
        
        
        
        retry=0
        while retry<10:
            try:
                
                payables_wb = xw.Book(inp_payables, update_links=True)
                break
            except Exception as e:
                time.sleep(2)
                retry+=1
                if retry==9:
                    raise e
        
        #######logger.info("Sheet Opened")
        # time.sleep(10)
        while True:
            try:
                # inp_sht = wb.sheets[0]
                inp_payables_sht = payables_wb.sheets[f"Unsettled Payables _{input_date}"]
                break
            except:
                time.sleep(10)
        

        # inp_sht.range('AB2').formula = '=O2+Q2'
        

        column_list = inp_payables_sht.range("A1").expand('right').value
        vendor_column = num_to_col_letters(column_list.index('Customer/Vendor Name')+1)
        vendor_column_num = column_list.index('Customer/Vendor Name')+1
        locationId_column = num_to_col_letters(column_list.index('Location Id')+1)
        locationId_column_num = column_list.index('Location Id')+1
        payab_last_column = num_to_col_letters(len(column_list))
        if payab_last_column != "AH":
            return "Input Payables column count change found, last column is not AH. Please fix and rerrun the job"

        inp_payables_sht.api.AutoFilterMode=False
        #Removing  MACQUARIE COMMODITIES (USA) INC. and all INTER-COMPANY PURCH/SALES vendor
        inp_payables_sht.api.Range(f"{vendor_column}1").AutoFilter(Feild:=vendor_column_num,Criteria1:="<>MACQUARIE COMMODITIES (USA) INC", Operator:=1, Criteria2:="<>INTER-COMPANY PURCH/SALES") #Removing macquarie and intercompany
        #Removing WPMEXICO Location ID
        inp_payables_sht.api.Range(f"{locationId_column}1").AutoFilter(Feild:=locationId_column_num,Criteria1:="<>WPMEXICO", Operator:=7) #Removing WPMEXICO
        last_row_payables = payab_sht.range(f'AI'+ str(payab_sht.cells.last_cell.row)).end('up').row
        payab_sht.range(f'A2:AH{last_row_payables}').clear_contents()
        inp_payables_sht.api.AutoFilter.Range.Copy()
        payab_sht.range("A1").api._PasteSpecial(Paste=win32c.PasteType.xlPasteValues)
        last_row_payables = payab_sht.range(f'A'+ str(payab_sht.cells.last_cell.row)).end('up').row
        wb.activate()
        payab_sht.activate()
        payab_sht.range(f'AI2:AK{last_row_payables}').select()
        wb.app.selection.api.FillDown()
        
        payables_wb.close()
        #################################Old code############################################
        # payables_df = pd.read_excel(inp_payables, sheet_name="Excl Macq & IC")
        # #Replacing data in Open Futures Sheet
        # last_row_payables = payab_sht.range(f'AI'+ str(payab_sht.cells.last_cell.row)).end('up').row
        # payab_sht.range(f'A2:AH{last_row_payables}').clear_contents()
        # payab_sht.range(f'A2').options(pd.DataFrame, header=False, index=False).value = payables_df
        # #Now calculating last row based on pasted data
        # last_row_payables = payab_sht.range(f'A'+ str(payab_sht.cells.last_cell.row)).end('up').row
        # wb.activate()
        # payab_sht.activate()
        # payab_sht.range(f'AI2:AK{last_row_payables}').select()
        # wb.app.selection.api.FillDown()
        # ############Unsettled Receivables part##################################
        # receivables_df = pd.read_excel(inp_receivables, sheet_name="Excl Macq & IC")
        # #Replacing data in Open Futures Sheet
        # last_row_receivables = receivables_sht.range(f'AF'+ str(receivables_sht.cells.last_cell.row)).end('up').row
        # receivables_sht.range(f'A2:AE{last_row_receivables}').clear_contents()
        # receivables_sht.range(f'A2').options(pd.DataFrame, header=False, index=False).value = receivables_df
        # #Now calculating last row based on pasted data
        # last_row_receivables = receivables_sht.range(f'A'+ str(receivables_sht.cells.last_cell.row)).end('up').row
        # wb.activate()
        # receivables_sht.activate()
        # receivables_sht.range(f'AF2:AH{last_row_receivables}').select()
        # wb.app.selection.api.FillDown()
        ##################################Unsettled Receivables Part###########################################
        retry=0
        while retry<10:
            try:
                receivables_wb = xw.Book(inp_receivables, update_links=False)
                break
            except Exception as e:
                time.sleep(2)
                retry+=1
                if retry==9:
                    raise e

        inp_receivables_sht = receivables_wb.sheets[0] #wb.sheets[0].name in 'Unsettled Receivables _'+input_date

        column_list = inp_receivables_sht.range("A1").expand('right').value
        vendor_column = num_to_col_letters(column_list.index('Customer/Vendor Name')+1)
        vendor_column_num = column_list.index('Customer/Vendor Name')+1
        locationId_column = num_to_col_letters(column_list.index('Location Id')+1)
        locationId_column_num = column_list.index('Location Id')+1
        receivable_last_column = num_to_col_letters(len(column_list))
        if receivable_last_column != "AE":
            return "Input Receivables column count change found, last column is not AE. Please fix and rerrun the job"
               
        inp_receivables_sht.api.AutoFilterMode=False
        ######logger.info("Removing  MACQUARIE COMMODITIES (USA) INC. and all INTER-COMPANY PURCH/SALES vendor")
        inp_receivables_sht.api.Range(f"{vendor_column}1").AutoFilter(Feild:=vendor_column_num,Criteria1:="<>MACQUARIE COMMODITIES (USA) INC", Operator:=1, Criteria2:="<>INTER-COMPANY PURCH/SALES") #Removing macquarie and intercompany
        ######logger.info("Removing WPMEXICO Location ID")
        inp_receivables_sht.api.Range(f"{locationId_column}1").AutoFilter(Feild:=locationId_column_num,Criteria1:="<>WPMEXICO", Operator:=7) #Removing WPMEXICO
        ######logger.info("Creating Excl IC & Macq and pasting data")
        inp_receivables_sht.api.AutoFilter.Range.Copy()
        receivables_sht.range("A1").api._PasteSpecial(Paste=win32c.PasteType.xlPasteValues)
        last_row_receivables = receivables_sht.range(f'A'+ str(receivables_sht.cells.last_cell.row)).end('up').row
        wb.activate()
        receivables_sht.activate()
        receivables_sht.range(f'AF2:AH{last_row_receivables}').select()
        wb.app.selection.api.FillDown()
        receivables_wb.close()
        ###################CTM Part#############################
        retry=0
        while retry<10:
            try: 
                ctm_wb = xw.Book(inp_ctm, update_links=False)
                break
            except Exception as e:
                time.sleep(2)
                retry+=1
                if retry==9:
                    raise e
        inp_ctm_sht = ctm_wb.sheets[0]
        column_list = inp_ctm_sht.range("A1").expand('right').value
        Customer_no_column=column_list.index('Customer')+1
        Customer_letter_column = num_to_col_letters(column_list.index('Customer')+1)
        Location_no_column=column_list.index('Location Id')+1
        Location_letter_column = num_to_col_letters(column_list.index('Location Id')+1)
        ctm_last_column = num_to_col_letters(len(column_list))
        if ctm_last_column != "AO":
            return "Input CTM column count change found, last column is not AO. Please fix and rerrun the job"
        inp_ctm_sht.api.Range(f"{Customer_letter_column}1").AutoFilter(Field:=f'{Customer_no_column}', Criteria1:=["<>MACQUARIE COMMODITIES (USA) INC."], Operator:=1,Criteria2=["<>INTER-COMPANY PURCH/SALES"])
        inp_ctm_sht.api.Range(f"{Location_letter_column}1").AutoFilter(Field:=f'{Location_no_column}', Criteria1:=["<>WPMEXICO"], Operator:=1)
        inp_ctm_last_row = inp_ctm_sht.range(f'A'+ str(ctm_sht.cells.last_cell.row)).end('up').row
        inp_ctm_sht.range(f"A2:AN{inp_ctm_last_row}").api.SpecialCells(win32c.CellType.xlCellTypeVisible).Copy()
        # inp_ctm_sht.api.AutoFilter.Range("A:AN").Copy()
        ctm_sht.range("A2").api._PasteSpecial(Paste=win32c.PasteType.xlPasteValues)
        last_row_ctm = ctm_sht.range(f'A'+ str(ctm_sht.cells.last_cell.row)).end('up').row
        wb.activate()
        ctm_sht.activate()
        ctm_sht.range(f'AO2:AP{last_row_ctm}').select()
        wb.app.selection.api.FillDown()
        ctm_wb.close()
        ########################Inventory Sheet####################################
        inventory_df = pd.read_excel(inp_inventory, sheet_name="INPUT DATA",header=2)
        inventory_df.columns = inventory_df.columns.str.rstrip()
        inventory_df = inventory_df.dropna(subset=["UOM MKT VALUE"])
        # last_row_inventory = inventory_sht.range(f'T'+ str(inventory_sht.cells.last_cell.row)).end('up').row
        # inventory_sht.range(f'A2:S{last_row_inventory}').clear_contents()
        inventory_sht.range(f'A4').options(pd.DataFrame, header=False, index=False).value = inventory_df
        #Now calculating last row based on pasted data
        last_row_inventory = inventory_sht.range(f'A'+ str(inventory_sht.cells.last_cell.row)).end('up').row
        wb.activate()
        inventory_sht.activate()
        inventory_sht.range(f'T4:U{last_row_inventory}').select()
        wb.app.selection.api.FillDown()
        wb.api.ActiveSheet.PivotTables(1).PivotCache().SourceData = f"'INVENTORY'!R3C1:R{last_row_inventory}C21"
        wb.api.ActiveSheet.PivotTables(1).PivotCache().Refresh()
        ######################STORAGE ACCRUAL JE############################################
        strg_accr_df = pd.read_excel(inp_strg_accr, sheet_name="INPUT DATA",header=2)
        strg_accr_df2 = pd.read_excel(inp_strg_accr, sheet_name="INPUT DATA",header=None)
        storage_col_list = storage_sht.range("A3").expand("right").value
        storage_sht.range(f'A4').options(pd.DataFrame, header=False, index=False).value = strg_accr_df
        #Entering Date
        storage_sht.range(f'A1').value = strg_accr_df2.iloc[0,0]
        last_row_strg_accr = storage_sht.range(f'A'+ str(storage_sht.cells.last_cell.row)).end('up').row
        wb.activate()
        storage_sht.activate()
        storage_sht.range(f'J4:K{last_row_strg_accr}').select()
        wb.app.selection.api.FillDown()
        ######################FREIGHT ACCRAUL TAB####################################
        retry=0
        while retry<10:
            try: 
                freight_wb = xw.Book(inp_freight_accr, update_links=False)
                break
            except Exception as e:
                time.sleep(2)
                retry+=1
                if retry==9:
                    raise e
        inp_freight_sht = freight_wb.sheets[0]
        freight_accr_df = inp_freight_sht.range("A1").expand("table").options(pd.DataFrame,header=1,
                                index=False).value
        # freight_accr_df = pd.read_excel(inp_freight_accr,sheet_name=0,header=0)
        freight_accr_df = freight_accr_df[freight_accr_df['Accrual Entered']!="C"].reset_index(drop=True)
        freight_accr_df = freight_accr_df[freight_accr_df['Location Name']!="WEST PLAINS MEXICO"].reset_index(drop=True)
        freight_accr_df['Comments'] = np.where(
            (freight_accr_df['Contract Location'].isna()), "Unapplied Tickets", "Applied to Contract")
        freight_sht.range(f'A2').options(pd.DataFrame, header=False, index=False).value = freight_accr_df
        last_row_freight_accr = freight_sht.range(f'A'+ str(freight_sht.cells.last_cell.row)).end('up').row
        wb.activate()
        freight_sht.activate()
        freight_sht.range(f'AA2:AB{last_row_freight_accr}').select()
        wb.app.selection.api.FillDown()
        freight_wb.close()
        """Repos Tab Part"""
        retry=0
        while retry<10:
            try:
                macq_wb = xw.Book(macq_file_path, update_links=False)
                break
            except Exception as e:
                time.sleep(2)
                retry+=1
                if retry==9:
                    raise e
        retry=0
        while retry<10:
            try:
                basis_wb = xw.Book(inp_basis_path, update_links=False)
                break
            except Exception as e:
                time.sleep(2)
                retry+=1
                if retry==9:
                    raise e
                
        basis_sheet = basis_wb.sheets["Active Pricings"]
        basis_wb.activate()
        basis_sheet.activate()
        basis_sheet.api.Cells.Find(What:="CORN", After:=basis_sheet.api.Application.ActiveCell,LookIn:=win32c.FindLookIn.xlFormulas,
                                   LookAt:=win32c.LookAt.xlPart, SearchOrder:=win32c.SearchOrder.xlByRows, SearchDirection:=win32c.SearchDirection.xlNext).Activate()
        corn_cell_value = basis_sheet.api.Application.ActiveCell.Address.replace("$","")
        corn_df = basis_sheet.range(f"{corn_cell_value}").expand('table').options(pd.DataFrame,header=1,index=False).value
        corn_df = corn_df[["Location", monthYear]]
        corn_dict = corn_df.set_index("Location")[monthYear].to_dict()
        ###HRW Location dict logic
        basis_sheet.api.Cells.Find(What:="HRW", After:=basis_sheet.api.Application.ActiveCell,LookIn:=win32c.FindLookIn.xlFormulas,
                                   LookAt:=win32c.LookAt.xlPart, SearchOrder:=win32c.SearchOrder.xlByRows, SearchDirection:=win32c.SearchDirection.xlNext).Activate()
        hrw_cell_value = basis_sheet.api.Application.ActiveCell.Address.replace("$","")
        hrw_df = basis_sheet.range(f"{hrw_cell_value}").expand('table').options(pd.DataFrame,header=1,index=False).value
        hrw_df = hrw_df[["Location", monthYear]]
        hrw_dict = hrw_df.set_index("Location")[monthYear].to_dict()
        hrw_dict["Omaha"] = hrw_dict["Omaha 12"] 


        basis_wb.close()

        ####Updating basis price in Macquarie Sheet
        
        macq_m_end_prices_sht = macq_wb.sheets["MONTH END PRICES"]
        
        macq_wb.activate()
        macq_m_end_prices_sht.activate()
        corn_last_row = macq_m_end_prices_sht.range(f'F'+ str(macq_m_end_prices_sht.cells.last_cell.row)).end('up').row
        hrw_last_row = macq_m_end_prices_sht.range(f'H'+ str(macq_m_end_prices_sht.cells.last_cell.row)).end('up').row
        corn_st_row = 4
        hrw_st_row = 4
        for row in range(corn_st_row,corn_last_row+1):
            try:
                macq_m_end_prices_sht.range(f"G{row}").value = corn_dict[macq_m_end_prices_sht.range(f"F{row}").value]
            except:
                try:
                    macq_m_end_prices_sht.range(f"G{row}").value = 0
                except Exception as e:
                    raise e
        for row in range(hrw_st_row,hrw_last_row+1):
            try:
                macq_m_end_prices_sht.range(f"I{row}").value = hrw_dict[macq_m_end_prices_sht.range(f"H{row}").value]
            except:
                try:
                    macq_m_end_prices_sht.range(f"I{row}").value = 0
                except Exception as e:
                    raise e

        ########Future Price Part#################
        hrw_map_df = pd.read_excel(mapping_loc,sheet_name = "Future Month HRW",header=3)
        hrw_mapping = hrw_map_df.set_index("Option Month")["Month"].to_dict()
        corn_map_df = pd.read_excel(mapping_loc,sheet_name = "Future Month YC",header=3)
        corn_mapping = corn_map_df.set_index("Option Month")["Month"].to_dict()

        corn_site_code = 300
        hrw_site_code = 348
        headers = {
                    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36",
                    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8",
                    "Accept-Encoding": "gzip, deflate, br",
                    "Accept-Language": "en-US,en;q=0.9",
                    }
        for site_code in [corn_site_code,hrw_site_code]:
            cme_date_url = f'https://www.cmegroup.com/CmeWS/mvc/Settlements/Futures/Settlements/{site_code}/FUT?strategy=DEFAULT&tradeDate={input_date3}'
            with requests.Session() as session:
                session.headers.update(headers)
                response = session.get(cme_date_url)
                if response.json()['empty']:
                    commodity = "Corn"
                    if site_code == hrw_site_code:
                        commodity = "HRW"
                    if messagebox.askyesno(messagebox.INFO,f"Data not found for {commodity} {input_date} on cme website, should I pick Latest Data?"):
                        cme_url = f'https://www.cmegroup.com/CmeWS/mvc/Settlements/Futures/TradeDate/{corn_site_code}'
                        date = session.get(cme_url).json()[1][0]
                        cme_date_url = f'https://www.cmegroup.com/CmeWS/mvc/Settlements/Futures/Settlements/{site_code}/FUT?strategy=DEFAULT&tradeDate={date}'
                        response= session.get(cme_date_url)
                #Again checking if already data present already or after picking latest date then only update else considered for manual entry
                if not response.json()['empty']:
                    data_dict = response.json()
                    df = pd.DataFrame(data_dict['settlements'])
                    df = df[["month", "settle"]]
                    df_dict = df.set_index("month")["settle"].to_dict()
                    form_dict = {"'6":"75", "'4":"50", "'2":"25", "'0":"00"}
                    ##Updating Future prices
                    if site_code == hrw_site_code:
                        # fut_start = macq_m_end_prices_sht.range(f'C3').end('down').end('down').row
                        fut_start = macq_m_end_prices_sht.range(f'B3').end('down').row
                        fut_end = macq_m_end_prices_sht.range(f'C{fut_start}').end('down').row
                    else:
                        fut_start = 3
                        fut_end = macq_m_end_prices_sht.range(f'C3').end('down').row
                    for row in range(fut_start,fut_end+1):
                        manual_entry=False
                        if macq_m_end_prices_sht.range(f"C{row}").value is not None:
                            if macq_m_end_prices_sht.range(f"C{row}").value.startswith("C"):
                                fut_value = df_dict[corn_mapping[macq_m_end_prices_sht.range(f"C{row}").value]]            
                            else:
                                fut_value = df_dict[hrw_mapping[macq_m_end_prices_sht.range(f"C{row}").value]]
                            if "'" in fut_value:
                                fut_value = int(fut_value.split("'")[0] + form_dict[("'" + fut_value.split("'")[1])])/10000
                            else:
                                fut_value = None
                                commodity = "Corn"
                                if site_code == hrw_site_code:
                                    commodity = "HRW"
                                messagebox.showinfo("Settle Price Not found",f"Settle price for {commodity} not found, click ok to proceed further")
                                manual_entry=True
                            if not manual_entry:
                                macq_m_end_prices_sht.range(f"D{row}").value = fut_value
                            else:
                                pass
            ####Updating date and refreshing Pivot##########################
            macq_m_end_je_sht = macq_wb.sheets["MONTH END JE UPDATED"]
            input_date4 = input_datetime
            input_date4 = datetime.strftime(input_date4, "X%m-X%d-%y").replace('X0','').replace('X','')
            try:
                macq_date_sht = macq_wb.sheets[input_date4]
            except:
                # messagebox.showinfo(messagebox.INFO,f"Sheet not found for {input_date4} so picking up latest sheet i.e {macq_wb.sheets[4].name}")
                return f"Sheet not found for {input_date4}"
                # macq_date_sht = macq_wb.sheets[4]
        #########Updating A1 A2 date in date sheet#############
        macq_wb.activate()
        macq_date_sht.activate()
        macq_date_sht.range("A1").value = input_datetime
        macq_date_sht.range("A2").value = input_datetime
        macq_wb.activate()
        macq_m_end_je_sht.activate()
        macq_m_end_je_sht.range("B2").value = input_datetime
        #########Pivot Refreshing Tables#######################
        date_sht_last_row = macq_date_sht.range(f'C'+ str(macq_date_sht.cells.last_cell.row)).end('up').row
        row_list = macq_date_sht.range(f"A1:A{date_sht_last_row}").value
        open_repo_row = row_list.index("OPEN REPOS")
        cur_mon_buyback_row = row_list.index("CURRENT MONTH BUYBACKS")
        open_repo_last_row =macq_date_sht.range(f"C{cur_mon_buyback_row}").end("up").row

        macq_wb.activate()
        macq_m_end_je_sht.activate()
        #1st Pivot
        macq_wb.api.ActiveSheet.PivotTables("PivotTable1").ChangePivotCache(macq_wb.api.PivotCaches().Create(SourceType=win32c.PivotTableSourceType.xlDatabase,
                                                 SourceData= f"'{macq_date_sht.name}'!R{open_repo_row}C1:R{open_repo_last_row}C42",
                                                 Version=8))
        # macq_wb.api.ActiveSheet.PivotTables(1).PivotCache().SourceData = f"'{macq_date_sht.name}'!R{open_repo_row}C1:R{open_repo_last_row}C42"
        macq_wb.api.ActiveSheet.PivotTables("PivotTable1").PivotCache().Refresh()
        #2nd Pivot
        macq_wb.api.ActiveSheet.PivotTables("PivotTable2").PivotCache().SourceData = f"'{macq_date_sht.name}'!R{cur_mon_buyback_row}C1:R{date_sht_last_row}C42"
        macq_wb.api.ActiveSheet.PivotTables("PivotTable2").PivotCache().Refresh()
        #3rd Pivot
        macq_wb.api.ActiveSheet.PivotTables("PivotTable3").PivotCache().SourceData = f"'{macq_date_sht.name}'!R{open_repo_row}C1:R{open_repo_last_row}C42"
        macq_wb.api.ActiveSheet.PivotTables("PivotTable3").PivotCache().Refresh()
        #4th Pivot
        macq_wb.api.ActiveSheet.PivotTables("PivotTable4").PivotCache().SourceData = f"'{macq_date_sht.name}'!R{open_repo_row}C1:R{date_sht_last_row}C42"
        macq_wb.api.ActiveSheet.PivotTables("PivotTable4").PivotCache().Refresh()

        
        #################Updating True False Formulas#####################################
        macq_wb.activate()
        macq_m_end_je_sht.activate()
        # pvt_from_rows= []
        # pvt_last_rows = []
        formula_list = [f'=IF(GETPIVOTDATA("MONTH END STORAGE",$B$last_row)=SUM(\'{macq_date_sht.name}\'!AC9:AC23),"TRUE","FALSE")', 
                        f'=IF(GETPIVOTDATA("STORAGE THRU PURCHASE DATE",$B$last_row)=SUM(\'{macq_date_sht.name}\'!X30:X42),"TRUE","FALSE")',
                        f'=IF(GETPIVOTDATA("Sum of MONTH END INVENTORY VALUE",$B$last_row)=SUM(\'{macq_date_sht.name}\'!AG9:AG17),"TRUE","FALSE")',
                        f'=IF(GETPIVOTDATA("CURRENT MONTH GROSS INTEREST",$B$last_row)=SUM(\'{macq_date_sht.name}\'!AB9:AB42),"TRUE","FALSE")'
                        ]
        finder_list = ['**MONTH END STORAGE ACCRUAL ON OPEN REPOS (REVERSING)', '**MONTH END STORAGE ON CURRENT MONTH BUYBACKS (NONREVERSING)',
                       '**MONTH END REPO INV. & LIAB (REV ENTRY)', '**MONTH END REPO INTEREST RECLASS (NR ENTRY)']
        for finder in range(len(finder_list)):

            macq_wb.api.ActiveSheet.PivotTables(f"PivotTable{finder+1}").PivotSelect("",win32c.PTSelectionMode.xlDataAndLabel,True)
            pvt_last_row = macq_wb.api.Application.Selection.Address.replace("$","").split(":")[1][1:]
            macq_m_end_je_sht.api.Cells.Find(What:=finder_list[finder], After:=macq_m_end_je_sht.api.Application.ActiveCell,LookIn:=win32c.FindLookIn.xlFormulas,
                                    LookAt:=win32c.LookAt.xlPart, SearchOrder:=win32c.SearchOrder.xlByRows, SearchDirection:=win32c.SearchDirection.xlNext).Activate()
            pvt_formula_row = macq_m_end_je_sht.api.Application.ActiveCell.Address.replace("$","")[1:]

            macq_m_end_je_sht.range(f"F{pvt_formula_row}").formula = formula_list[finder].replace("last_row", pvt_last_row)
        
        

        repos_sht.range("A:U").clear_contents()
        macq_wb.app.api.CutCopyMode=False
        macq_wb.api.ActiveSheet.Range("A:U").Copy()

        wb.activate()
        repos_sht.activate()
        
        repos_sht.api.Range(f"A1")._PasteSpecial(Paste=win32c.PasteType.xlPasteValuesAndNumberFormats,Operation=win32c.Constants.xlNone)
        ######################Updating prior week data in Current file############################
        retry=0
        while retry<10:
            try:
                prior_wb = xw.Book(prior_wb_loc, update_links=False)
                break
            except Exception as e:
                time.sleep(2)
                retry+=1
                if retry==9:
                    raise e

        prior_est_sht = prior_wb.sheets("West Plains Estimate")
        est_sht = wb.sheets("West Plains Estimate")
        prior_loc_est_sht = prior_wb.sheets("Locations Estimate")
        loc_est_sht = wb.sheets("Locations Estimate")
        prior_est_sht.range("C6:C8").copy()
        est_sht.range("D6:D8").api._PasteSpecial(Paste=win32c.PasteType.xlPasteValues,Operation=win32c.Constants.xlNone)
        #Copy Pasting formulas
        est_sht.range("C9:C12").copy()
        est_sht.range("D9:D12").paste()

        ##Location estimate sheet logic####
        prior_loc_est_sht.range("B4:B6").copy()
        loc_est_sht.range("C4:C6").api._PasteSpecial(Paste=win32c.PasteType.xlPasteValues,Operation=win32c.Constants.xlNone)
        #Copy Pasting formulas
        loc_est_sht.range("B7:B10").copy()
        loc_est_sht.range("C7:C10").paste()

        prior_loc_est_sht.range("B16:B18").copy()
        loc_est_sht.range("C16:C18").api._PasteSpecial(Paste=win32c.PasteType.xlPasteValues,Operation=win32c.Constants.xlNone)
        #Copy Pasting formulas
        loc_est_sht.range("B19:B22").copy()
        loc_est_sht.range("C19:C22").paste()

        prior_loc_est_sht.range("B28:B30").copy()
        loc_est_sht.range("C28:C30").api._PasteSpecial(Paste=win32c.PasteType.xlPasteValues,Operation=win32c.Constants.xlNone)
        #Copy Pasting formulas
        loc_est_sht.range("B31:B34").copy()
        loc_est_sht.range("C31:C34").paste()

        prior_loc_est_sht.range("B40:B42").copy()
        loc_est_sht.range("C40:C42").api._PasteSpecial(Paste=win32c.PasteType.xlPasteValues,Operation=win32c.Constants.xlNone)
        #Copy Pasting formulas
        loc_est_sht.range("B43:B46").copy()
        loc_est_sht.range("C43:C46").paste()

        prior_loc_est_sht.range("B52:B54").copy()
        loc_est_sht.range("C52:C54").api._PasteSpecial(Paste=win32c.PasteType.xlPasteValues,Operation=win32c.Constants.xlNone)
        #Copy Pasting formulas
        loc_est_sht.range("B55:B58").copy()
        loc_est_sht.range("C55:C58").paste()

        prior_loc_est_sht.range("B64:B66").copy()
        loc_est_sht.range("C64:C66").api._PasteSpecial(Paste=win32c.PasteType.xlPasteValues,Operation=win32c.Constants.xlNone)
        #Copy Pasting formulas
        loc_est_sht.range("B67:B70").copy()
        loc_est_sht.range("C67:C70").paste()







        ####Saving Workbook
        wb.save(output_loc)
        macq_wb.save(mac_output_loc)
        macq_wb.close()

        
        
        return f"{job_name} Report for {input_date} generated succesfully and previous output file picked is {prior_wb_loc}"
    except Exception as e:
        raise e
    finally:
        try:
            wb.app.quit()
        except:
            pass









def main():
    def on_closing():
        if messagebox.askokcancel("Quit", "Do you want to quit?"):
            
            root.destroy()
            sys.exit()
    def callback_2():
    
        # def report_callback_exception(self, exc, val, tb):
        #     showerror("Error", message=str(exc) + str(val) +str(tb))

        # try:
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
            if type(msg)==list:
                showwarning(title="Warning",message=msg[1])
                msg=msg[0]
            text_box.delete(1.0, "end")
            text_box.insert("end", f"\n{msg}", "center")
            submit_text.set("Submit")
            Rep_variable.set('Select')
            root.update()

            print()
        
        elif 'Select' in Rep_variable.get():
            text_box.delete(1.0, "end")
            text_box.insert("end", f"Please select job first", "center")


        root.update()
        # except Exception as e:
        #     Tk.report_callback_exception = report_callback_exception
        
        
    # def callback():
    #     try:
    #         threading.Thread(target=callback_2).start()
    #     except Exception as e:
    #         raise e
        
        
    def report_callback_exception(self, exc, val, tb):
        msg = traceback.format_exc()
        showerror("Error", message=msg)
        text_box.delete(1.0, "end")
        text_box.insert("end", str(exc), "center")
        submit_text.set("Submit")
        Rep_variable.set('Select')
        root.update()

    Tk.report_callback_exception = report_callback_exception
    frame_title.grid(row=0, column=1,pady=(24,0),columnspan=3, padx=(30,0))
    logo = PhotoImage(file = path + '\\'+'wp_logo.png')
    # logo = PhotoImage(file = path + '\\'+'wp_logo.png')


    title = Label(frame_title,image=logo, background='white')
    # title = Label(frame_title, text="Revelio Report Generator", font=("Algerian", 28), foreground='black', background="white")
    title.grid(column=1,row=0)

    root.protocol("WM_DELETE_WINDOW", on_closing)
    # input_date=None
    # output_date = None
    frame_options.grid(row=1,column=0, pady=30, padx=35, columnspan=2, rowspan=3)
    wp_job_ids = {'ABS':1,'MTD New Trades Report':mtd_new_trades,'AR Exposure E2E':ar_exposure,'AR Reports Exposure':ar_reports_exposure,'BBR':bbr,'CPR Report':cpr, 'Freight analysis':freight_analysis, 'CTM combined':ctm,'FIFO Report':fifo,'MTM Report':mtm_report,'Inventory MTM excel report summary':inv_mtm_excel_summ,
                    'MOC Interest Allocation':moc_interest_alloc,'Open AR':open_ar,'Open AP':open_ap, 'Unsettled Payable Report':unsetteled_payables,'Unsettled Receivable Report':unsetteled_receivables,
                    'Storage Month End Report':strg_month_end_report, "Month End BBR":bbr_monthEnd, "Bank Recons Report":bank_recons_rep, "Payables_GL_Entry_Monthly":payables_gl_entry_monthly,
                    "Receivables_GL_Entry_Monthly":receivables_gl_entry_monthly,"CTM_GL_Entry_Monthly":ctm_gl_entry_monthly, "Macquarie Accrual Entry":macq_accr_entry, "Ticket_N_Settlement_Report":tkt_n_settlement_summ,
                    "Payroll_Summary":payroll_summ,"Credit_Card_Entry":credit_card_entry, "Credit_Card_GL_Entry_Monthly":credit_card_gl,"Unsettled_AR_By_Reason":unsettled_ar_by_location_part1,
                    "Unsettled_AR_By_Location_with_Location":unsettled_ar_by_location_part2,"Open_AR_Monthly":open_ar_monthly,"Weekly_Estimate":weekly_estimate}
    # wp_job_ids = {'ABS':1,'BBR':bbr,'CPR Report':cpr, 'Freight analysis':freight_analysis, 'CTM combined':ctm,'MTM Report':mtm_report,
    #                 'MOC Interest Allocation':moc_interest_alloc,'Open AR':open_ar,'Open AP':open_ap, 'Unsettled Payable Report':unsetteled_payables,'Unsettled Receivable Report':unsetteled_receivables,
    #                 'Storage Month End Report':strg_month_end_report, "Month End BBR":bbr_monthEnd, "Bank Recons Report":bank_recons_rep}
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
    submit = Button(frame_submit, textvariable=submit_text, font = ("Book Antiqua bold", 12), bg="#20bebe", fg="white", height=1, width=14, command=callback_2, activebackground="#20bebb")
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

