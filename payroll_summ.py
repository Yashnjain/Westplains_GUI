import PyPDF2
from tabula import read_pdf
import os
from datetime import datetime, timedelta
import glob
from dateutil.relativedelta import relativedelta
import time
import xlwings as xw







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



def payroll_pdf_extractor(input_pdf, input_datetime, monthYear):
    main_dict = {}
   
    for loc in glob.glob(input_pdf):       #add month difference if ==2 then not consider that file
        file_date = loc.split()[-1].split(".pdf")[0].replace(".","-")
        file_datetime = datetime.strptime(loc.split()[-1].split(".pdf")[0],"%m.%d.%Y")
        file_date = datetime.strftime(file_datetime, "%d-%m-%Y")
        diff = relativedelta(input_datetime.replace(day=1),file_datetime.replace(day=1))
        diff = diff.months*(diff.years+1)
        if diff == 0 or diff == 1 or diff==-1:
            date_df = read_pdf(loc, pages = 1, guess = False, stream = True ,
                                pandas_options={'header':0}, area = ["30,290,120,415"], columns=["320"])[0]
            dates = date_df.iloc[0,1].split("to")
            monthYear1 = datetime.strftime(datetime.strptime(dates[0].strip(), "%m/%d/%Y"), "%b %y")
            monthYear2 = datetime.strftime(datetime.strptime(dates[1].strip(), "%m/%d/%Y"), "%b %y")
            if monthYear1 == monthYear or monthYear2 == monthYear:
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

                    medicare_ee = 0  #ER-Med	      Medicare -EE R
                    soc_sec_er = 0   #ER-SS           Social Security - ER
                    futa_nesui = 0   #FUTA             NESUI
                    suta_cosui = 0   #SUTA             COSUI
                    suta_wysui = 0   #SUTA             WYSUI
                    ffcra = 0        #FFCRA            Value Not Received Till Now ( Blank )
                    benefits = 0     #Benefits         Value Not Received Till Now ( Blank )
                    med_dent_vis = 0 # Med/Dent/Vis	   Total Value of Cafeteria 125 Deds
                    volutary = 0 #Voluntary            Sum of All Misc. Expenses with no Parent Name (Deduction Analysis )
                    garnish_chldi = 0 #Garnishment     Deduction Analysis – CHLDI
                    ee_401k = 0 #EE 401k                 Deduction Analysis 401K
                    er_401k = 0 #ER401k	               Deduction Analysis 4ROTH
                    ee_roth = 0 #EE Roth 	          Value Not Received Till Now ( Blank )
                    kln_401 = 0 #401KLN	                Deduction Analysis 401L1
                
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
                
                    for col in range(len(deduc_ana_df)):
                        if deduc_ana_df[deduc_ana_df.columns[0]][col] == "Cafeteria 125":
                            if deduc_ana_df.iloc[-1,0]!="Cafeteria 125":
                                while deduc_ana_df[deduc_ana_df.columns[0]][col] !="Total":
                                    col+=1
                                if "("  in deduc_ana_df[deduc_ana_df.columns[-1]][col] and ")" in deduc_ana_df[deduc_ana_df.columns[-1]][col]:
                                    med_dent_vis = float(deduc_ana_df[deduc_ana_df.columns[-1]][col].replace(",","").replace("(","").replace(")",""))
                                else:    
                                    med_dent_vis = float(deduc_ana_df[deduc_ana_df.columns[-1]][col].replace(",",""))*-1
                                break
                        
                        elif deduc_ana_df[deduc_ana_df.columns[0]][col] == "CHLD1":
                            if "("  in deduc_ana_df[deduc_ana_df.columns[-1]][col] and ")" in deduc_ana_df[deduc_ana_df.columns[-1]][col]:
                                garnish_chldi = float(deduc_ana_df[deduc_ana_df.columns[-1]][col].replace(",","").replace("(","").replace(")",""))
                            else:
                                garnish_chldi = float(deduc_ana_df[deduc_ana_df.columns[-1]][col].replace(",",""))*-1
                        elif deduc_ana_df[deduc_ana_df.columns[0]][col] == "401K":
                            if "("  in deduc_ana_df[deduc_ana_df.columns[-1]][col] and ")" in deduc_ana_df[deduc_ana_df.columns[-1]][col]:
                                ee_401k = float(deduc_ana_df[deduc_ana_df.columns[-1]][col].replace(",","").replace("(","").replace(")",""))
                            else:
                                ee_401k = float(deduc_ana_df[deduc_ana_df.columns[-1]][col].replace(",",""))*-1
                        elif deduc_ana_df[deduc_ana_df.columns[0]][col] == "4ROTH":
                            if "("  in deduc_ana_df[deduc_ana_df.columns[-1]][col] and ")" in deduc_ana_df[deduc_ana_df.columns[-1]][col]:
                                er_401k = float(deduc_ana_df[deduc_ana_df.columns[-1]][col].replace(",","").replace("(","").replace(")",""))
                            else:
                                er_401k = float(deduc_ana_df[deduc_ana_df.columns[-1]][col].replace(",",""))*-1
                        elif deduc_ana_df[deduc_ana_df.columns[0]][col] == "401L1":
                            if "("  in deduc_ana_df[deduc_ana_df.columns[-1]][col] and ")" in deduc_ana_df[deduc_ana_df.columns[-1]][col]:
                                kln_401 = float(deduc_ana_df[deduc_ana_df.columns[-1]][col].replace(",","").replace("(","").replace(")",""))
                            else:
                                kln_401 = float(deduc_ana_df[deduc_ana_df.columns[-1]][col].replace(",",""))*-1
                        else:
                            if deduc_ana_df[deduc_ana_df.columns[0]][col] != "Total":
                                if "("  in deduc_ana_df[deduc_ana_df.columns[-1]][col] and ")" in deduc_ana_df[deduc_ana_df.columns[-1]][col]:
                                    volutary += float(deduc_ana_df[deduc_ana_df.columns[-1]][col].replace(",","").replace("(","").replace(")",""))
                                else:
                                    volutary += float(deduc_ana_df[deduc_ana_df.columns[-1]][col].replace(",",""))*-1
                    if file_date in main_dict.keys():  
                        
                        main_dict[file_date][ada_group] = {"Gross":gross_value, "ER- SS":soc_sec_er, "ER - Med":medicare_ee, "FUTA":futa_nesui, "SUTA":suta_cosui+suta_wysui, "FFCRA": ffcra,
                            "Benefits":benefits, "Med/Dent/Vis":med_dent_vis, "Voluntary ":volutary, "Garnishment":garnish_chldi, "EE 401k ":ee_401k, "ER 401K":er_401k,
                            "EE Roth":ee_roth, "401KLN":kln_401}
                    else:  
                        main_dict[file_date] = {}
                        main_dict[file_date][ada_group] = {"Gross":gross_value, "ER- SS":soc_sec_er, "ER - Med":medicare_ee, "FUTA":futa_nesui, "SUTA":suta_cosui+suta_wysui, "FFCRA": ffcra,
                                "Benefits":benefits, "Med/Dent/Vis":med_dent_vis, "Voluntary ":volutary, "Garnishment":garnish_chldi, "EE 401k ":ee_401k, "ER 401K":er_401k,
                                "EE Roth":ee_roth, "401KLN":kln_401}
                    # main_dict[file_date] = {}
                    # main_dict[file_date][ada_group] = {"Gross":gross_value, "ER-SS":soc_sec_er, "ER-Med":medicare_ee, "FUTA":futa_nesui, "SUTA":suta_cosui+suta_wysui, "FFCRA": ffcra,
                    #             "Benefits":benefits, "Med/Dent/Vis":med_dent_vis, "Voluntary":volutary, "Garnishment":garnish_chldi, "EE 401k":ee_401k, "ER401k":er_401k,
                    #             "EE Roth":ee_roth, "401KLN":kln_401}
                    
        
    return main_dict

def payroll_summ(input_date, output_date):
    input_datetime = datetime.strptime(input_date,"%m.%d.%Y")
    monthYear = datetime.strftime(datetime.strptime(input_date, "%m.%d.%Y"), "%b %y")
    input_pdf = r"J:\WEST PLAINS\REPORT\Payroll summary accounting report\Raw Files" +f"\\Payroll Summary By Cost Center *.pdf"
    # input_pdf = r"C:\Users\imam.khan\OneDrive - BioUrja Trading LLC\Documents\WEST PLAINS\REPORT\Macquaire Accrual Entry\Raw Files" +f"\\Macq Statement_{input_date}.pdf"
    # if not os.path.exists(input_pdf):
    #         return(f"{input_pdf} PDF file not present for date {input_date}")
    input_xl = r"J:\WEST PLAINS\REPORT\Payroll summary accounting report\Raw Files" +f"\\Payroll by Dept - {monthYear}.xlsx"
    # input_xl = r"C:\Users\imam.khan\OneDrive - BioUrja Trading LLC\Documents\WEST PLAINS\REPORT\Macquaire Accrual Entry\Raw Files" +f"\\Macq Accrual_{input_date}.xlsx"
    if not os.path.exists(input_xl):
            return(f"{input_xl} Excel file not present for date {input_date}")
    template_xl = r"J:\WEST PLAINS\REPORT\Payroll summary accounting report\Raw Files" +f"\\Template.xlsx"
    # input_xl = r"C:\Users\imam.khan\OneDrive - BioUrja Trading LLC\Documents\WEST PLAINS\REPORT\Macquaire Accrual Entry\Raw Files" +f"\\Macq Accrual_{input_date}.xlsx"
    if not os.path.exists(template_xl):
            return(f"{template_xl} Excel file not present")
    output_location = r'J:\WEST PLAINS\REPORT\Payroll summary accounting report\Output Files'+f"\\Payroll by Dept - {monthYear}.xlsx"

    data = payroll_pdf_extractor(input_pdf, input_datetime, monthYear)

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
            inp_sht = wb.sheets["Detail"]
            break
        except Exception as e:
            time.sleep(2)
            retry+=1
            if retry ==9:
                raise e
    while retry < 10:
        try:
            t_wb=xw.Book(template_xl)
            break
        except Exception as e:
            time.sleep(2)
            retry+=1
            if retry ==9:
                raise e

    retry=0
    while retry < 10:
        try:
            t_sht = t_wb.sheets["Detail"]
            break
        except Exception as e:
            time.sleep(2)
            retry+=1
            if retry ==9:
                raise e

    inp_sht.range("A4:T4").expand("down").expand("down").delete()
    last_row=4
    first_row = 4
    init_chr = "F"
    last_column = inp_sht.range("F3").end('right').column-6 #considering f as intial col
    for pdf_data in range(len(data)-1,-1,-1):
        t_sht.range(f"A4:T4").expand("down").copy(inp_sht.range(f"A{first_row}"))#copying data from template
        last_row = inp_sht.range(f'A'+ str(inp_sht.cells.last_cell.row)).end('up').row
        #inserting dates
        inp_sht.range(f"C{first_row}:C{last_row}").value = list(data.keys())[pdf_data]
        
        for row in range(first_row,last_row+1):
            for col in range(last_column):
                try:
                    inp_sht.range(f"{chr(ord(init_chr)+col)}{row}").value = data[list(data.keys())[pdf_data]][inp_sht.range(f"A{row}").value][inp_sht.range(f"{chr(ord(init_chr)+col)}3").value]
                except:
                    inp_sht.range(f"{chr(ord(init_chr)+col)}{row}").value = 0
            #updating first row as last row
        first_row = last_row+1
    
    #Pivot part starts
    retry=0
    while retry < 10:
        try:
            p_sht = wb.sheets["Pivot"]
            break
        except Exception as e:
            time.sleep(2)
            retry+=1
            if retry ==9:
                raise e
    p_sht.activate()
    wb.api.ActiveSheet.PivotTables(1).PivotCache().SourceData = f"Detail!R3C1:R{last_row}C{last_column+6}" #Updateing data source, Removing initialization for f
    wb.api.ActiveSheet.PivotTables(1).PivotCache().Refresh()  

    p_last_row = p_sht.range(f'A'+ str(inp_sht.cells.last_cell.row)).end('up').row -1
    p_sht.range(f"C5:D{p_last_row}").copy(p_sht.range("G5"))

    #Updating Dates
    p_sht.range("K1").value = datetime.strftime(input_datetime.replace(day=1)-timedelta(days=1), "%d-%m-%Y") #Last Monthend
    p_sht.range("M1").value = datetime.strftime(input_datetime, "%d-%m-%Y") #Monthend

    
    wb.save(output_location)
    print()

    return f"Payroll Summary Report for {input_date} generated succesfully"
input_date = '03.11.2022'
msg = payroll_summ(input_date, output_date=None)
print(msg)
print()



