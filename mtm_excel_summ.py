import os, time
import xlwings as xw
from datetime import datetime
from tabula import read_pdf
import pandas as pd
from collections import defaultdict



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

        com_loc = list(com_loc[0].str.split('Commodity: ',expand=True)[1])
        # loc_dict = dict(zip(com_loc, [[]]*len(com_loc)))
        loc_dict = defaultdict(list)
        for page in range(1,len(com_loc)+1):
            df = read_pdf(f, pages = page, guess = False, stream = True ,
                            pandas_options={'header':0}, area = ["75,10,580,850"], columns=["65,85, 180,225, 260, 280,300,360,400,430,480,525,570,620,665,720"])
            df = pd.concat(df, ignore_index=True)
            ########logger.info("Filtering only required columns")
            df = df.iloc[:,[0,1,2,3,-2,-1]]
            # df = df[df['Offsite Name Cont. No.'].str.contains("Company Owned Risk:"),df['Offsite Name Cont. No.'].str.contains("Unpriced Sales:")]
            df = df[(df['Offsite Name Cont. No.'].str.contains("Company Owned Risk:")) | (df['Offsite Name Cont. No.'].str.contains("priced Sales:"))]
            # for i in df.loc[:,"Offsite Name Cont. No."]:

            df["Quantity.5"].fillna(0, inplace=True)
            df["Value.5"].fillna(0, inplace=True)

            df["Quantity.5"] = df["Quantity.5"].astype(str).str.replace("(","-").str.replace(",","").str.replace(")","").astype(float)
            df["Value.5"] = df["Value.5"].astype(str).str.replace("(","-").str.replace(",","").str.replace(")","").astype(float)

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
        
        if mtm_report:
            return loc_dict, hrw_fut, yc_fut
        else:
            return loc_dict
    except Exception as e:
        raise e








def inv_mtm_excel_summ(input_date, output_date):
    try:
        monthYear = datetime.strftime(datetime.strptime(input_date, "%m.%d.%Y"), "%B %Y")
        pdf_loc = r'J:\WEST PLAINS\REPORT\MTM reports\Raw Files\Inventory Market Valuation _'+input_date+'.pdf'
        # pdf_loc = r'C:\Users\imam.khan\OneDrive - BioUrja Trading LLC\Documents\WEST PLAINS\REPORT\MTM reports\Raw Files\Inventory Market Valuation _'+input_date+'.pdf'
        if not os.path.exists(pdf_loc):
            return(f"{pdf_loc} Pdf file not present for date {input_date}")
        input_xl = r'J:\WEST PLAINS\REPORT\Inv_MTM_Excel_Report_Summ\Raw Files\Inventory_MTMExcel_SummTemplate.xlsx'
        # input_xl = r'C:\Users\imam.khan\OneDrive - BioUrja Trading LLC\Documents\WEST PLAINS\REPORT\Inv_MTM_Excel_Report_Summ\Raw Files\Inventory_MTMExcel_SummTemplate.xlsx'
        if not os.path.exists(input_xl):
            return(f"{input_xl} Excel file not present for date {input_date}")

        # output_loc = r"C:\Users\imam.khan\OneDrive - BioUrja Trading LLC\Documents\WEST PLAINS\REPORT\FIFO reports\Output files" +f"\\Inventory MTM Excel Report {monthYear}.xlsx"
        output_loc = r'J:\WEST PLAINS\REPORT\Inv_MTM_Excel_Report_Summ\Output files' +f"\\Inventory MTM Excel Report {monthYear}.xlsx"



        loc_dict = inv_mtm_pdf_data_extractor(input_date,pdf_loc)

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
                try:
                    loc_dict[mtm_sht.range(f"D{i}").value][0].rename(index={'OMA COMM': 'TERMINAL'}, inplace=True)
                except:
                    pass
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

        mtm_wb.save(output_loc)

        return f"MTM report Generated for {input_date}"
    except Exception as e:
        raise e
    finally:
        try:
            mtm_wb.app.quit()
        except:
            pass







input_date = "03.31.2022"
output_date=None

msg = inv_mtm_excel_summ(input_date, output_date)
print(msg)
print()