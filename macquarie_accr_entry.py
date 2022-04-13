from codecs import lookup
from tabula import read_pdf
import PyPDF2
import pandas as pd
import xlwings as xw
import os, time
from datetime import datetime, timedelta


def mac_accr_pdf(input_pdf):
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
            df = read_pdf(input_pdf, pages = page+1, guess = False, stream = True ,
                        pandas_options={'header':0}, area = ["50,10,725,850"], columns=["195,280,430"])
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


def macq_accr_entry(input_date, output_date):
    try:
        xl_date = datetime.strftime(datetime.strptime(input_date, "%m.%d.%Y"), "%Y%m%d")
        next_date =datetime.strftime((datetime.strptime(input_date, "%m.%d.%Y")+timedelta(days=1)), "%Y%m%d")
        output_loc =  r"J:\WEST PLAINS\REPORT\Macquaire Accrual Entry\Output Files" +f"\\Macq Statement_{input_date}.xlsx"
        input_pdf = r"J:\WEST PLAINS\REPORT\Macquaire Accrual Entry\Raw Files" +f"\\Macq Statement_{input_date}.pdf"
        # input_pdf = r"C:\Users\imam.khan\OneDrive - BioUrja Trading LLC\Documents\WEST PLAINS\REPORT\Macquaire Accrual Entry\Raw Files" +f"\\Macq Statement_{input_date}.pdf"
        if not os.path.exists(input_pdf):
                return(f"{input_pdf} PDF file not present for date {input_date}")
        input_xl = r"J:\WEST PLAINS\REPORT\Macquaire Accrual Entry\Raw Files" +f"\\Macq Accrual_{input_date}.xlsx"
        # input_xl = r"C:\Users\imam.khan\OneDrive - BioUrja Trading LLC\Documents\WEST PLAINS\REPORT\Macquaire Accrual Entry\Raw Files" +f"\\Macq Accrual_{input_date}.xlsx"
        if not os.path.exists(input_xl):
                return(f"{input_xl} Excel file not present for date {input_date}")

        mapping_loc = r"J:\WEST PLAINS\REPORT\Macquaire Accrual Entry\Mapping.xlsx"

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


input_date = '01.31.2022'
msg = macq_accr_entry(input_date, output_date=None)
print(msg)
print()