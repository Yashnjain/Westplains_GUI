from cProfile import label
import PyPDF2
from tabula import read_pdf
import pandas as pd
import xlwings as xw
import time, os
from datetime import datetime, timedelta




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





def cash_colat(wb,bank_recons_loc, input_date_date):
    try:
        
        retry=0
        while retry < 10:
            try:
                bank_wb=xw.Book(bank_recons_loc)
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry ==10:
                    raise e

        while True:
            try:
                cash_colat_sht = wb.sheets["Cash Collateral"] #wb.sheets[0].name in 'Unsettled Receivables _'+input_date
                break
            except Exception as e:
                time.sleep(5)

        while True:
            try:
                bank_colat_sht = bank_wb.sheets["BANK REC"] #wb.sheets[0].name in 'Unsettled Receivables _'+input_date
                break
            except Exception as e:
                time.sleep(5)
        cash_colat_sht.range("A3").value = input_date_date
        # cash_colat_sht.range("B58").value = bank_colat_sht.range("B12").value
        # cash_colat_sht.range("E58").value = bank_colat_sht.range("B14").value
        cash_colat_sht.range("B12").value = bank_colat_sht.range("B58").value
        cash_colat_sht.range("B14").value = bank_colat_sht.range("E58").value

        bank_wb.close()
    except Exception as e:
        raise e

def ar_unsettled_by_tier(wb, unset_rec_loc):
    try:
        retry=0
        while retry < 10:
            try:
                unset_rec_wb=xw.Book(unset_rec_loc)
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry ==10:
                    raise e

        while True:
            try:
                xl_mac_n_ic = unset_rec_wb.sheets["Excl Macq & IC"] #wb.sheets[0].name in 'Unsettled Receivables _'+input_date
                break
            except Exception as e:
                time.sleep(5)
        last_row=xl_mac_n_ic.range(f'A' + str(xl_mac_n_ic.cells.last_cell.row)).end('up').row
        # 'J:\WEST PLAINS\REPORT\BBR Reports\Output\[Unsettled Receivables _02.14.2022.xlsx]Excl IC & Macq'!$A$1:$AJ$892
        unset_rec_wb.close()
        while True:
            try:
                ar_unsettled_by_tier_sht = wb.sheets["AR Unsettled ByTier"] #wb.sheets[0].name in 'Unsettled Receivables _'+input_date
                break
            except Exception as e:
                time.sleep(5)

        while True:
            try:
                ar_unsettled_by_tier_sht.select()
                break
            except Exception as e:
                time.sleep(5)
        
        # sht = wb.sheets["AR-Trade By Tier - Eligible"]
        wb.api.ActiveSheet.PivotTables(1).PivotCache().SourceData = f"'J:\\WEST PLAINS\\REPORT\\Unsettled Receivables\\Output Files\\[Unsettled Receivables _{input_date}.xlsx]Excl Macq & IC'!R1C1:R{last_row}C36"
        
          #f'Details!R1C1:R{len(new_rows)+1}C18' #Updateing data source
        wb.api.ActiveSheet.PivotTables(1).PivotCache().Refresh()
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
            acc_no = a[a.find('Account'):a.find('Account')+17]
            acc_no = acc_no.replace("Account: ","")
            print(f"account_num = {acc_no}, prev_acc = {prev_acc_no} and page is {page}")
            if acc_no == "":
                continue
            if prev_acc_no is None:
                prev_acc_no=acc_no #a[25:42]
            elif prev_acc_no != acc_no:
                print(page-1)
                
                print(acc_no)
                if str(prev_acc_no) in account_lst:
                    df = read_pdf(pdf_loc, pages = page, guess = False, stream = True ,
                                pandas_options={'header':0}, area = ["75,10,725,850"], columns=["180,280"])
                    df = pd.concat(df, ignore_index=True)
                    print(df)
                    amount_dict[prev_acc_no] = float(df.iloc[-1,-1].replace(",","")) 
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
                    print(prev_acc_no)
                    print()
                prev_acc_no = acc_no

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
                time.sleep(5)
        cell = 8
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
                time.sleep(5)
                retry+=1
                if retry ==10:
                    raise e

        while True:
            try:
                strg_acc_sht = strg_accr_wb.sheets[0] #wb.sheets[0].name in 'Unsettled Receivables _'+input_date
                break
            except Exception as e:
                time.sleep(5)

        while True:
            try:
                bbr_strg_acc_sht = wb.sheets["AR-Open Storage Rcbl"] #wb.sheets[0].name in 'Unsettled Receivables _'+input_date
                break
            except Exception as e:
                time.sleep(5)
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

def inv_whre_n_in_trans(wb, mtm_loc):
    try:
        retry=0
        while retry < 10:
            try:
                mtm_wb=xw.Book(mtm_loc)
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry ==10:
                    raise e

        while True:
            try:
                m_sht = mtm_wb.sheets[0] #wb.sheets[0].name in 'Unsettled Receivables _'+input_date
                break
            except Exception as e:
                time.sleep(5)

        while True:
            try:
                whre_sht = wb.sheets["Inventory Whre & In-Trans"] #wb.sheets[0].name in 'Unsettled Receivables _'+input_date
                break
            except Exception as e:
                time.sleep(5)
        while True:
            try:
                inv_oth_sht = wb.sheets["Inventory -Other"] #wb.sheets[0].name in 'Unsettled Receivables _'+input_date
                break
            except Exception as e:
                time.sleep(5)


        last_row=m_sht.range(f'A' + str(m_sht.cells.last_cell.row)).end('up').row
        main_loc = m_sht.range(f"A1:A{last_row}").value
        hrw_value=0
        yc_value = 0
        whre_sht.range(f"A3").value = datetime.strptime(input_date,"%m.%d.%Y")
        inv_oth_sht.range(f"A3").value = datetime.strptime(input_date,"%m.%d.%Y")
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


        whre_sht.range(f"C{hrw}").options(transpose=True).value = m_sht.range(f"C{hrw}").expand("down").value
        whre_sht.range(f"F{hrw}").options(transpose=True).value = m_sht.range(f"F{hrw}").expand("down").value
        whre_sht.range(f"I{hrw}").options(transpose=True).value = m_sht.range(f"I{hrw}:I{int(yc)-4}").value
        
        

        whre_sht.range(f"C{yc}").options(transpose=True).value = m_sht.range(f"C{yc}").expand("down").value
        whre_sht.range(f"F{yc}").options(transpose=True).value = m_sht.range(f"F{yc}").expand("down").value
        whre_sht.range(f"I{yc}").options(transpose=True).value = m_sht.range(f"I{yc}:I{int(other_loc_2)-5}").value

        whre_sht.range(f"C{other_loc_2}").options(transpose=True).value = m_sht.range(f"C{other_loc_2}").expand("down").value
        whre_sht.range(f"F{other_loc_2}").options(transpose=True).value = m_sht.range(f"F{other_loc_2}").expand("down").value


        inv_oth_sht.range(f"C{int(other_loc)-64}").options(transpose=True).value = m_sht.range(f"C{other_loc}:C{int(sunflwr)-6}").value
        inv_oth_sht.range(f"F{int(other_loc)-64}").options(transpose=True).value = m_sht.range(f"F{other_loc}:F{int(sunflwr)-6}").value
        
        inv_oth_sht.range(f"C{int(sunflwr)-64}").options(transpose=True).value = m_sht.range(f"C{sunflwr}").value
        inv_oth_sht.range(f"F{int(sunflwr)-64}").options(transpose=True).value = m_sht.range(f"F{sunflwr}").value


        mtm_wb.close()
        
        print()
    except Exception as e:
        raise e

def payables(wb, bbr_mapping_loc, open_ap_loc,unset_pay_loc):
    try:
        df = pd.read_excel(bbr_mapping_loc, usecols="A,B")   
        new_dict = dict(zip(df.iloc[:,0],df.iloc[:,1]))
        payab_df = pd.read_excel(bbr_mapping_loc, usecols="D,E")
        payab_dict = dict(zip(payab_df.iloc[:,0],payab_df.iloc[:,1]))
        retry=0
        while retry < 10:
            try:
                open_ap_wb=xw.Book(open_ap_loc)
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry ==10:
                    raise e

        while True:
            try:
                open_ap_sht = open_ap_wb.sheets["Pivot BB"] #wb.sheets[0].name in 'Unsettled Receivables _'+input_date
                break
            except Exception as e:
                time.sleep(5)
        while retry < 10:
            try:
                payab_wb=xw.Book(unset_pay_loc)
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry ==10:
                    raise e

        while True:
            try:
                payab_sht = payab_wb.sheets["Pivot BB"] #wb.sheets[0].name in 'Unsettled Receivables _'+input_date
                break
            except Exception as e:
                time.sleep(5)
        while True:
            try:
                bbr_payab_sht = wb.sheets["Payables"] #wb.sheets[0].name in 'Unsettled Receivables _'+input_date
                break
            except Exception as e:
                time.sleep(5)
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
                    bbr_payab_sht.range(f"{bbr_last_row+i}:{bbr_last_row+i}").insert()
            else:
                for i in range(len(bbr_loc) - len(open_ap_loc_lst)):
                    bbr_payab_sht.range(f"{bbr_last_row-i}:{bbr_last_row-i}").delete()
        else:
            pass

        i=10
        for loc in open_ap_loc_lst:
            bbr_payab_sht.range(f"A{i}").value = new_dict[loc]
            try:
                bbr_payab_sht.range(f"C{i}").value = dict_1[loc]
            except:
                bbr_payab_sht.range(f"C{i}").value = 0
            try:
                bbr_payab_sht.range(f"E{i}").value = dict_2[loc]
            except:
                bbr_payab_sht.range(f"E{i}").value = 0
            i+=1
        p_last_row = payab_sht.range("A4").end('down').row
        payab_loc_lst = payab_sht.range(f"A4:A{int(p_last_row)-1}").value
        total_col = payab_sht.range(f"D4:D{int(p_last_row)-1}").value

        

        dict_3 = dict(zip(payab_loc_lst,total_col))

        payb_loc = bbr_payab_sht.range(f"A{i}").end("down").end("down").row

        bbr_payb_loc_lst = bbr_payab_sht.range(f"A{payb_loc}").expand("down").value
        bbr_payb_loc_lst = bbr_payb_loc_lst[:-1]
        for loc in payab_loc_lst:
            bbr_payab_sht.range(f"A{i}").value = payab_dict[loc]
            try:
                bbr_payab_sht.range(f"C{payb_loc}").value = dict_3[loc]
            except:
                bbr_payab_sht.range(f"C{payb_loc}").value = 0
            payb_loc+=1

        open_ap_wb.close()
        payab_wb.close()

        
        print()
    except Exception as e:
        raise e












def bbr(input_date):
    try:
        output_location = r'J:\WEST PLAINS\REPORT\BBR Reports\Output files'+f"\\{input_date}_Borrowing Base Report.xlsx"
        input_date_date = datetime.strptime(input_date, "%m.%d.%Y").date()
        prev_month_year = datetime.strftime((datetime.strptime(input_date, "%m.%d.%Y").replace(day=1)-timedelta(days=1)),"%b %Y").upper()
        print(prev_month_year)
        input_xl = r"J:\WEST PLAINS\REPORT\BBR Reports\Raw Files" +f"\\{input_date}_Borrowing Base Report.xlsx"
        if not os.path.exists(input_xl):
                return(f"{input_xl} Excel file not present for date {input_date}")
        # account_lst = ["52311940", "523WP771", "523WP774", "523WP775", "523WP777", "523WP779", "523WP780", "523WP781", "523WP782", "523WP783", "523WP784", "523WP785", "523WP786", "523WP787", "523WP788", "523WP789", "523WP790", "523WP791", "523WP792", "523WP793", "523WP794", "523WP795", "523WPHLD"]
        pdf_loc = r"J:\WEST PLAINS\REPORT\BBR Reports\Raw Files\Macquarie Statement_"+input_date+".pdf"
        if not os.path.exists(pdf_loc):
                return(f"{pdf_loc} Pdf file not present for date {input_date}")

        bank_recons_loc = r"J:\WEST PLAINS\REPORT\BBR Reports\Raw Files\BANK RECONS_"+input_date+".xls"

        if not os.path.exists(bank_recons_loc):
                return(f"{bank_recons_loc} Excel file not present for date {input_date}")

        strg_accr_loc = r"J:\WEST PLAINS\REPORT\BBR Reports\Raw Files\STORAGE ACCRUAL "+prev_month_year+".xlsx"

        if not os.path.exists(strg_accr_loc):
                return(f"{strg_accr_loc} Excel file not present for date {input_date}")

        bbr_mapping_loc = r"J:\WEST PLAINS\REPORT\BBR Reports\bbr_payables_mapping.xlsx"

        if not os.path.exists(bbr_mapping_loc):
                return(f"{bbr_mapping_loc} Excel file not present for date {input_date}")




        unset_rec_loc = r'J:\WEST PLAINS\REPORT\Unsettled Receivables\Output files\Unsettled Receivables _'+input_date+'.xlsx'
        if not os.path.exists(unset_rec_loc):
            return(f"{unset_rec_loc} Excel file not present for date {input_date}")
        mtm_loc = r'J:\WEST PLAINS\REPORT\MTM reports\Output files\Inventory MTM_'+input_date+".xlsx"
        if not os.path.exists(mtm_loc):
            return(f"{mtm_loc} Excel file not present for date {input_date}")

        open_ap_loc = r'J:\WEST PLAINS\REPORT\Open AP\Output files\Open AP _'+input_date+'.xlsx'

        if not os.path.exists(open_ap_loc):
            return(f"{open_ap_loc} Excel file not present for date {input_date}")

        unset_pay_loc = r'J:\WEST PLAINS\REPORT\Unsettled Payables\Output files\Unsettled Payables _'+input_date+'.xlsx'

        if not os.path.exists(unset_pay_loc):
            return(f"{unset_pay_loc} Excel file not present for date {input_date}")
        # amount_dict = comm_acc_pdf_ext(account_lst, pdf_loc)
        # print(amount_dict)
        # print()

        retry=0
        while retry < 10:
            try:
                wb=xw.Book(input_xl)
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry ==9:
                    raise e

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
            suffix = ["st", "nd", "rd"][input_date_date % 10 - 1]
        cur_date= datetime.strftime(input_date_date, f"%B %d{suffix}, %Y")
        bbr_sht.range("A4").value = f'As of {cur_date} (the "Determination Date")'
        cash_colat(wb,bank_recons_loc, input_date_date)
        comm_acc_xl(wb, pdf_loc)
        ar_unsettled_by_tier(wb, unset_rec_loc)
        ar_open_storage_rcbl(wb, strg_accr_loc, input_date)
        inv_whre_n_in_trans(wb, mtm_loc)
        payables(wb, bbr_mapping_loc, open_ap_loc,unset_pay_loc)

        wb.save(output_location)
        print()
    except Exception as e:
        raise e
    finally:
        wb.app.quit()
    


inp_date = ["02.28.2022"]
for input_date in inp_date:
    bbr(input_date)
print("Done")


