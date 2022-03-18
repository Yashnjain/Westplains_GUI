from socketserver import BaseRequestHandler
from xlwings.constants import DeleteShiftDirection
import xlwings as xw
import os,time,sys,datetime,logging
from datetime import datetime, date, timedelta
import pandas as pd
import bu_alerts

# today = date.today().strftime("%m.%d.%Y")
# print(today)

output_location  = r'J:\WEST PLAINS\REPORT\CPR reports\Output files'+'\\Counter Party Risk Consolidated'
# current_date = datetime.strftime(date.today(), "%m.%d.%Y")
# current_date = "01.31.2022"

job_name = 'WESTPLAINS_CPR_REPORT'


# book_name = os.getcwd() +'\\Raw Files'+'\\Counter Party Risk Consolidated 02-21-2022.xlsx'
# book_name = os.getcwd() +'\\Raw Files'+'\\Counter Party Risk Consolidated'
# sheet_date =book_name.split()[-1].split('.')[0].replace('-','.')
book_name = r'J:\WEST PLAINS\REPORT\CPR reports\Raw Files'+'\\Counter Party Risk Consolidated'

# BB_bookname = os.getcwd() +'\\Raw Files' + '\\Counter Party Risk Consolidated'
BB_bookname = r'J:\WEST PLAINS\REPORT\CPR reports\Raw Files' + '\\Counter Party Risk Consolidated'

# UnsettledRec_book = os.getcwd()+'\\Unsettled Receivables' +'\\Unsettled Receivables _'
UnsettledRec_book = r'J:\WEST PLAINS\REPORT\Unsettled Receivables\Output files' + '\\Unsettled Receivables _'

# UnsettledPay_book = os.getcwd()+'\\Unsettled Payables' + '\\Unsettled Payables _'
UnsettledPay_book = r'J:\WEST PLAINS\REPORT\Unsettled Payables\Output files' + '\\Unsettled Payables _'

# Open_AR_book = os.getcwd() + '\\Open AR and AP' + '\\Open AR _'
Open_AR_book = r'J:\WEST PLAINS\REPORT\Open AR and AP\Output files' + '\\Open AR _'

# Open_AP_book = os.getcwd() + '\\Open AR and AP' + '\\Open AP _'
Open_AP_book = r'J:\WEST PLAINS\REPORT\Open AR and AP\Output files' + '\\Open AP _'

# CTM_book = os.getcwd() + '\\CTM Combined report'  + '\\CTM Combined _'
CTM_book = r'J:\WEST PLAINS\REPORT\CTM Combined report\Output files' + '\\CTM Combined _'


logger = logging.getLogger()
logger.setLevel(logging.INFO)
formatter =logging.Formatter('%(asctime)s:%(levelname)s:%(name)s:%(message)s')
receiver_email = "praveen.patel@biourja.com, imam.khan@biourja.com, devina.ligga@biourja.com, karan.khilnani@biourja.com, ayushi.joshi@biourja.com"
# receiver_email = "praveen.patel@biourja.com"

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

def main():
    input_date = "03.14.2022"
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
        #logger.info('Opening Unsettled Receivables Workbook')
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
        #logger.info('Opening Unsettled Payables Workbook')
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
        #logger.info('Opening BB Master +-25K report sheet')
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
        #logger.info('Delete rows with value between -25K to 25K')
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
        #logger.info('Refreshing all tab')  
        BB_wb.api.RefreshAll()
        print()
        wb.save(output_cpr)
        BB_wb.save(output_cpr_copy)

        return f"CPR Reports for {input_date} is generated"
    except Exception as e:
        # #logger.exception(str(e))
        raise e
    finally:
        try:
            wb.app.quit()
        except:
            pass
        


if __name__ == "__main__": 
    logfilename = bu_alerts.add_file_logging(
        logger, process_name=job_name.lower())
    logger.info('Execution Started')
    
    #start the process
    starttime=time.time()
    logger.warning('Start work at {} ...'.format(str(starttime)))
    try:
        prev_month = main()
        bu_alerts.send_mail(receiver_email = receiver_email, mail_subject=f"JOB SUCCESS {job_name} {prev_month}", mail_body = f"Job completed successfully, Attached logs", attachment_location=logfilename)
        # driver.quit()
    except Exception as e:
        print("Exception caught during execution: ",e)
        logger.exception(f'Exception caught during execution: {e}')
        # logger.info("sending failure mail")
          
        bu_alerts.send_mail(receiver_email = receiver_email, mail_subject ='JOB FAILED - {}'.format(job_name), mail_body = 'Error in main() details {}'.format(str(e)),attachment_location = logfilename)

    time_end = time.time()
    logger.warning('It took {} seconds to run.'.format(time_end - starttime))
    print('It took {} seconds to run.'.format(time_end - starttime))


