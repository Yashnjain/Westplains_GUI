import xlwings as xw
import xlwings.constants as win32c
import time
from datetime import date
import logging
import bu_alerts
import os
import datetime
from datetime import datetime, date

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

today_date=date.today()
# log progress --
logfile = os.getcwd() +"\\logs\\"+'OPENAR_Logfile'+str(today_date)+'.txt'

logging.basicConfig(filename=logfile, filemode='w',
                    format='%(asctime)s %(message)s')
logging.basicConfig(
    level=logging.INFO, 
    format='%(asctime)s [%(levelname)s] - %(message)s',
    filename=logfile)
logger = logging.getLogger()
logger.setLevel(logging.INFO)

input_date="02.21.2022"
output_date = "02.14.2022"
# input_sheet=os.getcwd()+f'\\Open AR _{input_date} - Production.xlsx' 
input_sheet = r'J:\WEST PLAINS\REPORT\Open AR\Raw files'+f'\\Open AR _{input_date} - Production.xlsx' 
# latest_output_date="01.31.2022"
# latest_output=r'C:\Users\Yashn.jain\Desktop\AR Referrence'+f'\\Open AR _{output_date} - Production.xlsx'
prev_output=r'J:\WEST PLAINS\REPORT\Open AR\Output files'+f'\\Open AR _{output_date} - Production.xlsx'
job_name = "Open_AR_Automation"
output_location = r'J:\WEST PLAINS\REPORT\Open AR\Output files'
receiver_email = "imam.khan@biourja.com, yashn.jain@biourja.com, devina.ligga@biourja.com, karan.khilnani@biourja.com, ayushi.joshi@biourja.com"
def main():
    try:
        if not os.path.exists(input_sheet):
            return(f"{input_sheet} Excel file not present for date {input_date}")
        if not os.path.exists(prev_output):
            return(f"{prev_output} Excel file not present for date {output_date}")    
        prev_month = datetime.strftime(datetime.strptime(input_date, "%m.%d.%Y"), "%B")
        logger.info("Opening operating workbook instance of excel")
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
        logger.info("Adding sheet to the same workbook")
        wb.sheets.add("Excl Macq & IC",after=wb.sheets[f"Open AR _{input_date} - Productio"]) 
        ws2=wb.sheets["Excl Macq & IC"]
        logger.info("Clearing its contents")
        ws2.cells.clear_contents()
        logger.info("Accessing Particular WorkBook[0]")
        ws1=wb.sheets[f"Open AR _{input_date} - Productio"]

        logger.info("Declaring Variables for columns and rows")
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


        logger.info("Applying Filter to the same workbook")
        ws1.api.Range(f"{Customer_letter_column}1").AutoFilter(Field:=f'{Customer_no_column}', Criteria1:=["<>MACQUARIE COMMODITIES (USA) INC."], Operator:=1,Criteria2=["<>INTER-COMPANY PURCH/SALES"])
        ws1.api.Range(f"{Location_letter_column}1").AutoFilter(Field:=f'{Location_no_column}', Criteria1:=["<>WPMEXICO"], Operator:=1)
        ws1.api.Range(f"{Total_AR_letter_column}1").AutoFilter(Field:=f'{Total_AR_no_column}', Criteria1:="<>0", Operator:=1)
        logger.info("Copying and pasting Worksheet")
        ws1.api.AutoFilter.Range.Copy()
        ws2.api.Paste()
        logger.info("Applying Autofit")
        ws2.autofit()

        logger.info("Declaring Variables for columns and rows")
        column_list = ws1.range("A1").expand('right').value
        Customer_column = num_to_col_letters(column_list.index('Customer')+1)
        Customer_column_num = column_list.index('Customer')+1

        logger.info("Copying Inter Company Data from inp sheet  to Intercompany Sheet")
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
            # logger.info("Deleting intercompant data from original sheet after copying it in previous code")
            # ws1.api.AutoFilterMode=False
            # time.sleep(1)
            # logger.info("Declaring Variables for columns and rows")
            # last_row = ws1.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row
            # # last_row+=1
            # column_list = ws1.range("A1").expand('right').value
            # Customer_no_column=column_list.index('Customer')+1
            # Customer_letter_column = num_to_col_letters(column_list.index('Customer')+1)
            # logger.info("Applying loop for deleting INTER-COMPANY PURCH/SALES")
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
        logger.info("Copying tier column from previous output sheet")   
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
        logger.info("Copy tier sheet AFTER the intercompany sheet of input book.")
        tier_sht.api.Copy(None, After=ws2.api)
        tier_wb.close()
        logger.info("Declaring Variables for columns and rows")
        ws3=wb.sheets['Tier']
        last_row = ws2.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row
        column_list = ws2.range("A1").expand('right').value
        Customer_no_column=column_list.index('Customer')+1
        Customer_letter_column = num_to_col_letters(column_list.index('Customer')+1)
        Customer_data = ws2.range(f"{Customer_letter_column}2").expand('down').value
        mylist = list(dict.fromkeys(Customer_data))
        logger.info("Declaring Variables for columns and rows")
        last_row = ws3.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row
        column_list = ws3.range("A1").expand('right').value
        Customer_Name_no_column=column_list.index('Customer Name')+1
        Customer_Name_letter_column = num_to_col_letters(column_list.index('Customer Name')+1)
        Customer_Name_data = ws3.range(f"{Customer_Name_letter_column}2").expand('down').value
        logger.info("Declaring Variables for columns and rows")
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
        logger.info("Declaring Variables for columns and rows")
        last_row = ws2.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row
        column_list = ws2.range("A1").expand('right').value
        Credit_Limit_no_column=column_list.index('Credit Limit')+1
        Credit_Limit_letter_column = num_to_col_letters(column_list.index('Credit Limit')+1)
        logger.info("Adding Tier Coloumn and inserting value and dragging them")
        ws2.api.Range(f"{Credit_Limit_letter_column}1").EntireColumn.Insert()
        ws2.range(f"{Credit_Limit_letter_column}1").value = "Tier"
        ws2.range(f"{Credit_Limit_letter_column}2").value ="=VLOOKUP(H2,Tier!A:B,2,0)"
        ws2.range(f"{Credit_Limit_letter_column}2").copy(ws2.range(f"{Credit_Limit_letter_column}2:{Credit_Limit_letter_column}{last_row}"))

        logger.info("Adding Worksheet for Pivot Table")
        wb.sheets.add("Pivot Summary",after=wb.sheets["Tier"])
        logger.info("Clearing New Worksheet")
        wb.sheets["Pivot Summary"].clear_contents()
        # ws4=wb.sheets["Pivot Summary"]
        logger.info("Declaring Variables for columns and rows")
        last_row = ws2.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row
        last_column = ws2.range('A1').end('right').last_cell.column
        last_column_letter=num_to_col_letters(ws2.range('A1').end('right').last_cell.column)
        logger.info("Creating Pivot table")
        PivotCache=wb.api.PivotCaches().Create(SourceType=win32c.PivotTableSourceType.xlDatabase, SourceData=f"\'Excl Macq & IC\'!R1C1:R{last_row}C{last_column}", Version=win32c.PivotTableVersionList.xlPivotTableVersion14)
        PivotTable = PivotCache.CreatePivotTable(TableDestination="'Pivot Summary'!R1C1", TableName="PivotTable1", DefaultVersion=win32c.PivotTableVersionList.xlPivotTableVersion14)
        logger.info("Adding particular Row Data in Pivot Table")
        PivotTable.PivotFields('Tier').Orientation = win32c.PivotFieldOrientation.xlRowField
        PivotTable.PivotFields('Tier').Position = 1
        PivotTable.PivotFields('Tier').RepeatLabels=True
        PivotTable.PivotFields('Customer').Orientation = win32c.PivotFieldOrientation.xlRowField
        logger.info("Adding particular Data Field in Pivot Table")
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
        logger.info("Changing No Format in Pivot Table")
        logger.info("Changing Table layout")
        PivotTable.RowAxisLayout(1)
        logger.info("Changing Table Style")
        PivotTable.TableStyle2 = ""

        # PivotTable.TableStyle2 = ""
        logger.info("Removing subtotal from Tier")
        PivotTable.PivotFields('Tier').Subtotals=(False, False, False, False, False, False, False, False, False, False, False, False)
        ws4=wb.sheets["Pivot Summary"]
        logger.info("Adding Worksheet Eligible")
        wb.sheets.add("Eligible",after=wb.sheets["Pivot Summary"])
        ws5=wb.sheets["Eligible"]
        logger.info("Declaring Variables for columns and rows and sheets")
        ws4=wb.sheets["Pivot Summary"]
        last_row = ws4.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row
        final=last_row-1
        last_column = ws4.range('A1').end('right').last_cell.column
        last_column_letter=num_to_col_letters(ws4.range('A1').end('right').last_cell.column)
        logger.info("Copying and pasting sheet to Eligible Worksheet")
        ws4.api.Range(f'A1:{last_column_letter}{final}').Copy()
        ws5.api.Paste()
        ws5.autofit()
        logger.info("Changing names of columns in new sheet")
        column_list = ws5.range("A1").expand('right').value
        changed_column_list=['Tier', 'Customer Name', 'Current', ' 1 - 10', ' 11 - 30', ' 31 - 60', ' 61 - 9999', 'Balance']
        i=0
        for values in column_list:
            values_column_no=column_list.index(values)+1
            values_letter_column = num_to_col_letters(column_list.index(values)+1)
            ws5.range(f"{values_letter_column}1").value = changed_column_list[i]
            ws3.range(f"{Tier_letter_column}{last_row_value}").font.name = 'Calibri'
            i+=1
        logger.info("Inserting extra Culumns,adding their values and dragging them")
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
        logger.info("Applying same previous operation for extra hidden columns")
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
        logger.info("CHECK FOR CASH CUSTOMER AND MAKING HIM INELIGIBLE")
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
        logger.info("Paste Special Values For Values In c & d")
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
        logger.info("Declaring Variables for columns and rows")
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
        logger.info("Starting loop for C column adjustment")
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
        logger.info("Adding variables")    
        last_row = ws5.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row
        last_row+=1
        logger.info("Starting loop for D column adjustment")
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
        logger.info("Adding Worksheet for Pivot Table")
        wb.sheets.add("Pivot BB",after=wb.sheets["Eligible"])
        logger.info("Clearing contents for new sheet")
        wb.sheets["Pivot BB"].clear_contents()
        ws6=wb.sheets["Pivot BB"]
        logger.info("Declaring Variables for columns and rows")
        last_row = ws5.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row
        last_column = ws5.range('A1').end('right').last_cell.column
        last_column_letter=num_to_col_letters(ws5.range('A1').end('right').last_cell.column)
        logger.info("Creating Pivot Table")
        PivotCache=wb.api.PivotCaches().Create(SourceType=win32c.PivotTableSourceType.xlDatabase, SourceData=f"\'Eligible\'!R1C1:R{last_row}C{last_column}", Version=win32c.PivotTableVersionList.xlPivotTableVersion14)
        PivotTable = PivotCache.CreatePivotTable(TableDestination="'Pivot BB'!R3C1", TableName="PivotTable1", DefaultVersion=win32c.PivotTableVersionList.xlPivotTableVersion14)
        logger.info("Adding particular Row in Pivot Table")

        PivotTable.PivotFields('Tier').Orientation = win32c.PivotFieldOrientation.xlRowField
        PivotTable.PivotFields('Tier').Position = 1
        PivotTable.PivotFields('Customer Name').Orientation = win32c.PivotFieldOrientation.xlRowField
        logger.info("Adding particular Data Field in Pivot Table")
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
        logger.info("Adding particular Page Field in Pivot Table")
        PivotTable.PivotFields('Eligiblity').Orientation = win32c.PivotFieldOrientation.xlPageField
        logger.info("Applying filter in Data Field in Pivot Table")
        PivotTable.PivotFields('Eligiblity').CurrentPage = "Eligible"
        logger.info("Changing No Format in Pivot Table")
        # PivotTable.RowAxisLayout(1)
        logger.info("Changing Table Style in Pivot Table")
        PivotTable.TableStyle2 = ""
        logger.info("Changing Table Layout in Pivot Table")
        PivotTable.RowAxisLayout(1)
        logger.info("Declaring Variables for columns and rows")
        last_row = ws5.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row
        last_column = ws5.range('A1').end('right').last_cell.column
        last_column_letter=num_to_col_letters(ws5.range('A1').end('right').last_cell.column)
        last_row2 = ws6.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row
        last_row2+=10
        logger.info("Creating Pivot Table")
        PivotCache=wb.api.PivotCaches().Create(SourceType=win32c.PivotTableSourceType.xlDatabase, SourceData=f"\'Eligible\'!R1C1:R{last_row}C{last_column}", Version=win32c.PivotTableVersionList.xlPivotTableVersion14)
        PivotTable = PivotCache.CreatePivotTable(TableDestination=f"'Pivot BB'!R{last_row2}C1", TableName="PivotTable2", DefaultVersion=win32c.PivotTableVersionList.xlPivotTableVersion14)
        logger.info("Adding particular data in RowField in Pivot Table")

        PivotTable.PivotFields('Tier').Orientation = win32c.PivotFieldOrientation.xlRowField
        PivotTable.PivotFields('Tier').Position = 1
        PivotTable.PivotFields('Customer Name').Orientation = win32c.PivotFieldOrientation.xlRowField
        logger.info("Adding particular Data Field in Pivot Table")
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
        logger.info("Adding particular Data Field in Pivot Table")
        PivotTable.PivotFields('Eligiblity').Orientation = win32c.PivotFieldOrientation.xlPageField
        logger.info("Applying filter in pagefield in Pivot Table")
        PivotTable.PivotFields('Eligiblity').CurrentPage = "Ineligible"
        logger.info("Changing No Format in Pivot Table")
        # PivotTable.RowAxisLayout(1)
        logger.info("Changing Table Style in Pivot Table")
        PivotTable.TableStyle2 = ""
        logger.info("Changing Layout for Pivot Table")
        PivotTable.RowAxisLayout(1)
        logger.info("Doing final adjustments for Sheets")
        ws6.autofit()
        wb.app.api.CutCopyMode=False
        wb.app.api.Autofilter=False
        wb.app.api.AutofilterMode=False

        last_col_num = ws1.range('A1').expand('right').last_cell.column 
        # last_col = num_to_col_letters(last_col_num) 
        last_row = ws2.range(f'A'+ str(ws2.cells.last_cell.row)).end('up').row 
        ###logger.info("Adding Worksheet for Pivot Table") 
        wb.sheets.add("For allocation entry",before=ws1) 
        ###logger.info("Creating Pivot table") 
        PivotCache=wb.api.PivotCaches().Create(SourceType=win32c.PivotTableSourceType.xlDatabase, SourceData=f'\'{ws1.name}\'!R1C1:R{last_row}C{last_col_num}', Version=win32c.PivotTableVersionList.xlPivotTableVersion14) 
        PivotTable = PivotCache.CreatePivotTable(TableDestination="'For allocation entry'!R3C1", TableName="PivotTable1", DefaultVersion=win32c.PivotTableVersionList.xlPivotTableVersion14)
         ###logger.info("Adding particular Row in Pivot Table") 
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
        wb.save(f"{output_location}\\Open AR _"+input_date+' - Production.xlsx')
        wb.app.quit()
        return prev_month
    except Exception as e:
        raise e
    finally:
        try:
            wb.app.quit()
        except:
            pass

if __name__ == "__main__":
    logfilename = bu_alerts.add_file_logging(logger, process_name=job_name.lower())
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

        # logger.info("mail sent")
    
    time_end = time.time()
    logger.warning('It took {} seconds to run.'.format(time_end - starttime))
    print('It took {} seconds to run.'.format(time_end - starttime))

#Logic 2
# last_row = ws5.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row    
# ws5.range(f"P2:P{last_row}").copy(ws5.range(f"E2:E{last_row}"))
# ws5.range(f"F2:F{last_row}").value=0
# ws5.range(f"G2:G{last_row}").value=0
#end










