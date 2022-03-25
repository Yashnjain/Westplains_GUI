import xlwings as xw
import xlwings.constants as win32c
import time
from datetime import date
import logging
import bu_alerts
import os
import datetime
from datetime import datetime, date
import xlwings.constants as win32c

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
def bbr_other_tabs(input_date, wb, input_ar, input_ctm):
    try:
        # input_date = "02.07.2022"
        # input_xl = r"J:\WEST PLAINS\REPORT\BBR Reports\Raw Files" +f"\\{input_date}_Borrowing Base Report.xlsx"
        # input_xl = r"C:\Users\Yashn.jain\Desktop\WEST PLAINS\REPORT\BBR Reports\Raw Files"+f"\\{input_date}_Borrowing Base Report.xlsx"
        # input_ar = r"J:\WEST PLAINS\REPORT\Open AR\Output files"+f"\\Open AR _{input_date} - Production.xlsx"
        # input_ar = r"C:\Users\Yashn.jain\Desktop\WEST PLAINS\REPORT\Open AR\Output files"+f"\\Open AR _{input_date} - Production.xlsx"
        # input_ctm = r"J:\WEST PLAINS\REPORT\CTM Combined report\Output files"+f"\\CTM Combined _{input_date}.xlsx"
        # input_ctm = r"C:\Users\Yashn.jain\Desktop\WEST PLAINS\REPORT\CTM Combined report\Output files"+f"\\CTM Combined _{input_date}.xlsx"
        # output_location=r"J:\WEST PLAINS\REPORT\BBR Reports\Output files"
        # output_location=r"C:\Users\Yashn.jain\Desktop\Sample_BBR"
        retry=0
        while retry < 10:
            try:
                ar_wb=xw.Book(input_ar)
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry ==10:
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
        #         time.sleep(5)
        #         retry+=1
        #         if retry ==10:
        #             raise e


        ws1=wb.sheets["AR-Trade By Tier - Eligible"]
        ws1.select()
        # pivotCount = wb.api.ActiveSheet.PivotTables().Count
        # #'\\Bio-India-FS\India sync$\WEST PLAINS\REPORT\BBR Reports\Raw Files\[Open AR _02.07.2022 - Production.xlsx]Eligible'!$A$1:$K$123
        # # 'Data 02.21.2022'!$A$1:$G$4731
        # #'\\Bio-India-FS\India sync$\WEST PLAINS\REPORT\BBR Reports\Raw Files\[Open AR _02.07.2022 - Production.xlsx]Eligible'!$A$1:$K$123
        # for j in range(1, pivotCount+1):
        #     wb.api.ActiveSheet.PivotTables("PivotTable1").PivotSelect("Tier[All]", win32c.PTSelectionMode.xlLabelOnly,True)
        #     # wb.api.ActiveSheet.PivotTables(j).PivotCache().SourceData = f"'J:\WEST PLAINS\REPORT\Open AR\Output files\[Open AR _{input_date} - Production]Eligible'!R1C1:R{last_row}C{total_column}"
        #     wb.api.ActiveSheet.PivotTables(j).PivotCache().Refresh()  

        ###logger.info("Adding Worksheet for Pivot Table")
        wb.sheets.add("AR-Trade By Tier - Eligible2",after=wb.sheets["Account Receivable Summary"])
        ###logger.info("Clearing contents for new sheet")
        wb.sheets["AR-Trade By Tier - Eligible2"].clear_contents()
        ws2=wb.sheets["AR-Trade By Tier - Eligible2"]
        ###logger.info("Declaring Variables for columns and rows")
        # last_row = ws5.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row
        # last_column = ws5.range('A1').end('right').last_cell.column
        # last_column_letter=num_to_col_letters(ws5.range('A1').end('right').last_cell.column)
        ###logger.info("Creating Pivot Table")
        PivotCache=wb.api.PivotCaches().Create(SourceType=win32c.PivotTableSourceType.xlDatabase, SourceData=f"'J:\\WEST PLAINS\\REPORT\\Open AR\\Output files\\[Open AR _{input_date} - Production.xlsx]Eligible'!R1C1:R{last_row}C{total_column}", Version=win32c.PivotTableVersionList.xlPivotTableVersion14)
        PivotTable = PivotCache.CreatePivotTable(TableDestination=f"'AR-Trade By Tier - Eligible2'!R7C1", TableName="PivotTable1", DefaultVersion=win32c.PivotTableVersionList.xlPivotTableVersion14)        ###logger.info("Adding particular Row in Pivot Table")
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

        ws1.api.Range("A1:A3").Copy()
        ws2.api.Paste()
        wb.app.api.CutCopyMode=False
        ws1.delete()
        ws2.name="AR-Trade By Tier - Eligible"
        ws3=wb.sheets["AR-Trade By Tier - Ineligible"]
        ws3.select()
        wb.sheets.add("AR-Trade By Tier - Ineligible2",after=wb.sheets["AR-Trade By Tier - Eligible"])
        ###logger.info("Clearing contents for new sheet")
        wb.sheets["AR-Trade By Tier - Ineligible2"].clear_contents()
        ws4=wb.sheets["AR-Trade By Tier - Ineligible2"]
        ###logger.info("Declaring Variables for columns and rows")
        # last_row = ws5.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row
        # last_column = ws5.range('A1').end('right').last_cell.column
        # last_column_letter=num_to_col_letters(ws5.range('A1').end('right').last_cell.column)
        ###logger.info("Creating Pivot Table")
        PivotCache=wb.api.PivotCaches().Create(SourceType=win32c.PivotTableSourceType.xlDatabase, SourceData=f"'J:\\WEST PLAINS\\REPORT\\Open AR\\Output files\\[Open AR _{input_date} - Production.xlsx]Eligible'!R1C1:R{last_row}C{total_column}", Version=win32c.PivotTableVersionList.xlPivotTableVersion14)
        PivotTable = PivotCache.CreatePivotTable(TableDestination=f"'AR-Trade By Tier - Ineligible2'!R7C1", TableName="PivotTable1", DefaultVersion=win32c.PivotTableVersionList.xlPivotTableVersion14)        ###logger.info("Adding particular Row in Pivot Table")
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

        ws3.api.Range("A1:A3").Copy()
        ws4.api.Paste()
        wb.app.api.CutCopyMode=False
        ws3.delete()
        ws4.name="AR-Trade By Tier - Ineligible"
        # ws5=wb.sheets['Detail CTM Non MCUI']

        retry=0
        while retry < 10:
            try:
                wb1=xw.Book(input_ctm)
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry ==10:
                    raise e 

        excl_sht = wb1.sheets("Excl Macq & IC")
        ##logger.info("Copy tier sheet AFTER the intercompany sheet of input book.")
        num_row = excl_sht.range('A1').end('down').row
        last_column = excl_sht.range('A1').end('right').last_cell.column
        last_column_letter=num_to_col_letters(last_column)
        excl_sht.range(f'A1:{last_column_letter}{num_row}').copy()
        wb.activate()
        ws5 = wb.sheets['Detail CTM Non MCUI']
        ws5.range('A1').paste()
        wb.app.api.CutCopyMode=False
        wb1.activate()
        wb1.close()
        wb.activate()
        ws6 = wb.sheets['Unrlz- Gains- Contracts Non MC']
        wb.sheets['Unrlz- Gains- Contracts Non MC'].select()

        #logger.info("Adding Worksheet for Pivot Table")
        wb.sheets.add("Unrlz- Gains- Contracts Non MC2",after=wb.sheets["Inventory -Other"])
        #logger.info("Clearing New Worksheet")
        wb.sheets["Unrlz- Gains- Contracts Non MC2"].clear_contents()
        ws7=wb.sheets["Unrlz- Gains- Contracts Non MC2"]
        #logger.info("Declaring Variables for columns and rows")
        last_column = ws5.range('A1').end('right').last_cell.column
        last_column_letter=num_to_col_letters(ws5.range('A1').end('right').last_cell.column)
        num_row = ws5.range('A1').end('down').row
        #logger.info("Creating Pivot table")
        PivotCache=wb.api.PivotCaches().Create(SourceType=win32c.PivotTableSourceType.xlDatabase, SourceData=f"\'Detail CTM Non MCUI\'!R1C1:R{num_row}C{last_column}", Version=win32c.PivotTableVersionList.xlPivotTableVersion14)
        PivotTable = PivotCache.CreatePivotTable(TableDestination="'Unrlz- Gains- Contracts Non MC2'!R7C1", TableName="PivotTable1", DefaultVersion=win32c.PivotTableVersionList.xlPivotTableVersion14)
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
        last_row2 = ws7.range(f'A'+ str(ws7.cells.last_cell.row)).end('up').row
        last_row2+=10
        #logger.info("Creating Pivot table")
        PivotCache=wb.api.PivotCaches().Create(SourceType=win32c.PivotTableSourceType.xlDatabase, SourceData=f"\'Detail CTM Non MCUI\'!R1C1:R{num_row}C{last_column}", Version=win32c.PivotTableVersionList.xlPivotTableVersion14)
        PivotTable = PivotCache.CreatePivotTable(TableDestination=f"'Unrlz- Gains- Contracts Non MC2'!R{last_row2}C1", TableName="PivotTable2", DefaultVersion=win32c.PivotTableVersionList.xlPivotTableVersion14)
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
        PivotTable.PivotFields('Ship Tier').CurrentPage = ">12 months"
        #logger.info("Changing No Format in Pivot Table")
        #logger.info("Changing Table layout")
        PivotTable.PivotFields('Location Id').Subtotals=(False, False, False, False, False, False, False, False, False, False, False, False)
        PivotTable.PivotFields('Commodity Id').Subtotals=(False, False, False, False, False, False, False, False, False, False, False, False)
        PivotTable.RowAxisLayout(1)
        #logger.info("Changing Table Style")
        PivotTable.TableStyle2 = ""
        wb.api.ActiveSheet.PivotTables("PivotTable2").InGridDropZones = True
        #logic for adding total
        last_row3 = ws7.range(f'A'+ str(ws7.cells.last_cell.row)).end('up').row 
        last_row3+=5
        ws7.range(f"E{last_row3}").value=f'=+GETPIVOTDATA("Gain/LossTotal",$A$7)+GETPIVOTDATA("Gain/LossTotal",$A${last_row2})'
        ws7.range(f"E{last_row3}").api.NumberFormat= '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'

        ws6.api.Range("A1:A3").Copy()
        ws7.api.Paste()
        ws7.api.Columns("C:C").ColumnWidth = 17
        wb.app.api.CutCopyMode=False
        ws6.delete()
        ws7.name="Unrlz- Gains- Contracts Non MC"
        # wb.save(f"{output_location}\\{input_date}_Borrowing Base Report_y.xlsx")
        # wb.app.quit()
    except Exception as e:
            raise e
    # finally:
    #         try:
    #             wb.app.quit()
    #         except:
    #             pass

main()