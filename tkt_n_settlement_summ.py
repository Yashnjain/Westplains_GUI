import os, time
from sqlalchemy import table
import xlwings as xw
from datetime import datetime
import pandas as pd

from payroll_summ import num_to_col_letters

def set_borders(border_range):
    for border_id in range(7,13):
        border_range.api.Borders(border_id).LineStyle=1
        border_range.api.Borders(border_id).Weight=2

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






def tkt_n_settlement_summ(input_date, output_date):
    try:
        monthYear = datetime.strftime(datetime.strptime(input_date, "%m.%d.%Y"), "%b %Y")
        Year = datetime.strftime(datetime.strptime(input_date, "%m.%d.%Y"), "%Y")

        tkt_query_xl = r"J:\WEST PLAINS\REPORT\Ticket And Settlement Summary\Raw Files" +f"\\Ticket Query {Year}.xlsx"
        # input_xl = r"C:\Users\imam.khan\OneDrive - BioUrja Trading LLC\Documents\WEST PLAINS\REPORT\Macquaire Accrual Entry\Raw Files" +f"\\Macq Accrual_{input_date}.xlsx"
        if not os.path.exists(tkt_query_xl):
            return(f"{tkt_query_xl} Excel file not present for year {Year}")

        settlement_xl = r"J:\WEST PLAINS\REPORT\Ticket And Settlement Summary\Raw Files\SETTLEMENT MAKER.xlsx"
        if not os.path.exists(settlement_xl):
            return(f"{settlement_xl} Excel file not present")


        template_xl = r"J:\WEST PLAINS\REPORT\Ticket And Settlement Summary\Raw Files\Ticket Query monYearTemplate.xlsx"
        if not os.path.exists(template_xl):
            return(f"{template_xl} Excel file not present")

        output_file =  r"J:\WEST PLAINS\REPORT\Ticket And Settlement Summary\Output Files"+f"\\Tickets and Settlement {monthYear}"
        det_output_file = r"J:\WEST PLAINS\REPORT\Ticket And Settlement Summary\Output Files"+f"\\Ticket Query {monthYear} Details"


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
                tkt_sht = tkt_wb.sheets["2021"]
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
        tkt_sht.range(f"A1:M{last_row-1}").copy()
        tkt_ent_sht.range("A1").paste(paste="values_and_number_formats") #pasting only values
        tkt_last_row = tkt_ent_sht.range(f'A'+ str(tkt_ent_sht.cells.last_cell.row)).end('up').row
        
        #adding Add by column by copy pasting add_by column already present in column K
        i=0
        while tkt_ent_sht.range(chr(ord("M")-i)+"1").value != "add_by":
            i+=1
        add_by_col = chr(ord("M")-i)
        tkt_ent_sht.range("N1").value = "Add By"
        tkt_ent_sht.range(f"{add_by_col}2").expand("down").copy(tkt_ent_sht.range("N2"))
        tkt_wb.close()

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
                set_sht = set_wb.sheets["Sheet1"]
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

        #Refreshing Pivots
        while retry < 10:
            try:
                tkt_p_sht = wb.sheets["Ticket Summary (2)"]
                break
            except Exception as e:
                time.sleep(2)
                retry+=1
                if retry ==9:
                    raise e
        tkt_p_sht.activate()

        pivotCount = wb.api.ActiveSheet.PivotTables().Count
            
        for j in range(1, pivotCount+1):
            wb.api.ActiveSheet.PivotTables(j).PivotCache().SourceData = f"'{tkt_ent_sht.name}'!R1C1:R{tkt_last_row}C14" #14 for N col
            wb.api.ActiveSheet.PivotTables(j).PivotCache().Refresh()

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
            wb.api.ActiveSheet.PivotTables(j).PivotCache().SourceData = f"'{inp_set_sht.name}'!R1C1:R{inp_set_last_row}C12" #12 for L col
            wb.api.ActiveSheet.PivotTables(j).PivotCache().Refresh()

        #Combining data for summary tab
        left_df = tkt_p_sht.range('A2:B2').options(pd.DataFrame, 
                                header=1,
                                index=False, 
                                expand='down').value[:-1]
        left_df.columns = ["Row Labels", "Tickets"]
        right_df = set_p_sht.range('A1:B1').options(pd.DataFrame, 
                                header=1,
                                index=False, 
                                expand='down').value[:-1]
        right_df.columns = ["Row Labels", "Settlements"]

        merged_df = left_df.merge(right_df, on='Row Labels', how='outer')

        #inserting merged data in sheet 1
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
            wb.api.ActiveSheet.PivotTables(j).PivotCache().SourceData = f"'Sheet1'!R1C1:R{len(merged_df)}C3" #3 for C col
            wb.api.ActiveSheet.PivotTables(j).PivotCache().Refresh()

    
        

        new_wb = xw.Book()
        time.sleep(1)
        new_wb.sheets[0].name = "Ticket Summary"
        tkt_p_sht.activate()
        # tkt_p_sht.api.Range(tkt_p_sht.api.Cells.SpecialCells(12).Address).Copy()
        # new_wb.activate()
        # new_wb.sheets['Ticket Summary'].activate()
        # new_wb.sheets['Ticket Summary'].api.Range("A1").Select()
        # new_wb.sheets['Ticket Summary'].range("A1").api.Range("A1").PasteSpecial(Paste=	-4163)    #xlPasteValues
        # new_wb.sheets['Ticket Summary'].autofit(axis="columns")

        #Generating data for final file
        #Ticket Summary Data sheets
        tkt_df1 = tkt_p_sht.range('A2').options(pd.DataFrame, 
                                header=1,
                                index=False, 
                                expand='table').value
        tkt_df1.columns = ["Add By", "Total"]

        next_table_row = tkt_p_sht.range("A2").end("down").end("down").row+1
        last_row = tkt_p_sht.range(f'A'+ str(tkt_p_sht.cells.last_cell.row)).end('up').row
        tkt_df2 = tkt_p_sht.range(f"A{next_table_row}:A{last_row}").options(pd.DataFrame, 
                                header=1,
                                index=False, 
                                expand='right').value
        table3_col = num_to_col_letters(tkt_p_sht.range("B2").end("right").column)
        last_row = tkt_p_sht.range(f'F'+ str(tkt_p_sht.cells.last_cell.row)).end('up').row
        tkt_df3 = tkt_p_sht.range(f"{table3_col}2:{table3_col}{last_row}").options(pd.DataFrame, 
                                header=1,
                                index=False, 
                                expand='right').value
        
        #pasting data in final workbook Ticket Summary Sheet
        
        new_wb.sheets['Ticket Summary'].range("A1").options(pd.DataFrame, header=1, index=False, expand='table').value = tkt_df1
        new_wb.sheets['Ticket Summary'].range(f"A{len(tkt_df1)+6}").options(pd.DataFrame, header=1, index=False, expand='table').value = tkt_df2
        new_wb.sheets['Ticket Summary'].range(f"{table3_col}1").options(pd.DataFrame, header=1, index=False, expand='table').value = tkt_df3
        new_wb.sheets['Ticket Summary'].autofit(axis="columns")

        #setting Borders for 1st table
        border_range = new_wb.sheets['Settlement Summary'].range("A1").expand("table")
        set_borders(border_range)

        #setting Borders for 2nd table
        border_range = new_wb.sheets['Settlement Summary'].range(f"A{len(tkt_df1)+6}").expand("table")
        set_borders(border_range)

        #setting Borders for 3rd table
        border_range = new_wb.sheets['Settlement Summary'].range(f"{table3_col}1").expand("table")
        set_borders(border_range)




        #Settlement Summary Data sheets
        new_wb.sheets.add("Settlement Summary",after=new_wb.sheets[f"Ticket Summary"]) 



        set_df1 = set_p_sht.range('A2').options(pd.DataFrame, 
                                header=False,
                                index=False,
                                expand='table').value
        

        
        table2_col = num_to_col_letters(set_p_sht.range("B2").end("right").column)
        last_row = set_p_sht.range(chr(ord(table2_col)+1)+ str(set_p_sht.cells.last_cell.row)).end('up').row
        set_df2 = set_p_sht.range(f"{table2_col}2:{table2_col}{last_row}").options(pd.DataFrame, 
                                header=False,
                                index=False, 
                                expand='right').value
        
        #pasting data in final workbook Settlement Summary Sheet
        
        new_wb.sheets['Settlement Summary'].range("A2").options(pd.DataFrame, header=False, index=False, expand='table').value = set_df1
        new_wb.sheets['Settlement Summary'].range(f"A{len(set_df1)+5}").options(pd.DataFrame, header=False, index=False, expand='table').value = set_df2
        
        new_wb.sheets['Settlement Summary'].autofit(axis="columns")

        #setting Borders for 1st table
        border_range = new_wb.sheets['Settlement Summary'].range("A2").expand("table")
        set_borders(border_range)

        #setting Borders for 2nd table
        border_range = new_wb.sheets['Settlement Summary'].range(f"A{len(set_df1)+5}").expand("table")
        set_borders(border_range)
        

        #Settlement Summary Data sheets
        new_wb.sheets.add("Consolidated Summary",after=new_wb.sheets[f"Settlement Summary"]) 



        summ_df = summ_p_sht.range('A3').options(pd.DataFrame, 
                                header=1,
                                index=False,
                                expand='table').value
        summ_df.columns = ['User', 'Sum of Tickets', 'Sum of Settlements']

        new_wb.sheets['Consolidated Summary'].range("A1").options(pd.DataFrame, header=1, index=False, expand='table').value = summ_df
        new_wb.sheets['Consolidated Summary'].autofit(axis="columns")
        #setting Borders
        border_range = new_wb.sheets['Consolidated Summary'].range("A1").expand("table")
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




    
    






input_date = '12.31.2021'
msg = tkt_n_settlement_summ(input_date, output_date=None)
print(msg)
print()