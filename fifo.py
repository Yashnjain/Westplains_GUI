
import xlwings.constants as win32c
from datetime import datetime
import time
import glob, os
import xlwings as xw
import xlwings.constants as win32c
import pandas as pd
from tabula import read_pdf

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


def other_loc_extractor(input_pdf):
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
            if location == "NGREEL":
                location = "NORTH GREELEY"
            if location == "BROWNSVILL":
                location = "BROWNSVILLE"
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
    # for row in df.iterrows():
    #     print(row)

    print()

def fifo(input_date, output_date):
    location = ["HRW", "YC"]
    inp_date = datetime.strftime(datetime.strptime(input_date, "%m.%d.%Y"), "%m.%d.%y")
    monthYear = datetime.strftime(datetime.strptime(input_date, "%m.%d.%Y"), "%B %Y")
    for loc in location:
        input_xl = r"J:\WEST PLAINS\REPORT\FIFO reports\Raw Files" +f"\\Inventory on site {loc}_{inp_date}.xlsx"
        # input_xl = r"C:\Users\imam.khan\OneDrive - BioUrja Trading LLC\Documents\WEST PLAINS\REPORT\FIFO reports\Raw Files" +f"\\Inventory on site {loc}_{inp_date}.xlsx"
        if not os.path.exists(input_xl):
                return(f"{input_xl} Excel file not present for date {input_date}")
        
        input_mtm = r"J:\WEST PLAINS\REPORT\MOC Interest allocation\Raw Files" +f"\\Inventory MTM Excel Report {monthYear}.xlsx"
        # input_mtm = r"C:\Users\imam.khan\OneDrive - BioUrja Trading LLC\Documents\WEST PLAINS\REPORT\MOC Interest allocation\Raw files" +f"\\Inventory MTM Excel Report {monthYear}.xlsx"
        if not os.path.exists(input_mtm):
                return(f"{input_mtm} Excel file not present for date {input_date}")

        input_mapping = r"J:\WEST PLAINS\REPORT\FIFO reports" +f"\\Sub_Loc_Mapping.xlsx"
        # input_mapping = r"C:\Users\imam.khan\OneDrive - BioUrja Trading LLC\Documents\WEST PLAINS\REPORT\FIFO reports" +f"\\Sub_Loc_Mapping.xlsx"
        if not os.path.exists(input_mapping):
                return(f"{input_mapping} Excel file not present for date")
        
        input_pdf = r"J:\WEST PLAINS\REPORT\FIFO reports\Raw Files" +f"\\Inventory Trial Balance_{inp_date}.pdf"
        # input_mtm = r"C:\Users\imam.khan\OneDrive - BioUrja Trading LLC\Documents\WEST PLAINS\REPORT\MOC Interest allocation\Raw files" +f"\\Inventory MTM Excel Report {monthYear}.xlsx"
        if not os.path.exists(input_pdf):
                return(f"{input_pdf} Excel file not present for date {input_date}")

        # input_yc = r"J:\WEST PLAINS\REPORT\FIFO reports\Raw Files" +f"\\Inventory on site YC_{inp_date}.xlsx"
        # if not os.path.exists(input_yc):
        #         return(f"{input_yc} Excel file not present for date {input_date}")

        ouput_loc = r"J:\WEST PLAINS\REPORT\FIFO reports\Output files" +f"\\Inventory on site {loc}_{inp_date}.xlsx"
        # output_loc = r"C:\Users\imam.khan\OneDrive - BioUrja Trading LLC\Documents\WEST PLAINS\REPORT\FIFO reports\Output files" +f"\\Inventory on site {loc}_{inp_date}.xlsx"
        # ouput_yc = r"J:\WEST PLAINS\REPORT\FIFO reports\Output files" +f"\\Inventory on site YC_{inp_date}.xlsx"
        mtm_ouput_loc = r"J:\WEST PLAINS\REPORT\FIFO reports\Output files" +f"\\Inventory MTM Excel Report {monthYear}.xlsx"

        

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
        inp_sht.range("1:1").insert()
        inp_sht.api.AutoFilterMode=False

        last_column = num_to_col_letters(inp_sht.range("A2").end('right').column)
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
        inp_sht.range(f"{qty_col}:{qty_col}").insert()
        inp_sht.range(f"{qty_col}2").value = "Qty"
        inp_sht.range(f"{value_col}:{value_col}").insert()
        inp_sht.range(f"{value_col}2").value = "Value"
        inp_sht.range(f"{price_col}:{price_col}").insert()
        inp_sht.range(f"{price_col}2").value = "Price"

        #Filtering Inter-Company and putting their quantiy to 0
        inp_sht.select()
        inp_sht.api.AutoFilterMode=False
        inp_sht.api.Range(f"{cust_name_col}2").AutoFilter(Field:=cust_name_col_num,Criteria1:="INTER-COMPANY PURCH/SALES")
        l_row = inp_sht.range(f"{quantityCol}2").end('down').row
        # rng = inp_sht.api.Range(f"{quantityCol}2:{quantityCol}{l_row}").SpecialCells(12).Address.split(",")[1:]
        inp_sht.api.Range(f"{quantityCol}3:{quantityCol}{l_row}").SpecialCells(12).Select()
        wb.app.selection.value = 0
        # inp_sht.range(f"{quantityCol}2").value = 0
        # inp_sht.range(f"{quantityCol}2").copy()
        # inp_sht.range(f"{quantityCol}2").expand("down").select()
        # wb.app.selection.paste()
        
        # inp_sht.range(f"{quantityCol}2").value = "Quantity"

        inp_sht.api.Range(f"{unit_cost_col}2").AutoFilter(Field:=unit_cost_col_num,Criteria1:=">=1", Operator:=2,Criteria2:="0")
        l_row = inp_sht.range(f"{invValCol}2").end('down').row
        inp_sht.api.Range(f"{invValCol}3:{invValCol}{l_row}").SpecialCells(12).Select()
        wb.app.selection.value = 0
        # wb.app.selection.value = 0
        # inp_sht.range(f"{invValCol}2").value = 0
        # inp_sht.range(f"{invValCol}").copy()
        # inp_sht.range(f"{invValCol}2").expand("down").select()
        # wb.app.selection.paste()
        
        # inp_sht.range(f"{invValCol}2").value = "Inventory Value"
        
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
                    mtm_sht = mtm_wb.sheets["MTM Excel Summary"]
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
            if loc == "HRW":
                sub_loc = mtm_sht.range("B116:B119").value
                for s_loc in range(len(sub_loc)):
                    if sub_loc[s_loc] == columns_2[key]:
                        new_sht.range("O1").value = mtm_sht.range(f"E{s_loc+116}").value
                        new_sht.range("P1").value = mtm_sht.range(f"F{s_loc+116}").value
            else:
                sub_loc = mtm_sht.range("B121:B124").value
                for s_loc in range(len(sub_loc)):
                    if sub_loc[s_loc] == columns_2[key]:
                        new_sht.range("O1").value = mtm_sht.range(f"E{s_loc+121}").value
                        new_sht.range("P1").value = mtm_sht.range(f"F{s_loc+121}").value
            
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
            mtm_last_row = mtm_sht.range("F6").end("down").row
            mtm_sht.api.Range(f"F6").AutoFilter(Field:=6,Criteria1:=loc, Operator:=7)
            columns_1[key]=columns_1[key].replace("NGREEL","NORTH GREELEY")
            columns_1[key]=columns_1[key].replace("BROWNSVILL","BROWNSVILLE")
            
            mtm_sht.api.Range(f"E6").AutoFilter(Field:=5,Criteria1:=columns_1[key].split(','), Operator:=7)

            mtm_sht.api.Range(f"J7:J{mtm_last_row}").SpecialCells(12).Select()
            mtm_wb.app.selection.value = new_sht.range(f"R{i}").value
        if loc == "HRW":
            mtm_sht.api.AutoFilterMode=False
            mtm_sht.api.Range(f"F6").AutoFilter(Field:=6,Criteria1:='<>HRW', Operator:=1, Criteria2:='<>YC')
            
            try:
                rng_lst = mtm_sht.api.Range(f"F6:F{mtm_last_row}").SpecialCells(12).Address.split(",")
            except:
                rng_lst = list(mtm_sht.api.Range(f"F6:F{mtm_last_row}").SpecialCells(12).Address)
            # 
            for rng in rng_lst:
                rng.split("$")[2].replace(':','')
                for i in range(int(rng.split("$")[2].replace(':','')), int(rng.split("$")[-1])+1):
                    if i == 6:
                        continue
                    if mtm_sht.range(f"I{i}").value is not None and mtm_sht.range(f"I{i}").value != 0:
                        print(i)
                        try:
                            if mtm_sht.range(f"F{i}").value == "YGS":
                                mtm_sht.range(f"J{i}").value = loc_dict["SORGHUM"][mtm_sht.range(f"E{i}").value]
                            else:
                                mtm_sht.range(f"J{i}").value = loc_dict[mtm_sht.range(f"F{i}").value][mtm_sht.range(f"E{i}").value]
                        except:
                            pass



            print()
        
        wb.save(ouput_loc)
        wb.close()
    mtm_sht.api.AutoFilterMode=False    
    mtm_wb.api.ActiveSheet.PivotTables(1).PivotCache().Refresh()   
    mtm_wb.api.ActiveSheet.PivotTables(2).PivotCache().Refresh()  
    mtm_wb.save(mtm_ouput_loc)
    mtm_wb.app.quit()
    print()

    

    
    return f"Fifo reports Genrated for {input_date}"


input_date = "01.31.2022"
output_date=None

msg = fifo(input_date, output_date)
print(msg)
print()
