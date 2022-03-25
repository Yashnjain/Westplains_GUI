import xlwings as xw 
import pandas as pd
import numpy as np
import os
from datetime import datetime


def moc_get_df_from_input_excel(input_dir, mtm_file, open_ap_file, open_ar_file,unsettled_pay_file, unsettled_recev_file):
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
            ws_mtm = wb_mtm.sheets['MTM Excel Summary']
            last_row = ws_mtm.range(f'A'+ str(ws_mtm.cells.last_cell.row)).end('up').row
            first_row  = ws_mtm.range(f"A{last_row}").end('up').last_cell.row
            req_index = first_row + 1
            df_mtm = pd.read_excel(mtm_file,sheet_name='MTM Excel Summary', usecols="A,B", skiprows=req_index)   
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
        finally:
            try:
                wb_mtm.app.quit()
            except Exception as e:
                print(e)
        
        """"This is the code for Open AP files"""
        try:
            inner_keys = ['Alliance/Hay Springs','Gering','Omaha','Johnstown','KC','BROWNSVILL']
            inner_dict = {}.fromkeys(inner_keys)
            # df_ap = pd.read_excel(open_ap_file,sheet_name='For allocation entry',usecols="A,B", skiprows=2)
            df_ap = pd.read_excel(open_ap_file,sheet_name = 0, usecols="A,B", skiprows=2)

            new_dict = dict(zip(df_ap.iloc[:,0],df_ap.iloc[:,1]))
            inner_dict['Alliance/Hay Springs'] = new_dict['HAYSPRG']
            inner_dict['Gering'] = new_dict.get('GERING')
            inner_dict['Omaha'] = new_dict.get('TERMINAL')
            inner_dict['Johnstown'] = new_dict.get('OMA COMM') + new_dict.get('JTELEV')
            inner_dict['KC'] = new_dict.get('KANSAS CTY')
            inner_dict['BROWNSVILL'] = new_dict.get('BROWNSVILL')
            req_dict['Accounts Payable'] = inner_dict
        except Exception as e:
            print(e)
            print("The format of input file is wrong for Open AP or the file does not exist. Please enter the correct format")
        
        """"This is the code for Open AR files"""
        try:
            inner_keys = ['Alliance/Hay Springs','Gering','Omaha','Johnstown','KC','BROWNSVILL']
            inner_dict = {}.fromkeys(inner_keys)
            # df_ar = pd.read_excel(open_ar_file, sheet_name='For allocation entry',usecols="A,B", skiprows=2)
            df_ar = pd.read_excel(open_ar_file, sheet_name = 0, usecols="A,B", skiprows=2)
            new_dict = dict(zip(df_ar.iloc[:,0],df_ar.iloc[:,1]))
            inner_dict['Alliance/Hay Springs'] = new_dict['HAYSPRG']
            inner_dict['Gering'] = new_dict.get('GERING')
            inner_dict['Omaha'] = new_dict.get('TERMINAL')
            inner_dict['Johnstown'] = new_dict.get('OMA COMM') + new_dict.get('JTELEV')
            inner_dict['KC'] = new_dict.get('KANSAS CTY')
            inner_dict['BROWNSVILL'] = new_dict.get('BROWNSVILL')
            req_dict['Open A/R'] = inner_dict
        except Exception as e:
            print(e)
            print("The format of input file is wrong for Open AR or the file does not exist. Please enter the correct format")
        
        """This is the code for Unsettled Payables files"""
        try:
            inner_keys = ['Alliance/Hay Springs','Gering','Omaha','Johnstown','KC','BROWNSVILL']
            inner_dict = {}.fromkeys(inner_keys)
            # df_pay = pd.read_excel(unsettled_pay_file, sheet_name = 'For allocation entry', usecols="A,B", skiprows=2)
            df_pay = pd.read_excel(unsettled_pay_file, sheet_name = 0, usecols="A,B", skiprows=2)
            new_dict = dict(zip(df_pay.iloc[:,0],df_pay.iloc[:,1]))
            inner_dict['Alliance/Hay Springs'] = new_dict['HAY SPRINGS - WEST PLAINS, LLC']
            inner_dict['Gering'] = new_dict.get('GERING - WEST PLAINS, LLC')
            inner_dict['Omaha'] = new_dict.get('OMAHA TERMINAL ELEVATOR - WEST PLAINS, LLC')
            inner_dict['Johnstown'] = new_dict.get('OMAHA COMM - WEST PLAINS, LLC') + new_dict.get('JOHNSTOWN - WEST PLAINS, LLC')
            inner_dict['KC'] = new_dict.get('KANSAS CTY')
            inner_dict['BROWNSVILL'] = new_dict.get('BROWNSVILLE - WEST PLAINS, LLC')
            req_dict['Unsettled A/P'] = inner_dict
        except Exception as e:
            print(e)
            print("The format of input file is wrong for Unsettled A/P or the file does not exist. Please enter the correct format")
            
        """This is the code for Unsettled Receivables"""
        try:
            inner_keys = ['Alliance/Hay Springs','Gering','Omaha','Johnstown','KC','BROWNSVILL']
            inner_dict = {}.fromkeys(inner_keys)
            # df_recev = pd.read_excel(unsettled_recev_file, sheet_name = 'For allocation entry', usecols="A,B", skiprows=2)
            df_recev = pd.read_excel(unsettled_recev_file, sheet_name = 0, usecols="A,B", skiprows=2)
            new_dict = dict(zip(df_recev.iloc[:,0],df_recev.iloc[:,1]))
            inner_dict['Alliance/Hay Springs'] = new_dict['HAY SPRINGS - WEST PLAINS, LLC']
            inner_dict['Gering'] = new_dict.get('GERING - WEST PLAINS, LLC')
            inner_dict['Omaha'] = new_dict.get('OMAHA TERMINAL ELEVATOR - WEST PLAINS, LLC')
            inner_dict['Johnstown'] = new_dict.get('OMAHA COMM - WEST PLAINS, LLC') + new_dict.get('JOHNSTOWN - WEST PLAINS, LLC')
            inner_dict['KC'] = new_dict.get('KANSAS CTY')
            inner_dict['BROWNSVILL'] = new_dict.get('BROWNSVILLE - WEST PLAINS, LLC')
            req_dict['Unsettled A/R'] = inner_dict
        except Exception as e:
            print(e)
            print("The format of input file is wrong for Unsettled A/R or the file does not exist. Please enter the correct format")
            
        
        main_df = pd.DataFrame(req_dict)
        print("Main dataframe created")
        return main_df
    except Exception as e:
        print(e)
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
                ws_alloc.range('P9:P15').options(transpose=True).value = main_df.values[5]

                ws_alloc.range('E9:p15').api.NumberFormat = '_("$"* #,##0_);_("$"* (#,##0);_("$"* "-"??_);_(@_)'

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
                        elif key == 'P17':
                            ws_alloc.range('P29:P35').options(transpose=True).value = main_df.values[5]
                else:        
                    ws_alloc.range('E29:E35').options(transpose=True).value = main_df.values[0]
                    ws_alloc.range('F29:F35').options(transpose=True).value = main_df.values[1]
                    ws_alloc.range('G29:G35').options(transpose=True).value = main_df.values[2]
                    ws_alloc.range('I29:I35').options(transpose=True).value = main_df.values[3]
                    ws_alloc.range('J29:J35').options(transpose=True).value = main_df.values[4]
                    ws_alloc.range('P29:P35').options(transpose=True).value = main_df.values[5]

                # ws_alloc.range('E37:p37').formula = '=+E29+E30+E31-E32-E33-E34-E35'
                # ws_alloc.range('E39:p39').formula = '=E37/$Q$37'
                # ws_alloc.range('E40:p40').formula = '=E39*$E$62'

                ws_alloc.range('E29:p35').api.NumberFormat = '_("$"* #,##0_);_("$"* (#,##0);_("$"* "-"??_);_(@_)'
                wb_alloc.save(output_dir + '\\' + file.replace(file.split('_')[1],input_date) + '.xls')                
                print(f"MOC Allocment file generated for {input_date}")
    except Exception as e:
        print("Template file was not found or some other issue occured")
        print(e)
    finally:
        try:
            wb_alloc.app.quit()
        except Exception as e:
            print(e)


def moc_interest_alloc(input_date, output):
    input_dir = r"J:\WEST PLAINS\REPORT"
    output_dir = r"J:\WEST PLAINS\REPORT\MOC Interest allocation\Output Files"
    template_dir = r"J:\WEST PLAINS\REPORT\MOC Interest allocation\Raw files\template"
    dt = datetime.strptime(input_date,"%m.%d.%Y")
    mtm_input_date = dt.strftime("%B %Y")
    mtm_file = input_dir + '\\MOC Interest allocation\Raw files\\Inventory MTM Excel Report ' + mtm_input_date + '.xlsx'
    open_ap_file = input_dir +"\\Open AP\\Output files\\" + "Open AP _" + input_date + '.xlsx'
    open_ar_file = input_dir +"\\Open AR\\Output files\\" + "Open AR _" + input_date + ' - Production.xlsx'
    unsettled_pay_file = input_dir +"\\Unsettled Payables\\Output files\\" + "Unsettled Payables _" + input_date + '.xlsx'
    unsettled_recev_file = input_dir +"\\Unsettled Receivables\\Output files\\" + "Unsettled Receivables _" + input_date + '.xlsx'
    main_df = moc_get_df_from_input_excel(input_dir, mtm_file, open_ap_file, open_ar_file,unsettled_pay_file, unsettled_recev_file)
    update_moc_excel(main_df, template_dir, output_dir, input_date)



if __name__ == '__main__':
    
    
    input_date = input("Enter the input date in 'm.d.Y' format for the file name ")
    moc_allocation(input_date)

    print("Completed")