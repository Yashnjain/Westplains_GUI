import PyPDF2
from tabula import read_pdf
import os
from datetime import datetime


def payroll_pdf_extractor(input_pdf):
    acc_no = None
    pdfReader = PyPDF2.PdfFileReader(input_pdf)
    ada_group = []
    location = []
    for page in range(pdfReader.numPages - 1):
        
        pageObj = pdfReader.getPage(page)
        a=pageObj.extractText()
        ada_group.append(a.split('Totals for Department: ')[1].split()[0])
        location.append(a.split('Totals for Department: ')[1].split()[2])
    df = read_pdf(input_pdf, pages = 'all', guess = False, stream = True ,
                        pandas_options={'header':0}, area = ["150,5,560,850"], columns=["65,120,145,200,330,380,430,470,700,750"])[0]
    print(df)
    state_fed_df = df.iloc[:,:4]
    state_fed_df = state_fed_df[state_fed_df[state_fed_df.columns[0]].notna()]
    state_taxable_df = df.iloc[:,4:8]
    state_taxable_df = state_taxable_df[state_taxable_df[state_taxable_df.columns[0]].notna()]
    deduc_ana_df = df.iloc[:,8:]
    deduc_ana_df = deduc_ana_df[deduc_ana_df[deduc_ana_df.columns[0]].notna()]
    return a

def payroll_summ(input_date, output_date):
    monthYear = datetime.strftime(datetime.strptime(input_date, "%m.%d.%Y"), "%b %y")
    input_pdf = r"J:\WEST PLAINS\REPORT\Payroll summary accounting report\Raw Files" +f"\\Payroll Summary By Cost Center {input_date}.pdf"
    # input_pdf = r"C:\Users\imam.khan\OneDrive - BioUrja Trading LLC\Documents\WEST PLAINS\REPORT\Macquaire Accrual Entry\Raw Files" +f"\\Macq Statement_{input_date}.pdf"
    if not os.path.exists(input_pdf):
            return(f"{input_pdf} PDF file not present for date {input_date}")
    input_xl = r"J:\WEST PLAINS\REPORT\Payroll summary accounting report\Raw Files" +f"\\Payroll by Dept - {monthYear}.xlsx"
    # input_xl = r"C:\Users\imam.khan\OneDrive - BioUrja Trading LLC\Documents\WEST PLAINS\REPORT\Macquaire Accrual Entry\Raw Files" +f"\\Macq Accrual_{input_date}.xlsx"
    if not os.path.exists(input_xl):
            return(f"{input_xl} Excel file not present for date {input_date}")


    data = payroll_pdf_extractor(input_pdf)

    return f"Payroll Summary Report for {input_date} generated succesfully"
input_date = '03.11.2022'
msg = payroll_summ(input_date, output_date=None)
print(msg)
print()



