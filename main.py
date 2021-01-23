import xlsxwriter
import keyboard
import time
import os

out = "results/"
company = "company.csv"
year = "year.csv"

sheets = 5
excel_path = "C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
formula = '=@RDP.Data("INSERT","TR.HoldingsDate;TR.EarliestHoldingsDate;TR.PctOfSharesOutHeld;TR.SharesHeld;TR.SharesHeldValue;TR.InvestorFullName;TR.FilingType;TR.InvParentType;TR.OwnTrnverRating;TR.OwnTrnverRatingCode;TR.OwnTurnover;TR.NbrOfInstrHeldByI"&"nv;TR.NbrOfInstrBoughtByInv;TR.NbrOfInstrSoldByInv;TR.PctPortfolio;TR.InvestorType;TR.InvInvestmentStyleCode;TR.InvInvmtOrientation;TR.InvestorRegion;TR.InvAddrCountry","SDate=YEAR-12-31 CH=Fd RH=IN",A1)'

def createXslc(comapny_id, year_list):
    workbook = xlsxwriter.Workbook(out+comapny_id+'.xlsx')

    for y in year_list:
        worksheet = workbook.add_worksheet()
        content = formula.replace("INSERT",comapny_id).replace("YEAR",y)
        worksheet.write('A1', content)
    
    workbook.close()

def openFiles(companies):
    for c in companies:
        f_path = out + c + ".xlsx"
        f_path = os.path.abspath(f_path)
        os.system('start "{}" "{}"'.format(excel_path,f_path))
        time.sleep(20)
        keyboard.send('ctrl+s')
        time.sleep(5)
        keyboard.send('alt+F4')
        time.sleep(3)

def readFile(fname):
    result = []
    with open(fname, 'r', encoding='utf-8') as file_in:
        for line in file_in:
            result.append(line.strip().replace('\n',""))
    return result

if __name__ == "__main__":
    companies = readFile(company)
    years = readFile(year)

    for c in companies:
        createXslc(c,years)
    
    openFiles(companies)
   
    print("Done!")