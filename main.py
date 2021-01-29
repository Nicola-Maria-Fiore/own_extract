from openpyxl import load_workbook
import xlsxwriter
import keyboard
import time
import sys
import os

out = "results/Excel Empty/"
out_original = "results/Excel Original/"
do_output = "results/"
company = "resources/list of firms.csv"
year = "resources/period.csv"
conf = "resources/conf.txt"

excel_path = None
times = []
formula = None

def createXslc(comapny_id, year_list):
    for o in [out,out_original]:
        workbook = xlsxwriter.Workbook(o+comapny_id+'.xlsx')

        for y in year_list:
            worksheet = workbook.add_worksheet(y)
            content = formula.replace("INSERT",comapny_id).replace("YEAR",y)
            worksheet.write('A1', content)
        
        workbook.close()

def openFiles(companies):
    for c in companies:
        f_path = out + c + ".xlsx"
        f_path = os.path.abspath(f_path)
        os.system('start "{}" "{}"'.format(excel_path,f_path))
        time.sleep(times[0])
        keyboard.send('ctrl+s')
        time.sleep(times[1])
        keyboard.send('alt+F4')
        time.sleep(times[2])

def readFile(fname):
    result = []
    with open(fname, 'r', encoding='utf-8') as file_in:
        for line in file_in:
            result.append(line.strip().replace('\n',""))
    return result

def getConf(s):
    return ((s.split("-->"))[1]).strip()


#def addRIC(companies):
#    for c in companies:
#       f_path = out + c + ".xlsx"
#        workbook = load_workbook(f_path)
#        for worksheet in workbook.worksheets:
#            worksheet['A2'] = "RIC"
#        workbook.save(f_path)


def createDo(companies, years):
    content = ""
    for c in companies:
        for y in years:
            content += 'import "{}.xlsx", sheet("{}") cellrange(A2) firstrow \n save "{}_{}.dta" \n'.format(c,y,c,y)
    

if __name__ == "__main__":
    companies = readFile(company)
    years = readFile(year)

    configurations = readFile(conf)
    excel_path = getConf(configurations[0])
    formula = getConf(configurations[4])
    times = configurations[1:4]
    for i in range(0, len(times)):
        times[i] = int(getConf(times[i]))  

    if len(sys.argv)>1 and sys.argv[1] == "-a": 
        for c in companies:
            createXslc(c,years) 
    elif len(sys.argv)>1 and sys.argv[1] == "-b":     
        #openFiles(companies)
        addRIC(companies)
    else:
        print("Read instructions") 
   
    print("Done!")