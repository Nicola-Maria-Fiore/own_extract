import eikon as ek
from numpy import e
import pandas as pd
from time import sleep
from datetime import datetime

variables_file = "resources/variables.csv"
out_folder = "results/csv/"
app_key_file = "resources/appkey.txt"
time_file = "results/time.csv" 

importDo_file = "results/Do Files/import.do"
appendDo_file = "results/Do Files/append.do"

def getFields():
    cols = []
    df = pd.read_csv(variables_file)
    for index, _ in df.iterrows():
        cols.append(df.loc[index,"variables"])
    
    return cols


def EikonDownloader(firms,years):
    with open(app_key_file, 'r') as f:
        ek.set_app_key(f.read().strip())

    importDo = ""
    appendDo = ""

    variables = getFields()
    count = 1
    start_time = datetime.today().strftime('%Y-%m-%d-%H:%M:%S')
    for f in firms:
        for y in years:
            print("{} - Downloading {}_{}".format(str(count),f,y))
            
            fname = f + "_" + y
            importDo += 'import delimited "{}.csv", varnames(1) case(lower) stringcols(_all)  \n save "{}.dta" \n clear all \n'.format(fname,fname)
            
            if appendDo=="": 
                appendDo += 'use "{}_{}.dta" \nsave "ownership.dta"\n'.format(f,y)
            else:
                appendDo += 'append using "{}_{}.dta" \n'.format(f,y)

            try:
                data, err = ek.get_data(f, variables, parameters={"SDate":""+y+"-12-31", "CH":"Fd", "RH":"IN"})
                if err==None:
                    fname = f + "_" + y + ".csv"
                    data.to_csv(out_folder+fname)
                    sleep(0.3)
                else:
                    raise Exception("Error")
            except Exception as e:
                print("Error at {} {}".format(f,y))
                print(e)        
        count += 1
    end_time = datetime.today().strftime('%Y-%m-%d')

    time_content = "start time,"+start_time+"\nend_time,"+end_time
    with open(time_file, "w") as f:
        f.write(time_content)

    with open(importDo_file, "w") as f:
        f.write(importDo)

    with open(appendDo_file, "w") as f:
        f.write(appendDo+'save "ownership.dta, replace"')