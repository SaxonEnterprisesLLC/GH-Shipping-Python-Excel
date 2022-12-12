# William Leonard, Saxon Enterprises LLC
# Contact: wfleonard@saxonenterprises.net 732-673-4260
# CSV to Excel program for GH AR Shipping Data 
# csv_to_excel.py

import warnings
warnings.simplefilter(action='ignore', category=FutureWarning)
import pandas as pd
import shippers_menu as menu
import time as t

def csvToExcel(csvFile, excelFile, col, vendor):
    
    Col1 = col
    dropCol="Shipping System / Adjustment"
    dropColName = "Import Shipment Detail"
    pubCharged = "Published Charge"

    if vendor == "FEDEX": 
        try:
            csv_file = pd.read_csv(csvFile)
            # remove , in Col1 so that we can change to float
            csv_file[Col1] = csv_file[Col1].replace(',', '', regex=True)
            csv_file[Col1] = csv_file[Col1].replace('"', '', regex=True)
            csv_file = csv_file.astype({Col1:'float'})
            fedex_file = pd.ExcelWriter(excelFile)
            csv_file.to_excel(fedex_file, index=False)
            fedex_file.save()
        except FileNotFoundError:
            print(f"File does not exist!! {csvFile}, Exiting Program\n\n")
            t.sleep(3)
            menu.main()
    elif vendor == "DHL":
        try:
            csv_file = pd.read_csv(csvFile)
            # remove , in Col1 so that we can change to float
            # csv_file[Col1] = csv_file[Col1].str.replace(',', '')
            # csv_file = csv_file.astype({Col1:'float'})
            dhl_file = pd.ExcelWriter(excelFile)
            csv_file.to_excel(dhl_file, index=False)
            dhl_file.save()
        except FileNotFoundError:
            print(f"File does not exist!! {csvFile}, Exiting Program\n\n")
            t.sleep(3)
            menu.main()
    elif vendor == "UPS":
        try:
            csv_file = pd.read_csv(csvFile)
            # remove , in Col1 so that we can change to float
            csv_file[Col1] = csv_file[Col1].str.replace(',', '')
            csv_file[Col1] = csv_file[Col1].str.replace('$', '')
            csv_file[Col1] = csv_file[Col1].str.replace('(', '')
            csv_file[Col1] = csv_file[Col1].str.replace(')', '')
            csv_file = csv_file.astype({Col1:'float'})
            csv_file = csv_file[csv_file[dropCol].str.contains(dropColName)==False]
            csv_file = csv_file[csv_file[pubCharged].isnull()]
            #csv_file.drop(csv_file[(csv_file["Published Charge"] == None)].index,axis=0,inplace=True)
            #csv_file = csv_file[csv_file[pubCharged].str.isalnum()==True]
            ups_file = pd.ExcelWriter(excelFile)
            csv_file.to_excel(ups_file, index=False)
            ups_file.save()
        except FileNotFoundError:
            print(f"File does not exist!! {csvFile}, Exiting Program\n\n")
            t.sleep(2)
            menu.main()
    else:
        pass
    
