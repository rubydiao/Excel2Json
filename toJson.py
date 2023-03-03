import pandas as pd
import openpyxl
import json
import traceback
import sys

# try:
path_input = r"C:/Users/narutchai/Desktop/Selenium/hw4.xlsx"
# ที่อยู่ไฟล์ input
try:
    
    load_sheetname = openpyxl.load_workbook(path_input)
    sheet_name_list = load_sheetname.sheetnames
    # สร้างตัวแปลมาเก็บชื่อชีทเป็นลิส

    dic_result = list()
    #Last result json list

    for sheet_name in sheet_name_list:
        #Loop from sheet

        excel_data_df = pd.read_excel(path_input, sheet_name=sheet_name)
        excel_data_df = excel_data_df.dropna(how='all')

        
        if excel_data_df.columns[0].__contains__("Unnamed"):
                excel_data_df.columns = excel_data_df.iloc[0]
                excel_data_df = excel_data_df.loc[excel_data_df["effdate"]!="effdate"]
                
        else:
                pass
        #init var search table "interest rate"
        startRow_interest = 0
        for get_interest_row in range(len(excel_data_df)):
            if str(excel_data_df.iloc[get_interest_row,0]) == "interest rate":
                startRow_interest = get_interest_row
            else:
                pass

        # df interest
        interest_df = excel_data_df.iloc[startRow_interest:,:]
        interest_df.columns = excel_data_df.iloc[startRow_interest]
        interest_df = interest_df.loc[interest_df["interest rate"]!="interest rate"]
        interest_df = interest_df.dropna(axis=1,how='all')

        # df campaign1 filter from df interest
        excel_data_df = excel_data_df.iloc[:startRow_interest]

        #Format Date to YMD format
        excel_data_df['effdate'] = pd.to_datetime(excel_data_df["effdate"])
        excel_data_df['effdate'] = excel_data_df['effdate'].dt.strftime('%Y/%m/%d')
        filter_excel_data_df = excel_data_df[["filename","effdate","remark"]]


        column = interest_df.columns.tolist()
        temp = interest_df.values.tolist()
        temp.insert(0,column)
        case_list = []

        for dataRow in filter_excel_data_df.values:
            case = {filter_excel_data_df.columns[0]: dataRow[0], filter_excel_data_df.columns[1]: dataRow[1], filter_excel_data_df.columns[2]:dataRow[2] }
            case_list.append(case)
        thisdict = {
            excel_data_df.columns[0]: {
            'data':case_list,
            interest_df.columns[0]:temp
            }

        }    
        dic_result.append(thisdict)

    #Write to JSON File
    jsonString = json.dumps(dic_result)

    #ตั้งชื่อไฟล์JSON
    name_JSON = "JSONdata.json"
    jsonFile = open(name_JSON, "w")
    jsonFile.write(jsonString)
    jsonFile.close()
except Exception as e:
    exc_type, exc_value, exc_tb = sys.exc_info()
    tb = traceback.TracebackException(exc_type, exc_value, exc_tb)
    print(''.join(tb.format_exception_only()))
