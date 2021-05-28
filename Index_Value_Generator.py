# ******************************************************************************
# If we OCR will not give any code we need to add the code according to index
# ******************************************************************************

# IMPORTING LIBRARIES
import time
import pandas as pd
import pandas
import xlrd


# Getting index plus value
def Indexplus_Data(path, code):
    import re
    cleanedList1 = []

    value = 0

    mapingcode = (code)
    # print(mapingcode[0])

    xls = pandas.ExcelFile(path)
    # print(path)

    # Runnig for loop to consider sheet name
    sheets = xls.sheet_names
    for n in sheets:
        # print(sheets)

        data = pd.read_excel(xls, n)

        for column in data:

            columndata = data[column]
            k = (columndata.values)

            for m in k:

                if m == 1:

                    v2 = ['BOST', 'BOX', 'BOY', 'BON', ]
                    v = data.columns.get_loc(column)
                    x = v + 1
                    y = v + 2

                    val = data.iloc[10, y]
                    # print(val)

                    for i in v2:

                        if i in str(val):
                            # print(n)

                            v = data.columns.get_loc(column)
                            # print(v)

                            columndata = (data.iloc[:, v + 1]).tolist()
                            cleanedList = [x for x in columndata if str(x) != 'nan']

                            for i in cleanedList:

                                if i == mapingcode[-1]:
                                    ind = (cleanedList.index(i))
                                    value1 = (cleanedList[ind + 1])

                                    if len(value1) >= 11:

                                        cleanedList1.append(re.findall(r'\d+', value1))
                                        value = cleanedList1[0][0]

                                    else:
                                        value = value1

    return value

# Getting minos index minos data from csv

def Indexminos_Data(path, code):

    import re
    cleanedList1 = []

    value = 0

    mapingcode = (code)
    # print(mapingcode[0])

    xls = pandas.ExcelFile(path)
    # print(path)


    sheets = xls.sheet_names

    # Runnig for loop to consider sheet name
    for n in sheets:
        # print(sheets)

        data = pd.read_excel(xls, n)

        for column in data:

            columndata = data[column]
            k = (columndata.values)

            for m in k:

                if m == 1:

                    v2 = ['BOST', 'BOX', 'BOY', 'BON', ]
                    v = data.columns.get_loc(column)
                    x = v + 1
                    y = v + 2

                    val = data.iloc[10, y]
                    # print(val)

                    for i in v2:

                        if i in str(val):
                            # print(n)

                            v = data.columns.get_loc(column)
                            # print(v)

                            columndata = (data.iloc[:, v + 1]).tolist()
                            cleanedList = [x for x in columndata if str(x) != 'nan']

                            for i in cleanedList:

                                if i == mapingcode[-1]:
                                    ind = (cleanedList.index(i))
                                    value1 = (cleanedList[ind - 1])

                                    if len(value1) >= 11:

                                        cleanedList1.append(re.findall(r'\d+', value1))

                                        value = cleanedList1[0][0]

                                    else:
                                        value = value1

    return value


# Testing

# path = ("/home/jsw/Documents/Master sheet.xlsx")
# code = ['SE10079666465']

# path = ("/home/jsw/Documents/Wt2.xlsx")
# code = ['SE14809']

# x = Indexplus_Data(path,code)
# print(x)

# x = Indexminos_Data(path,code)
# print(x)