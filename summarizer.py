import pandas as pd
from xlsxwriter.utility import xl_rowcol_to_cell

def summarize(f):

    lol1 = []
    lol2 = []
    temp = []
    # sum1 = sum2 = sum3 = sum4 = sum5 = sum6 = isum1 = isum2 = isum3 = isum4 = isum5 = isum6 = 0

    # Copy xlsx(!!!! not .xlsm) file in this folder and set path(line 10) according to its name. Then set name in line
    # 111 and ensure(!!!) that a file of the same name is not present in this folder.

    df = pd.ExcelFile(f)
    tabnames = df.sheet_names
    for tab in tabnames:
        sheet = df.parse(sheet_name=tab, skipfooter=45, usecols="A,G,H,I,O", header=None)
        colA = sheet.iloc[:,0]
        colG = sheet.iloc[:,1]
        colH = sheet.iloc[:,2]
        colI = sheet.iloc[:,3]
        colO = sheet.iloc[:,4]

        for x in colA:
            if "Invoice No".casefold() in str(x).casefold():
                temp.append(x[-3:])
            if "Invoice Date".casefold() in str(x).casefold():
                temp.append(x[-10:])
            if "Name:".casefold() in str(x).casefold():
                temp.append(x[6:])
            if "GSTIN".casefold() in str(x).casefold():
                temp.append(x[7:])
            if "TOTAL".casefold() in str(x).casefold():
                pos = colA[colA == x].index
                val1 = val2 = val3 = val4 = val5 = val6 = 0

                if "IGST".casefold() in tab.casefold():
                    val1 = colH.loc[pos].values[0]
                    temp.append(val1)

                    if str(colI.loc[pos].values[0]).casefold() != "NaN".casefold():
                        val2 = colI.loc[pos].values[0]
                        temp.append(val2)
                    else:
                        temp.append(val2)

                    val3 = colO.loc[pos + 1].values[0]
                    temp.append(val3)

                    val4 = colO.loc[pos + 2].values[0]
                    temp.append(val4)

                    val5 = 0
                    temp.append(val5)

                    val6 = colO.loc[pos + 3].values[0]
                    temp.append(val6)

                    lol2.append(temp)

                else:
                    val1 = colG.loc[pos].values[0]
                    temp.append(val1)

                    if str(colH.loc[pos].values[0]).casefold() != "NaN".casefold():
                        val2 = colH.loc[pos].values[0]
                        temp.append(val2)

                    else:
                        temp.append(val2)

                    val3 = colO.loc[pos + 1].values[0]
                    temp.append(val3)

                    val4 = colO.loc[pos + 2].values[0]
                    temp.append(val4)

                    val5 = colO.loc[pos + 3].values[0]
                    temp.append(val5)

                    val6 = colO.loc[pos + 5].values[0]
                    temp.append(val6)

                    lol1.append(temp)

                temp = []
                break

    nrows1 = len(lol1)
    nrows2 = len(lol2)

    writer = pd.ExcelWriter('Summary.xlsx', engine='xlsxwriter') # pylint: disable=abstract-class-instantiated
    workbook = writer.book
    worksheets = []

    if lol1:
        lol1 = list(map(list, zip(*lol1)))
        lol1[0] = list(map(int, lol1[0]))
        df1 = pd.DataFrame(
            {'DATE': lol1[1], 'INVOICE NO': lol1[0], 'PARTICULARS': lol1[2], 'GSTIN': lol1[3], 'AMOUNT': lol1[4],
            'TRANSPORT': lol1[5], 'NET': lol1[6], 'CGST @ 9%': lol1[7], 'SGST @ 9%': lol1[8], 'GROSS': lol1[9]})
        df1.to_excel(writer, 'CGST & SGST', index=False)
        worksheet1 = writer.sheets['CGST & SGST']
        worksheets.append(worksheet1)

    if lol2:
        lol2 = list(map(list, zip(*lol2)))
        lol2[0] = list(map(int, lol2[0]))
        df2 = pd.DataFrame(
            {'DATE': lol2[1], 'INVOICE NO': lol2[0], 'PARTICULARS': lol2[2], 'GSTIN': lol2[3], 'AMOUNT': lol2[4],
            'TRANSPORT': lol2[5], 'NET': lol2[6], 'IGST @ 18%': lol2[7], 'SGST': lol2[8], 'GROSS': lol2[9]})
        df2.to_excel(writer, 'IGST', index=False)
        worksheet2 = writer.sheets['IGST']
        worksheets.append(worksheet2)

    format1 = workbook.add_format({'num_format': '0.00'})
    invFormat = workbook.add_format({'align': 'center'})
    totalFormat = workbook.add_format({'bold': True})

    for ws in worksheets:

        ws.set_column('A:A', 12)
        ws.set_column('B:B', 12, invFormat)
        ws.set_column('C:C', 35)
        ws.set_column('D:D', 20)
        ws.set_column('E:J', 10, format1)

        ws.write(nrows1 + 1, 3, 'Total', totalFormat)
        for column in range(4, 10):
            cell_location = xl_rowcol_to_cell(nrows1 + 1, column)
            start_range = xl_rowcol_to_cell(1, column)
            end_range = xl_rowcol_to_cell(nrows1, column)
            formula = "=SUM({:s}:{:s})".format(start_range, end_range)
            ws.write_formula(cell_location, formula, totalFormat)
        nrows1 = nrows2

    writer.save()
    print("Success!!")