# import xlwings as xw
# import pyodbc

# file_path = r"C:\Users\ylee\Desktop\dev\AUL\MONTHLYREPORT\etl\Claims Detail_RFJ Group Append.xlsx"


# with xw.App(visible=True) as Excel:
#     # Open the workbook and append the data   
#     xw.App.display_alerts = True
#     wb = xw.Book(file_path)

#     # Refresh all Pivot Tables within the workbook
#     wb.sheets['pivot'].select()
#     wb.api.ActiveSheet.PivotTables('PivotTable1').PivotCache().refresh()
#     # wb.api.RefreshAll()
#     # for wsh in wb.sheets:
#     #     for pivottable in wsh.api.PivotTables():
#     #         pivottable.PivotCache().refresh()

#     wb.save()





# for driver in pyodbc.drivers():
#     print(driver)



# """
# SQL Server
# Microsoft Access Driver (*.mdb, *.accdb)
# Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)
# Microsoft Access Text Driver (*.txt, *.csv)
# Microsoft Access dBASE Driver (*.dbf, *.ndx, *.mdx)
# SQL Server Native Client 11.0
# ODBC Driver 13 for SQL Server
# PostgreSQL ANSI(x64)
# PostgreSQL Unicode(x64)
# ODBC Driver 17 for SQL Server
# Amazon Redshift (x64)
# """
