import openpyxl
AL = openpyxl.reader.excel.load_workbook(filename="all.xlsx", data_only=True)
AL.active = 0
sheet_zero = AL.active
for i in range(2,60):
   print( '{"type": "Feature","properties": {"name":', sheet_zero['H'+str(i)].value,'},"geometry": {"type": "Point","coordinates": [',sheet_zero['F'+str(i)].value,',',sheet_zero['E'+str(i)].value,'] }}, ','\n')
