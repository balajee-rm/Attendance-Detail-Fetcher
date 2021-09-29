import xlrd
from xlwt import Workbook
from datetime import datetime as dt
 
# Location of Surabhi Attendance file same for all
loc = ("Surabhi Attendance.xls")
# Location of erp register data downloaded and converted to xls format
# any erp register downloaded should be saved as erp.xls
loc1 = ("erp.xls")

# To open Workbook and sheet
wb = xlrd.open_workbook(loc)
wb1 = xlrd.open_workbook(loc1)
wbw = Workbook()
sheet = wb.sheet_by_index(0)
erp_sheet = wb1.sheet_by_index(0)
sheet1 = wbw.add_sheet('Sheet 1')
 
l2 = [];
for erp_c in range (15, erp_sheet.ncols):
	l2.append(erp_sheet.cell_value(0, erp_c));
d1 = [];
h1 = [];
for l1 in l2:
	d1.append(l1.split(" h ")[0]);
	h1.append(l1.split(" h ")[1]);

out_r = 0;
sheet1.write(out_r, 0, 'Student ID')
sheet1.write(out_r, 1, 'Date')
sheet1.write(out_r, 2, 'Hour')

for er in range (erp_sheet.nrows):
	for sr in range (sheet.nrows):
		if (sheet.cell_value(sr, 0) == erp_sheet.cell_value(er, 2)):
			date = dt.utcfromtimestamp(((int(sheet.cell_value(sr, 1)))- 25569) * 86400.0).strftime('%d/%m/%y');
			for i in range (len(d1)):
				if (d1[i] == str(date)):
					if (h1[i] == str(int(sheet.cell_value(sr, 2)))):
						print("Student ID: " + str(int(erp_sheet.cell_value(er, 2))) + ", date = " + d1[i] + ", hour = " + h1[i]);
						out_r = out_r + 1;
						sheet1.write(out_r, 0, str(int(erp_sheet.cell_value(er, 2))))
						sheet1.write(out_r, 1, d1[i])
						sheet1.write(out_r, 2, h1[i])

# Close output.xls, if it is opend already for saving
# It shows the corresponding student's date and hour of current section processing
wbw.save('output.xls')
