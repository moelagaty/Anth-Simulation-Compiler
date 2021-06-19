import sys, subprocess, xlsxwriter
from datetime import datetime
subprocess.call("./repos/build/bin/currentRates "+' '.join(sys.argv[1:]), shell=True)
output=[['   ', 'Manual', 'Auto', 'Ratio', 'Mult Manual', 'Mult Auto', 'Strategy Manual', 'Strategy Auto', 'Time Manual', 'Time Auto']]
for line in subprocess.check_output("./repos/build/bin/currentRates "+' '.join(sys.argv[1:]), encoding='UTF-8', shell=True).splitlines():
	data=[string.strip(':') for string in line.strip().split(' ') if len(string.strip())>0]
	if not data[0].startswith('T'): continue
	output.append(data)

with xlsxwriter.Workbook(datetime.now().strftime('%Y%m%d_%H%M%S')+'.xlsx') as workbook:
	worksheet = workbook.add_worksheet('Simulation')
	worksheet.set_column(1, 2, 10)
	worksheet.set_column(3, 3, 9)
	worksheet.set_column(4, 4, 13)
	worksheet.set_column(5, 5, 10)
	worksheet.set_column(6, 7, 20)
	worksheet.set_column(8, 8, 13)
	worksheet.set_column(9, 9, 11)
	for row_num, data in enumerate(output):
		worksheet.write_row(row_num, 0, data)
