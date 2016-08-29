#coding=utf-8
	
import os;
os.system('color 1F');
os.system('mode con cols=40 lines=15');
	
try:
	import xlrd;
	import xlwt;
	
	print '''

    *******************************    
    *                             *    
    *  XLS combiner by Mike Wang  *    
    *                             *    
    *******************************    



[info]''';
	print 'Running, please wait...';
	
	new_workbook = xlwt.Workbook();
	
	dir = os.listdir("./input/");
	dir.sort();
	
	pre_workbook = xlrd.open_workbook("./input/" + dir[0]);
	pre_sheets = pre_workbook.sheets()
	nsheets = len(pre_sheets);
	
	for e in range(nsheets):
		new_sheet = new_workbook.add_sheet(pre_sheets[e].name);
		new_row = 0;
		for fname in dir:
			workbook = xlrd.open_workbook("./input/" + fname);
			sheet = workbook.sheet_by_index(e);
			for i in range(sheet.nrows):
				for j in range(len(sheet.row(i))):
					new_sheet.write(new_row + i, j, sheet.row(i)[j].value);
			new_row += sheet.nrows;
	
	new_workbook.save('output.xls');
	
	#Success
	os.system('color 2F');
	print 'DONE! Please check output.xls!';
	
except:
	os.system('color 4F');
	import time;
	import traceback;
	print 'An error occurred. See err.log.';
	errtime = time.time();
	errlog = traceback.format_exc();
	fo = open("err.log", "a");
	fo.write('Unix timestamp: '+str(errtime)+'\n');
	fo.write(errlog+'\n');
	fo.close();
	
finally:
	raw_input('Press Enter to exit...');
