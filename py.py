import os,csv,xlwt
wb = xlwt.Workbook();p=0;xx=None;xres = []
worksheet = wb.add_sheet('Sheet 1');listtxt=open('list.txt','r');listsite=listtxt.readlines();lensite=len(listsite)
xox=None
try:
	for x in range (lensite):
		listsite[x] = listsite[x].replace('\n', '')
		fileproc=open(listsite[x]);fileread=csv.reader(fileproc);filedata=list(fileread);lendata=len(filedata)
		worksheet.write_merge(p, p + 1, 0, 4,listsite[x].replace('data/',''))
		p = p + 2;datalist=[]
		for i in range(lendata):
			if filedata[i][0] == ' ' and i != 5:
				break
			elif i >= 0 and i < 8:
				continue
			else:
				datalist.insert(i,int(filedata[i][2]))
		lendatalist = len(datalist)
		sums = sum(datalist);avg = sums / lendatalist;mx = max(datalist);mi = min(datalist)
		res = [mx,mi,avg,sums];xres.append(res)
except Exception:
	print("data ada yg gak valid bos");xx=0
fnt = xlwt.Font();fnt.name = 'Calibri';fnt.colour_index = 0;fnt.height = 180;fnt.bold = False
borders = xlwt.Borders();borders.left = 1;borders.right = 1;borders.top = 1;borders.bottom = 1
al = xlwt.Alignment();al.horz = xlwt.Alignment.HORZ_CENTER;al.vert = xlwt.Alignment.VERT_CENTER
style = xlwt.XFStyle();style.font = fnt;style.borders = borders;style.alignment = al
p=0;o=0
if xx == 0 :
	pass
else :
	for x in range(lensite):
		for z in range(0,4):
			worksheet.write_merge(p, p+1,o+5,o+5, xres[x][z],style)
			o=o+1
		o=0;p=p+2
	wb.save('output.xls');print("sukses bosq")
