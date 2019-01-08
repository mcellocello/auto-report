import os,csv,xlwt,fnmatch
target=fnmatch.filter(os.listdir('.'),'Weekly_Report-Top_AP_by_Client_Count_*.csv')
lisx=["BGR01","DPK01","SRG01","SBY01","BKS02","JKT03","PLB01","LLG01","JBR01",
	  "JKB03","JKB02","JKB01","TGR03","JKS01","JKS02","JKT01","JKP02","CKR01",
	  "JKU01","CKR03","BDG02","TGR04","BDG01","BLI02","BLI01","BLI03","MDN03",
	  "MDN06","MDN01","CKR02","JKP01","JOG01","PLB02","TGR07","BGR03","TGR06",
	  "TGR08"]

wb = xlwt.Workbook();worksheet = wb.add_sheet('Sheet 1');raw=open(target[0],'r')
rawread=csv.reader(raw);filedata=list(rawread);lendata=len(filedata)
datalist=[];datahost=[]

fnt = xlwt.Font();fnt.name = 'Calibri';fnt.colour_index = 0;fnt.height = 180;fnt.bold = False
borders = xlwt.Borders();borders.left = 1;borders.right = 1;borders.top = 1;borders.bottom = 1
al = xlwt.Alignment();al.horz = xlwt.Alignment.HORZ_CENTER;al.vert = xlwt.Alignment.VERT_CENTER
style = xlwt.XFStyle();style.font = fnt;style.borders = borders;style.alignment = al
n=0;ln=len(lisx);p=0

for i in range(lendata):
	if filedata[i][0] == ' ' and i != 5:
		break
	elif i >= 0 and i < 8:
		continue
	else:
		datalist.append(filedata[i][7]);datahost.append(filedata[i][0])

for j in range(len(datahost)):
	for i in range(len(datahost)):
		if n==ln:
			break
		elif lisx[n] in datahost[i] and n < ln:
			print("ada",datahost[i],"jumlah",datalist[i])
			worksheet.write_merge(p, p + 1, 0, 0, datahost[i], style)
			worksheet.write_merge(p, p + 1, 1, 1, datalist[i], style)
			p=p+2;n=n+1
			break
		else:
			if n!=ln:
				continue
			else:
				break

wb.save("output-top-ap.xls")
