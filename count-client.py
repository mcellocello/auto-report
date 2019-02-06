import xlwt,os,csv,fnmatch

wb=xlwt.Workbook();p=0;xres=[];worksheet=wb.add_sheet('Sheet 1')
listsite=open('list-x.txt','r').readlines();lensite=len(listsite)
try:
	for x in range (lensite):
		listsite[x]=listsite[x].replace('\n', '')
		target = fnmatch.filter(os.listdir('./count/'), listsite[x])
		raw = open('count/' + target[0], 'r')
		filedata = list(csv.reader(raw));lendata=len(filedata)
		worksheet.write_merge(p, p+1,0,4,target);p=p+2;datalist=[]
		for i in range(lendata):
			if filedata[i][0] == ' ' and i!=5:
				break
			elif i>=0 and i<8:
				continue
			else:
				datalist.insert(i,int(filedata[i][2]))
		lendatalist=len(datalist);sums=sum(datalist)
		avg=sums/lendatalist;mx=max(datalist);mi=min(datalist)
		res=[mx,mi,avg,sums];xres.append(res)
		print(target[0],res)


	fnt=xlwt.Font();fnt.name='Calibri';fnt.colour_index=0;fnt.height=180;fnt.bold=False
	borders=xlwt.Borders();borders.left=1;borders.right=1;borders.top=1;borders.bottom=1
	al=xlwt.Alignment();al.horz=xlwt.Alignment.HORZ_CENTER;al.vert=xlwt.Alignment.VERT_CENTER
	style=xlwt.XFStyle();style.font=fnt;style.borders=borders;style.alignment=al;p=0;o=0

	for x in range(lensite):
		for z in range(0,4):
			worksheet.write_merge(p, p+1,o+5,o+5, xres[x][z],style)
			o=o+1
		o=0;p=p+2

	wb.save('output-count-client.xls');
	print('\nSukses bosq ==>> "output-count-client.xls"')

except IndexError as error:
    print(listsite[x])
