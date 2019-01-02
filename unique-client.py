import os,csv,xlwt,fnmatch
wb = xlwt.Workbook();worksheet = wb.add_sheet('Sheet 1')
fileo=open('list-unique.txt','r');listfile=list(fileo.readlines());p=0

fnt = xlwt.Font();fnt.name = 'Calibri';fnt.colour_index = 0;fnt.height = 180;fnt.bold = False
borders = xlwt.Borders();borders.left = 1;borders.right = 1;borders.top = 1;borders.bottom = 1
al = xlwt.Alignment();al.horz = xlwt.Alignment.HORZ_CENTER;al.vert = xlwt.Alignment.VERT_CENTER
style = xlwt.XFStyle();style.font = fnt;style.borders = borders;style.alignment = al

for i in range (len(listfile)):
    listfile[i]=listfile[i].replace('\n', '')
    target = fnmatch.filter(os.listdir('./unik/'), listfile[i])
    raw=open('unik/'+target[0],'r')
    fileread = list(csv.reader(raw))
    worksheet.write_merge(p,p+1,0,2,target[0])
    worksheet.write_merge(p,p+1,3,3,fileread[7][0],style)
    worksheet.write_merge(p,p+1,4,4,fileread[11][1],style)
    p=p+2

wb.save("output-unique-client.xls");print("\nsukses")
