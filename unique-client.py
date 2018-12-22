import os,csv,xlwt
wb = xlwt.Workbook();worksheet = wb.add_sheet('Sheet 1')
fileo=open('list-unique.txt','r');listfile=list(fileo.readlines());p=0

for i in range (len(listfile)):
    print(listfile[i].replace("unik/",""))
    raw=open(listfile[i].replace("\n",""),'r')
    fileread = list(csv.reader(raw))
    worksheet.write_merge(p,p+1,0,2,listfile[i].replace("unik/",""))
    worksheet.write_merge(p,p+1,3,3,fileread[7][0])
    worksheet.write_merge(p,p+1,4,4,fileread[11][1])
    p=p+2

wb.save("output-unique-client.xls");print("\nsukses")