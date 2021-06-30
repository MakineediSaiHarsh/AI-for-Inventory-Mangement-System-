import xlrd
import xlsxwriter
path="C:\\Users\\SAI HARSH\\Desktop\\hackathon\\train.xlsx"

excel_workbook = xlrd.open_workbook(path)
excel_worksheet = excel_workbook.sheet_by_index(0)

'''print(excel_worksheet.cell_value(0,0))
s=excel_worksheet.cell_value(1,0)[:4]
s=int(s)
if(s<2014):
    print(s)'''

dic={}
for i in range(1,11):
    d={}
    for j in range(1,51):
        #d[j]=[0,0]
        d[j]=1000000
    dic[i]=d

for row in range(1,excel_worksheet.nrows):
    s = excel_worksheet.cell_value(row, 0)[:4]
    s=int(s)
    if(s==2013):
        if(dic[int(excel_worksheet.cell_value(row,1))][int(excel_worksheet.cell_value(row,2))]> int(excel_worksheet.cell_value(row,3))):
            dic[int(excel_worksheet.cell_value(row, 1))][int(excel_worksheet.cell_value(row, 2))]=int(excel_worksheet.cell_value(row,3))
        #dic[int(excel_worksheet.cell_value(row,1))][int(excel_worksheet.cell_value(row,2))][0]+=int(excel_worksheet.cell_value(row,3))
        #dic[int(excel_worksheet.cell_value(row, 1))][int(excel_worksheet.cell_value(row, 2))][1]+=1


new_path="C:\\Users\\SAI HARSH\\Desktop\\hackathon\\minimum.xlsx"
new_workbook= xlsxwriter.Workbook(new_path)
new_worksheet=new_workbook.add_worksheet()

new_worksheet.write(0,0,"Shop")
new_worksheet.write(0,1,"Item")
new_worksheet.write(0,2,"Minimum")
k=1
for i in range(1,11):
    for j in range(1,51):
            new_worksheet.write(k,0,i)
            new_worksheet.write(k,1, j)

            #x = dic[i][j][0] / dic[i][j][1]
            #x=round(x)

            new_worksheet.write(k,2,dic[i][j])
            k=k+1
            print("Shop "+str(i)+" item "+str(j)+" average= ",end="")
            print(dic[i][j])
new_workbook.close()