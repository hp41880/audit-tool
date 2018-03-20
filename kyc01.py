import glob,openpyxl,pandas as pd,numpy as np,csv
path="D:\\Ekyc\\"

# This loop will add content from each excel file into a dict of sets
files=glob.glob(path+"/*.xlsx")
s1={}
for filename in files:
    wb = openpyxl.load_workbook(filename)
    s1[filename.rstrip(".xlsx").lstrip(path)]=set()
    for row in wb.active.rows:
       for cell in row:
           cv=cell.value
           s1[filename.rstrip(".xlsx").lstrip(path)].add(cv)



#   This will update u1 by intersecting it with all the other sets
u1=s1[filename.rstrip(".xlsx").lstrip(path)]
for i in s1:
    j=s1[i]
    u1=j&u1
    print(i,'---',len(u1))
    
u1=s1[filename.rstrip(".xlsx").lstrip(path)]    
s2=list() 
for i in s1:
    s2.append(s1[i])

conv=[]

for i in range(len(s2)):
    temp=[]
    for j in range(len(s2)):
        temp.append(s2[i]&s2[j])
    conv.append(temp)

superset=set()    
for i in range(len(conv)):
    for j in range(len(conv)):
        if i!=j:
            superset=superset|conv[i][j] 
print(superset)
print(len(superset))
    
# save u1 to excel file
header = [u'field name', u'regular expression to find this field name', u'rowminus', u'rowplus',u'colminus',u'colplus',u'regex to find field value']
wb = openpyxl.Workbook()
dest_filename = path+'scripts\\empty_book.xlsx'
ws1 = wb.active
ws1.title = "range names"
ws1.append(header)
for row in u1:
    ws1.append([row])
wb.save(filename = dest_filename)


wb=openpyxl.load_workbook(path+"scripts\\empty_book.xlsx")
ws=wb.active

df=pd.DataFrame(columns=header)
df = pd.read_excel(path+"scripts\\empty_book.xlsx", encoding = 'utf8')
df1=df[df['regular expression to find this field name'].notnull()]
df1=df1.reset_index()


df2=pd.DataFrame(columns=df1['field name'])
l=[]
for excel_filename in files:
    d={}
    for i in range(len(df1)):
        reg_ex=df1['regular expression to find this field name'][i]
        rowminus=df1['rowminus'][i].astype(np.int64)
        rowplus=df1['rowplus'][i].astype(np.int64)
        colminus=df1['colminus'][i].astype(np.int64)
        colplus=df1['colplus'][i].astype(np.int64)
        z=ss(excel_filename,reg_ex)
        if len(z)>0:
            row=z[0][0]
            col=z[0][1]
        else:
            print("following reg_ex not in excel file--",reg_ex,"---",excel_filename)
#            continue
        surr_text=surrounding_text(row,col,rowminus,rowplus,colminus,colplus,excel_filename)
        CAF_N=re.compile(df1['regex to find field value'][i])
        
        if len(CAF_N.findall(surr_text))>0:
            d[reg_ex]=CAF_N.findall(surr_text)[0]
            dff=pd.DataFrame([d])
#        print(dff)                           
#        print(surr_text)
#        print("\n")
        else:
            d[reg_ex]=CAF_N.findall(surr_text)
            dff=pd.DataFrame([d])
    df2=df2.append(dff)
        
print(df2)
    
    