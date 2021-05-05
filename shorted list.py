import xlrd 
  
loc = (r'C:\Users\ADMIN\Documents\FaceRecognition-master\dimension.xlsx') 
  
wb = xlrd.open_workbook(loc) 
sheet = wb.sheet_by_index(0) 
sheet.cell_value(0, 0)
l=[]
for i in range(sheet.nrows):
     l.append(sheet.cell_value(i, 0))
print(l)
temp=[]
for i in range(1,len(l)):
     if l[i] not in temp:
          temp.append(l[i])

a=0
f=[]
for j in temp:
     for i in range(1,len(l)):
     
          if(l[i]==j):
               a=a+1
     
     if(a>5):
          f.append(j)
          a=0
     if(a<5):
          a=0
print(f)
          
               

               

    

        
    
    
