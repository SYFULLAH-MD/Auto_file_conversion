from datetime import datetime
import pandas as pd

p=pd.ExcelFile('created1.xlsx')
p.sheet_names
for i in p.sheet_names:
    file=pd.read_excel(p,sheet_name=i)
    date = datetime.now().strftime("%Y%m%d%I%M%S")
    file.to_csv(i+"_"+f"{date}"+'.txt', sep="|",index=0)

file = open("ABC1001_"+f"{date}"+".txt","r")
Counter = 0
Content = file.read()
CoList = Content.split("\n")

for i in CoList:
    if i:
        Counter += 1

print("Number of Rows",Counter, file=open("count.txt","a"))
