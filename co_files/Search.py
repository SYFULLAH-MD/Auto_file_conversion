import pandas as pd
p=pd.read_excel('ABC1001.xls',header=0)
list1 = list(p['Vendor Account Number']) 
mylist = [str(int) for int in list1]
def Search(mylist, n, key):    
    for i in range(0, n):  
        if (mylist[i] == key):  
            return i  
    return -1  
     
key = input()
  
n = len(mylist)  
res = Search(mylist, n, key)  
if(res == -1):  
    print("Element not found")  
else:
    print("Element found in row: ", res+1)
