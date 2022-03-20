from datetime import datetime
import pandas as pd
import numpy as np
import xlsxwriter
import openpyxl
import glob
import os

file=glob.glob('C:/Users/Lenovo/Desktop/newdoc/*.xls')
for filename in file:
    p=pd.read_excel(filename)

list1 = list(p['ZIP/Postal code']) 
mylist = [str(int) for int in list1]
for i in range(len(mylist)):
        mylist[i]=mylist[i].replace(mylist[i],mylist[i].zfill(6))

p['ZIP/Postal code']=p['ZIP/Postal code'].astype(str)
df = pd.DataFrame(np.array([mylist]).T)
df.columns =['CUSTOMER_ZIP/POSTAL_CODE']

columns = ['CUSTOMER_LAST_NAME','CUSTOMER_FIRST_NAME','CUSTOMER_ADDRESS_1','CUSTOMER_ADDRESS_2','CUSTOMER_CITY',
           'CUSTOMER_STATE/PROVINCE','CUSTOMER_COUNTRY','CUSTOMER_ZIP/POSTAL_CODE','CUSTOMER_ZIP_CODE_+_FOUR',
           'CUSTOMER_EMAIL_ADDRESS','CUSTOMER_PHONE_NUMBER','CUSTOMER_ID','VENDOR_ACCOUNT_NUMBER','SUBSCRIPTION_TYPE',
           'SUBSCRIPTION_ORDER_DATE','SUBSCRIPTION_START_DATE','SUBSCRIPTION_CANCEL_DATE','TRAIL_SUBSCRIPTION_START_DATE',
           'TRAIL_SUBSCRIPTION_END_DATE','STATUS','PRODUCT','PRODUCT_PRICE','DISCOUNT_PRICE','DEVICE','DEVICE_VERSION',
           'PAYMENT_PERIOD_START_DATE','PAYMENT_PERIOD_END_DATE','PAYMENT_DATE','PAYMENT_AMOUNT','REFUND_DATE','REFUND_AMOUNT',
           'CUSTOMER_LAST_4-DIGIT_CC','CUSTOMER_CC_TYPE']

df=pd.DataFrame(columns=columns)
for i in ['Email Address','Subscriber ID','Vendor Account Number','subscrition cancel date mm-dd-yyyy','Phone Number:xxx-xxx-xxxx or xxxxxxxxxx']:
    p[i]=p[i].astype(str)
    p[i]=p[i].str.replace('nan','')

df['CUSTOMER_ID'] = p['Subscriber ID'].map(lambda s:" "+s)
df['CUSTOMER_LAST_NAME']=p["Last Name"].map(lambda x:" "+x)
df['CUSTOMER_FIRST_NAME']=p["First Name"].map(lambda y:" "+y)
df['STATUS']=p["Status:ACTIVE or EXPIRED"].map(lambda z:" "+z)
df['CUSTOMER_EMAIL_ADDRESS'] = p["Email Address"].map(lambda h:" "+h)
#df['CUSTOMER_ZIP/POSTAL_CODE'] = p['ZIP/Postal code'].map(lambda l:" "+l)
df['VENDOR_ACCOUNT_NUMBER'] = p['Vendor Account Number'].map(lambda v:" "+v)
df['SUBSCRIPTION_CANCEL_DATE'] = p['subscrition cancel date mm-dd-yyyy'].map(lambda d:" "+d)
df['CUSTOMER_PHONE_NUMBER'] = p["Phone Number:xxx-xxx-xxxx or xxxxxxxxxx"].map(lambda p:" "+p)

eight_column = df.pop('CUSTOMER_ZIP/POSTAL_CODE')
df.insert(7, 'CUSTOMER_ZIP/POSTAL_CODE', eight_column)

st=pd.concat([df])
file=glob.glob('C:/Users/Lenovo/Desktop/newdoc/*.xls')
for filename in file:
    fn = os.path.splitext(os.path.basename(filename))[0]
    st.to_excel('./created1.xlsx', engine='xlsxwriter',index=False)
    fl=openpyxl.load_workbook("created1.xlsx")
    fl_sheet = fl['Sheet1']
    fl_sheet.title = fn
    fl.save("created1.xlsx")
    p1=pd.ExcelFile('created1.xlsx')
    p1.sheet_names
    for i in p1.sheet_names:
        file=pd.read_excel(p1,sheet_name=i)
        date = datetime.now().strftime("%Y%m%d%I%M%S")
        file.to_csv(i+"_"+f"{date}"+'.txt', sep="|",index=0)
        file1 = open(i+"_"+f"{date}"+".txt","r")
        Counter = 0
        Content = file1.read()
        CoList = Content.split("\n")
        for i1 in CoList:
            if i1:
                Counter += 1
        print("Number of Rows",Counter, file=open(i+"_"+f"{date}"+".txt"+".done","a"))
