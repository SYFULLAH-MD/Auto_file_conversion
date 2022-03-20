import pandas as pd
import numpy as np
p=pd.read_excel('ABC1001.xls',header=0)
list1 = list(p['ZIP/Postal code']) 
mylist = [str(i) for i in list1]
for i in range(len(mylist)):
        mylist[i]=mylist[i].replace(mylist[i],mylist[i].zfill(6))

p['ZIP/Postal code']=p['ZIP/Postal code'].astype(str)
df = pd.DataFrame(np.array([mylist]).T)
df.columns =['CUSTOMER_ZIP/POSTAL_CODE']

columns = ['CUSTOMER_LAST_NAME','CUSTOMER_FIRST_NAME','CUSTOMER_ADDRESS_1','CUSTOMER_ADDRESS_2','CUSTOMER_CITY','CUSTOMER_STATE/PROVINCE',
           'CUSTOMER_COUNTRY','CUSTOMER_ZIP/POSTAL_CODE','CUSTOMER_ZIP_CODE_+_FOUR','CUSTOMER_EMAIL_ADDRESS',
           'CUSTOMER_PHONE_NUMBER','CUSTOMER_ID','VENDOR_ACCOUNT_NUMBER','SUBSCRIPTION_TYPE','SUBSCRIPTION_ORDER_DATE','SUBSCRIPTION_START_DATE',
           'SUBSCRIPTION_CANCEL_DATE','TRAIL_SUBSCRIPTION_START_DATE','TRAIL_SUBSCRIPTION_END_DATE','STATUS','PRODUCT','PRODUCT_PRICE','DISCOUNT_PRICE',
           'DEVICE','DEVICE_VERSION','PAYMENT_PERIOD_START_DATE',
           'PAYMENT_PERIOD_END_DATE','PAYMENT_DATE','PAYMENT_AMOUNT','REFUND_DATE','REFUND_AMOUNT','CUSTOMER_LAST_4-DIGIT_CC','CUSTOMER_CC_TYPE']
df=pd.DataFrame(columns=columns)
df['CUSTOMER_LAST_NAME']=p["Last Name"].map(lambda x:" "+x)
df['CUSTOMER_FIRST_NAME']=p["First Name"].map(lambda y:" "+y)
#df['CUSTOMER_ZIP/POSTAL_CODE'] = p['ZIP/Postal code'].map(lambda l:" "+l)
df['CUSTOMER_ZIP/POSTAL_CODE'] = pd.DataFrame(mylist)
p['Email Address']=p['Email Address'].astype(str)
df['CUSTOMER_EMAIL_ADDRESS'] = p["Email Address"].map(lambda h:" "+h)
p['Phone Number:xxx-xxx-xxxx or xxxxxxxxxx']=p['Phone Number:xxx-xxx-xxxx or xxxxxxxxxx'].astype(str)
df['CUSTOMER_PHONE_NUMBER'] = p["Phone Number:xxx-xxx-xxxx or xxxxxxxxxx"].map(lambda p:" "+p)
p['Subscriber ID']=p['Subscriber ID'].astype(str)
df['CUSTOMER_ID'] = p['Subscriber ID'].map(lambda s:" "+s)
p['Vendor Account Number']=p['Vendor Account Number'].astype(str)
df['VENDOR_ACCOUNT_NUMBER'] = p['Vendor Account Number'].map(lambda v:" "+v)
p['subscrition cancel date mm-dd-yyyy']=p['subscrition cancel date mm-dd-yyyy'].astype(str)
df['SUBSCRIPTION_CANCEL_DATE'] = p['subscrition cancel date mm-dd-yyyy'].map(lambda d:" "+d)
df['STATUS']=p["Status:ACTIVE or EXPIRED"].map(lambda z:" "+z)
eight_column = df.pop('CUSTOMER_ZIP/POSTAL_CODE')
df.insert(7, 'CUSTOMER_ZIP/POSTAL_CODE', eight_column)

st=pd.concat([df])
st.to_excel('./created1.xlsx', sheet_name='ABC1001',index=False)
