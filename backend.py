import pandas as pd
import re
import numpy as np

#
req_cols = ["Shipping Phone Number","Customer Name","Created at","Item Name" ,"Seller SKU","Unit Price","Paid Price","Variation"]

# Reading the csv file
df_new = pd.read_csv('~/test.csv' , sep=';' , encoding='utf-8' ,usecols=req_cols)


#spliting variation data
get_variation = df_new['Variation']
for part in get_variation:
        v = str(part)
        data = re.split(r'[:,]' , v)
        color = (data[1])
        size =  (data[4])

df_new[['Color' , 'Size']] = get_variation.str.split(',' ,expand=True)

#insert data
df_new.insert(8 , "Payment Method" , "Credit")
df_new.insert(9 , "Courier Type" , "Local")
df_new.insert(10 , "Sales Type" , "MarketPlace Daraz")
df_new.insert(11, "Billing Amount" , df_new['Unit Price'])

#Deleting Variation Column
del df_new['Variation']

#changing headers name
df_after_rename = df_new.rename(columns={
"Shipping Phone Number" : "Contact Number",
"Customer Name" : "Customer Name",
"Created at" : "Date",
"Item Name": "Product Details",
"Seller SKU" : "Code",
"Unit Price" : "Net Sale Price",
"Paid Price" : "Due Amount"
})

# saving xlsx file
write_excel_file = pd.ExcelWriter('/storage/emulated/0/data.xlsx')

df_after_rename.to_excel(write_excel_file,index=False,index_label='Order Serial',sheet_name="CRM Sheet")

write_excel_file.save()
