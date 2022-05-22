import pandas as pd
import re
import numpy as np

# Columns
req_cols = ["Shipping Phone Number","Customer Name","Created at","Item Name" ,"Seller SKU","Unit Price","Paid Price","Variation","Order Number"]

# Reading the csv file
df_new = pd.read_csv('test.csv' , sep=';' , encoding='utf-8' ,usecols=req_cols)


# Fabric Data 
df_new['Fabric'] = df_new['Unit Price'].replace({310: 'Card', 330: 'Combed Dawah', 350: 'Combed Regular', 490: 'Viscos', 550: 'Pleated'})

# Spliting Variation data

# Color Columm
color = df_new['Variation'].str.split(',' ,expand=True)[0]
df_new['Color'] = color.str.split(':' , expand=True)[1]

# Size Column
size = df_new['Variation'].str.split(',' ,expand=True)[1]
df_new['Size'] = size.str.split(':' ,expand=True)[2]

# Deleting Variation Column
del df_new['Variation']

# Insert data
df_new.insert(11, "Sales Type" , "MarketPlace Daraz")
df_new.insert(12, "Due Amount" , df_new['Unit Price'])
df_new.insert(13 , "Payment Method" , "Credit")
df_new.insert(14, "Courier Type" , "Local")
df_new.insert(15 , "Daraz Order No" , df_new['Order Number'])

# Changing headers name
df_after_rename = df_new.rename(columns={
"Shipping Phone Number" : "Contact Number",
"Customer Name" : "Customer Name",
"Created at" : "Date",
"Item Name": "Product Details",
"Seller SKU" : "Code",
"Unit Price" : "Net Sale Price",
"Paid Price" : "Due Amount"
})

# Saving xlsx file
write_excel_file = pd.ExcelWriter('data.xlsx')

df_after_rename.to_excel(write_excel_file,index=False,index_label='Order Serial',sheet_name="CRM Sheet")

write_excel_file.save()
