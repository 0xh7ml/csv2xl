from flask import Flask , render_template , request , redirect ,send_file
import pandas as pd
import re
import numpy as np
import openpyxl
app = Flask(__name__)
@app.route("/")
def index():
    return render_template('index.html')

@app.route("/data" , methods=['GET','POST'])
def data():
    req_cols = ["Shipping Phone Number","Customer Name","Created at","Item Name" ,"Seller SKU","Unit Price","Paid Price","Variation"]
    if request.method == "POST":
        csv = request.files['file']
        if csv:
            df_new = pd.read_csv(csv,sep=';',encoding='utf-8',usecols=req_cols)
            #spliting variation data
            get_variation = df_new['Variation']
            for part in get_variation:
                v = str(part)
                data = re.split(r'[:,]' , v)
                result = []
                color = result.append(data[1])
                size =  result.append(data[4])
                df_new[['Color']] = color
                df_new[['Size']] = size
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
            "Paid Price" : "Due Amount"})
            # saving xlsx file
            write_excel_file = pd.ExcelWriter('data.xlsx')
            df_after_rename.to_excel(write_excel_file,index=False,sheet_name="CRM Sheet")
            write_excel_file.save()
        return redirect('/download')
@app.route('/download')
def downloadFile():
    download_path = "data.xlsx"
    return send_file(download_path, as_attachment=True)
if __name__ == '__main__':
    app.run(debug=False) 
