from flask import Flask, flash, request, redirect, render_template, session, url_for
from werkzeug.utils import secure_filename
import os
import pandas as pd
import numpy as np

UPLOAD_FOLDER = 'static'
ALLOWED_EXTENSIONS = {'xlsx', 'xlsm'}

app = Flask(__name__, template_folder='templates', static_folder='static')
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['SECRET_KEY'] = 'I_LOVE_INDIA'

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def read_data_from_files(folder_path):
    data_frames = []
    for file_name in os.listdir(folder_path):
        if file_name.endswith('.xlsx') or file_name.endswith('.xlsm'):
            file_path = os.path.join(folder_path, file_name)
            df = pd.read_excel(file_path, sheet_name='Sheet1')
            data_frames.append(df.astype(str))
    return data_frames

def merge_dataframes_outer_merge(dfs):
    # Ensure consistent data types for merging
    for df in dfs:
        df = df.applymap(lambda x: str(x))  # Convert all columns to strings

    merged_df = dfs[0]
    
    # Loop through the rest of the uploaded DataFrames for left merge
    for df in dfs[1:]:
        # Perform an outer merge
        merged_df = pd.merge(merged_df, df, how='outer')
    
    # Convert all columns to strings in the merged DataFrame
    merged_df = merged_df.applymap(lambda x: str(x))
    
    return merged_df

@app.route('/')
def index():
    return render_template('index.html')

current_directory = os.path.dirname(os.path.abspath(__file__))

@app.route('/upload_file', methods=['POST'])
def upload_file():
    if request.method == 'POST':
        uploaded_files = request.files.getlist('file')  # Use getlist to get multiple files

        data_frames = []  # List to store DataFrames from all uploaded files

        for uploaded_df in uploaded_files:
            if uploaded_df and allowed_file(uploaded_df.filename):
                data_filename = secure_filename(uploaded_df.filename)
                file_path = os.path.join(app.config['UPLOAD_FOLDER'], data_filename)
                uploaded_df.save(file_path)

                # Read the Excel file into a DataFrame and append to the list
                df = pd.read_excel(file_path, sheet_name='Sheet1')
                data_frames.append(df.astype(str))
         # Check for invalid mobile numbers
        invalid_numbers = pd.concat([
            df.loc[~df['Ship Phone1'].astype(str).str.match(r'^\d{10}$')]
            for df in data_frames
        ])
        invalid_pincodes = pd.concat([
            df.loc[~df['Ship Pincode'].astype(str).str.match(r'^\d{6}$')]
            for df in data_frames
        ])
        
        if not any('Ship Pincode'in df.columns for df in data_frames):
            flash('Pincode is missing in atleast one uploaded sheet')
         # Check if 'Ship Phone1' column exists in any of the DataFrames
        if not any('Ship Phone1' in df.columns for df in data_frames):
            flash('Mobile number is missing in at least one uploaded sheet')
        

        # Merge DataFrames
        merged_data_frame = merge_dataframes_outer_merge(data_frames)

        # Save the merged DataFrame to a file
        merged_data_path = os.path.join(current_directory, 'merged_data.xlsx')
        merged_data_frame.to_excel(merged_data_path, index=False)

        if not invalid_numbers.empty:
            flash("Invalid mobile numbers found in the 'Ship Phone1' column.")
            
        if not invalid_pincodes.empty:
            flash("Invalid pincode numbers found in Ship Pincode column.")
            
        else:
            flash('Files upload successful!')

    return render_template('index.html')


@app.route('/next')
def next():
    return render_template('orderupload.html')

@app.route('/orderupload', methods=['POST', 'GET'])
def orderupload():
    if request.method == 'POST':
        uploaded_df = request.files['file']  # Upload order format empty file with header fields
        
        if uploaded_df and allowed_file(uploaded_df.filename):
            data_filename = secure_filename(uploaded_df.filename)
            uploaded_df.save(os.path.join(app.config['UPLOAD_FOLDER'], data_filename))
            
        flash('Files upload successful!')

    return render_template('orderupload.html')

@app.route('/show_data')
def showdata():
    session['UPLOAD_FOLDER'] = app.config['UPLOAD_FOLDER']

    file_path_df1 = os.path.join(app.config['UPLOAD_FOLDER'], 'Order_Feed.xlsx')
    file_path_df2 = 'merged_data.xlsx'

    df1 = pd.read_excel(file_path_df1, header=1, parse_dates=['Order Date']).astype(str)
    
    df2 = pd.read_excel(file_path_df2).astype(str)
 
    # Retain only rows where 'Order Qty' has some value
    df1 = df1[df1['Order Qty'].notna()]

    for column in df1.columns:
        if column in df2.columns:
            df1[column] = df2[column]

    output_file_path = os.path.join(current_directory, 'Order_Feed_Final.xlsx')
    
    df1['Order Date'] = pd.to_datetime(df1['Order Date'])

    # Convert the date column to yymmdd format
    df1['Order Date'] = pd.to_datetime(df1['Order Date'], format='%m%d%Y').dt.strftime('%y%m%d')

    
    df1['External Order Number']=df1['Order Date']+"VR"+df1['Customer Code']
    # Replacing Nan values with constant values specified by DTDC.
    values= {"Order Type":"Prepaid",
             "Order Owner":"DTDC",
             "Order Currency":"INR",
             "Site Location":"DTDC Mumbai",
            "Ship Email1":"Noreply@akzonobel.com",
            "Ship Country":"INDIA",
            "Bill Address Same as Ship Address Flag":"Y",
            "Status":"Confirmed",
            "UOM":"01",
            "Unit Cost":"0.01",
            "Contact Person":"S MD Zakeer Hussain/N B Divakar",
            "Bill To Name":"Akzonobel india Ltd",
            "Bill To Address 1":"Plot No: 62P, Hoskote Industrial Area, Hoskote",
            "Bill To Address 2":"Bangalore - 562114",
            "Bill To Phone1":"8884831312"}
    #df1.replace('NaN', np.nan, inplace=True)
    df1.fillna(value=values, inplace=True)
    
        
    
    # shift column 'Name' to first position 
    first_column = df1.pop('External Order Number') 
  
# insert column using insert(position,column_name, 
# first_column) function 
    df1.insert(0, 'External Order Number', first_column) 
    df1.sort_values(by=['External Order Number'], inplace=True)
    
    df1.to_excel(output_file_path, index=False)
    
    static_folder_path = os.path.join(current_directory, 'static')
    for file_name in os.listdir(static_folder_path):
        file_path = os.path.join(static_folder_path, file_name)
        try:
            if os.path.isfile(file_path):
                os.remove(file_path)
        except Exception as e:
            print(f"Error deleting file {file_path}: {e}")
    result_html = df1.to_html()
    
    return render_template('result.html', table=result_html)


if __name__ == '__main__':
    app.run(debug=True,  port=8080,host="0.0.0.0")
