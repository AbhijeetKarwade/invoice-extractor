from flask import Flask, render_template, send_from_directory, request, jsonify
import os
import pandas as pd
from werkzeug.utils import secure_filename
import json
import math
import logging

app = Flask(__name__, static_folder='static', template_folder='static')

# Configuration for file uploads
UPLOAD_FOLDER = 'Uploads'
ALLOWED_EXTENSIONS = {'xlsx'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Ensure upload directory exists
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def convert_nans(obj):
    """Convert NaN values to None (which becomes null in JSON)"""
    if isinstance(obj, float) and math.isnan(obj):
        return None
    elif isinstance(obj, dict):
        return {k: convert_nans(v) for k, v in obj.items()}
    elif isinstance(obj, list):
        return [convert_nans(item) for item in obj]
    return obj

def process_excel_file(filepath):
    """Process the Excel file and return the processed data in memory"""
    try:
        # Use context manager to ensure the file is properly closed
        with pd.ExcelFile(filepath) as excel_file:
            # Read the two sheets, specifying header row (row 3, so header=2)
            # and skip the empty row (row 4) by using skiprows=[3]
            df_custom = pd.read_excel(
                excel_file, 
                sheet_name='Custom Report', 
                header=2,  # Column names are in row 3 (0-indexed as 2)
                skiprows=[3]  # Skip row 4 (0-indexed as 3) which is empty
            )
            df_items = pd.read_excel(
                excel_file, 
                sheet_name='Item Details', 
                header=2,  # Column names are in row 3 (0-indexed as 2)
                skiprows=[3]  # Skip row 4 (0-indexed as 3) which is empty
            )

            # Clean column names by stripping whitespace
            df_custom.columns = df_custom.columns.str.strip()
            df_items.columns = df_items.columns.str.strip()

            # Function to remove empty columns
            def remove_empty_columns(df):
                # Check which columns are completely empty (all NaN or empty strings)
                empty_cols = [col for col in df.columns 
                             if df[col].isna().all() or (df[col].astype(str).str.strip().eq('').all())]
                # Drop empty columns
                return df.drop(columns=empty_cols)

            # Remove empty columns from both dataframes
            df_custom = remove_empty_columns(df_custom)
            df_items = remove_empty_columns(df_items)

            # Select key columns and rename for consistency
            df_custom_key = df_custom[['Date', 'Reference No']].rename(columns={'Date': 'date', 'Reference No': 'Ref_No'})
            df_items_key = df_items[['Date', 'Invoice No./Txn No.']].rename(columns={'Date': 'date', 'Invoice No./Txn No.': 'Ref_No'})

            # Combine key columns without removing duplicates
            parent_table = pd.concat([df_custom_key, df_items_key], ignore_index=True)
            parent_table['date'] = pd.to_datetime(parent_table['date'], dayfirst=True, errors='coerce')

            # Prepare remaining columns from both sheets
            custom_remaining = df_custom.drop(columns=['Date', 'Reference No'])
            items_remaining = df_items.drop(columns=['Date', 'Invoice No./Txn No.'])

            # Add Ref_No and date back for merging
            custom_remaining['Ref_No'] = df_custom['Reference No']
            custom_remaining['date'] = df_custom['Date']
            items_remaining['Ref_No'] = df_items['Invoice No./Txn No.']
            items_remaining['date'] = df_items['Date']

            # Convert dates to datetime for consistent merging
            custom_remaining['date'] = pd.to_datetime(custom_remaining['date'], dayfirst=True, errors='coerce')
            items_remaining['date'] = pd.to_datetime(items_remaining['date'], dayfirst=True, errors='coerce')

            # Merge with parent_table using Ref_No and date to preserve duplicates
            parent_table = parent_table.merge(custom_remaining, on=['Ref_No', 'date'], how='outer')
            parent_table = parent_table.merge(items_remaining, on=['Ref_No', 'date'], how='outer')

            # Identify string and numeric columns, excluding 'Balance' for special handling
            string_columns = [col for col in parent_table.columns if parent_table[col].dtype == 'object']
            numeric_columns = [col for col in parent_table.columns if parent_table[col].dtype in ['float64', 'int64'] and col != 'Balance']

            # Fill missing values: "NaN" for string columns, 0.00 for numeric columns (except Balance)
            parent_table[string_columns] = parent_table[string_columns].fillna('NaN')
            for col in numeric_columns:
                parent_table[col] = parent_table[col].apply(lambda x: 0.00 if pd.isna(x) else x)

            # Handle 'Balance' column separately: keep NaN as is, preserve explicit 0
            if 'Balance' in parent_table.columns:
                # Ensure explicit 0 values remain as 0, NaN remains NaN
                parent_table['Balance'] = parent_table['Balance'].apply(lambda x: x if not pd.isna(x) else pd.NA)

            # Sort by date and Ref_No
            parent_table = parent_table.sort_values(by=['date', 'Ref_No'])

            # Format date to DD/MM/YYYY
            parent_table['date'] = parent_table['date'].dt.strftime('%d/%m/%Y')

            # Remove duplicates based on all columns
            parent_table = parent_table.drop_duplicates(keep='first')

            # Convert the processed DataFrame to a dictionary
            processed_data = parent_table.to_dict(orient='records')
            processed_data = convert_nans(processed_data)
            
            # Get row counts for each sheet
            sheet_stats = {
                'customReportRows': len(df_custom),
                'itemDetailsRows': len(df_items)
            }
            
            return processed_data, list(parent_table.columns), sheet_stats
    except Exception as e:
        logger.error(f"Error processing Excel file: {str(e)}")
        return None, str(e), None

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/print')
def print_view():
    return render_template('print.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400
    
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        
        try:
            # Save the uploaded file
            file.save(filepath)
            
            # Process the Excel file and get data in memory
            processed_data, error_or_headers, sheet_stats = process_excel_file(filepath)
            
            if isinstance(error_or_headers, str):  # It's an error message
                return jsonify({'error': error_or_headers}), 500
            
            return jsonify({
                'message': 'File successfully processed',
                'data': processed_data,
                'headers': error_or_headers,  # In this case, it's the headers list
                'sheetStats': sheet_stats
            })
        except Exception as e:
            logger.error(f"Error during file processing: {str(e)}")
            return jsonify({'error': str(e)}), 500
        finally:
            # Ensure the file is deleted after processing
            try:
                if os.path.exists(filepath):
                    os.remove(filepath)
                    logger.info(f"Successfully deleted temporary file: {filepath}")
            except Exception as e:
                logger.error(f"Error deleting file {filepath}: {str(e)}")
                # Continue even if deletion fails - the file will be cleaned up later
    
    return jsonify({'error': 'Invalid file type'}), 400

@app.route('/static/<path:filename>')
def static_files(filename):
    return send_from_directory(app.static_folder, filename)

if __name__ == '__main__':
    app.run(debug=True)