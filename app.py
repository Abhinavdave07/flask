# from flask import Flask, render_template, request, redirect, send_file
# import pandas as pd
# import os
# import re
# from openpyxl.utils import get_column_letter

# # Function to sanitize sheet names by removing invalid characters
# def sanitize_sheet_name(sheet_name):
#     # Replace invalid characters with an underscore
#     return re.sub(r'[\\/*?:[\]]', '_', sheet_name)
# app = Flask(__name__)
# @app.route('/')
# def upload_file():
#     return render_template('1.html')

# @app.route('/upload', methods=['POST'])  
# def process_file():
#     if 'file' not in request.files:
#         return "No file part"
#     file = request.files['file']
#     if file.filename == '':
#         return "No selected file"
#     if file:
#         # Save the uploaded file
#         file_path = os.path.join('uploads', file.filename)
#         file.save(file_path)
        
#         # Read the Excel file
#         df = pd.read_excel(file_path)

#         # Convert relevant columns to lowercase for case-insensitive comparison
#         df['Have you done Internship'] = df['Have you done Internship'].str.lower()
#         df['Have you got any stipend during the Internship?'] = df['Have you got any stipend during the Internship?'].str.lower()

#         # Filter based on 'Have you done Internship' column
#         internship_done_df = df[df['Have you done Internship'] == 'yes']
#         internship_not_done_df = df[df['Have you done Internship'] == 'no']

#         # Filter based on stipend status
#         stipend_received_df = internship_done_df[internship_done_df['Have you got any stipend during the Internship?'] == 'yes']
#         no_stipend_df = internship_done_df[internship_done_df['Have you got any stipend during the Internship?'] == 'no']

#         # Calculate statistics
#         total_students = len(df)
#         total_internships_done = len(internship_done_df)
#         total_internships_not_done = len(internship_not_done_df)
#         total_stipend_received = len(stipend_received_df)
#         total_no_stipend = len(no_stipend_df)

#         # Create a DataFrame for statistics
#         stats_data = {
#             'Statistic': ['Total Students', 'Internships Done', 'Internships Not Done', 'Stipend Received', 'No Stipend'],
#             'Count': [total_students, total_internships_done, total_internships_not_done, total_stipend_received, total_no_stipend]
#         }
#         stats_df = pd.DataFrame(stats_data)

#         # Write categorized data to an Excel file (without organization-wise sheets)
#         output_file_path = 'categorized_students.xlsx'
#         with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
#             internship_done_df.to_excel(writer, sheet_name='Internships_Done', index=False)
#             internship_not_done_df.to_excel(writer, sheet_name='Internships_Not_Done', index=False)
#             stipend_received_df.to_excel(writer, sheet_name='Stipend_Received', index=False)
#             no_stipend_df.to_excel(writer, sheet_name='No_Stipend', index=False)
            
#             # Write the statistics sheet
#             stats_df.to_excel(writer, sheet_name='Statistics', index=False)

#             # Adjust column widths for all sheets
#             workbook = writer.book
#             for sheet_name in writer.sheets:
#                 worksheet = writer.sheets[sheet_name]
#                 for col in worksheet.columns:
#                     max_length = 0
#                     column = col[0].column_letter  # Get the column name
#                     for cell in col:
#                         try:
#                             if len(str(cell.value)) > max_length:
#                                 max_length = len(str(cell.value))
#                         except:
#                             pass
#                     adjusted_width = max_length + 2  # Add some extra space
#                     worksheet.column_dimensions[column].width = adjusted_width

#         # Send the categorized file back to the user
#         return send_file(output_file_path, as_attachment=True)

# if __name__ == '__main__':
#     app.run(debug=True)

from flask import Flask, render_template, request, redirect, send_file
import pandas as pd
import os
import re
from openpyxl.utils import get_column_letter

# Function to sanitize sheet names by removing invalid characters
def sanitize_sheet_name(sheet_name):
    return re.sub(r'[\\/*?:[\]]', '_', sheet_name)

app = Flask(__name__)

@app.route('/')
def upload_file():
    # No error by default
    return render_template('1.html', error=None)

@app.route('/upload', methods=['POST'])
def process_file():
    try:
        if 'file' not in request.files:
            error = "No file part"
            return render_template('1.html', error=error)
        
        file = request.files['file']
        if file.filename == '':
            error = "No selected file"
            return render_template('1.html', error=error)
        
        if file:
            if not os.path.exists('uploads'):
                os.makedirs('uploads')

            file_path = os.path.join('uploads', file.filename)
            file.save(file_path)

            try:
                df = pd.read_excel(file_path)
            except Exception as e:
                error = f"Error reading the Excel file: {e}"
                return render_template('1.html', error=error)

            required_columns = ['Have you done Internship', 'Have you got any stipend during the Internship?']
            if not all(col in df.columns for col in required_columns):
                error = f"Missing required columns: {', '.join(required_columns)}"
                return render_template('1.html', error=error)

            df['Have you done Internship'] = df['Have you done Internship'].str.lower()
            df['Have you got any stipend during the Internship?'] = df['Have you got any stipend during the Internship?'].str.lower()

            internship_done_df = df[df['Have you done Internship'] == 'yes']
            internship_not_done_df = df[df['Have you done Internship'] == 'no']
            stipend_received_df = internship_done_df[internship_done_df['Have you got any stipend during the Internship?'] == 'yes']
            no_stipend_df = internship_done_df[internship_done_df['Have you got any stipend during the Internship?'] == 'no']

            total_students = len(df)
            total_internships_done = len(internship_done_df)
            total_internships_not_done = len(internship_not_done_df)
            total_stipend_received = len(stipend_received_df)
            total_no_stipend = len(no_stipend_df)

            stats_data = {
                'Statistic': ['Total Students', 'Internships Done', 'Internships Not Done', 'Stipend Received', 'No Stipend'],
                'Count': [total_students, total_internships_done, total_internships_not_done, total_stipend_received, total_no_stipend]
            }
            stats_df = pd.DataFrame(stats_data)

            output_file_path = 'categorized_students.xlsx'
            with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
                internship_done_df.to_excel(writer, sheet_name='Internships_Done', index=False)
                internship_not_done_df.to_excel(writer, sheet_name='Internships_Not_Done', index=False)
                stipend_received_df.to_excel(writer, sheet_name='Stipend_Received', index=False)
                no_stipend_df.to_excel(writer, sheet_name='No_Stipend', index=False)
                stats_df.to_excel(writer, sheet_name='Statistics', index=False)

                workbook = writer.book
                for sheet_name in writer.sheets:
                    worksheet = writer.sheets[sheet_name]
                    for col in worksheet.columns:
                        max_length = 0
                        column = col[0].column_letter
                        for cell in col:
                            try:
                                if len(str(cell.value)) > max_length:
                                    max_length = len(str(cell.value))
                            except:
                                pass
                        adjusted_width = max_length + 2
                        worksheet.column_dimensions[column].width = adjusted_width

            return send_file(output_file_path, as_attachment=True)

    except Exception as e:
        error = f"An error occurred: {e}"
        return render_template('1.html', error=error)

if __name__ == '__main__':
    app.run(debug=True)
