from flask import Flask, render_template, request, redirect, send_file
import pandas as pd
import os


app = Flask(__name__)

@app.route('/')
def upload_file():
    return render_template('1.html')

@app.route('/upload', methods=['POST'])  
def process_file():
    if 'file' not in request.files:
        return "No file part"
    file = request.files['file']
    if file.filename == '':
        return "No selected file"
    if file:
        # Save the uploaded file
        file_path = os.path.join('uploads', file.filename)
        file.save(file_path)
        
        # Read the Excel file
        df = pd.read_excel(file_path)

        # Convert the 'Have you done Internship' column to lowercase for case-insensitive comparison
        df['Have you done Internship'] = df['Have you done Internship'].str.lower()
        df['Have you got any stipend during the Internship?'] = df['Have you got any stipend during the Internship?'].str.lower()

        # Filter based on 'Have you done Internship' column
        internship_done_df = df[df['Have you done Internship'] == 'yes']
        internship_not_done_df = df[df['Have you done Internship'] == 'no']

        # Filter based on stipend status
        stipend_received_df = internship_done_df[internship_done_df['Have you got any stipend during the Internship?'] == 'yes']
        no_stipend_df = internship_done_df[internship_done_df['Have you got any stipend during the Internship?'] == 'no']

        # Group courses based on organization

        # Write categorized data to an Excel file
        output_file_path = 'categorized_students.xlsx'
        with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
            internship_done_df.to_excel(writer, sheet_name='Internships_Done', index=False)
            internship_not_done_df.to_excel(writer, sheet_name='Internships_Not_Done', index=False)
            stipend_received_df.to_excel(writer, sheet_name='Stipend_Received', index=False)
            no_stipend_df.to_excel(writer, sheet_name='No_Stipend', index=False)


        # Send the categorized file back to the user
        return send_file(output_file_path, as_attachment=True)

# Function to sanitize sheet names by removing invalid characters



if __name__ == '__main__':
    app.run(debug=True)
