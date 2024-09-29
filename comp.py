import pandas as pd

# Load the Excel file
file_path = 'test.xlsx'  # Replace with your actual file path
df = pd.read_excel(file_path)

# Step 1: Filter students who have done any internship
internship_done_df = df[df['Have you done any internship'].str.lower() == 'yes']
internship_not_done_df = df[df['Have you done any internship'].str.lower() == 'no']

# Step 2: Categorize by type of internship (unpaid, paid, course)
unpaid_internship_df = internship_done_df[internship_done_df['Type of course /internship'].str.lower() == 'unpaid internship']
paid_internship_df = internship_done_df[internship_done_df['Type of course /internship'].str.lower() == 'paid internship']
course_df = internship_done_df[internship_done_df['Type of course /internship'].str.lower() == 'course']

# Step 3: Further categorize courses by platform
courses_by_platform = course_df.groupby('PLATFORM')

# Save each category into different Excel sheets
with pd.ExcelWriter('categorized_students.xlsx') as writer:
    internship_done_df.to_excel(writer, sheet_name='Internships_Done', index=False)
    internship_not_done_df.to_excel(writer, sheet_name='Internships_Not_Done', index=False)
    unpaid_internship_df.to_excel(writer, sheet_name='Unpaid_Internships', index=False)
    paid_internship_df.to_excel(writer, sheet_name='Paid_Internships', index=False)
    for platform, group in courses_by_platform:
        group.to_excel(writer, sheet_name=f'{platform}_Courses', index=False)

print("Categorized Excel file created successfully.")
