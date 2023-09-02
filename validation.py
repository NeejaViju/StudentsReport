import os
import datetime
import json
import pandas as pd


# Path to the parent directory containing student folders
parent_folder_path = "Students"

# ... (previous part of your code)

# Define the student data
student_data = {
    "PPP001": "Mohamed Hasir",
    "PPP002": "Ganesh Kumar R",
    "PPP003": "Deepa N",
    "PPP004": "Nt. Nallathayammal",
    "PPP005": "Prasanth Govindaraj",
    "PPP006": "Murali T",
    "PPP007": "LEEMAN THOMAS",
    "PPP008": "Vimal Nadarajan",
    "PPP009": "Saravanan Selvam",
    "PPP010": "Srinivasan SR",
    "PPP011": "David Raj",
    "PPP012": "Yogesh Kumar JG",
    "PPP013": "Aravindhan Selvaraj",
    "PPP014": "Naveen Bromiyo A R",
    "PPP016": "Madhan Karthick",
    "PPP017": "Pavithra Selvaraj",
    "PPP018": "Sindhu Laheri Uthaya Surian",
    "PPP019": "Nalina Athinamilagi",
    "PPP020": "Nithya Naveen",
    "PPF001": "Ranjitha",
    "PPF002": "Suganthi Ramaraj",
    "PPF004": "Swathipriya",
    "PPF005": "Jumana",
    "PPF006": "Indira Priyadharshini",
    "PPF007": "Riyas ahamed J"
}

# ... (rest of your code)

# Load the expected files configuration
with open("expected_files_config.json", "r") as config_file:
    expected_files_config = json.load(config_file)

# Function to validate week folder contents
def validate_week_folder(week_folder_path, expected_files):
    files_in_folder = os.listdir(week_folder_path)
    files_in_folder_stripped = [f.strip() for f in files_in_folder]
    present_files = [file for file in expected_files if any(file.lower() == f.lower() for f in files_in_folder_stripped)]
    missing_files = [file for file in expected_files if not any(file.lower() == f.lower() for f in files_in_folder_stripped)]
    return present_files, missing_files


# ... (previous part of your code)

# Specific week you want to check
specific_week = "Week02"

# Get current date and time
current_datetime = datetime.datetime.now()
current_datetime_str = current_datetime.strftime("%Y-%m-%d %I:%M:%S %p")  # Format with AM/PM

# Create a list to store report data
report_data = []

for student_id, student_name in student_data.items():
    student_folder_path = os.path.join(parent_folder_path, f"{student_id} - {student_name}")
    week_folder_name = specific_week
    week_folder_path = os.path.join(student_folder_path, week_folder_name)

    if os.path.exists(student_folder_path) and os.path.isdir(student_folder_path) and os.path.exists(week_folder_path) and os.path.isdir(week_folder_path):
        expected_files = expected_files_config.get(week_folder_name, {}).get(student_id, [])
        present_files, missing_files = validate_week_folder(week_folder_path, expected_files)
        missing_files_str = ', '.join(missing_files)
        if len(present_files) == len(expected_files):
            completion_status = "Completed"
        else:
            completion_status = "Pending"
        report_data.append([student_id, student_name, week_folder_name, missing_files_str, completion_status])
    else:
        report_data.append([student_id, student_name, week_folder_name, "Folder or data not found", "", ""])

# Convert the report_data list to a pandas DataFrame
report_df = pd.DataFrame(report_data, columns=["Student ID", "Student Name", "Week", "Pending Task", "Completion Status"])

# Save DataFrame to an Excel file
report_excel_filename = f"{specific_week}_report.xlsx"

# Create a Pandas Excel writer using XlsxWriter as the engine
with pd.ExcelWriter(report_excel_filename, engine='xlsxwriter') as writer:
    # Write the DataFrame data to XlsxWriter
    report_df.to_excel(writer, sheet_name='Report', index=False)
    
    # Get the xlsxwriter workbook and worksheet objects
    workbook = writer.book
    worksheet = writer.sheets['Report']

    # Add a header format.
    header_format = workbook.add_format({
        'bold': True,
        'text_wrap': True,
        'valign': 'top',
        'fg_color': '#007bff',
        'font_color': 'white',
        'border': 1
    })

    # Set the column widths and apply the header format.
    for col_num, value in enumerate(report_df.columns.values):
        worksheet.write(0, col_num, value, header_format)
        column_len = max(report_df[value].astype(str).apply(len).max(), len(value))
        col_width = column_len + 2
        worksheet.set_column(col_num, col_num, col_width)
    
    # Add the specific week to the worksheet
    worksheet.write(len(report_df) + 2, 0, f'Week: {specific_week}')
    
    # Add the current date and time to the worksheet
    worksheet.write(len(report_df) + 3, 0, f'Generated: {current_datetime_str}')
    
    # Add cell formats for completed students' names and statuses (green background, white text)
    green_format = workbook.add_format({'bg_color': 'green', 'font_color': 'white', 'bold': True})

    # Iterate through the DataFrame to format cells based on completion status
    for row_num, completion_status in enumerate(report_df["Completion Status"]):
        if completion_status == "Completed":
            worksheet.write(row_num + 1, 1, report_df.at[row_num, "Student Name"], green_format)
            worksheet.write(row_num + 1, 4, completion_status, green_format)
    
# Print a message indicating Excel report generation
print(f"Excel report generated: {report_excel_filename}")
