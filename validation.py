import os
import datetime
import pandas as pd

parent_folder_path = "Students"

# Default tasks for each week
default_tasks = {
    "Week01": ["Git_Task","Index_File_Updation","create_Html_file_on_Name","dulingo_update"],
    "Week02": ["create_wordpress_blog_and_7articles","update_linkedin_with-photo","create_canva-menu","download_figma_and_install"],
    # ... default tasks for other weeks
}

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
    "PPP015": "Kalai Selvi",
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
    "PPF007": "Riyas ahamed J",
}

weeks_to_report = ["Week01", "Week02"]  # Add other weeks as needed

def is_file_present(expected_file, files_in_folder):
    return any(
        expected_file.lower() == file_in_folder.lower()
        for file_in_folder in files_in_folder
    )

def validate_week_folder(week_folder_path, expected_files):
    files_in_folder = os.listdir(week_folder_path)
    files_in_folder_stripped = [
        os.path.splitext(f)[0].strip().lower() for f in files_in_folder
    ]
    present_files = [
        file for file in expected_files if is_file_present(file, files_in_folder_stripped)
    ]
    missing_files = [
        file for file in expected_files if not is_file_present(file, files_in_folder_stripped)
    ]
    return present_files, missing_files

for specific_week in weeks_to_report:

    current_datetime = datetime.datetime.now()
    current_datetime_str = current_datetime.strftime("%Y-%m-%d %I:%M:%S %p")

    report_data = []

    for student_id, student_name in student_data.items():
        student_folder_path = os.path.join(parent_folder_path, f"{student_id} - {student_name}")
        week_folder_path = os.path.join(student_folder_path, specific_week)

        if os.path.exists(week_folder_path) and os.path.isdir(week_folder_path):
            expected_files = default_tasks.get(specific_week, [])
            present_files, missing_files = validate_week_folder(week_folder_path, expected_files)
            missing_files_str = ", ".join(missing_files)
            completion_status = "Completed" if not missing_files else "Pending"
            report_data.append(
                [student_id, student_name, specific_week, missing_files_str, completion_status]
            )
        else:
            report_data.append([student_id, student_name, specific_week, "Folder not found", ""])

    report_df = pd.DataFrame(
        report_data,
        columns=["Student ID", "Student Name", "Week", "Pending Task", "Completion Status"]
    )

    report_excel_filename = f"{specific_week}_report.xlsx"

    # Begin the Excel writing and formatting segment
    with pd.ExcelWriter(report_excel_filename, engine="xlsxwriter") as writer:
        report_df.to_excel(writer, sheet_name="Report", index=False)

        workbook = writer.book
        worksheet = writer.sheets["Report"]

        # Header formatting
        header_format = workbook.add_format({
            "bold": True,
            "text_wrap": True,
            "valign": "top",
            "fg_color": "#007bff",
            "font_color": "white",
            "border": 1
        })

        # Completed tasks formatting
        green_format = workbook.add_format({
            "bg_color": "green",
            "font_color": "white",
            "bold": True
        })

        # Formatting column widths and headers
        for col_num, value in enumerate(report_df.columns.values):
            worksheet.write(0, col_num, value, header_format)
            column_len = max(report_df[value].astype(str).apply(len).max(), len(value))
            col_width = column_len + 2
            worksheet.set_column(col_num, col_num, col_width)

        # Styling specific columns based on completion status
        for row_num, completion_status in enumerate(report_df["Completion Status"], start=1):
            if completion_status == "Completed":
                worksheet.write(row_num, report_df.columns.get_loc("Student Name"), report_df.iloc[row_num-1]["Student Name"], green_format)
                worksheet.write(row_num, report_df.columns.get_loc("Completion Status"), completion_status, green_format)

        # Write additional data at the end of the report
        worksheet.write(len(report_df) + 2, 0, f"Week: {specific_week}")
        worksheet.write(len(report_df) + 3, 0, f"Generated: {current_datetime_str}")

    print(f"Excel report generated: {report_excel_filename}")