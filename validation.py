import os
import datetime
import pandas as pd
import xlsxwriter
import subprocess
import plotly.graph_objects as go

#---------automated git pull code 
# try:
#     result = subprocess.run(['git', 'pull'], check=True, text=True, capture_output=True)
#     print(result.stdout)
#     if result.stderr:
#         print("Error output:", result.stderr)
# except subprocess.CalledProcessError as e:
#     print(f"Error pulling from git: {e}")
#     if e.stderr:
#         print("Detailed error:", e.stderr.strip())
#     # Decide how to handle the error: e.g., exit the script
#     exit(1)

parent_folder_path = "Students"

# Default tasks for each week
default_tasks = {
    "Week01": ["Git_Task","Index_File_Updation","create_Html_file_on_Name","dulingo_update"],
    "Week02": ["create_wordpress_blog_and_7articles","update_linkedin_with-photo","create_canva-menu","download_figma_and_install"],
    "Week03" :["Fibonacci_Sequence","Calculator","Tic_Tac_Toe","Generative AI"],
    "Week04": ["Error messages_200 OK_404 Not Found","Google Chrome Extensions","Tweet_AI tool_Futurepedia","Download_Install_ Google Chrome Canary Version"],
    "Week05": ["Create Framer Site","Create A Snake Game","Create Paper Prototype","Create Social Media Profile Using CSS"] ,  
    "Week06": ["summary of Fermi's paradox","summary of Drake's Equation","Create a table  using CSS Grid","create small project using CSS Flexbox"],
    "Week07": ["create small project using CSS box sizing","create small project using CSS Box Shadow","create small project ussing CSS Border Radius","create small project using CSS Justify content"],
    "Week08": ["Learn prompt Engineering","Create sidebar","Dig Hugging Face","Learn Javascript"]
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

weeks_to_report = ["Week01", "Week02", "Week03", "Week04", "Week05", "Week06", "Week07", "Week08"]  # Add other weeks as needed

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

      # Calculate streak for each student separately
    report_df["Streak"] = report_df.groupby('Student ID')["Completion Status"].transform(
        lambda x: (x == "Completed").astype(int).cumsum()
    )

    current_dir = os.path.dirname(os.path.abspath(__file__))
    folder_path = os.path.join(current_dir,'Reports')

    if not os.path.exists(folder_path):
        os.makedirs(folder_path)


    # Create the report filename
    report_excel_filename = os.path.join(folder_path, f"{specific_week}_report.xlsx")







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


#analysis Report----------------------------------------------------

def analyze_report(specific_week):
    report_excel_filename = f"{specific_week}_report.xlsx"
    
    # Load the generated report
    report_df = pd.read_excel(report_excel_filename)

    analysis_dict = {"Week": specific_week}
    
    # 1. Number and Percentage of Completed Students
    completed_students = report_df[report_df["Completion Status"] == "Completed"]
    num_completed_students = len(completed_students)
    percent_completed_students = (num_completed_students / len(report_df)) * 100

    analysis_dict["Completed Students"] = num_completed_students
    analysis_dict["Completed Percentage"] = percent_completed_students

    # 2. Number and Percentage of Pending Students
    pending_students = report_df[report_df["Completion Status"] == "Pending"]
    num_pending_students = len(pending_students)
    percent_pending_students = (num_pending_students / len(report_df)) * 100

    analysis_dict["Pending Students"] = num_pending_students
    analysis_dict["Pending Percentage"] = percent_pending_students

    # 3. Tasks Most Frequently Pending
    all_pending_tasks = report_df["Pending Task"].dropna().str.split(", ").sum()
    task_counts = pd.Series(all_pending_tasks).value_counts()
    for task, count in task_counts.items():
        analysis_dict[task] = count

    return analysis_dict

# Define the weeks you want to analyze
weeks_to_analyze = ["Week01", "Week02"]  # Add or remove weeks as per your data

results = []
for week in weeks_to_analyze:
    week_analysis = analyze_report(week)
    results.append(week_analysis)

# Convert the list of dictionaries to DataFrame
df_results = pd.DataFrame(results)

# Save to Excel
with pd.ExcelWriter('Analysis_Report.xlsx') as writer:
    df_results.to_excel(writer, sheet_name="Analysis", index=False)

print("Analysis saved to Analysis_Report.xlsx")

def create_chart(df_results):
    # Create a bar chart with completed and pending students
    fig = go.Figure()
    fig.add_trace(go.Bar(x=df_results["Week"], y=df_results["Completed Students"], name='Completed Students', marker_color='green'))
    fig.add_trace(go.Bar(x=df_results["Week"], y=df_results["Pending Students"], name='Pending Students', marker_color='red'))

    # Update layout for better appearance
    fig.update_layout(
        title='Students Status Analysis',
        xaxis=dict(title='Week'),
        yaxis=dict(title='Number of Students'),
        barmode='group'
    )
    
    # Convert plotly figure to HTML and return
    return fig.to_html(full_html=False)

# Generate the chart
chart_html = create_chart(df_results)

# Define a basic Bootstrap template for the HTML report
# Update HTML_TEMPLATE to include a placeholder for the chart
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Analysis Report</title>
    <link href="https://maxcdn.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css" rel="stylesheet">
</head>
<body>
    <div class="container">
        <h1 class="my-4 text-center">Analysis Report</h1>
        {chart}
        <div class="table-responsive mt-5">
            {table}
        </div>
    </div>
</body>
</html>
"""

# Convert the analysis DataFrame to HTML
html_content = df_results.to_html(classes='table table-bordered table-hover', table_id='analysisTable')

# Using JavaScript to ensure the table takes the full width
html_content += """
<script>
    document.getElementById('analysisTable').style.width = '100%';
</script>
"""

# Replace the placeholders in the template with the table and the chart
html_report = HTML_TEMPLATE.format(table=html_content, chart=chart_html)

# Define the name for the HTML report
html_report_filename = "Analysis_Report.html"

# Save the HTML content to a file
with open(html_report_filename, 'w', encoding='utf-8') as file:
    file.write(html_report)

print(f"Analysis saved to {html_report_filename}")

