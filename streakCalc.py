import pandas as pd
import os

# List of files
reports_directory = "Reports"
files = [os.path.join(reports_directory, f"Week{str(i).zfill(2)}_report.xlsx") for i in range(1, 53)]

# Read the first file to initialize the final dataframe
if os.path.exists(files[0]):
    df_final = pd.read_excel(files[0], engine='openpyxl')

    # Ensure 'Student ID' column doesn't contain any NaN values
    df_final.dropna(subset=['Student ID'], inplace=True)
    
    # Filter students by ID pattern
    df_final = df_final[df_final['Student ID'].str.startswith(('PPP', 'PPF'))]
    
    df_final = df_final[['Student ID', 'Student Name', 'Streak', 'Completion Status']]
    df_final.rename(columns={'Completion Status': 'Week 1', 'Streak': 'Streak Week 1'}, inplace=True)

    # Loop through other files and merge streak and completion status data
    for idx, file in enumerate(files[1:], 2):
        if os.path.exists(file):  # Check if file exists before reading
            df = pd.read_excel(file, engine='openpyxl')
            
            # Ensure 'Student ID' column doesn't contain any NaN values
            df.dropna(subset=['Student ID'], inplace=True)
            
            df = df[df['Student ID'].str.startswith(('PPP', 'PPF'))]
            df_data = df[['Student ID', 'Streak', 'Completion Status']]
            df_data = df_data.rename(columns={'Streak': f'Streak Week {idx}', 'Completion Status': f'Week {idx}'})
            df_final = pd.merge(df_final, df_data, on="Student ID", how="outer")

    # Calculate Total Streak
    streak_cols = [col for col in df_final.columns if 'Streak' in col]
    df_final['Total Streak'] = df_final[streak_cols].sum(axis=1)

    # Create Weekly Streak Sheet Data
    df_weekly_streak = df_final[['Student ID', 'Student Name'] + streak_cols]
    
    # Create Status & Streak Sheet Data
    status_cols = [col for col in df_final.columns if 'Week' in col and 'Streak' not in col]
    df_status_streak = df_final[['Student ID', 'Student Name'] + status_cols + ['Total Streak']]
    
    # Save to a new Excel file with two sheets
    with pd.ExcelWriter("Sample_Task_name/Consolidated_Report.xlsx", engine='openpyxl') as writer:
        df_status_streak.to_excel(writer, sheet_name='Status & Streak', index=False)
        df_weekly_streak.to_excel(writer, sheet_name='Weekly Streak', index=False)

else:
    print(f"File {files[0]} not found!")
