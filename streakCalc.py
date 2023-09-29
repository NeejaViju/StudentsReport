import pandas as pd
import os
# Function to get streak
def calculate_streak(row):
    streak = 0
    for status in reversed(row):
        if status == "Done":
            streak += 1
        else:
            return streak
    return streak

# List of files
files = [f"Week{str(i).zfill(2)}_report.xlsx" for i in range(1, 53)]

# Read the first file to initialize the final dataframe
df_final = pd.read_excel(files[0], engine='openpyxl')
df_final = df_final[['Student ID', 'Student Name', 'Completion Status']]
df_final.rename(columns={'Completion Status': 'Week 1'}, inplace=True)

# Loop through other files and merge data
for idx, file in enumerate(files[1:], 2):
    if os.path.exists(file):  # Check if file exists before reading
        df = pd.read_excel(file, engine='openpyxl')
        df = df[['Student ID', 'Completion Status']]
        df.rename(columns={'Completion Status': f"Week {idx}"}, inplace=True)
        df_final = pd.merge(df_final, df, on="Student ID", how="outer")

# Calculate the streak for each student
df_final["Streak"] = df_final.iloc[:, 2:].apply(calculate_streak, axis=1)

# Save to a new Excel file
df_final.to_excel("Sample_Task_name/Consolidated_Report.xlsx", index=False, engine='openpyxl')
