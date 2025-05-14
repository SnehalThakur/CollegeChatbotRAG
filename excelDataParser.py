import pandas as pd

# Load the Excel file
file_path = r'C:\Users\snehal\PycharmProjects\ChatbotRAG\data\UG CLASS CTECH TT_odd (24-25)_ 3RD-5TH-7TH SEM_TO STAFF_V3_9-08-2024-w.e.f. 12-08-2024.xls'
sheet_name = 'Sheet1'  # Replace with the correct sheet name if different

# Read the Excel file
df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)

# Initialize variables to store the timetable data
timetable_data = []

# Define the days of the week
days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']

# Extract faculty mapping from the table
faculty_mapping = {}
for i, row in df.iterrows():
    if pd.notna(row[0]) and "Course code: Course Name" in str(row[0]):
        # This is the start of the faculty mapping table
        for j in range(i + 1, len(df)):
            if pd.notna(df.iloc[j, 0]):
                subject_abbr = df.iloc[j, 2]  # Subject Abbreviation
                faculty_name = df.iloc[j, 3]  # Faculty Full Name
                faculty_mapping[subject_abbr] = faculty_name
            else:
                break  # End of faculty mapping table

# Iterate through the rows to extract the timetable information
for i, row in df.iterrows():
    if row[0] in days:
        day = row[0]
        for j in range(1, len(row)):
            if pd.notna(row[j]):
                time_slot = df.iloc[i - 1, j]
                subject_info = row[j]
                if isinstance(subject_info, str):
                    subject_parts = subject_info.split('\n')
                    subject_abbr = subject_parts[0].strip()
                    classroom = subject_parts[2].strip() if len(subject_parts) > 2 else ''
                    # Get the full faculty name from the mapping
                    teacher = faculty_mapping.get(subject_abbr, 'Unknown Teacher')
                    timetable_data.append([day, time_slot, subject_abbr, teacher, classroom])

# Create a DataFrame from the extracted data
timetable_df = pd.DataFrame(timetable_data, columns=['Day', 'Time', 'Subject', 'Teacher', 'Classroom'])

# Write the DataFrame to a text file
output_file = 'timetable.txt'
with open(output_file, 'w') as f:
    for _, row in timetable_df.iterrows():
        f.write(f"{row['Day']}, {row['Time']}, {row['Subject']}, {row['Teacher']}, {row['Classroom']}\n")

print(f"Timetable has been written to {output_file}")