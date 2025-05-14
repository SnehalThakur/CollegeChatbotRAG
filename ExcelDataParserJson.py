import os

import xlrd
import json
from datetime import datetime
import re
from pathlib import Path
import xlrd
import json
from datetime import datetime
import re


def get_merged_cells(sheet):
    """Get all merged cell ranges in the sheet"""
    return [(merged.rlo, merged.rhi, merged.clo, merged.chi)
            for merged in sheet.merged_cells]


def is_merged_cell(merged_cells, row_idx, col_idx):
    """Check if a cell is part of a merged range"""
    for rlo, rhi, clo, chi in merged_cells:
        if rlo <= row_idx < rhi and clo <= col_idx < chi:
            return (rlo, clo)  # Return top-left cell of merged range
    return None


def extract_section_classroom(cell_value):
    """Extract section (between 'Section :-' and 'CLASSROOM:')
       and classroom (after 'CLASSROOM:')"""
    text = str(cell_value)
    section = classroom = ""

    section_marker = "Section :-"
    classroom_marker = "CLASSROOM:"

    section_start = text.find(section_marker)
    classroom_start = text.find(classroom_marker)

    if section_start != -1 and classroom_start != -1:
        section = text[section_start + len(section_marker):classroom_start].strip()
        classroom = text[classroom_start + len(classroom_marker):].strip()
    elif section_start != -1:
        section = text[section_start + len(section_marker):].strip()
    elif classroom_start != -1:
        classroom = text[classroom_start + len(classroom_marker):].strip()

    section = re.sub(r'[^a-zA-Z0-9]', '', section)
    classroom = classroom.split()[0] if classroom else ""

    return section, classroom


def parse_timetable_sheet(sheet):
    timetable_data = []
    info_row_indexes = [4, 35, 69, 103, 136]  # 0-based for rows 5,36,70,104,137
    current_semester = current_section = current_classroom = ""
    merged_cells = get_merged_cells(sheet)

    for row_idx in range(sheet.nrows):
        # Check if this is an info row
        if row_idx in info_row_indexes:
            semester_cell = get_merged_cell_value(sheet, row_idx, 1)
            current_semester = str(semester_cell).replace("Semester :-", "").strip()

            section_classroom_cell = get_merged_cell_value(sheet, row_idx, 5)
            current_section, current_classroom = extract_section_classroom(section_classroom_cell)
            continue

        # Process timetable rows (Monday to Friday)
        day = str(sheet.cell_value(row_idx, 0)).strip()
        if day in ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]:
            time_slots = {
                2: "09:00-10:00", 3: "10:00-11:00", 4: "11:00-12:00",
                5: "12:00-1:00", 6: "1:00-2:00", 7: "2:00-3:00",
                8: "3:00-4:00", 9: "4:00-5:00"
            }

            for col_idx, time_slot in time_slots.items():
                if col_idx >= sheet.ncols:
                    continue

                # Check if this cell is merged
                merged_info = is_merged_cell(merged_cells, row_idx, col_idx)
                if merged_info:
                    merged_rlo, merged_clo = merged_info
                    # Only process if we're at the top-left cell of the merged range
                    if row_idx != merged_rlo or col_idx != merged_clo:
                        continue

                    # Get the merged cell value
                    cell_value = sheet.cell_value(merged_rlo, merged_clo)

                    # Calculate how many periods this spans (width of merged columns)
                    merged_width = 1
                    for rlo, rhi, clo, chi in merged_cells:
                        if rlo == merged_rlo and clo == merged_clo:
                            merged_width = chi - clo
                            break

                    # Create entries for each time slot in the merged range
                    for i in range(merged_width):
                        current_col = col_idx + i
                        if current_col in time_slots:
                            current_time = time_slots[current_col]

                            if not cell_value or str(cell_value).strip().upper() == "RECESS":
                                continue

                            parts = [p.strip() for p in str(cell_value).split("\n") if p.strip()]
                            subject_abbr = parts[0].split(":")[0].strip() if ":" in parts[0] else parts[0]

                            faculty = ""
                            room = current_classroom

                            if len(parts) > 1 and "(" in parts[1]:
                                faculty = parts[1].strip("()")
                            elif len(parts) > 1:
                                faculty = parts[1]

                            if len(parts) > 2:
                                room = parts[2]

                            timetable_data.append({
                                "semester": current_semester,
                                "section": current_section,
                                "classroom": room,
                                "day": day,
                                "time": current_time,
                                "period": current_col - 2,
                                "subject": subject_abbr,
                                "faculty": faculty,
                                "room": room
                            })
                else:
                    # Normal (non-merged) cell processing
                    cell_value = sheet.cell_value(row_idx, col_idx)
                    if not cell_value or str(cell_value).strip().upper() == "RECESS":
                        continue

                    parts = [p.strip() for p in str(cell_value).split("\n") if p.strip()]
                    subject_abbr = parts[0].split(":")[0].strip() if ":" in parts[0] else parts[0]

                    faculty = ""
                    room = current_classroom

                    if len(parts) > 1 and "(" in parts[1]:
                        faculty = parts[1].strip("()")
                    elif len(parts) > 1:
                        faculty = parts[1]

                    if len(parts) > 2:
                        room = parts[2]

                    timetable_data.append({
                        "semester": current_semester,
                        "section": current_section,
                        "classroom": room,
                        "day": day,
                        "time": time_slot,
                        "period": col_idx - 2,
                        "subject": subject_abbr,
                        "faculty": faculty,
                        "room": room
                    })

    return timetable_data


# [Rest of the code remains the same as previous implementation]
# parse_subjects_section(), parse_excel_file(), and main() functions
# stay exactly the same as in the previous complete solution


def get_merged_cell_value(sheet, row_idx, col_idx):
    """Get value from merged cell"""
    for merged in sheet.merged_cells:
        if merged[0] <= row_idx < merged[1] and merged[2] <= col_idx < merged[3]:
            return sheet.cell_value(merged[0], merged[2])
    return sheet.cell_value(row_idx, col_idx)


def extract_section_classroom(cell_value):
    """Extract section (between 'Section :-' and 'CLASSROOM:')
       and classroom (after 'CLASSROOM:')"""
    text = str(cell_value)
    section = classroom = ""

    # Find the positions of the markers
    section_marker = "Section :-"
    classroom_marker = "CLASSROOM:"

    section_start = text.find(section_marker)
    classroom_start = text.find(classroom_marker)

    if section_start != -1 and classroom_start != -1:
        # Extract section (between the markers)
        section = text[section_start + len(section_marker):classroom_start].strip()
        # Extract classroom (after classroom marker)
        classroom = text[classroom_start + len(classroom_marker):].strip()
    elif section_start != -1:
        # Only section marker found
        section = text[section_start + len(section_marker):].strip()
    elif classroom_start != -1:
        # Only classroom marker found
        classroom = text[classroom_start + len(classroom_marker):].strip()

    # Clean up any remaining whitespace or special characters
    section = re.sub(r'[^a-zA-Z0-9]', '', section)
    classroom = classroom.split()[0] if classroom else ""

    return section, classroom


def parse_timetable_sheet(sheet):
    timetable_data = []
    # 0-based indexes for rows 5,36,70,104,137
    info_row_indexes = [4, 35, 69, 103, 136]
    current_semester = current_section = current_classroom = ""

    for row_idx in range(sheet.nrows):
        # Check if this is an info row
        if row_idx in info_row_indexes:
            # Get semester from merged columns B-E (indexes 1-4)
            semester_cell = get_merged_cell_value(sheet, row_idx, 1)
            current_semester = str(semester_cell).replace("Semester :-", "").strip()

            # Get section and classroom from merged columns F-H (indexes 5-7)
            section_classroom_cell = get_merged_cell_value(sheet, row_idx, 5)
            current_section, current_classroom = extract_section_classroom(section_classroom_cell)
            continue

        # Process timetable rows (Monday to Friday)
        day = str(sheet.cell_value(row_idx, 0)).strip()
        if day in ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]:
            # Time slots mapping to columns (0-based index)
            time_slots = {
                2: "09:00-10:00",  # Column C
                3: "10:00-11:00",  # Column D
                4: "11:00-12:00",  # Column E
                5: "12:00-1:00",  # Column F
                6: "1:00-2:00",  # Column G (Recess)
                7: "2:00-3:00",  # Column H
                8: "3:00-4:00",  # Column I
                9: "4:00-5:00"  # Column J
            }

            for col_idx, time_slot in time_slots.items():
                if col_idx >= sheet.ncols:  # Skip if column doesn't exist
                    continue

                cell_value = get_merged_cell_value(sheet, row_idx, col_idx)
                if not cell_value or str(cell_value).strip().upper() == "RECESS":
                    continue

                # Parse subject, faculty, and room
                parts = [p.strip() for p in str(cell_value).split("\n") if p.strip()]
                subject_abbr = parts[0].split(":")[0].strip() if ":" in parts[0] else parts[0]

                faculty = ""
                room = current_classroom

                if len(parts) > 1 and "(" in parts[1]:
                    faculty = parts[1].strip("()")
                elif len(parts) > 1:
                    faculty = parts[1]

                if len(parts) > 2:
                    room = parts[2]

                timetable_data.append({
                    "semester": current_semester.strip(),
                    "section": current_section.strip(),
                    "classroom": room.strip(),
                    "day": day.strip(),
                    "time": time_slot.strip(),
                    "period": col_idx - 2,  # 0-based period index
                    "subject": subject_abbr.strip(),
                    "faculty": faculty.strip(),
                    "room": room.strip()
                })

    return timetable_data


def parse_subjects_section(sheet):
    subject_data = []
    current_type = None

    # Find the THEORY SUBJECT/PRACTICAL section
    found_section = False
    for row_idx in range(sheet.nrows):
        row = [get_merged_cell_value(sheet, row_idx, col) for col in range(sheet.ncols)]
        if not any(row):  # Skip empty rows
            continue

        # Check for section headers
        if row[0] and "THEORY SUBJECT" in str(row[0]):
            current_type = "theory"
            found_section = True
            continue
        elif row[0] and "PRACTICAL" in str(row[0]):
            current_type = "practical"
            found_section = True
            continue

        if not found_section:
            continue

        # Skip non-data rows
        if not row[0] or "Course code:" in str(row[0]):
            continue

        # Parse subject information
        course_code = ""
        subject_abbr = ""
        faculty_name = ""
        faculty_abbr = ""

        # Course code (merged columns A and B)
        if row[0] and row[1]:
            course_code = f"{row[0]}: {row[1]}" if row[1] else str(row[0])
        elif row[0]:
            course_code = str(row[0])

        # Subject abbreviation (column D)
        if len(row) > 3 and row[3]:
            subject_abbr = str(row[3])

        # Faculty name (column E)
        if len(row) > 4 and row[4]:
            faculty_name = str(row[4])

        # Faculty abbreviation (column F)
        if len(row) > 5 and row[5]:
            faculty_abbr = str(row[5])

        # For practicals (merged columns G, H, I)
        if current_type == "practical" and len(row) > 6:
            supporting_staff = []
            if row[6] and row[7] and row[8]:
                supporting_staff = [str(row[6]), str(row[7]), str(row[8])]
            elif row[6] and row[7]:
                supporting_staff = [str(row[6]), str(row[7])]
            elif row[6]:
                supporting_staff = [str(row[6])]

            if supporting_staff:
                faculty_name += " (" + ", ".join(supporting_staff) + ")"

        if course_code and subject_abbr:
            subject_data.append({
                "course_code": course_code.strip(),
                "subject_abbreviation": subject_abbr.strip(),
                "subject_type": current_type.strip(),
                "faculty_name": faculty_name.strip(),
                "faculty_abbreviation": faculty_abbr.strip()
            })

    return subject_data


def parse_excel_file(file_path):
    try:
        workbook = xlrd.open_workbook(file_path)
        all_timetable_data = []
        all_subject_data = []

        for sheet_name in workbook.sheet_names():
            sheet = workbook.sheet_by_name(sheet_name)

            # Parse timetable data from each sheet
            timetable_data = parse_timetable_sheet(sheet)
            all_timetable_data.extend(timetable_data)

            # Parse subject data from each sheet
            subject_data = parse_subjects_section(sheet)
            all_subject_data.extend(subject_data)

        return {
            "timetable": all_timetable_data,
            "subjects": all_subject_data
        }
    except xlrd.XLRDError as e:
        raise Exception(f"Excel file parsing error: {str(e)}")
    except Exception as e:
        raise Exception(f"Error processing file: {str(e)}")


def save_uploaded_file(uploaded_file, save_dir="uploaded_files"):
    """Save uploaded file to disk and return the file path"""
    try:
        # Create directory if it doesn't exist
        Path(save_dir).mkdir(parents=True, exist_ok=True)

        # Create file path
        file_path = os.path.join(save_dir, uploaded_file.name)

        # Save file
        with open(file_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        return file_path
    except Exception as e:
        print(f"Error saving file: {e}")
        return None


def excelToJsonConverter(filepath):
    # Input Excel file path
    input_file = r'C:\Users\snehal\PycharmProjects\ChatbotRAG\data\UG CLASS CTECH TT_odd (24-25)_ 3RD-5TH-7TH SEM_TO STAFF_V3_9-08-2024-w.e.f. 12-08-2024.xls'
    if filepath:
        input_file = filepath
    try:
        print(f"Processing {input_file}...")
        result = parse_excel_file(input_file)

        # Generate timestamp for output files
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        # Prepare output dictionary
        output_files = {
            "timetable": {
                "filename": f"timetable_{timestamp}.json",
                "filepath": os.path.abspath(f"timetable_{timestamp}.json")
            },
            "subjects": {
                "filename": f"subjects_{timestamp}.json",
                "filepath": os.path.abspath(f"subjects_{timestamp}.json")
            }
        }

        # Save timetable data
        timetable_output = f"timetable_{timestamp}.json"
        with open(timetable_output, 'w', encoding='utf-8') as f:
            json.dump(result["timetable"], f, indent=2, ensure_ascii=False)
        print(f"Timetable data saved to {timetable_output}")

        # Save subject data
        subjects_output = f"subjects_{timestamp}.json"
        with open(subjects_output, 'w', encoding='utf-8') as f:
            json.dump(result["subjects"], f, indent=2, ensure_ascii=False)
        print(f"Subject data saved to {subjects_output}")

        print("Processing completed successfully!")

        return output_files


    except Exception as e:
        print(f"Error: {str(e)}")
        print("Make sure you have xlrd installed (pip install xlrd)")


if __name__ == "__main__":
    # Install xlrd if not available
    try:
        import xlrd
    except ImportError:
        print("xlrd package is required. Installing...")
        import subprocess
        import sys

        subprocess.check_call([sys.executable, "-m", "pip", "install", "xlrd"])

    excelToJsonConverter()