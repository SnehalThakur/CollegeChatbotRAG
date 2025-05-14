import json
from collections import defaultdict


class TimetableProcessor:
    def __init__(self):
        self.timetable_data = []
        self.subjects_data = []

    def load_json(self, timetable_path, subjects_path):
        """Load both JSON files"""
        with open(timetable_path, 'r') as f:
            self.timetable_data = json.load(f)
        with open(subjects_path, 'r') as f:
            self.subjects_data = json.load(f)

    def process_data(self):
        """Organize data by day -> semester -> lectures"""
        organized_data = defaultdict(lambda: defaultdict(list))
        abbreviations = {}

        for entry in self.timetable_data:
            # Get semester or default to "Other"
            semester = entry.get("semester", "Other")

            # Find matching subject info
            subject_info = next(
                (item for item in self.subjects_data
                 if item["subject_abbreviation"] == entry["subject"]),
                {
                    "course_code": entry["subject"],
                    "faculty_name": entry.get("faculty", ""),
                    "faculty_abbreviation": entry.get("faculty", "")
                }
            )

            # Create lecture entry
            lecture = {
                "time": entry["time"],
                "subject": entry["subject"],
                "subject_full": subject_info["course_code"],
                "faculty": subject_info["faculty_abbreviation"],
                "faculty_full": subject_info["faculty_name"],
                "room": entry.get("room", ""),
                "section": entry.get("section", "")
            }

            organized_data[entry["day"]][semester].append(lecture)

            # Build abbreviations dictionary
            if entry["subject"] not in abbreviations:
                faculty_display = (f"{subject_info['faculty_name']} ({subject_info['faculty_abbreviation']})"
                                   if subject_info['faculty_name'] else subject_info['faculty_abbreviation'])
                abbreviations[entry["subject"]] = {
                    "full_form": subject_info["course_code"],
                    "faculty": faculty_display
                }

        return organized_data, abbreviations

    def generate_structured_text(self):
        """Generate clean text output organized by day and semester"""
        organized_data, abbreviations = self.process_data()

        # Define display order
        day_order = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
        semester_order = ["3", "5", "7", "Other"]

        output_lines = ["TIMETABLE STRUCTURED DATA\n"]

        # Add timetable by day and semester
        for day in day_order:
            if day in organized_data:
                output_lines.append(f"\n=== {day.upper()} ===")

                for semester in semester_order:
                    if semester in organized_data[day]:
                        output_lines.append(f"\nSemester {semester}:")

                        for lecture in sorted(organized_data[day][semester], key=lambda x: x["time"]):
                            section = f" (Section {lecture['section']})" if lecture["section"] else ""
                            output_lines.append(
                                f"{lecture['time']}: {lecture['subject']} - {lecture['faculty_full']}" +
                                f" - Room {lecture['room']}{section}"
                            )

        # Add abbreviations
        output_lines.append("\n\nABBREVIATIONS:")
        for abbrev, details in sorted(abbreviations.items()):
            output_lines.append(f"{abbrev}: {details['full_form']} - {details['faculty']}")

        return "\n".join(output_lines)


if __name__ == "__main__":
    processor = TimetableProcessor()
    processor.load_json('timetable_20250501_144326.json', 'subjects_20250501_144326.json')
    structured_text = processor.generate_structured_text()

    # Save to text file (or pass directly to your model)
    with open('timetable_structured.txt', 'w') as f:
        f.write(structured_text)

    print("Structured timetable data generated successfully.")