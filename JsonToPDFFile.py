import pdfkit
import json
from collections import defaultdict
import os
import platform


class TimetableGenerator:
    def __init__(self):
        self.wkhtmltopdf_path = self._find_wkhtmltopdf()
        self.config = pdfkit.configuration(wkhtmltopdf=self.wkhtmltopdf_path) if self.wkhtmltopdf_path else None

    def _find_wkhtmltopdf(self):
        """Find wkhtmltopdf executable path based on OS."""
        paths = []
        if platform.system() == 'Windows':
            paths = [
                r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe',
                r'C:\Program Files (x86)\wkhtmltopdf\bin\wkhtmltopdf.exe'
            ]
        elif platform.system() in ['Linux', 'Darwin']:
            paths = [
                '/usr/bin/wkhtmltopdf',
                '/usr/local/bin/wkhtmltopdf',
                '/opt/homebrew/bin/wkhtmltopdf'
            ]

        for path in paths:
            if os.path.exists(path):
                return path

        try:
            if platform.system() == 'Windows':
                import subprocess
                result = subprocess.run(['where', 'wkhtmltopdf'], capture_output=True, text=True).stdout.strip()
            else:
                result = os.popen('which wkhtmltopdf').read().strip()

            if result and os.path.exists(result):
                return result
        except:
            return None

    def load_json(self, file_path):
        """Load JSON data from file."""
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except FileNotFoundError:
            raise FileNotFoundError(f"JSON file not found at {file_path}")
        except json.JSONDecodeError:
            raise ValueError(f"Invalid JSON format in {file_path}")

    def organize_by_semester_day(self, timetable_data, subjects_data):
        """Organize data by day, then by semester."""
        organized_data = defaultdict(lambda: defaultdict(list))
        abbreviations = {}

        for entry in timetable_data:
            semester = entry.get("semester", "Other")
            subject_info = next(
                (item for item in subjects_data if item["subject_abbreviation"] == entry["subject"]),
                {
                    "course_code": entry["subject"],
                    "faculty_name": entry.get("faculty", ""),
                    "faculty_abbreviation": entry.get("faculty", "")
                }
            )

            organized_entry = {
                "time": entry["time"],
                "subject": entry["subject"],
                "subject_full": subject_info["course_code"],
                "faculty": subject_info["faculty_abbreviation"],
                "faculty_full": subject_info["faculty_name"],
                "room": entry.get("room", ""),
                "section": entry.get("section", "")
            }

            organized_data[entry["day"]][semester].append(organized_entry)

            if entry["subject"] not in abbreviations:
                faculty_display = (f"{subject_info['faculty_name']} ({subject_info['faculty_abbreviation']})"
                                   if subject_info['faculty_name'] else subject_info['faculty_abbreviation'])
                abbreviations[entry["subject"]] = {
                    "full_form": subject_info["course_code"],
                    "faculty": faculty_display
                }

        return organized_data, abbreviations

    def generate_html(self, organized_data, abbreviations):
        """Generate HTML with semester subsections under each day."""
        day_order = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
        semester_order = ["3", "5", "7", "Other"]  # Customize as needed

        html_content = """
        <!DOCTYPE html>
        <html>
        <head>
            <title>Semester-wise Timetable</title>
            <meta charset="UTF-8">
            <style>
                body {
                    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                    line-height: 1.6;
                    margin: 0;
                    padding: 20px;
                    color: #333;
                }
                .container {
                    max-width: 1000px;
                    margin: 0 auto;
                }
                h1 {
                    text-align: center;
                    color: #2c3e50;
                    margin-bottom: 30px;
                    padding-bottom: 10px;
                    border-bottom: 2px solid #3498db;
                }
                .day-section {
                    margin-bottom: 40px;
                }
                .day-header {
                    color: #3498db;
                    margin-top: 25px;
                    padding-bottom: 5px;
                    border-bottom: 1px solid #ddd;
                }
                .semester-section {
                    margin: 15px 0 25px 20px;
                    padding: 10px;
                    background-color: #f5f5f5;
                    border-radius: 5px;
                }
                .semester-header {
                    color: #2c3e50;
                    font-weight: bold;
                    margin-bottom: 10px;
                }
                .lecture-list {
                    list-style-type: none;
                    padding-left: 0;
                }
                .lecture-item {
                    margin: 8px 0;
                    padding: 8px 12px;
                    background-color: white;
                    border-radius: 4px;
                    display: flex;
                    flex-wrap: wrap;
                    align-items: center;
                }
                .time {
                    font-weight: bold;
                    color: #7f8c8d;
                    min-width: 120px;
                }
                .subject {
                    font-weight: bold;
                    color: #2c3e50;
                    min-width: 150px;
                }
                .faculty {
                    flex-grow: 1;
                }
                .room {
                    color: #16a085;
                    font-style: italic;
                    min-width: 100px;
                    text-align: right;
                }
                .section {
                    background-color: #e3f2fd;
                    padding: 2px 6px;
                    border-radius: 3px;
                    font-size: 0.9em;
                    margin-left: 10px;
                }
                .abbrev-section {
                    margin-top: 50px;
                    page-break-before: always;
                    padding-top: 20px;
                }
                .abbrev-section h2 {
                    color: #2c3e50;
                    border-bottom-color: #2c3e50;
                }
                .abbrev-list {
                    column-count: 2;
                    column-gap: 30px;
                }
                .abbrev-item {
                    break-inside: avoid;
                    margin-bottom: 10px;
                    padding: 5px;
                }
                .abbrev-item strong {
                    color: #3498db;
                }
                @media print {
                    body {
                        padding: 0;
                        font-size: 11pt;
                    }
                    .container {
                        max-width: 100%;
                    }
                    .semester-section {
                        page-break-inside: avoid;
                    }
                }
            </style>
        </head>
        <body>
            <div class="container">
                <h1>Semester-wise Timetable</h1>
        """

        # Add timetable organized by day and semester
        for day in day_order:
            if day in organized_data:
                html_content += f"""
                <div class="day-section">
                    <h2 class="day-header">{day}</h2>
                """

                for semester in semester_order:
                    if semester in organized_data[day]:
                        html_content += f"""
                        <div class="semester-section">
                            <div class="semester-header">Semester {semester}</div>
                            <ul class="lecture-list">
                        """

                        for entry in sorted(organized_data[day][semester], key=lambda x: x["time"]):
                            section_display = f'<span class="section">Sec {entry["section"]}</span>' if entry[
                                "section"] else ""
                            html_content += f"""
                            <li class="lecture-item">
                                <span class="time">{entry['time']}</span>
                                <span class="subject">{entry['subject']}</span>
                                <span class="faculty">{entry['faculty_full']}</span>
                                <span class="room">Room {entry['room']}</span>
                                {section_display}
                            </li>
                            """

                        html_content += """
                            </ul>
                        </div>
                        """

                html_content += "</div>"

        # Add abbreviations section
        html_content += """
                <div class="abbrev-section">
                    <h2>Abbreviations and Faculty</h2>
                    <div class="abbrev-list">
        """

        for abbrev, details in sorted(abbreviations.items()):
            html_content += f"""
                        <div class="abbrev-item">
                            <strong>{abbrev}:</strong> {details['full_form']} - {details['faculty']}
                        </div>
            """

        html_content += """
                    </div>
                </div>
            </div>
        </body>
        </html>
        """

        return html_content

    def generate_pdf(self, output_file="semester_timetable.pdf"):
        """Main method to generate PDF."""
        if not self.config:
            print("Error: wkhtmltopdf not found. Please install it first.")
            print("Windows: Download from https://wkhtmltopdf.org/downloads.html")
            print("Mac: brew install wkhtmltopdf")
            print("Linux: sudo apt-get install wkhtmltopdf")
            return False

        try:
            timetable_data = self.load_json('timetable_20250501_144326.json')
            subjects_data = self.load_json('subjects_20250501_144326.json')

            organized_data, abbreviations = self.organize_by_semester_day(timetable_data, subjects_data)
            html_content = self.generate_html(organized_data, abbreviations)

            options = {
                'encoding': 'UTF-8',
                'quiet': '',
                'page-size': 'A4',
                'margin-top': '15mm',
                'margin-right': '15mm',
                'margin-bottom': '15mm',
                'margin-left': '15mm',
                'enable-local-file-access': None
            }

            pdfkit.from_string(html_content, output_file, configuration=self.config, options=options)
            print(f"Successfully generated PDF: {output_file}")
            return True

        except FileNotFoundError as e:
            print(f"Error: {str(e)}")
        except ValueError as e:
            print(f"Error: {str(e)}")
        except Exception as e:
            print(f"An unexpected error occurred: {str(e)}")

        return False


if __name__ == "__main__":
    generator = TimetableGenerator()
    generator.generate_pdf("semester_wise_timetable.pdf")