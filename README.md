# Schedule2ics4Calvin
 - This is a script that converts Calvin course schedule downloaded from Workday to .ics, which you can import your schedule into Calendar.
 - Export your courses as EXCEL from workday.calvin.edu
 - Make sure the name is "View_My_Courses.xlsx"
 - INSTALL PYTHON IN YOUR COMPUTER
 - RUN ```pip3 install -r requirements.txt``` in cmd/terminal
 - RUN the script ```python3 app.py``` and IMPORT "courses.ics" to your Calendar
### NOW! You can access the web at [Schedule2ICS](http://127.0.0.1/).

---

## Project Overview

Schedule2ics4Calvin is a small Python and Flask project that converts a Calvin University Workday course schedule export into an `.ics` calendar file. The generated `courses.ics` file can be imported into common calendar applications such as Apple Calendar, Google Calendar, Outlook, and other apps that support iCalendar files.

The app is designed for Calvin students who want to move their class schedule from Workday into a personal calendar without manually creating each class meeting.

## What It Does

- Accepts an Excel schedule export from Workday.
- Reads course meeting details from the `View My Courses` worksheet.
- Extracts course names, dates, meeting days, start times, end times, and locations.
- Converts recurring class meetings into iCalendar events.
- Returns a downloadable `courses.ics` file.

## Requirements

- Python 3
- pip
- A Calvin Workday course schedule exported as an Excel `.xlsx` file

Python dependencies are listed in `requirements.txt`.

## Installation

Clone or download this repository, then install the required Python packages:

```bash
pip3 install -r requirements.txt
```

If your system uses `python` and `pip` instead of `python3` and `pip3`, use those commands instead.

## How to Export Your Schedule from Workday

1. Go to `workday.calvin.edu`.
2. Open your course schedule view.
3. Export your courses as an Excel file.
4. Save the file as `View_My_Courses.xlsx`.

The converter expects the workbook to contain a sheet named `View My Courses`.

## Running the Web App

Start the Flask app:

```bash
python3 app.py
```

By default, the app runs on port `15646`:

```text
http://127.0.0.1:15646/
```

Open that address in your browser, upload your Workday Excel file, and download the generated `courses.ics` file.

## Importing the Calendar File

After downloading `courses.ics`, import it into your preferred calendar app.

Common options include:

- Apple Calendar: File > Import
- Google Calendar: Settings > Import & export
- Outlook: Add calendar > Upload from file

Exact import steps may vary depending on your calendar app and device.

## Project Structure

```text
.
├── app.py
├── courses2ics.py
├── requirements.txt
├── README.md
├── static/
│   └── illus_webstation_enabled.jpg
└── templates/
    ├── courses2schedule.html
    ├── error.html
    └── index.html
```

### Main Files

- `app.py`: Flask application, upload route, file validation, and `.ics` download response.
- `courses2ics.py`: Core conversion logic for parsing Workday course data and writing iCalendar events.
- `templates/courses2schedule.html`: Upload page for the Excel file.
- `templates/index.html`: Landing page.
- `requirements.txt`: Python package dependencies.

## Expected Input Format

The converter expects an Excel workbook with a worksheet named:

```text
View My Courses
```

The current parser is tailored to the column layout used by Calvin Workday exports. If Workday changes the exported spreadsheet format, the parser in `courses2ics.py` may need to be updated.

## Output

The app generates an iCalendar file named:

```text
courses.ics
```

Each course meeting is written as a recurring weekly calendar event using the `America/Detroit` timezone.

## Troubleshooting

### The upload fails

Make sure the file is an `.xlsx` file and that it was exported from Calvin Workday.

### The app cannot find the worksheet

Make sure the workbook contains a sheet named `View My Courses`.

### Calendar events look incorrect

Check that the Workday export includes meeting days, times, dates, and locations. If the Workday export format has changed, the parsing logic in `courses2ics.py` may need adjustment.

### The server address in the original notes does not load

The app currently runs on port `15646`, so use:

```text
http://127.0.0.1:15646/
```

instead of only `http://127.0.0.1/`.

## Notes

This project is specific to Calvin University Workday schedule exports. It may not work with schedule files from other schools or with manually edited spreadsheets unless they match the expected Workday format.
