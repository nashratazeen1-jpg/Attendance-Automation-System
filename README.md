# Attendance Automation System

A Python-based **Attendance Automation System** that processes attendance data from Excel files, automatically calculates student-wise attendance, generates monthly result files, and prepares a final semester Progressive Attendance (PA) Sheet.

The main goal is to **reduce faculty workload, save time, and minimize manual calculation errors** caused by repeatedly applying Excel formulas.

## Project Overview 🚀

In the traditional process, attendance data is maintained in Excel and calculations such as attendance percentage, monthly results, and semester-wise PA sheets are prepared manually.

This project automates that workflow using Python.

Instead of manually applying formulas and preparing reports for each month, the system:

- Takes attendance data from an Excel file.
- Processes the data automatically.
- Calculates student-wise attendance percentages.
- Generates monthly attendance result files.
- Combines monthly results into a final semester PA Sheet.

### `monthly.py`

Handles the **monthly attendance calculation**.

It processes the attendance Excel data for a selected month, performs the required calculations, and generates the corresponding monthly result Excel file.

Example monthly outputs:

```text
dec_result.xlsx
jan_result.xlsx
feb_result.xlsx
mar_result.xlsx
```

### `pa_sheet.py`

Handles the **semester-level attendance processing**.

It takes the generated monthly result files and combines the attendance data to create the final **Progressive Attendance (PA) Sheet**, including the overall semester attendance information.

Example:

```text
PA_Sheet_Result.xlsx
```

## How the System Works ⚒️

### Step 1 — Provide Input

The faculty provides the attendance data in the required **Excel format**.

### Step 2 — Select Month

The system processes the attendance according to the selected month.

The supported month types are:

- 28 Days
- 30 Days
- 31 Days
- 15 Days

### Step 3 — Automatic Processing

`monthly.py` reads and processes the Excel attendance data automatically.

It calculates student-wise attendance and generates the required monthly result file.

### Step 4 — Monthly Result Generation

The system can generate separate result files for different months, such as:

- December
- January
- February
- March

Each result file contains the corresponding student attendance percentage.

### Step 5 — Generate Final PA Sheet

`pa_sheet.py` combines the generated monthly attendance result files and calculates the overall semester attendance.

The final output is:

```text
PA_Sheet_Result.xlsx
```

This produces a structured PA Sheet ready for academic use and submission.

---

## Key Features

- **Attendance Automation** — Automates calculations that would otherwise require manual Excel formulas.
- **Monthly Result Generation** — Creates separate attendance result files for each month.
- **Student-wise Attendance Calculation** — Calculates attendance percentage for individual students.
- **Semester PA Sheet Generation** — Combines monthly results into one final semester report.
- **Reduced Faculty Workload** — Removes repetitive manual calculation and report preparation work.
- **Reduced Errors** — Minimizes errors caused by manual calculations and formula handling.
- **Simple Workflow** — Designed to make the process easier for non-technical users.

---

## Technology Used

- **Python** — Core processing and automation
- **Pandas** — Reading, processing, and analyzing Excel attendance data
- **Tkinter** — Graphical User Interface (GUI)
- **Microsoft Excel** — Input and output format for attendance records

---

## System Flow 🪜

```text
Attendance Excel File
        │
        ▼
   Select Month
        │
        ▼
   monthly.py
        │
        ▼
Process Attendance Data
        │
        ▼
Calculate Attendance %
        │
        ▼
Monthly Result Excel Files
        │
        ▼
   pa_sheet.py
        │
        ▼
Combine Monthly Results
        │
        ▼
Final Semester PA Sheet
```

## Example Output

### Monthly Result

```text
dec_result.xlsx
jan_result.xlsx
feb_result.xlsx
mar_result.xlsx
```

### Final Semester Result

```text
PA_Sheet_Result.xlsx
```

The final PA Sheet combines the monthly attendance information and provides the overall semester attendance data.

## Future Scope

The project can be further extended with:

- Biometric or RFID attendance integration
- Real-time attendance tracking
- Web-based attendance dashboard
- Cloud-based attendance data management
