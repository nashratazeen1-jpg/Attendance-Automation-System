# Attendance Automation System

A Python-based **Attendance Automation System** that processes attendance data from Excel files, automatically calculates **student-wise slot-based attendance**, generates monthly result files, and prepares a final semester Progressive Attendance (PA) Sheet.

The main goal is to **reduce faculty workload, save time, and minimize manual calculation errors** caused by repeatedly applying Excel formulas.

## 🚀 Project Overview:

In the traditional process, attendance data is maintained in Excel and calculations such as attendance percentage, monthly results, and semester-wise PA sheets are prepared manually.

This project automates that workflow using Python.

Instead of manually applying formulas and preparing reports for each month, the system:

- Takes attendance data from an Excel file.
- Processes the data automatically.
- Calculates **Morning, Afternoon, and Evening slot-wise attendance** for each student.
- Calculates individual slot percentages and the **overall attendance percentage**.
- Generates monthly attendance result files.
- Combines monthly results into a final semester PA Sheet.

### `monthly.py`

Handles the **monthly slot-based attendance calculation**.

It processes the attendance Excel data for a selected month and calculates:

- **Morning (MRST)** attended slots and percentage.
- **Afternoon (AFST)** attended slots and percentage.
- **Evening (EVST)** attended slots and percentage.
- **Total attended slots** across all three slots.
- **Overall attendance percentage** based on the total available slots.

The total available slots depend on the selected month type:

```text
28 Days → 24 College Days → 72 Total Slots
30 Days → 26 College Days → 78 Total Slots
31 Days → 27 College Days → 81 Total Slots
15 Days → 15 College Days → 45 Total Slots
```

After processing, the system generates the corresponding monthly result Excel file.

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

## ⚒️ How the System Works:

### Step 1 — Provide Input

The faculty provides the attendance data in the required **Excel format**.

### Step 2 — Select Month

The system processes the attendance according to the selected month.

The supported month types are:

- 28 Days
- 30 Days
- 31 Days
- 15 Days

Each month type has a predefined number of college days and total available attendance slots.

### Step 3 — Automatic Processing

`monthly.py` reads and processes the Excel attendance data automatically.

It calculates **Morning, Afternoon, and Evening attended slots**, their individual percentages, the total attended slots, and the overall student attendance percentage.

### Step 4 — Monthly Result Generation

The system can generate separate result files for different months, such as:

- December
- January
- February
- March

Each result file contains the student's **slot-wise attendance, total attended slots, and overall attendance percentage**.

### Step 5 — Generate Final PA Sheet

`pa_sheet.py` combines the generated monthly attendance result files and calculates the overall semester attendance.

The final output is:

```text
PA_Sheet_Result.xlsx
```

This produces a structured PA Sheet ready for academic use and submission.

## ✨ Key Features:

- **Attendance Automation** — Automates calculations that would otherwise require manual Excel formulas.
- **Slot-wise Attendance Calculation** — Calculates Morning, Afternoon, and Evening attended slots and their percentages.
- **Monthly Result Generation** — Creates separate attendance result files for each month.
- **Student-wise Attendance Calculation** — Calculates total attended slots and overall attendance percentage for individual students.
- **Semester PA Sheet Generation** — Combines monthly results into one final semester report.
- **Reduced Faculty Workload** — Removes repetitive manual calculation and report preparation work.
- **Reduced Errors** — Minimizes errors caused by manual calculations and formula handling.
- **Simple Workflow** — Designed to make the process easier for non-technical users.

## 🤔 Why Slot-wise Attendance?

Slot-wise attendance helps faculties identify which time slot a student is most frequently absent in. This makes it easier to spot attendance patterns and, when required, discuss them with students or parents for better understanding and follow-up.

## 🤖 Technology Used:

- **Python** — Core processing and automation
- **Pandas** — Reading, processing, and analyzing Excel attendance data
- **Tkinter** — Graphical User Interface (GUI)
- **Microsoft Excel** — Input and output format for attendance records

## 🪜 System Flow:

```text
Attendance Excel File
        │
        ▼
   Select Month Type
        │
        ▼
   monthly.py
        │
        ▼
Process Slot-wise Attendance
(Morning / Afternoon / Evening)
        │
        ▼
Calculate Slot-wise %
        │
        ▼
Calculate Total Slots & Overall %
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

## 🔷 Example Output:

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


