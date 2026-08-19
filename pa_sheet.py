import pandas as pd
import tkinter as tk
import os


# ---------------- CUSTOM POPUP WINDOW ----------------
def custom_popup(title, msg, file_name=None):
    pop = tk.Tk()
    pop.title(title)
    pop.geometry("520x260")
    pop.resizable(False, False)

    pop.configure(bg="white")

    tk.Label(pop, text=title, font=("Segoe UI", 18, "bold"), bg="white", fg="#2c3e50").pack(pady=15)
    tk.Label(pop, text=msg, font=("Segoe UI", 13), bg="white", fg="#34495e").pack(pady=10)

    if file_name:
        tk.Label(pop, text=f"📄 File: {file_name}", font=("Segoe UI", 12, "bold"),
                 bg="white", fg="#27ae60").pack(pady=5)

    # Open Folder button
    def open_folder():
        folder_path = os.getcwd()
        os.startfile(folder_path)

    btn_frame = tk.Frame(pop, bg="white")
    btn_frame.pack(pady=20)

    tk.Button(btn_frame, text="OK", font=("Segoe UI", 12, "bold"),
              width=12, bg="#27ae60", fg="white", relief="flat",
              command=pop.destroy).grid(row=0, column=0, padx=10)

    tk.Button(btn_frame, text="Open Folder", font=("Segoe UI", 12, "bold"),
              width=12, bg="#2980b9", fg="white", relief="flat",
              command=open_folder).grid(row=0, column=1, padx=10)

    pop.mainloop()


# ---------- READ PA SHEET (only basic student info) ----------
try:
    pa = pd.read_excel("input.xlsx", sheet_name="PA Sheet", header=4)
except FileNotFoundError:
    custom_popup("Error", "input.xlsx file not found!\nPlease keep it in same folder.")
    exit()

# Keep only required columns (Old data) pa is user defined
try:
    pa = pa[["SR. No.", "ROLL NO.", "NAME OF STUDENT"]]
except KeyError:
    custom_popup("Error", "Columns not found in input.xlsx!\nCheck SR. No., ROLL NO., NAME OF STUDENT")
    exit()


# ---------- READ MONTHLY RESULT FILES ----------
try:
    dec = pd.read_excel("dec_result.xlsx")
    jan = pd.read_excel("jan_result.xlsx")
    feb = pd.read_excel("feb_result.xlsx")
    mar = pd.read_excel("march_result.xlsx")
except FileNotFoundError:
    custom_popup(
        "Error",
        "One or more monthly result files are missing!\n\nRequired:\n"
        "dec_result.xlsx\njan_result.xlsx\nfeb_result.xlsx\nmarch_result.xlsx"
    )
    exit()


# Clean column names
for df in [dec, jan, feb, mar]:
    df.columns = df.columns.str.strip()


# ---------- ADD MONTHLY DATA ----------

# DEC
pa["Dec MRST (15)"] = dec["MRST (15)"]
pa["Dec MRST %"] = dec["MRST %"]

pa["Dec AFST (15)"] = dec["AFST (15)"]
pa["Dec AFST %"] = dec["AFST %"]

pa["Dec EVST (15)"] = dec["EVST (15)"]
pa["Dec EVST %"] = dec["EVST %"]

pa["Dec Total (45)"] = dec["Total (45)"]
pa["Dec Attendance %"] = dec["Attendance %"]


# JAN
pa["Jan MRST (27)"] = jan["MRST (27)"]
pa["Jan MRST %"] = jan["MRST %"]

pa["Jan AFST (27)"] = jan["AFST (27)"]
pa["Jan AFST %"] = jan["AFST %"]

pa["Jan EVST (27)"] = jan["EVST (27)"]
pa["Jan EVST %"] = jan["EVST %"]

pa["Jan Total (81)"] = jan["Total (81)"]
pa["Jan Attendance %"] = jan["Attendance %"]


# FEB
pa["Feb MRST (24)"] = feb["MRST (24)"]
pa["Feb MRST %"] = feb["MRST %"]

pa["Feb AFST (24)"] = feb["AFST (24)"]
pa["Feb AFST %"] = feb["AFST %"]

pa["Feb EVST (24)"] = feb["EVST (24)"]
pa["Feb EVST %"] = feb["EVST %"]

pa["Feb Total (72)"] = feb["Total (72)"]
pa["Feb Attendance %"] = feb["Attendance %"]


# MAR
pa["Mar MRST (27)"] = mar["MRST (27)"]
pa["Mar MRST %"] = mar["MRST %"]

pa["Mar AFST (27)"] = mar["AFST (27)"]
pa["Mar AFST %"] = mar["AFST %"]

pa["Mar EVST (27)"] = mar["EVST (27)"]
pa["Mar EVST %"] = mar["EVST %"]

pa["Mar Total (81)"] = mar["Total (81)"]
pa["Mar Attendance %"] = mar["Attendance %"]


# ---------- YEAR TOTAL & ATTENDANCE ----------
pa["Sem Total (279)"] = (
    pa["Dec Total (45)"] +
    pa["Jan Total (81)"] +
    pa["Feb Total (72)"] +
    pa["Mar Total (81)"]
)

pa["Sem Attendance %"] = (pa["Sem Total (279)"] / 279 * 100).round(2)


# ---------- SAVE TO NEW FILE ----------
output_file = "PA_Sheet_Result.xlsx"
pa.to_excel(output_file, index=False)

custom_popup(
    "✅ Success",
    "PA Result Sheet Generated Successfully!",
    file_name=output_file
)
