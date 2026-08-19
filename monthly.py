# Code for Taking Monthly Slot Based Result
import time
import pandas as pd
import os

import tkinter as tk
from tkinter import filedialog


# ---------------- CUSTOM POPUP WINDOW ----------------
def custom_popup(title, msg, file_name=None):
    pop = tk.Tk()
    pop.title(title)
    pop.geometry("520x260")
    pop.resizable(False, False)

    pop.configure(bg="white")

    tk.Label(pop, text=title, font=("Segoe UI", 18, "bold"), bg="white", fg="#2c3e50").pack(pady=15)
    tk.Label(pop, text=msg, font=("Segoe UI", 13), bg="white", fg="#34495e").pack(pady=10)

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


# ---------------- GUI FOR OPTION SELECTION ----------------
def choose_month():
    root = tk.Tk()
    root.title("Attendance Automation System")
    root.geometry("620x420")
    root.resizable(False, False)
    root.configure(bg="white")

    choice = {"op": None}

    tk.Label(root, text="Attendance Automation System",
             font=("Segoe UI", 18, "bold"), bg="white", fg="#2c3e50").pack(pady=15)

    tk.Label(root, text="Select Month-Type Attendance:",
             font=("Segoe UI", 13), bg="white", fg="#34495e").pack(pady=5)

    def set_choice(val):
        choice["op"] = val
        root.destroy()

    btn_style = {
        "font": ("Segoe UI", 12, "bold"),
        "width": 55,
        "height": 2,
        "bg": "#f1f1f1",
        "fg": "#2c3e50",
        "relief": "flat",
        "activebackground": "#dcdcdc"
    }

    tk.Button(root, text="a) 28 Days Month  (Total 24 College Days)",
              command=lambda: set_choice("a"), **btn_style).pack(pady=8)

    tk.Button(root, text="b) 30 Days Month  (Total 26 College Days)",
              command=lambda: set_choice("b"), **btn_style).pack(pady=8)

    tk.Button(root, text="c) 31 Days Month  (Total 27 College Days)",
              command=lambda: set_choice("c"), **btn_style).pack(pady=8)

    tk.Button(root, text="d) 15 Days Month  (Total 15 College Days)",
              command=lambda: set_choice("d"), **btn_style).pack(pady=8)

    tk.Label(root, text="(Window will close after selection)",
             font=("Segoe UI", 10), bg="white", fg="gray").pack(pady=15)

    root.mainloop()
    return choice["op"]


# ---------------- START PROGRAM ----------------
print("\n")
print("\t\t\t\t|||||||> ============================================================== <|||||||")
print("\t\t\t\t\t\t Attendance Automation System - Monthly & Yearly")
print("\t\t\t\t\t\t\t Developed By: Shaikh Nashra Tazeen")
print("\t\t\t\t|||||||> ============================================================== <|||||||\n\n")

time.sleep(1)

# GUI open
op = choose_month()

# If user closes window without selecting
if op is None:
    print("No option selected. Program closed.")
    exit()


# ---------------- MAIN LOGIC ----------------
match op:

    case 'a':
        print("You Entered Option (a)..,There are Total 24 College Days!...\n")
        time.sleep(1)

        xl = input("Enter the Excel File of Student's Attendance: ")
        print("\n\n")

        try:
            df = pd.read_excel(xl, header=2)
            df.columns = df.columns.str.strip()
            print(df)
        except FileNotFoundError:
            custom_popup("Error", "Excel file not found!\nCheck file name.")
            exit()

        df["MRST %"] = ((df["MRST (24)"] / 24) * 100).round(2)
        df["AFST %"] = ((df["AFST (24)"] / 24) * 100).round(2)
        df["EVST %"] = ((df["EVST (24)"] / 24) * 100).round(2)

        df["Total (72)"] = df["MRST (24)"] + df["EVST (24)"] + df["AFST (24)"]

        total_slots = 72
        df["Attendance %"] = (df["Total (72)"] / total_slots) * 100
        df["Attendance %"] = df["Attendance %"].round(2)

        final_df = df[[
            "SR. No.", "ROLL NO.", "NAME OF STUDENT",
            "MRST (24)", "MRST %", "AFST (24)", "AFST %",
            "EVST (24)", "EVST %", "Total (72)", "Attendance %"
        ]]

        output_file = "feb_result.xlsx"
        final_df.to_excel(output_file, index=False)

        time.sleep(1)

        custom_popup("Success", f"Attendance Generated Successfully!\nFile: {output_file}")

    case 'b':
        print("You Entered Option (b)..,There are Total 26 College Days!...\n")
        time.sleep(1)

        xl = input("Enter the Excel File of Attendance: ")
        print("\n\n")

        try:
            df = pd.read_excel(xl, header=2)
            df.columns = df.columns.str.strip()
            print(df)
        except FileNotFoundError:
            custom_popup("Error", "Excel file not found!\nCheck file name.")
            exit()

        df["MRST %"] = ((df["MRST (26)"] / 26) * 100).round(2)
        df["AFST %"] = ((df["AFST (26)"] / 26) * 100).round(2)
        df["EVST %"] = ((df["EVST (26)"] / 26) * 100).round(2)

        df["Total (78)"] = df["MRST (26)"] + df["AFST (26)"] + df["EVST (26)"]

        total_slots = 78
        df["Attendance %"] = (df["Total (78)"] / total_slots) * 100
        df["Attendance %"] = df["Attendance %"].round(2)

        final_df = df[[
            "SR. No.", "ROLL NO.", "NAME OF STUDENT",
            "MRST (26)", "MRST %", "AFST (26)", "AFST %",
            "EVST (26)", "EVST %", "Total (78)", "Attendance %"
        ]]

        output_file = "30_result.xlsx"
        final_df.to_excel(output_file, index=False)

        time.sleep(1)

        custom_popup("Success", f"Attendance Generated Successfully!\nFile: {output_file}")

    case 'c':
        print("You Entered Option (c)..,There are Total 27 College Days!...\n")
        time.sleep(1)

        xl = input("Enter the Excel File of Attendance: ")
        print("\n\n")

        try:
            df = pd.read_excel(xl, header=2)
            df.columns = df.columns.str.strip()
            print(df)
        except FileNotFoundError:
            custom_popup("Error", "Excel file not found!\nCheck file name.")
            exit()

        df["MRST %"] = ((df["MRST (27)"] / 27) * 100).round(2)
        df["AFST %"] = ((df["AFST (27)"] / 27) * 100).round(2)
        df["EVST %"] = ((df["EVST (27)"] / 27) * 100).round(2)

        df["Total (81)"] = df["MRST (27)"] + df["AFST (27)"] + df["EVST (27)"]

        total_slots = 81
        df["Attendance %"] = (df["Total (81)"] / total_slots) * 100
        df["Attendance %"] = df["Attendance %"].round(2)

        final_df = df[[
            "SR. No.", "ROLL NO.", "NAME OF STUDENT",
            "MRST (27)", "MRST %", "AFST (27)", "AFST %",
            "EVST (27)", "EVST %", "Total (81)", "Attendance %"
        ]]

        output_file = "march_result.xlsx"
        final_df.to_excel(output_file, index=False)

        time.sleep(1)

        custom_popup("Success", f"Attendance Generated Successfully!\nFile: {output_file}")

    case 'd':
        print("You Entered Option (d)..,There are Total 15 College Days!...\n")
        time.sleep(1)

        xl = input("Enter the Excel File of Attendance: ")
        print("\n\n")

        try:
            df = pd.read_excel(xl, header=2)
            df.columns = df.columns.str.strip()
            print(df)
        except FileNotFoundError:
            custom_popup("Error", "Excel file not found!\nCheck file name.")
            exit()

        df["MRST %"] = ((df["MRST (15)"] / 15) * 100).round(2)
        df["AFST %"] = ((df["AFST (15)"] / 15) * 100).round(2)
        df["EVST %"] = ((df["EVST (15)"] / 15) * 100).round(2)

        df["Total (45)"] = df["MRST (15)"] + df["AFST (15)"] + df["EVST (15)"]

        total_slots = 45
        df["Attendance %"] = (df["Total (45)"] / total_slots) * 100
        df["Attendance %"] = df["Attendance %"].round(2)

        final_df = df[[
            "SR. No.", "ROLL NO.", "NAME OF STUDENT",
            "MRST (15)", "MRST %", "AFST (15)", "AFST %",
            "EVST (15)", "EVST %", "Total (45)", "Attendance %"
        ]]

        output_file = "dec_result.xlsx"
        final_df.to_excel(output_file, index=False)

        time.sleep(1)

        custom_popup("Success", f"Attendance Generated Successfully!\nFile: {output_file}")

    case _:
        custom_popup("Invalid Option", "You Entered Invalid Option!\nPlease Try Again!")
        print("You Entered Invalid Option!...Please Try Again!.")
