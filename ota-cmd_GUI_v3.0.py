import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import subprocess
import pandas as pd
import os
from PIL import Image, ImageTk
import datetime
import matplotlib.pyplot as plt
from io import BytesIO
import openpyxl
import time
import logging
import threading
import matplotlib
import shutil
import sys
import requests

# Use the Agg backend for Matplotlib
matplotlib.use('Agg')

# Determine the script directory
script_dir = os.path.dirname(os.path.abspath(__file__))

# Constants
JAR_PATH = os.path.join(script_dir, "ota-cmdutil-0.0.1-SNAPSHOT.jar")
LOGO_PATH = os.path.join(script_dir, "logo.png")
SAMPLE_COMMANDS_PATH = os.path.join(script_dir, "Sample_Commands.xlsx")
IMG_SIZE = (120, 110)
OUTPUT_FORMAT = "%Y%m%d_%H%M%S"

# Global variable to control the execution
abort_flag = False

# Logging setup
logging.basicConfig(filename='ota_test_automation.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def browse_excel_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        entry_excel_file_path.delete(0, tk.END)
        entry_excel_file_path.insert(0, file_path)

def validate_inputs(excel_file_path, imei, tester_name):
    if not excel_file_path:
        messagebox.showerror("Error", "No file selected.")
        return False
    if not imei.isdigit() or len(imei) != 15:
        messagebox.showerror("Error", "IMEI must be a 15-digit integer value.")
        return False
    if not tester_name:
        messagebox.showerror("Error", "Please enter the Tester Name.")
        return False
    return True

def check_internet():
    url = "http://www.google.com"
    timeout = 5
    for _ in range(3):
        try:
            requests.get(url, timeout=timeout)
            return True
        except requests.ConnectionError:
            time.sleep(1)
    return False

def execute_commands():
    global abort_flag
    label_execution_stats.config(text="Executed: 0  Passed: 0  Failed: 0 Error: 0")
    button_browse.config(state=tk.DISABLED)
    button_reset.config(state=tk.DISABLED)
    excel_file_path = entry_excel_file_path.get()
    imei = entry_imei.get()
    tester_name = entry_tester.get()
    versionInfo = verinf()

    if not validate_inputs(excel_file_path, imei, tester_name):
        button_execute.config(state=tk.NORMAL)
        button_browse.config(state=tk.NORMAL)
        button_reset.config(state=tk.NORMAL)
        return

    if not check_internet():
        messagebox.showerror("Internet Not Available", "Internet not available. Click OK to exit the application.")
        root.destroy()
        return

    output_folder = os.path.dirname(excel_file_path)
    start_time = datetime.datetime.now()
    timestamp = start_time.strftime(OUTPUT_FORMAT)
    output_file_excel = os.path.join(output_folder, f"{imei}_OTA_Cmd_TestResult_{timestamp}.xlsx")
    output_file_txt = os.path.join(output_folder, f"{imei}_OTA_Cmd_Test_Logs_{timestamp}.txt")

    try:
        df = pd.read_excel(excel_file_path)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to read Excel file: {e}")
        logging.error(f"Failed to read Excel file: {e}")
        button_execute.config(state=tk.NORMAL)
        return

    commands_executed = commands_passed = commands_failed = Error_count = 0
    total_commands = len(df)

    progress_bar.pack(pady=10)
    progress_var.set(0)
    label_progress_percentage.pack(pady=10)
    label_execution_stats.pack(pady=5)
    label_output_file_path.pack(pady=5)

    rows = []
    try:
        with open(output_file_txt, "w") as f_out_txt, pd.ExcelWriter(output_file_excel) as writer:
            for i, command in enumerate(df['Commands']):
                if abort_flag:
                    break
                creationflags = subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0
                result = subprocess.run(["java", "-jar", JAR_PATH, imei, command], capture_output=True, text=True, shell=False, creationflags=creationflags)
                output = result.stdout
                cmd = f'java -jar {JAR_PATH} {imei} {command}'
                commands_executed += 1
                relevant_output = parse_output(output)
                expected_response = df.at[i, 'Expected Response']
                test_case_result = compare_responses(expected_response, relevant_output)
                if test_case_result == 'Pass':
                    commands_passed += 1
                elif test_case_result == 'Check_IMEI':
                    Error_count += 1
                    if Error_count == 5 :
                        abort_execution_for_Invalid_IMEI()
                elif test_case_result == 'Device Offline':
                    Error_count += 1
                elif test_case_result == 'MQTT Error':
                    Error_count += 1
                else:
                    commands_failed += 1

                rows.append([i + 1, datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), command, expected_response, relevant_output, test_case_result])
                log_command_execution(f_out_txt, cmd, output)

                update_progress_bar(commands_executed, total_commands)
                time.sleep(4)	#delay between two commands

            save_to_excel(writer, rows)
            insert_summary_chart(writer, imei, tester_name, start_time, commands_passed, commands_failed, Error_count, versionInfo)
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")
        logging.error(f"An error occurred: {e}")
    finally:
        display_execution_summary(output_folder, commands_executed, commands_passed, commands_failed, Error_count)
        button_execute.config(state=tk.NORMAL)
        button_abort.config(state=tk.DISABLED)
        button_reset.config(state=tk.NORMAL)
        button_browse.config(state=tk.NORMAL)

def parse_output(output):
    if "Response:" in output:
        return output.split("Response:", 1)[1].strip()
    elif "operation failed" in output:
        return "Device is offline"
    elif "Read timed out" in output:
        return "MQTT response issue"
    elif "404001" in output:
        return "Response not received, Check IMEI"
    else:
        return "No response found"

def compare_responses(expected_response, relevant_output):
    if expected_response == relevant_output:
        return 'Pass'
    elif relevant_output == "Device is offline":
        return 'Device Offline'
    elif relevant_output == "MQTT response issue":
        return 'MQTT Error'
    elif relevant_output == "Response not received, Check IMEI":
        return 'Check_IMEI'
    else:
        return 'Fail'
    
def verinf():
    imei = entry_imei.get()
    versionInfo = "RSW0x.x.x"
    creationflags = subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0           
    result = subprocess.run(["java", "-jar", JAR_PATH, imei, "GET VERINF"], capture_output=True, text=True, shell=False, creationflags=creationflags)
    output = result.stdout
    lines = output.splitlines()
    for line in lines:
        if line.startswith("Response:"):
            parts = line.split(',')
            if len(parts) > 2:
                versionInfo = parts[2]
    return versionInfo

def log_command_execution(f_out_txt, cmd, output):
    print(cmd, file=f_out_txt)
    for line in output.splitlines():
        print(line, file=f_out_txt)
    print("\n", file=f_out_txt)

def update_progress_bar(commands_executed, total_commands):
    progress_var.set((commands_executed / total_commands) * 100)
    label_progress_percentage.config(text=f"{int((commands_executed / total_commands) * 100)}%")
    root.update_idletasks()

def save_to_excel(writer, rows):
    df_output = pd.DataFrame(rows, columns=['Serial Number', 'Timestamp', 'Commands', 'Expected Response', 'Actual Response', 'Test Case Result'])
    df_output.to_excel(writer, index=False)

def insert_summary_chart(writer, imei, tester_name, start_time, commands_passed, commands_failed, Error_count, versionInfo): 
    end_time = datetime.datetime.now()
    total_duration = end_time - start_time
    formatted_duration = f"{total_duration.seconds//3600}Hrs {(total_duration.seconds%3600)//60}Mins {total_duration.seconds%60}Secs"
    counts = [commands_passed, commands_failed, Error_count]
    labels = ['Passed', 'Failed', 'Error_count']
    plt.figure(figsize=(6, 6))
    plt.pie(counts, labels=labels, autopct='%1.0f%%')
    plt.text(0, 1.5, 'Test Summary', fontsize=12, ha='center')
    plt.text(0, 1.4, f'Device IMEI: {imei}', fontsize=10, ha='center')
    plt.text(0, 1.3, f'MCU Software Version: {versionInfo}', fontsize=10, ha='center')
    plt.text(0, 1.2, f'Commands Execution Duration: {formatted_duration}', fontsize=10, ha='center')
    plt.text(0, -1.4, f'Tested by: {tester_name}', fontsize=12, ha='center')

    # Save the pie chart to bytes
    img_bytes = BytesIO()
    plt.savefig(img_bytes, format='png')
    img_bytes.seek(0)

    # Insert the pie chart into the Excel file
    image = openpyxl.drawing.image.Image(img_bytes)
    worksheet = writer.sheets['Sheet1']  # Change the sheet name if necessary
    worksheet.add_image(image, 'H2')

def display_execution_summary(output_folder, commands_executed, commands_passed, commands_failed, Error_count):
    label_output_file_path.config(text=f"Output file saved at:\n{output_folder}")
    label_execution_stats.config(text=f"Executed: {commands_executed:3}  Passed: {commands_passed:3}  Failed: {commands_failed:3}  Error: {Error_count:3}")
    messagebox.showinfo("Execution Completed", f"Command execution completed.\n\nExecuted: {commands_executed}\nPassed: {commands_passed}\nFailed: {commands_failed}\nError: {Error_count}\n\nResults saved in: {output_folder}")
    button_open_folder.config(state=tk.NORMAL)
    logging.info("All commands executed")

def open_output_folder():
    output_folder = os.path.dirname(entry_excel_file_path.get())
    os.startfile(output_folder)

def start_execution():
    global abort_flag
    abort_flag = False
    button_execute.config(state=tk.DISABLED)
    button_abort.config(state=tk.NORMAL)
    button_reset.config(state=tk.DISABLED)
    button_browse.config(state=tk.DISABLED)
    threading.Thread(target=execute_commands).start()

def abort_execution_for_Invalid_IMEI():
    global abort_flag
    # Show a error dialog
    messagebox.showerror("Check IMEI", "Response not received, Application aborted")
    abort_flag = True
        
def abort_execution():
    global abort_flag
    # Show a confirmation dialog
    response = messagebox.askyesno("Abort", "Are you sure you want to abort?")
    if response:  # If user clicked 'Yes'
        # Handle the abort action here
        abort_flag = True
    else:
        abort_flag = False

def reset_fields():
    entry_excel_file_path.delete(0, tk.END)
    entry_imei.delete(0, tk.END)
    entry_tester.delete(0, tk.END)
    label_execution_stats.config(text="Executed: 0  Passed: 0  Failed: 0 Error: 0")
    update_progress_bar(0, 100)

def exit_application():
    if messagebox.askokcancel("Quit", "Do you want to quit?"):
        root.destroy()

def open_help():
    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="Sample_Commands.xlsx", title="Save Sample Commands File")
    if save_path:
        shutil.copy(SAMPLE_COMMANDS_PATH, save_path)
        messagebox.showinfo("Download Complete", f"The sample commands Excel file has been downloaded and saved to:\n{save_path}")

def contact_us():
    # Functionality for the Contact Us tab
    # Create a new window for displaying contact information
    contact_window = tk.Toplevel()
    contact_window.title("Contact Us")
    
    # Contact information labels
    label_contact_info = tk.Label(contact_window, text="Contact Information:", font=("Calibri", 14, "bold"))
    label_contact_info.pack(pady=10)

    label_email = tk.Label(contact_window, text="Email: jayrahul@gmail.com", font=("Calibri", 12))
    label_email.pack()

    label_phone = tk.Label(contact_window, text="Phone: +91-9890012345", font=("Calibri", 12))
    label_phone.pack()

    label_address = tk.Label(contact_window, text="Address: Aundh,Pune(MS)– 411007, India", font=("Calibri", 12))
    label_address.pack()

#     # Center the contact window
#     contact_window.geometry("+{}+{}".format(
#         root.winfo_rootx() + root.winfo_reqwidth() // 2 - contact_window.winfo_reqwidth() // 2,
#         root.winfo_rooty() + root.winfo_reqheight() // 2 - contact_window.winfo_reqheight() // 2))

# Create the main window
root = tk.Tk()
root.title("OTA Commands Test Automation Tool v3.0")
root.configure(bg='black')
root.config(highlightbackground="gray", highlightthickness=3)
root.geometry("800x600")
root.iconphoto(False, tk.PhotoImage(file=LOGO_PATH))
root.state('zoomed')  # Open in maximized mode

# Menu Bar
menu_bar = tk.Menu(root)

# File Menu
file_menu = tk.Menu(menu_bar, tearoff=0)
file_menu.add_command(label="Execute", command=start_execution)
file_menu.add_command(label="Abort", command=abort_execution)
file_menu.add_command(label="Reset", command=reset_fields)
file_menu.add_separator()
file_menu.add_command(label="Exit", command=exit_application)

# Help Menu
help_menu = tk.Menu(menu_bar, tearoff=0)
help_menu.add_command(label="Sample Commands Excel File", command=open_help)

# Contact Us Menu
contact_menu = tk.Menu(menu_bar, tearoff=0)
contact_menu.add_command(label="Contact Us", command=contact_us)

menu_bar.add_cascade(label="File", menu=file_menu)
menu_bar.add_cascade(label="Help", menu=help_menu)
menu_bar.add_cascade(label="Contact Us", menu=contact_menu)

root.config(menu=menu_bar)

# Load and display logo
logo_img = Image.open(LOGO_PATH).resize(IMG_SIZE, Image.LANCZOS)
logo_img = ImageTk.PhotoImage(logo_img)
label_logo = tk.Label(root, image=logo_img, bg='black')
label_logo.pack(pady=10)

# Excel file input frame
# Widgets
frame_excel_file = tk.Frame(root, bg='black')
frame_excel_file.pack(pady=10)
label_excel_file = tk.Label(frame_excel_file, text="Excel File Path:", bg='black', fg='white', font=("Calibri", 12))
label_excel_file.grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)

entry_excel_file_path = tk.Entry(frame_excel_file, width=40, font=("Calibri", 12))
entry_excel_file_path.grid(row=0, column=1, padx=5, pady=5)

button_browse = tk.Button(frame_excel_file, text="Browse", command=browse_excel_file, font=("Calibri", 12))
button_browse.grid(row=0, column=2, padx=5, pady=5)

# Define a common width for labels and entries
label_width = 12
entry_width = 25

# IMEI input frame
frame_imei = tk.Frame(root, bg='black')
frame_imei.pack(pady=10)
label_imei = tk.Label(frame_imei, text="IMEI:", bg='black', fg='white', font=("Calibri", 12), width=label_width, anchor='w')
label_imei.pack(side=tk.LEFT, padx=5)
entry_imei = tk.Entry(frame_imei, width=entry_width, font=("Calibri", 12))
entry_imei.pack(side=tk.LEFT, padx=5)

# Tester name input frame
frame_tester = tk.Frame(root, bg='black')
frame_tester.pack(pady=10)
label_tester = tk.Label(frame_tester, text="Tester Name:", bg='black', fg='white', font=("Calibri", 12), width=label_width, anchor='w')
label_tester.pack(side=tk.LEFT, padx=5)
entry_tester = tk.Entry(frame_tester, width=entry_width, font=("Calibri", 12))
entry_tester.pack(side=tk.LEFT, padx=5)

# Execute, Abort, Reset buttons
frame_buttons = tk.Frame(root, bg='black')
frame_buttons.pack(pady=10)
button_width = 16
button_execute = tk.Button(frame_buttons, text="Execute Commands", width=button_width, command=start_execution, font=("Calibri", 12))
button_execute.pack(side=tk.LEFT, padx=5)
button_abort = tk.Button(frame_buttons, text="Abort", width=button_width, command=abort_execution, state=tk.DISABLED, font=("Calibri", 12))
button_abort.pack(side=tk.LEFT, padx=5)
button_reset = tk.Button(frame_buttons, text="Reset", width=button_width, command=reset_fields, state=tk.NORMAL, font=("Calibri", 12))
button_reset.pack(side=tk.LEFT, padx=5)

# Progress bar and labels
progress_var = tk.DoubleVar()
progress_bar = ttk.Progressbar(root, orient="horizontal", maximum=100, length=400, mode="determinate", variable=progress_var)
label_progress_percentage = tk.Label(root, text="", bg='black', fg='white', font=("Calibri", 12))
label_execution_stats = tk.Label(root, text="Executed: 0  Passed: 0  Failed: 0 Error: 0", bg='black', fg='white', font=("Calibri", 12))
label_output_file_path = tk.Label(root, text="", wraplength=300, bg='black', fg='white', font=("Calibri", 12))
button_open_folder = tk.Button(root, text="Open Output Folder", command=open_output_folder, state=tk.DISABLED, font=("Calibri", 12))

# Pack elements
progress_bar.pack(pady=10)
label_progress_percentage.pack(pady=5)
label_execution_stats.pack(pady=5)
label_output_file_path.pack(pady=5)

# Set initial focus
entry_excel_file_path.focus()
button_open_folder.pack(pady=5)
# Team and copyright labels
label_Team_name = tk.Label(root, text="Designed & Developed by: Mr. Rahul J. Shimpi", bg='black', fg='white', font=("Calibri", 12))
label_Team_name.pack(pady=(235, 5))
label_copyright = tk.Label(root, text="© 2024 rjs", bg='black', fg='white', font=("Calibri", 10))
label_copyright.pack(pady=(0, 5))

root.mainloop()