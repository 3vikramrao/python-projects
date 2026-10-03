import html
import os
import random
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import win32com.client as win32

# ============================================================
# GLOBALS
# ============================================================

df = None
current_file_path = None
last_loaded_timestamp = None
auto_running = False
scheduled_job = None
outlook_app = None

# ============================================================
# OUTLOOK CLIENT
# ============================================================

def get_outlook():
    global outlook_app
    if outlook_app is None:
        outlook_app = win32.Dispatch("Outlook.Application")
    return outlook_app

# ============================================================
# FILE TIMESTAMP
# ============================================================

def get_file_timestamp(path):
    try:
        return os.path.getmtime(path)
    except OSError:
        return None

# ============================================================
# RESET APP
# ============================================================

def reset_app_state():
    global df, current_file_path, last_loaded_timestamp

    df = None
    current_file_path = None
    last_loaded_timestamp = None

    excel_path_var.set("")
    status_var.set("Memory cleared")

    for dropdown in [
        fname_col_dropdown,
        lname_col_dropdown,
        email_col_dropdown,
        company_col_dropdown,
        title_col_dropdown
    ]:
        dropdown["values"] = []
        dropdown.set("")

    subject_box.delete(0, tk.END)
    body_box.delete("1.0", tk.END)

# ============================================================
# LOAD EXCEL
# ============================================================

def load_excel_columns(file_path):
    global df, current_file_path, last_loaded_timestamp

    try:
        timestamp = get_file_timestamp(file_path)

        if file_path != current_file_path or timestamp != last_loaded_timestamp:
            # fillna("") prevents Excel blank cells from rendering as literal "nan"
            df = pd.read_excel(file_path, engine="openpyxl").fillna("")

            current_file_path = file_path
            last_loaded_timestamp = timestamp

            cols = list(df.columns)

            fname_col_dropdown["values"] = cols
            lname_col_dropdown["values"] = cols
            email_col_dropdown["values"] = cols
            company_col_dropdown["values"] = cols
            title_col_dropdown["values"] = cols

            # Auto-detect common email column name
            for col in cols:
                if "email" in str(col).lower():
                    email_col_var.set(col)
                    break

            status_var.set(f"Loaded {len(df)} records successfully.")

    except Exception as e:
        messagebox.showerror("Error", str(e))

# ============================================================
# BROWSE FILE
# ============================================================

def browse_file():
    file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )

    if not file_path:
        return

    excel_path_var.set(file_path)
    load_excel_columns(file_path)

# ============================================================
# FORCE RELOAD
# ============================================================

def force_reload():
    global df, last_loaded_timestamp

    if not current_file_path:
        messagebox.showerror("Error", "No file loaded")
        return

    try:
        df = pd.read_excel(current_file_path, engine="openpyxl").fillna("")
        last_loaded_timestamp = get_file_timestamp(current_file_path)
        status_var.set("File reloaded successfully")
    except Exception as e:
        messagebox.showerror("Error", str(e))

# ============================================================
# TOKEN REPLACEMENT
# ============================================================

def replace_tokens(text, row):
    # Support mapped legacy keys
    mapping = {
        "{FirstName}": str(row.get(fname_col_var.get(), "")),
        "{LastName}": str(row.get(lname_col_var.get(), "")),
        "{Email}": str(row.get(email_col_var.get(), "")),
        "{Company}": str(row.get(company_col_var.get(), "")),
        "{Title}": str(row.get(title_col_var.get(), ""))
    }

    for key, value in mapping.items():
        text = text.replace(key, value)

    # Dynamic support for any column by {Column_Name}
    for col in row.index:
        text = text.replace(f"{{{col}}}", str(row[col]))

    return text

# ============================================================
# OUTLOOK DRAFT / SEND
# ============================================================

def create_draft(row, send_email=False):
    try:
        email_col = email_col_var.get()
        if not email_col or email_col not in row:
            raise ValueError("Please select a valid Email column.")

        outlook = get_outlook()

        subject_template = subject_box.get().strip()
        body_template = body_box.get("1.0", tk.END).strip()

        subject = replace_tokens(subject_template, row)
        body = replace_tokens(body_template, row)
        target_email = str(row[email_col]).strip()

        mail = outlook.CreateItem(0)

        # Load default signature
        mail.Display()
        existing_signature = mail.HTMLBody

        # Escape special characters before converting linebreaks to HTML
        escaped_body = html.escape(body)
        formatted_body = "<br>".join(escaped_body.splitlines())

        is_sim = simulation_mode.get()

        if is_sim:
            sim_addr = sim_email_var.get().strip()
            mail.To = sim_addr if sim_addr else target_email
            mail.Subject = f"[SIMULATION] {subject}"
        else:
            mail.To = target_email
            mail.Subject = subject

        mail.HTMLBody = formatted_body + existing_signature

        if send_email:
            if is_sim:
                mail.Save()
                status_var.set(f"SIMULATION: Saved draft for {target_email}")
            else:
                mail.Send()
                status_var.set(f"Sent email to {target_email}")
        else:
            mail.Display()

    except Exception as e:
        messagebox.showerror("Outlook Error", str(e))
        raise e

# ============================================================
# PREVIEW ALL
# ============================================================

def preview_messages():
    if df is None or df.empty:
        messagebox.showerror("Error", "Load an Excel file first")
        return

    email_col = email_col_var.get()
    if not email_col or email_col not in df.columns:
        messagebox.showerror("Error", "Select an Email column before previewing.")
        return

    win = tk.Toplevel(root)
    win.title("Preview All Emails")
    win.geometry("800x600")

    txt = tk.Text(win, wrap="word")
    txt.pack(fill="both", expand=True)

    subject_template = subject_box.get().strip()
    body_template = body_box.get("1.0", tk.END).strip()

    for idx, row in df.iterrows():
        subject = replace_tokens(subject_template, row)
        body = replace_tokens(body_template, row)

        txt.insert(
            tk.END,
            f"ROW {idx+1}\n"
            f"TO: {row[email_col]}\n"
            f"SUBJECT: {subject}\n\n"
            f"BODY:\n{body}\n"
            f"{'='*60}\n\n"
        )

# ============================================================
# REVIEW WINDOW
# ============================================================

def show_email_popup(start_idx):
    global auto_running, scheduled_job

    popup = tk.Toplevel(root)
    popup.title("Email Review & Automation")
    popup.geometry("850x650")

    current_idx = [start_idx]

    text_widget = tk.Text(popup, wrap="word")
    text_widget.pack(fill="both", expand=True)

    def refresh():
        row = df.iloc[current_idx[0]]
        subject = replace_tokens(subject_box.get().strip(), row)
        body = replace_tokens(body_box.get("1.0", tk.END).strip(), row)
        email_col = email_col_var.get()
        email = row.get(email_col, "NO EMAIL FOUND")

        text_widget.delete("1.0", tk.END)
        text_widget.insert(
            tk.END,
            f"Record {current_idx[0]+1} of {len(df)}\n"
            f"TO: {email}\n"
            f"SUBJECT: {subject}\n"
            f"{'-'*50}\n\n"
            f"{body}\n"
        )

    refresh()

    button_frame = tk.Frame(popup)
    button_frame.pack(pady=10)

    def previous_record():
        if current_idx[0] > 0:
            current_idx[0] -= 1
            refresh()

    def next_record():
        if current_idx[0] < len(df) - 1:
            current_idx[0] += 1
            refresh()

    def open_draft():
        row = df.iloc[current_idx[0]]
        create_draft(row, send_email=False)

    def auto_send_loop():
        global auto_running, scheduled_job

        if not auto_running:
            return

        try:
            row = df.iloc[current_idx[0]]
            create_draft(row, send_email=True)

            if current_idx[0] < len(df) - 1:
                current_idx[0] += 1
                refresh()

                # Calculate delay and convert to milliseconds
                delay = random.randint(60, 300)
                status_var.set(f"Waiting {delay}s before next email...")
                scheduled_job = popup.after(delay * 1000, auto_send_loop)
            else:
                auto_running = False
                messagebox.showinfo("Completed", "All emails processed successfully.")

        except Exception as e:
            auto_running = False
            messagebox.showerror("Automation Error", str(e))

    def start_auto():
        global auto_running
        if auto_running:
            return
        auto_running = True
        auto_send_loop()

    def pause_auto():
        global auto_running, scheduled_job
        auto_running = False
        if scheduled_job:
            popup.after_cancel(scheduled_job)
            scheduled_job = None
        status_var.set("Automation paused")

    def on_window_close():
        pause_auto()
        popup.destroy()

    popup.protocol("WM_DELETE_WINDOW", on_window_close)

    tk.Button(button_frame, text="Previous", command=previous_record).grid(row=0, column=0, padx=5)
    tk.Button(button_frame, text="Next", command=next_record).grid(row=0, column=1, padx=5)
    tk.Button(button_frame, text="Create Outlook Draft", bg="green", fg="white", command=open_draft).grid(row=0, column=2, padx=5)
    tk.Button(button_frame, text="Auto Send", bg="blue", fg="white", command=start_auto).grid(row=0, column=3, padx=5)
    tk.Button(button_frame, text="Pause", bg="orange", command=pause_auto).grid(row=0, column=4, padx=5)

# ============================================================
# START PROCESS
# ============================================================

def start_process():
    if df is None or df.empty:
        messagebox.showerror("Error", "Load Excel first")
        return

    if not email_col_var.get() or email_col_var.get() not in df.columns:
        messagebox.showerror("Error", "Please map a valid Email column.")
        return

    try:
        start_row = int(start_row_var.get().strip()) - 1
        if start_row < 0 or start_row >= len(df):
            messagebox.showerror("Error", f"Start row must be between 1 and {len(df)}")
            return
        show_email_popup(start_row)
    except ValueError:
        messagebox.showerror("Error", "Start row must be an integer.")

# ============================================================
# GUI LAYOUT
# ============================================================

root = tk.Tk()
simulation_mode = tk.BooleanVar(value=True)

root.title("Outlook Mail Merge Assistant")
root.geometry("900x850")

excel_path_var = tk.StringVar()
status_var = tk.StringVar(value="Ready")
start_row_var = tk.StringVar(value="1")
sim_email_var = tk.StringVar(value="test@example.com")

fname_col_var = tk.StringVar()
lname_col_var = tk.StringVar()
email_col_var = tk.StringVar()
company_col_var = tk.StringVar()
title_col_var = tk.StringVar()

# --- File Selection ---
file_frame = tk.Frame(root)
file_frame.pack(fill="x", padx=10, pady=5)

tk.Label(file_frame, text="Excel File:").pack(anchor="w")
tk.Entry(file_frame, textvariable=excel_path_var).pack(fill="x", pady=2)

top_buttons = tk.Frame(file_frame)
top_buttons.pack(pady=5)
tk.Button(top_buttons, text="Browse", command=browse_file).grid(row=0, column=0, padx=5)
tk.Button(top_buttons, text="Reload", command=force_reload).grid(row=0, column=1, padx=5)
tk.Button(top_buttons, text="Reset", bg="orange", command=reset_app_state).grid(row=0, column=2, padx=5)

# --- Column Mappings ---
mapping_frame = tk.LabelFrame(root, text="Column Mappings (Dynamic tags like {Header_Name} also work)")
mapping_frame.pack(fill="x", padx=10, pady=5)

tk.Label(mapping_frame, text="First Name:").grid(row=0, column=0, padx=5, pady=3, sticky="w")
fname_col_dropdown = ttk.Combobox(mapping_frame, textvariable=fname_col_var, width=25)
fname_col_dropdown.grid(row=0, column=1, padx=5, pady=3)

tk.Label(mapping_frame, text="Last Name:").grid(row=0, column=2, padx=5, pady=3, sticky="w")
lname_col_dropdown = ttk.Combobox(mapping_frame, textvariable=lname_col_var, width=25)
lname_col_dropdown.grid(row=0, column=3, padx=5, pady=3)

tk.Label(mapping_frame, text="Email Column:").grid(row=1, column=0, padx=5, pady=3, sticky="w")
email_col_dropdown = ttk.Combobox(mapping_frame, textvariable=email_col_var, width=25)
email_col_dropdown.grid(row=1, column=1, padx=5, pady=3)

tk.Label(mapping_frame, text="Company:").grid(row=1, column=2, padx=5, pady=3, sticky="w")
company_col_dropdown = ttk.Combobox(mapping_frame, textvariable=company_col_var, width=25)
company_col_dropdown.grid(row=1, column=3, padx=5, pady=3)

tk.Label(mapping_frame, text="Title:").grid(row=2, column=0, padx=5, pady=3, sticky="w")
title_col_dropdown = ttk.Combobox(mapping_frame, textvariable=title_col_var, width=25)
title_col_dropdown.grid(row=2, column=1, padx=5, pady=3)

# --- Row Controls ---
row_frame = tk.Frame(root)
row_frame.pack(fill="x", padx=10, pady=5)
tk.Label(row_frame, text="Start Row (1-based index):").pack(side="left")
tk.Entry(row_frame, textvariable=start_row_var, width=8).pack(side="left", padx=5)

# --- Email Content ---
tk.Label(root, text="Subject Template:").pack(anchor="w", padx=10)
subject_box = tk.Entry(root)
subject_box.pack(fill="x", padx=10, pady=2)

tk.Label(root, text="Body Template (Tokens: {FirstName}, {LastName}, {Email}, {Company}, {Title} or Any {Column}):").pack(anchor="w", padx=10, pady=(5, 0))
body_box = tk.Text(root, height=12)
body_box.pack(fill="both", expand=True, padx=10, pady=2)

# --- Options & Execution ---
options_frame = tk.LabelFrame(root, text="Sending Options")
options_frame.pack(fill="x", padx=10, pady=5)

tk.Checkbutton(
    options_frame,
    text="Simulation Mode (Saves drafts; redirects recipient to test address)",
    variable=simulation_mode,
    fg="blue"
).pack(anchor="w", padx=5)

sim_box_frame = tk.Frame(options_frame)
sim_box_frame.pack(fill="x", padx=5, pady=2)
tk.Label(sim_box_frame, text="Simulation Test Address:").pack(side="left")
tk.Entry(sim_box_frame, textvariable=sim_email_var, width=30).pack(side="left", padx=5)

button_bar = tk.Frame(root)
button_bar.pack(pady=5)
tk.Button(button_bar, text="Preview All", command=preview_messages).grid(row=0, column=0, padx=10)
tk.Button(button_bar, text="Start Review / Auto Send", bg="green", fg="white", command=start_process).grid(row=0, column=1, padx=10)

# --- Status Bar ---
tk.Label(root, textvariable=status_var, bd=1, relief="sunken", anchor="w").pack(side="bottom", fill="x")

root.mainloop()