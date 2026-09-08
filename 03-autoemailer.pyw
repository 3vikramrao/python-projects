import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path
import win32com.client as win32
import os
import gc

# ============================================================
# GLOBALS
# ============================================================

df = None
current_file_path = None
last_loaded_timestamp = None

# ============================================================
# FILE TIMESTAMP
# ============================================================

def get_file_timestamp(path):
    try:
        return os.path.getmtime(path)
    except:
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
        name_col_dropdown,
        email_col_dropdown,
        company_col_dropdown,
        title_col_dropdown
    ]:
        dropdown["values"] = []
        dropdown.set("")

    subject_box.delete(0, tk.END)
    body_box.delete("1.0", tk.END)

    gc.collect()

# ============================================================
# LOAD EXCEL
# ============================================================

def load_excel_columns(file_path):

    global df, current_file_path, last_loaded_timestamp

    try:

        timestamp = get_file_timestamp(file_path)

        if file_path != current_file_path or timestamp != last_loaded_timestamp:

            df = pd.read_excel(file_path, engine="openpyxl")

            current_file_path = file_path
            last_loaded_timestamp = timestamp

            cols = list(df.columns)

            name_col_dropdown["values"] = cols
            email_col_dropdown["values"] = cols
            company_col_dropdown["values"] = cols
            title_col_dropdown["values"] = cols

            status_var.set("Excel loaded successfully")

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

        df = pd.read_excel(
            current_file_path,
            engine="openpyxl"
        )

        last_loaded_timestamp = get_file_timestamp(
            current_file_path
        )

        status_var.set("File reloaded")

    except Exception as e:
        messagebox.showerror("Error", str(e))

# ============================================================
# TOKEN REPLACEMENT
# ============================================================

def replace_tokens(text, row):

    mapping = {
        "{Name}": str(row.get(name_col_var.get(), "")),
        "{Email}": str(row.get(email_col_var.get(), "")),
        "{Company}": str(row.get(company_col_var.get(), "")),
        "{Title}": str(row.get(title_col_var.get(), ""))
    }

    for key, value in mapping.items():
        text = text.replace(key, value)

    return text

# ============================================================
# OUTLOOK DRAFT
# ============================================================

def create_draft(row):

    try:

        outlook = win32.Dispatch("Outlook.Application")

        subject_template = subject_box.get()

        body_template = body_box.get(
            "1.0",
            tk.END
        )

        subject = replace_tokens(
            subject_template,
            row
        )

        body = replace_tokens(
            body_template,
            row
        )

        email = str(
            row[email_col_var.get()]
        )

        mail = outlook.CreateItem(0)

        # Force Outlook to load default signature
        mail.Display()
        existing_signature = mail.HTMLBody

        mail.To = email
        mail.Subject = subject

        mail.HTMLBody = f"""
        <html>
        <body style='font-family:Calibri;font-size:11pt'>
        {body.replace(chr(10), '<br>')}
        </body>
        </html>
        """

        mail.HTMLBody = (
            body.strip().replace("\n", "<br>")
            #+ "<br>"
            + existing_signature
        )

        mail.Display()

    except Exception as e:

        messagebox.showerror(
            "Outlook Error",
            str(e)
        )

# ============================================================
# PREVIEW ALL
# ============================================================

def preview_messages():

    if df is None:
        messagebox.showerror(
            "Error",
            "Load Excel first"
        )
        return

    win = tk.Toplevel(root)

    win.title("Preview All Emails")
    win.geometry("800x600")

    txt = tk.Text(
        win,
        wrap="word"
    )

    txt.pack(
        fill="both",
        expand=True
    )

    subject_template = subject_box.get()

    body_template = body_box.get(
        "1.0",
        tk.END
    )

    for idx, row in df.iterrows():

        subject = replace_tokens(
            subject_template,
            row
        )

        body = replace_tokens(
            body_template,
            row
        )

        txt.insert(
            tk.END,
            f"""
ROW {idx+1}

TO:
{row[email_col_var.get()]}

SUBJECT:
{subject}

BODY:
{body}

================================================

"""
        )

# ============================================================
# REVIEW WINDOW
# ============================================================

def show_email_popup(start_idx):

    popup = tk.Toplevel(root)
    popup.title("Email Review")
    popup.geometry("850x650")

    current_idx = [start_idx]

    text_widget = tk.Text(
        popup,
        wrap="word"
    )

    text_widget.pack(
        fill="both",
        expand=True
    )

    def refresh():

        row = df.iloc[current_idx[0]]

        subject = replace_tokens(
            subject_box.get(),
            row
        )

        body = replace_tokens(
            body_box.get("1.0", tk.END),
            row
        )

        email = row[email_col_var.get()]

        text_widget.delete(
            "1.0",
            tk.END
        )

        text_widget.insert(
            tk.END,
            f"""
Record {current_idx[0]+1} of {len(df)}

TO:
{email}

SUBJECT:
{subject}

------------------------------------------------

{body}
"""
        )

    refresh()

    button_frame = tk.Frame(popup)
    button_frame.pack(pady=10)

    def previous_record():

        if current_idx[0] > 0:
            current_idx[0] -= 1
            refresh()

    def next_record():

        if current_idx[0] < len(df)-1:
            current_idx[0] += 1
            refresh()

    def open_draft():

        row = df.iloc[current_idx[0]]

        create_draft(row)

    tk.Button(
        button_frame,
        text="Previous",
        command=previous_record
    ).grid(
        row=0,
        column=0,
        padx=10
    )

    tk.Button(
        button_frame,
        text="Next",
        command=next_record
    ).grid(
        row=0,
        column=1,
        padx=10
    )

    tk.Button(
        button_frame,
        text="Create Outlook Draft",
        bg="green",
        fg="white",
        command=open_draft
    ).grid(
        row=0,
        column=2,
        padx=10
    )

# ============================================================
# START PROCESS
# ============================================================

def start_process():

    if df is None:
        messagebox.showerror(
            "Error",
            "Load Excel first"
        )
        return

    try:

        start_row = int(
            start_row_var.get()
        ) - 1

        show_email_popup(start_row)

    except Exception as e:

        messagebox.showerror(
            "Error",
            str(e)
        )

# ============================================================
# GUI
# ============================================================

root = tk.Tk()

root.title(
    "Outlook Mail Merge Assistant"
)

root.geometry("900x750")
root.minsize(900, 750)

excel_path_var = tk.StringVar()
status_var = tk.StringVar()
start_row_var = tk.StringVar(value="1")

name_col_var = tk.StringVar()
email_col_var = tk.StringVar()
company_col_var = tk.StringVar()
title_col_var = tk.StringVar()

# ============================================================
# FILE
# ============================================================

tk.Label(
    root,
    text="Excel File"
).pack()

tk.Entry(
    root,
    textvariable=excel_path_var,
    width=80
).pack()

top_buttons = tk.Frame(root)
top_buttons.pack(pady=5)

tk.Button(
    top_buttons,
    text="Browse",
    command=browse_file
).grid(row=0, column=0, padx=5)

tk.Button(
    top_buttons,
    text="Reload",
    command=force_reload
).grid(row=0, column=1, padx=5)

tk.Button(
    top_buttons,
    text="Reset",
    bg="orange",
    command=reset_app_state
).grid(row=0, column=2, padx=5)

# ============================================================
# COLUMN MAPPINGS
# ============================================================

mapping_frame = tk.LabelFrame(root, text="Column Mapping")
mapping_frame.pack(fill="x", padx=10, pady=5)

# Row 1
tk.Label(mapping_frame, text="Name Column").grid(
    row=0, column=0, padx=5, pady=5, sticky="w"
)

name_col_dropdown = ttk.Combobox(
    mapping_frame,
    textvariable=name_col_var,
    width=35
)
name_col_dropdown.grid(
    row=0, column=1, padx=5, pady=5
)

tk.Label(mapping_frame, text="Email Column").grid(
    row=0, column=2, padx=5, pady=5, sticky="w"
)

email_col_dropdown = ttk.Combobox(
    mapping_frame,
    textvariable=email_col_var,
    width=35
)
email_col_dropdown.grid(
    row=0, column=3, padx=5, pady=5
)

# Row 2
tk.Label(mapping_frame, text="Company Column").grid(
    row=1, column=0, padx=5, pady=5, sticky="w"
)

company_col_dropdown = ttk.Combobox(
    mapping_frame,
    textvariable=company_col_var,
    width=35
)
company_col_dropdown.grid(
    row=1, column=1, padx=5, pady=5
)

tk.Label(mapping_frame, text="Title Column").grid(
    row=1, column=2, padx=5, pady=5, sticky="w"
)

title_col_dropdown = ttk.Combobox(
    mapping_frame,
    textvariable=title_col_var,
    width=35
)
title_col_dropdown.grid(
    row=1, column=3, padx=5, pady=5
)


# ============================================================
# START ROW
# ============================================================

tk.Label(
    root,
    text="Start Row"
).pack()

tk.Entry(
    root,
    textvariable=start_row_var,
    width=10
).pack()

# ============================================================
# SUBJECT
# ============================================================

tk.Label(
    root,
    text="Subject Template"
).pack()

subject_box = tk.Entry(
    root,
    width=100
)

subject_box.pack(
    padx=10,
    pady=5
)

# ============================================================
# BODY
# ============================================================

tk.Label(
    root,
    text="""
Available Tokens:
{Name}
{Email}
{Company}
{Title}
"""
).pack()

body_box = tk.Text(
    root,
    height=10,
    width=100
)

body_box.pack(
    padx=10,
    pady=5
)
# ============================================================
# LOAD OUTLOOK SIGNATURES
# ============================================================

def get_outlook_signatures():

    signatures = []

    try:

        sig_path = Path.home() / "AppData/Roaming/Microsoft/Signatures"

        if sig_path.exists():

            for file in sig_path.glob("*.htm"):
                signatures.append(file.stem)

    except Exception as e:
        print(e)

    return signatures

# ============================================================
# ACTION BUTTONS
# ============================================================

button_frame = tk.Frame(root)
button_frame.pack(pady=10)

tk.Button(
    button_frame,
    text="Preview All",
    command=preview_messages
).grid(row=0, column=0, padx=10)

tk.Button(
    button_frame,
    text="Start Review",
    bg="green",
    fg="white",
    command=start_process
).grid(row=0, column=1, padx=10)

# ============================================================
# STATUS
# ============================================================

tk.Label(
    root,
    textvariable=status_var
).pack()

root.mainloop()