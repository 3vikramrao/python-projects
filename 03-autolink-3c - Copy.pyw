import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pyperclip
import webbrowser
import time
import os
import gc

df = None
original_template = ""
user_cancelled = False

# ✅ NEW (professional reload)
current_file_path = None
last_loaded_timestamp = None


# -------------------------------------------------------------------
# FILE TIMESTAMP
# -------------------------------------------------------------------
def get_file_timestamp(path):
    try:
        return os.path.getmtime(path)
    except:
        return None


# -------------------------------------------------------------------
# RESET APP STATE
# -------------------------------------------------------------------
def reset_app_state():
    global df, user_cancelled, current_file_path, last_loaded_timestamp

    df = None
    user_cancelled = False
    current_file_path = None
    last_loaded_timestamp = None

    excel_path_var.set("")
    name_col_var.set("")
    link_col_var.set("")
    start_row_var.set("1")
    status_var.set("Memory cleared.")

    name_col_dropdown["values"] = []
    link_col_dropdown["values"] = []
    name_col_dropdown.set("")
    link_col_dropdown.set("")

    message_box.delete("1.0", tk.END)

    gc.collect()


# -------------------------------------------------------------------
# CHROME SETUP
# -------------------------------------------------------------------
chrome_paths = [
    "C:/Program Files/Google/Chrome/Application/chrome.exe",
    "C:/Program Files (x86)/Google/Chrome/Application/chrome.exe"
]

chrome_browser = None
for path in chrome_paths:
    if os.path.exists(path):
        webbrowser.register('chrome', None, webbrowser.BackgroundBrowser(path))
        chrome_browser = webbrowser.get('chrome')
        break

if chrome_browser is None:
    chrome_browser = webbrowser


# -------------------------------------------------------------------
# LOAD EXCEL (SMART RELOAD)
# -------------------------------------------------------------------
def load_excel_columns(file_path):
    global df, current_file_path, last_loaded_timestamp

    try:
        file_timestamp = get_file_timestamp(file_path)

        if file_path != current_file_path or file_timestamp != last_loaded_timestamp:

            df = pd.read_excel(file_path, engine="openpyxl")

            current_file_path = file_path
            last_loaded_timestamp = file_timestamp

            col_names = list(df.columns)

            name_col_var.set("")
            link_col_var.set("")
            name_col_dropdown.set("")
            link_col_dropdown.set("")

            name_col_dropdown["values"] = col_names
            link_col_dropdown["values"] = col_names

            status_var.set("✅ Fresh Excel loaded.")

        else:
            status_var.set("⚠ File unchanged. Using existing data.")

    except Exception as e:
        messagebox.showerror("Error", f"Load failed:\n{e}")


# -------------------------------------------------------------------
# FORCE RELOAD BUTTON
# -------------------------------------------------------------------
def force_reload_file():
    global df, last_loaded_timestamp

    if not current_file_path:
        messagebox.showerror("Error", "No file loaded.")
        return

    try:
        df = pd.read_excel(current_file_path, engine="openpyxl")
        last_loaded_timestamp = get_file_timestamp(current_file_path)

        status_var.set("🔄 File reloaded successfully.")

    except Exception as e:
        messagebox.showerror("Error", f"Reload failed:\n{e}")


# -------------------------------------------------------------------
# BROWSE
# -------------------------------------------------------------------
def browse_file():
    file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )

    if not file_path:
        return

    reset_app_state()

    global current_file_path
    current_file_path = file_path

    excel_path_var.set(file_path)
    load_excel_columns(file_path)


# -------------------------------------------------------------------
# PREVIEW
# -------------------------------------------------------------------
def preview_messages():
    if df is None:
        messagebox.showerror("Error", "Load Excel first.")
        return

    name_col = name_col_var.get()
    link_col = link_col_var.get()
    template = message_box.get("1.0", tk.END).strip()

    win = tk.Toplevel(root)
    win.title("Preview")
    win.geometry("600x500")

    txt = tk.Text(win, wrap="word")
    txt.pack(expand=True, fill="both")

    for idx, row in df.iterrows():
        name = str(row[name_col]).strip()
        url = str(row[link_col]).strip()
        msg = template.replace("{name}", name)

        txt.insert(tk.END, f"Row {idx+1}/{len(df)}\n{name}\n{url}\n\n{msg}\n\n---\n\n")


# -------------------------------------------------------------------
# POPUP
# -------------------------------------------------------------------
def show_confirmation_popup(idx, name_col, link_col, template):
    popup = tk.Toplevel(root)
    popup.title("Open LinkedIn?")
    popup.geometry("520x470")

    global user_cancelled

    def on_close():
        global user_cancelled
        user_cancelled = True
        popup.destroy()

    popup.protocol("WM_DELETE_WINDOW", on_close)

    text_widget = tk.Text(popup, wrap="word", height=14, padx=10, pady=10)
    text_widget.pack()

    def refresh(i):
        text_widget.config(state="normal")
        text_widget.delete("1.0", tk.END)

        row = df.iloc[i]
        name = str(row[name_col]).strip()
        url = str(row[link_col]).strip()
        msg = template.replace("{name}", name)

        text_widget.insert("1.0",
            f"Row {i+1} of {len(df)}\n\n"
            f"Name: {name}\n"
            f"URL:\n{url}\n\n"
            f"Message:\n{msg}"
        )

        text_widget.config(state="disabled")
        return name, url, msg

    current_idx = idx
    current_name, current_url, current_msg = refresh(current_idx)

    def next_row():
        nonlocal current_idx, current_name, current_url, current_msg
        if current_idx < len(df)-1:
            current_idx += 1
            current_name, current_url, current_msg = refresh(current_idx)

    def prev_row():
        nonlocal current_idx, current_name, current_url, current_msg
        if current_idx > 0:
            current_idx -= 1
            current_name, current_url, current_msg = refresh(current_idx)

    def open_profile():
        pyperclip.copy(current_msg)
        chrome_browser.open(current_url)
        popup.destroy()

    btns = tk.Frame(popup)
    btns.pack(pady=10)

    tk.Button(btns, text="PREVIOUS", command=prev_row).grid(row=0, column=0, padx=5)
    tk.Button(btns, text="NEXT", command=next_row).grid(row=0, column=1, padx=5)
    tk.Button(btns, text="OPEN PROFILE", bg="green", fg="white",
              command=open_profile).grid(row=1, column=0, padx=5, pady=10)
    tk.Button(btns, text="SKIP", command=popup.destroy).grid(row=1, column=1, padx=5)

    popup.transient(root)
    popup.grab_set()
    root.wait_window(popup)


# -------------------------------------------------------------------
# START PROCESS
# -------------------------------------------------------------------
def start_process():
    global user_cancelled, original_template

    if df is None:
        messagebox.showerror("Error", "Load file first.")
        return

    name_col = name_col_var.get()
    link_col = link_col_var.get()
    template = message_box.get("1.0", tk.END).strip()

    original_template = template
    user_cancelled = False

    start_row = int(start_row_var.get()) - 1

    for idx in range(start_row, len(df)):
        if user_cancelled:
            break

        show_confirmation_popup(idx, name_col, link_col, template)

        if user_cancelled:
            break

        time.sleep(0.2)

    if user_cancelled:
        status_var.set("Process stopped.")
    else:
        messagebox.showinfo("Done", "Completed.")


# -------------------------------------------------------------------
# GUI
# -------------------------------------------------------------------
root = tk.Tk()
root.title("LinkedIn Tool (Pro Reload)")
root.geometry("750x500")

excel_path_var = tk.StringVar()
name_col_var = tk.StringVar()
link_col_var = tk.StringVar()
status_var = tk.StringVar()
start_row_var = tk.StringVar(value="1")

tk.Label(root, text="Excel File").pack()
tk.Entry(root, textvariable=excel_path_var, width=60).pack()

top_btns = tk.Frame(root)
top_btns.pack(pady=5)

tk.Button(top_btns, text="Browse", command=browse_file).grid(row=0, column=0, padx=5)
tk.Button(top_btns, text="Reset Memory",
          command=reset_app_state,
          bg="orange").grid(row=0, column=1, padx=5)
tk.Button(top_btns, text="Reload File",
          command=force_reload_file,
          bg="#0078D7", fg="white").grid(row=0, column=2, padx=5)

tk.Label(root, text="Name Column").pack()
name_col_dropdown = ttk.Combobox(root, textvariable=name_col_var)
name_col_dropdown.pack()

tk.Label(root, text="LinkedIn Column").pack()
link_col_dropdown = ttk.Combobox(root, textvariable=link_col_var)
link_col_dropdown.pack()

tk.Label(root, text="Start Row").pack()
tk.Entry(root, textvariable=start_row_var, width=10).pack()

tk.Label(root, text="Message Template").pack()
message_box = tk.Text(root, height=8, width=70)
message_box.pack()

btns = tk.Frame(root)
btns.pack(pady=10)

tk.Button(btns, text="Preview", command=preview_messages).grid(row=0, column=0, padx=10)
tk.Button(btns, text="Start", command=start_process,
          bg="green", fg="white").grid(row=0, column=1, padx=10)

tk.Label(root, textvariable=status_var).pack()

root.mainloop()