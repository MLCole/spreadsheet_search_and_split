import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import re


class DataSegmenterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Location Segmenter v1.3")
        self.root.geometry("550x350")

        # --- WINDOW FOCUS CONTROLS ---
        # Forces window to the front on launch
        self.root.attributes('-topmost', True)
        self.root.focus_force()

        # Tkinter Variables
        self.file_path = tk.StringVar()
        self.dir_path = tk.StringVar()
        self.search_term = tk.StringVar()
        self.status_msg = tk.StringVar(value="Ready to process.")
        self.submitted = False

        # --- UI LAYOUT ---
        tk.Label(root, text="Spreadsheet Segmenter", font=("Arial", 14, "bold")).pack(pady=15)

        # File Selection
        f_frame = tk.Frame(root)
        f_frame.pack(fill="x", padx=20)
        tk.Label(f_frame, text="Source File:").pack(side="left")
        tk.Entry(f_frame, textvariable=self.file_path, width=40).pack(side="left", padx=5)
        tk.Button(f_frame, text="Browse", command=self.browse_file).pack(side="left")

        # Output Folder Selection
        d_frame = tk.Frame(root)
        d_frame.pack(fill="x", padx=20, pady=10)
        tk.Label(d_frame, text="Output Folder:").pack(side="left")
        tk.Entry(d_frame, textvariable=self.dir_path, width=38).pack(side="left", padx=5)
        tk.Button(d_frame, text="Browse", command=self.browse_directory).pack(side="left")

        # Search Term Input
        s_frame = tk.Frame(root)
        s_frame.pack(fill="x", padx=20)
        tk.Label(s_frame, text="Search Term:").pack(side="left")
        entry = tk.Entry(s_frame, textvariable=self.search_term, width=25)
        entry.pack(side="left", padx=5)
        entry.bind('<Return>', lambda e: self.submit())  # Allow 'Enter' key to trigger

        # Status Bar
        tk.Label(root, textvariable=self.status_msg, fg="blue", font=("Arial", 9, "italic")).pack(pady=10)

        # Buttons
        btn_frame = tk.Frame(root)
        btn_frame.pack(pady=15)
        tk.Button(btn_frame, text="Generate Chunk", bg="#2ecc71", fg="white", width=18, command=self.submit).pack(
            side="left", padx=10)
        tk.Button(btn_frame, text="Exit Program", bg="#e74c3c", fg="white", width=12, command=self.root.destroy).pack(
            side="left", padx=10)

    def browse_file(self):
        p = filedialog.askopenfilename(filetypes=[("Spreadsheets", "*.csv *.xlsx *.xls")])
        if p: self.file_path.set(p)

    def browse_directory(self):
        p = filedialog.askdirectory()
        if p: self.dir_path.set(p)

    def submit(self):
        if not self.file_path.get() or not self.search_term.get():
            messagebox.showwarning("Incomplete", "Please select a source file and enter a search term.")
            return
        self.submitted = True
        self.root.quit()  # Break the mainloop to trigger processing


def safe_load(path):
    """
    Loads CSV/Excel with a fallback chain to handle encoding errors (0x96).
    """
    try:
        ext = path.lower().split('.')[-1]
        if ext == 'csv':
            # Chain of encodings to handle Windows-1252/Latin1 characters
            for enc in ['utf-8-sig', 'latin1', 'cp1252']:
                try:
                    return pd.read_csv(path, encoding=enc, low_memory=False, dtype=str)
                except (UnicodeDecodeError, LookupError):
                    continue
            raise Exception("Unrecognized character encoding in CSV.")
        else:
            return pd.read_excel(path, engine='openpyxl', dtype=str)
    except Exception as e:
        messagebox.showerror("Read Error", f"Could not open file: {e}")
        return None


# --- MAIN EXECUTION ---
if __name__ == "__main__":
    root = tk.Tk()
    app = DataSegmenterGUI(root)

    while True:
        root.mainloop()  # Open/Restore GUI

        # If the user closed the window or hit exit
        if not app.submitted:
            break

        app.status_msg.set("Processing data... please wait.")
        root.update_idletasks()

        # 1. Load Data
        df = safe_load(app.file_path.get())

        if df is not None:
            # 2. Whole Word Search Logic
            term = app.search_term.get().strip()
            # \b ensures we match the whole word (e.g., TX1 matches 'TX1 Site' but not 'TX10')
            pattern = rf'\b{re.escape(term)}\b'

            # Scan every cell; na=False prevents crashes on empty columns
            mask = df.apply(lambda row: row.astype(str).str.contains(pattern, case=False, regex=True, na=False).any(),
                            axis=1)
            results = df[mask]

            if results.empty:
                messagebox.showinfo("No Results", f"No whole-word matches found for '{term}'")
            else:
                # 3. Custom Naming Logic
                # Get filename without path or extension
                orig_base = os.path.splitext(os.path.basename(app.file_path.get()))[0]
                # Replace spaces in search term with underscores
                safe_term = term.replace(" ", "_")
                ext = os.path.splitext(app.file_path.get())[1].lower()

                # Determine save folder (fallback to source folder if not selected)
                out_dir = app.dir_path.get() if app.dir_path.get() else os.path.dirname(app.file_path.get())
                out_name = f"{orig_base}_{safe_term}{ext}"
                final_out = os.path.join(out_dir, out_name)

                # 4. Save file
                try:
                    if ext == '.csv':
                        results.to_csv(final_out, index=False)
                    else:
                        results.to_excel(final_out, index=False, engine='openpyxl')

                    messagebox.showinfo("Success", f"Found {len(results)} rows.\nSaved as: {out_name}")
                    app.search_term.set("")  # Reset search for the next location
                except Exception as e:
                    messagebox.showerror("Save Error", f"Failed to save file:\n{e}")

        # Reset flag and status to allow next run
        app.status_msg.set("Ready for next search.")
        app.submitted = False

    # Fully close the window on exit
    try:
        root.destroy()
    except:
        pass
