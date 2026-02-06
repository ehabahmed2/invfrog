"""
PDF Invoice Extractor & Organizer
Main Application Entry Point
"""

import os
import sys
import threading
import queue
import shutil
import csv
import traceback
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
from typing import List

# Import our logic module
import parser
from parser import ParseResult, Status, NamingScheme

# --- Helpers ---
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class SplashScreen:
    def __init__(self, root):
        self.root = root
        self.window = tk.Toplevel(root)
        self.window.overrideredirect(True) # Remove border
        
        # Center splash
        w, h = 500, 300
        ws = self.root.winfo_screenwidth()
        hs = self.root.winfo_screenheight()
        x = (ws/2) - (w/2)
        y = (hs/2) - (h/2)
        self.window.geometry(f'{w}x{h}+{int(x)}+{int(y)}')
        
        # Colors
        bg_color = "#2c3e50"
        self.window.configure(bg=bg_color)
        
        # Try loading logo
        try:
            from PIL import Image, ImageTk
            img_path = resource_path(os.path.join('assets', 'logo.png'))
            if os.path.exists(img_path):
                load = Image.open(img_path)
                load = load.resize((100, 100))
                self.render = ImageTk.PhotoImage(load)
                img = tk.Label(self.window, image=self.render, bg=bg_color, bd=0)
                img.pack(pady=(40, 10))
        except Exception:
            # Fallback if PIL not installed or image missing
            tk.Label(self.window, text="[LOGO]", bg=bg_color, fg="white").pack(pady=50)

        tk.Label(self.window, text="PDF Invoice Extractor", 
                 font=("Segoe UI", 20, "bold"), bg=bg_color, fg="white").pack()
        tk.Label(self.window, text="Loading...", 
                 font=("Segoe UI", 10), bg=bg_color, fg="#bdc3c7").pack(pady=10)
        
        self.window.update()

    def close(self):
        self.window.destroy()

class InvoiceApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Invoice Extractor & Organizer")
        self.root.geometry("1000x750")
        
        # Data
        self.input_folder = tk.StringVar()
        self.output_folder = tk.StringVar()
        self.naming_var = tk.StringVar(value=NamingScheme.INVOICE_NUMBER.value)
        self.organize_by_month = tk.BooleanVar(value=False)
        self.dry_run = tk.BooleanVar(value=True)
        
        self.queue = queue.Queue()
        self.results: List[ParseResult] = []
        self.processing = False
        
        self.setup_ui()
        
    def setup_ui(self):
        # Styles
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("Header.TLabel", font=('Segoe UI', 12, 'bold'))
        style.configure("Success.TLabel", foreground="green")
        
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # --- Top Section: Header ---
        header = ttk.Frame(main_frame)
        header.pack(fill=tk.X, pady=(0,10))
        ttk.Label(header, text="Invoice Extractor", style="Header.TLabel").pack(side=tk.LEFT)
        
        # --- Folder Selection ---
        grp_folders = ttk.LabelFrame(main_frame, text="Locations", padding=10)
        grp_folders.pack(fill=tk.X, pady=5)
        
        # Input
        f1 = ttk.Frame(grp_folders)
        f1.pack(fill=tk.X, pady=2)
        ttk.Label(f1, text="Input PDF Folder:", width=15).pack(side=tk.LEFT)
        ttk.Entry(f1, textvariable=self.input_folder, width=50).pack(side=tk.LEFT, padx=5)
        ttk.Button(f1, text="Browse", command=self.browse_input).pack(side=tk.LEFT)
        
        # Output
        f2 = ttk.Frame(grp_folders)
        f2.pack(fill=tk.X, pady=2)
        ttk.Label(f2, text="Output Folder:", width=15).pack(side=tk.LEFT)
        ttk.Entry(f2, textvariable=self.output_folder, width=50).pack(side=tk.LEFT, padx=5)
        ttk.Button(f2, text="Browse", command=self.browse_output).pack(side=tk.LEFT)
        
        # --- Settings Panel ---
        grp_settings = ttk.LabelFrame(main_frame, text="Organizer Settings", padding=10)
        grp_settings.pack(fill=tk.X, pady=5)
        
        # Naming Scheme
        ttk.Label(grp_settings, text="Naming Scheme:", font=('Segoe UI', 9, 'bold')).pack(anchor=tk.W)
        frm_radio = ttk.Frame(grp_settings)
        frm_radio.pack(fill=tk.X, pady=2)
        ttk.Radiobutton(frm_radio, text="Invoice Number (INV_12345.pdf)", 
                        variable=self.naming_var, value=NamingScheme.INVOICE_NUMBER.value).pack(side=tk.LEFT, padx=10)
        ttk.Radiobutton(frm_radio, text="Vendor Name (Acme_INV_12345.pdf)", 
                        variable=self.naming_var, value=NamingScheme.VENDOR_NAME.value).pack(side=tk.LEFT, padx=10)
        ttk.Radiobutton(frm_radio, text="Original Filename", 
                        variable=self.naming_var, value=NamingScheme.ORIGINAL_FILENAME.value).pack(side=tk.LEFT, padx=10)
        
        # Options
        frm_opts = ttk.Frame(grp_settings)
        frm_opts.pack(fill=tk.X, pady=5)
        ttk.Checkbutton(frm_opts, text="Organize into folders by Year-Month", variable=self.organize_by_month).pack(side=tk.LEFT, padx=10)
        
        # Dry Run
        self.chk_dry = ttk.Checkbutton(frm_opts, text="DRY RUN (Preview only - No files copied)", variable=self.dry_run)
        self.chk_dry.pack(side=tk.LEFT, padx=20)
        
        # --- Action Bar ---
        btn_frame = ttk.Frame(main_frame, padding=10)
        btn_frame.pack(fill=tk.X)
        
        self.btn_start = ttk.Button(btn_frame, text="START PROCESSING", command=self.start_processing)
        self.btn_start.pack(side=tk.LEFT, ipadx=20, ipady=5)
        
        self.btn_open = ttk.Button(btn_frame, text="Open Output Folder", command=self.open_output, state=tk.DISABLED)
        self.btn_open.pack(side=tk.LEFT, padx=10)
        
        self.lbl_status = ttk.Label(btn_frame, text="Ready")
        self.lbl_status.pack(side=tk.RIGHT)
        
        # --- Progress ---
        self.progress = ttk.Progressbar(main_frame, mode='determinate')
        self.progress.pack(fill=tk.X, pady=5)
        
        # --- Results Table ---
        cols = ("orig", "new", "inv", "date", "amt", "status", "reason")
        self.tree = ttk.Treeview(main_frame, columns=cols, show='headings', selectmode='browse')
        
        self.tree.heading("orig", text="Original File")
        self.tree.heading("new", text="Proposed Filename")
        self.tree.heading("inv", text="Invoice #")
        self.tree.heading("date", text="Date")
        self.tree.heading("amt", text="Total")
        self.tree.heading("status", text="Status")
        self.tree.heading("reason", text="Details")
        
        self.tree.column("orig", width=150)
        self.tree.column("new", width=200)
        self.tree.column("inv", width=80)
        self.tree.column("date", width=80)
        self.tree.column("amt", width=70)
        self.tree.column("status", width=70)
        self.tree.column("reason", width=200)
        
        # Scrollbar
        sb = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

        # Tags for colors
        self.tree.tag_configure('OK', background='#d4edda') # Light Green
        self.tree.tag_configure('PARTIAL', background='#fff3cd') # Light Yellow
        self.tree.tag_configure('SKIPPED', background='#f8d7da') # Light Red

        # Start Queue Checker
        self.root.after(100, self.process_queue)

    def browse_input(self):
        d = filedialog.askdirectory()
        if d: 
            self.input_folder.set(d)
            if not self.output_folder.get():
                self.output_folder.set(d) # Default output to input

    def browse_output(self):
        d = filedialog.askdirectory()
        if d: self.output_folder.set(d)
        
    def open_output(self):
        d = self.output_folder.get()
        if os.path.exists(d):
            os.startfile(d)

    def start_processing(self):
        inp = self.input_folder.get()
        out = self.output_folder.get()
        
        if not inp or not os.path.isdir(inp):
            messagebox.showerror("Error", "Select a valid input folder.")
            return
            
        if not out:
            out = inp
            
        self.processing = True
        self.btn_start.config(state=tk.DISABLED)
        self.tree.delete(*self.tree.get_children())
        self.results = []
        self.btn_open.config(state=tk.DISABLED)
        
        # Config options
        scheme = NamingScheme(self.naming_var.get())
        by_month = self.organize_by_month.get()
        is_dry = self.dry_run.get()
        
        # Start Thread
        t = threading.Thread(target=self.worker, args=(inp, out, scheme, by_month, is_dry))
        t.start()
        
    def worker(self, input_dir, output_dir, scheme, by_month, is_dry):
        files = [f for f in os.listdir(input_dir) if f.lower().endswith('.pdf')]
        total = len(files)
        
        # Ensure output dir exists if not dry run
        if not is_dry and not os.path.exists(output_dir):
            try:
                os.makedirs(output_dir)
            except Exception as e:
                self.queue.put(("error", f"Cannot create output folder: {e}"))
                return

        for idx, filename in enumerate(files):
            self.queue.put(("progress", (idx / total) * 100, f"Processing {idx+1}/{total}: {filename}"))
            
            fpath = os.path.join(input_dir, filename)
            
            try:
                # 1. Parse
                res = parser.process_single_pdf(fpath, filename)
                
                # 2. Determine Paths
                res.proposed_filename = parser.generate_proposed_filename(filename, res.data, scheme)
                
                # If skipped, target path is N/A
                if res.status == parser.Status.SKIPPED:
                    res.target_path = "N/A"
                else:
                    target_full = parser.calculate_target_path(output_dir, res.proposed_filename, res.data, by_month)
                    
                    # 3. Handle File Operations (Rename/Copy)
                    if not is_dry:
                        try:
                            # Calculate safe unique path
                            safe_path = parser.get_safe_unique_path(target_full)
                            res.target_path = safe_path
                            
                            # Create subfolders
                            os.makedirs(os.path.dirname(safe_path), exist_ok=True)
                            
                            # COPY (Safer than move)
                            shutil.copy2(fpath, safe_path)
                            
                        except OSError as e:
                            res.status = parser.Status.SKIPPED
                            res.reason = f"Permission/Write Error: {e}"
                        except Exception as e:
                            res.status = parser.Status.SKIPPED
                            res.reason = f"System Error: {str(e)[:50]}"
                    else:
                        # Dry Run: Just show where it WOULD go
                        res.target_path = target_full + " (preview)"

                self.results.append(res)
                self.queue.put(("row", res))
                
            except Exception as e:
                # Catch-all
                tb = traceback.format_exc()
                print(tb) # To console for debug
                # Log to queue?
                
        # Generate Reports
        self.save_reports(output_dir, is_dry)
        
        self.queue.put(("done", is_dry))

    def save_reports(self, folder, is_dry):
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        name = "preview_run.csv" if is_dry else f"extraction_results_{ts}.csv"
        
        try:
            with open(os.path.join(folder, name), 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f)
                
                # 1. REMOVED "Proposed Filename" and "Target Path" from this list:
                writer.writerow(["Original Filename", "Invoice Number", "Date", "Total Amount", "Status", "Reason"])
                
                for r in self.results:
                    # 2. REMOVED r.proposed_filename and r.target_path from this list:
                    writer.writerow([
                        r.filename, 
                        r.data.get('invoice_number'), 
                        r.data.get('date'), 
                        r.data.get('total_amount'), 
                        r.status.value, 
                        r.reason
                    ])
        except: pass
        
        # ... (Keep the rest of the error logging code exactly as it was) ...
        if any(r.status == parser.Status.SKIPPED for r in self.results):
            try:
                with open(os.path.join(folder, "errors.log"), "a", encoding='utf-8') as f:
                    f.write(f"\n--- Run {ts} ---\n")
                    for r in [x for x in self.results if x.status == parser.Status.SKIPPED]:
                        f.write(f"{r.filename}: {r.reason}\n")
            except: pass
    
    
    def process_queue(self):
        try:
            while True:
                msg = self.queue.get_nowait()
                mtype = msg[0]
                
                if mtype == "progress":
                    val, txt = msg[1], msg[2]
                    self.progress['value'] = val
                    self.lbl_status.config(text=txt)
                    
                elif mtype == "row":
                    r = msg[1]
                    # Insert into tree
                    self.tree.insert("", "end", values=(
                        r.filename, r.proposed_filename, 
                        r.data.get('invoice_number'), r.data.get('date'), r.data.get('total_amount'),
                        r.status.value, r.reason
                    ), tags=(r.status.value,))
                    self.tree.yview_moveto(1)
                    
                elif mtype == "done":
                    is_dry = msg[1]
                    self.processing = False
                    self.btn_start.config(state=tk.NORMAL)
                    self.btn_open.config(state=tk.NORMAL)
                    self.progress['value'] = 100
                    
                    summary = f"Processed {len(self.results)} files."
                    if is_dry:
                        summary += " DRY RUN Complete (No files moved)."
                        messagebox.showinfo("Dry Run Complete", f"{summary}\nCheck 'preview_run.csv' in output folder.")
                    else:
                        messagebox.showinfo("Complete", f"{summary}\nFiles organized successfully.")
                    
                    self.lbl_status.config(text="Done")
                    
                elif mtype == "error":
                    messagebox.showerror("Error", msg[1])
                    self.processing = False
                    self.btn_start.config(state=tk.NORMAL)
                    
        except queue.Empty:
            pass
        finally:
            self.root.after(100, self.process_queue)

if __name__ == "__main__":
    # Ensure high DPI awareness
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except: pass

    root = tk.Tk()
    
    # Show splash briefly
    splash = SplashScreen(root)
    root.after(2500, splash.close) # Close after 2.5s
    root.after(2500, lambda: root.deiconify()) # Show main window
    
    root.withdraw() # Hide main window initially
    app = InvoiceApp(root)
    root.mainloop()