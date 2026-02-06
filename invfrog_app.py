"""
InvFrog - PDF Invoice Extractor & Organizer
A desktop tool to extract structured invoice data from machine-generated PDFs,
rename and organize them into folders.
"""

import os
import sys
import csv
import shutil
import threading
import queue
import traceback
from datetime import datetime
from pathlib import Path
from typing import List, Optional
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd
from parser import (
    process_single_pdf, ParseResult, Status, NamingScheme,
    generate_proposed_filename, generate_target_path, get_unique_filepath,
    sanitize_filename
)


def resource_path(relative_path: str) -> str:
    """Get absolute path to resource, works for dev and PyInstaller."""
    if getattr(sys, 'frozen', False):
        base = sys._MEIPASS
    else:
        base = os.path.abspath(os.path.dirname(__file__))
    return os.path.join(base, relative_path)


class SplashScreen:
    """Splash screen displayed during app startup."""
    
    def __init__(self):
        self.root = tk.Tk()
        self.root.overrideredirect(True)
        
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        width, height = 500, 500
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        self.root.geometry(f"{width}x{height}+{x}+{y}")
        
        self.root.configure(bg='#0a0a0a')
        
        self.load_splash_image()
        
        self.root.after(2500, self.close)
    
    def load_splash_image(self):
        """Load and display splash image."""
        try:
            from PIL import Image, ImageTk
            splash_path = resource_path('assets/splash.jpg')
            if os.path.exists(splash_path):
                img = Image.open(splash_path)
                img = img.resize((500, 500), Image.Resampling.LANCZOS)
                self.splash_img = ImageTk.PhotoImage(img)
                label = tk.Label(self.root, image=self.splash_img, bg='#0a0a0a')
                label.pack(fill=tk.BOTH, expand=True)
                return
        except ImportError:
            pass
        except Exception:
            pass
        
        frame = tk.Frame(self.root, bg='#0a0a0a')
        frame.pack(fill=tk.BOTH, expand=True)
        
        title = tk.Label(frame, text="InvFrog", font=('Segoe UI', 36, 'bold'),
                        fg='#00ff88', bg='#0a0a0a')
        title.pack(pady=(150, 20))
        
        subtitle = tk.Label(frame, text="PDF Invoice Extractor & Organizer",
                           font=('Segoe UI', 14), fg='#888888', bg='#0a0a0a')
        subtitle.pack(pady=10)
        
        loading = tk.Label(frame, text="Initializing...",
                          font=('Segoe UI', 11), fg='#00aa66', bg='#0a0a0a')
        loading.pack(pady=(50, 0))
    
    def close(self):
        self.root.destroy()
    
    def run(self):
        self.root.mainloop()


class InvFrogApp:
    """Main application class."""
    
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("InvFrog - PDF Invoice Extractor & Organizer")
        self.root.geometry("1100x800")
        self.root.minsize(900, 700)
        
        self.selected_folder = tk.StringVar(value="")
        self.output_folder = tk.StringVar(value="")
        self.naming_scheme = tk.StringVar(value="invoice_number")
        self.organize_by_month = tk.BooleanVar(value=False)
        self.dry_run_mode = tk.BooleanVar(value=True)
        self.open_after_extraction = tk.BooleanVar(value=True)
        
        self.status_queue = queue.Queue()
        self.is_processing = False
        self.results: List[ParseResult] = []
        self.last_excel_path: str = ""
        self.copied_count = 0
        
        self.logo_img = None
        
        self.setup_styles()
        self.create_widgets()
        self.check_queue()
    
    def setup_styles(self):
        style = ttk.Style()
        style.theme_use('clam')
        
        style.configure('Title.TLabel', font=('Segoe UI', 16, 'bold'), padding=5)
        style.configure('Subtitle.TLabel', font=('Segoe UI', 10), foreground='#666666')
        style.configure('Header.TLabel', font=('Segoe UI', 11, 'bold'))
        style.configure('Status.TLabel', font=('Segoe UI', 10))
        style.configure('Summary.TLabel', font=('Segoe UI', 11, 'bold'))
        style.configure('Big.TButton', font=('Segoe UI', 11), padding=(20, 10))
        style.configure('Action.TButton', font=('Segoe UI', 10), padding=(12, 6))
        style.configure('Green.TButton', font=('Segoe UI', 11, 'bold'), padding=(20, 10))
        
        style.configure('Treeview', font=('Segoe UI', 9), rowheight=26)
        style.configure('Treeview.Heading', font=('Segoe UI', 9, 'bold'))
        
        style.configure('Settings.TLabelframe', font=('Segoe UI', 10, 'bold'))
        style.configure('Settings.TLabelframe.Label', font=('Segoe UI', 10, 'bold'))
    
    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding=15)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        self.create_header(main_frame)
        self.create_folder_section(main_frame)
        self.create_settings_panel(main_frame)
        self.create_action_buttons(main_frame)
        self.create_progress_section(main_frame)
        self.create_treeview(main_frame)
        self.create_summary_section(main_frame)
        self.create_bottom_buttons(main_frame)
        self.create_help_text(main_frame)
    
    def create_header(self, parent):
        header_frame = ttk.Frame(parent)
        header_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.load_logo(header_frame)
        
        title_frame = ttk.Frame(header_frame)
        title_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        title_label = ttk.Label(title_frame, text="InvFrog", style='Title.TLabel',
                               foreground='#228B22')
        title_label.pack(anchor='w')
        
        subtitle_label = ttk.Label(
            title_frame,
            text="Extract invoice data from machine-generated PDFs to Excel. Rename & organize files.",
            style='Subtitle.TLabel',
            wraplength=700
        )
        subtitle_label.pack(anchor='w')
    
    def load_logo(self, parent):
        """Load and display logo image."""
        try:
            from PIL import Image, ImageTk
            logo_path = resource_path('assets/logo.jpg')
            if os.path.exists(logo_path):
                img = Image.open(logo_path)
                img = img.resize((80, 60), Image.Resampling.LANCZOS)
                self.logo_img = ImageTk.PhotoImage(img)
                logo_label = tk.Label(parent, image=self.logo_img)
                logo_label.pack(side=tk.LEFT, padx=(0, 15))
        except ImportError:
            pass
        except Exception:
            pass
    
    def create_folder_section(self, parent):
        folder_frame = ttk.Frame(parent)
        folder_frame.pack(fill=tk.X, pady=8)
        
        input_frame = ttk.Frame(folder_frame)
        input_frame.pack(fill=tk.X, pady=3)
        
        ttk.Label(input_frame, text="Input Folder:", width=12).pack(side=tk.LEFT)
        self.input_entry = ttk.Entry(input_frame, textvariable=self.selected_folder, width=55, state='readonly')
        self.input_entry.pack(side=tk.LEFT, padx=(5, 10), fill=tk.X, expand=True)
        self.browse_input_btn = ttk.Button(input_frame, text="Browse...",
                                           command=self.browse_input_folder, style='Action.TButton')
        self.browse_input_btn.pack(side=tk.LEFT)
        
        output_frame = ttk.Frame(folder_frame)
        output_frame.pack(fill=tk.X, pady=3)
        
        ttk.Label(output_frame, text="Output Folder:", width=12).pack(side=tk.LEFT)
        self.output_entry = ttk.Entry(output_frame, textvariable=self.output_folder, width=55, state='readonly')
        self.output_entry.pack(side=tk.LEFT, padx=(5, 10), fill=tk.X, expand=True)
        self.browse_output_btn = ttk.Button(output_frame, text="Browse...",
                                            command=self.browse_output_folder, style='Action.TButton')
        self.browse_output_btn.pack(side=tk.LEFT)
        ttk.Label(output_frame, text="(Default: Input folder)", style='Subtitle.TLabel').pack(side=tk.LEFT, padx=5)
    
    def create_settings_panel(self, parent):
        settings_frame = ttk.LabelFrame(parent, text="Settings", padding=10, style='Settings.TLabelframe')
        settings_frame.pack(fill=tk.X, pady=10)
        
        left_frame = ttk.Frame(settings_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        ttk.Label(left_frame, text="Naming Scheme:", style='Header.TLabel').pack(anchor='w')
        
        schemes = [
            ("Invoice Number (Recommended)", "invoice_number"),
            ("Vendor Name", "vendor_name"),
            ("Original Filename", "original_filename"),
        ]
        for text, value in schemes:
            rb = ttk.Radiobutton(left_frame, text=text, variable=self.naming_scheme, value=value)
            rb.pack(anchor='w', padx=15, pady=1)
        
        right_frame = ttk.Frame(settings_frame)
        right_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(30, 0))
        
        ttk.Label(right_frame, text="Options:", style='Header.TLabel').pack(anchor='w')
        
        organize_cb = ttk.Checkbutton(
            right_frame,
            text="Organize into folders by Year-Month (YYYY-MM)",
            variable=self.organize_by_month
        )
        organize_cb.pack(anchor='w', padx=15, pady=2)
        
        dry_run_cb = ttk.Checkbutton(
            right_frame,
            text="Dry Run (Preview only - no files copied)",
            variable=self.dry_run_mode
        )
        dry_run_cb.pack(anchor='w', padx=15, pady=2)
        
        open_cb = ttk.Checkbutton(
            right_frame,
            text="Open Excel file after extraction",
            variable=self.open_after_extraction
        )
        open_cb.pack(anchor='w', padx=15, pady=2)
    
    def create_action_buttons(self, parent):
        btn_frame = ttk.Frame(parent)
        btn_frame.pack(fill=tk.X, pady=10)
        
        self.extract_btn = ttk.Button(
            btn_frame,
            text="Start Extraction & Organize",
            command=self.start_extraction,
            style='Green.TButton',
            state='disabled'
        )
        self.extract_btn.pack(side=tk.LEFT)
        
        self.progress_label = ttk.Label(btn_frame, text="", style='Status.TLabel')
        self.progress_label.pack(side=tk.LEFT, padx=20)
    
    def create_progress_section(self, parent):
        self.progress_bar = ttk.Progressbar(parent, mode='determinate', length=400)
        self.progress_bar.pack(fill=tk.X, pady=5)
    
    def create_treeview(self, parent):
        tree_frame = ttk.Frame(parent)
        tree_frame.pack(fill=tk.BOTH, expand=True, pady=8)
        
        columns = ('filename', 'proposed', 'invoice_num', 'date', 'total', 'status', 'reason')
        self.tree = ttk.Treeview(tree_frame, columns=columns, show='headings', height=10)
        
        self.tree.heading('filename', text='Original Filename')
        self.tree.heading('proposed', text='Proposed New Name')
        self.tree.heading('invoice_num', text='Invoice #')
        self.tree.heading('date', text='Date')
        self.tree.heading('total', text='Total')
        self.tree.heading('status', text='Status')
        self.tree.heading('reason', text='Details')
        
        self.tree.column('filename', width=180, minwidth=120)
        self.tree.column('proposed', width=200, minwidth=150)
        self.tree.column('invoice_num', width=100, minwidth=70)
        self.tree.column('date', width=90, minwidth=70)
        self.tree.column('total', width=90, minwidth=70)
        self.tree.column('status', width=70, minwidth=60)
        self.tree.column('reason', width=200, minwidth=100)
        
        v_scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        h_scrollbar = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        self.tree.grid(row=0, column=0, sticky='nsew')
        v_scrollbar.grid(row=0, column=1, sticky='ns')
        h_scrollbar.grid(row=1, column=0, sticky='ew')
        
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        
        self.tree.tag_configure('ok', background='#d4edda')
        self.tree.tag_configure('partial', background='#fff3cd')
        self.tree.tag_configure('skipped', background='#f8d7da')
    
    def create_summary_section(self, parent):
        summary_frame = ttk.LabelFrame(parent, text="Summary", padding=10)
        summary_frame.pack(fill=tk.X, pady=8)
        
        self.summary_text = tk.StringVar(value="Select a folder and click 'Start Extraction' to begin.")
        summary_label = ttk.Label(summary_frame, textvariable=self.summary_text, style='Summary.TLabel')
        summary_label.pack(anchor='w')
    
    def create_bottom_buttons(self, parent):
        action_frame = ttk.Frame(parent)
        action_frame.pack(fill=tk.X, pady=8)
        
        self.open_folder_btn = ttk.Button(
            action_frame,
            text="Open Output Folder",
            command=self.open_output_folder,
            style='Action.TButton',
            state='disabled'
        )
        self.open_folder_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        self.export_preview_btn = ttk.Button(
            action_frame,
            text="Export Preview CSV",
            command=self.export_preview_csv,
            style='Action.TButton',
            state='disabled'
        )
        self.export_preview_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        self.export_errors_btn = ttk.Button(
            action_frame,
            text="Export Error Report",
            command=self.export_error_report,
            style='Action.TButton',
            state='disabled'
        )
        self.export_errors_btn.pack(side=tk.LEFT)
    
    def create_help_text(self, parent):
        help_text = (
            "Tip: Enable 'Dry Run' to preview file names before copying. "
            "Only machine-generated PDFs are supported (text must be selectable). "
            "Files are COPIED (not moved) to preserve originals."
        )
        help_label = ttk.Label(parent, text=help_text, style='Subtitle.TLabel', wraplength=1050)
        help_label.pack(anchor='w', pady=(5, 0))
    
    def browse_input_folder(self):
        folder = filedialog.askdirectory(title="Select Folder Containing PDF Invoices")
        if folder:
            self.selected_folder.set(folder)
            if not self.output_folder.get():
                self.output_folder.set(folder)
            self.extract_btn.configure(state='normal')
            self.tree.delete(*self.tree.get_children())
            self.summary_text.set("Ready to process. Click 'Start Extraction' to begin.")
    
    def browse_output_folder(self):
        folder = filedialog.askdirectory(title="Select Output Folder")
        if folder:
            self.output_folder.set(folder)
    
    def start_extraction(self):
        if self.is_processing:
            return
        
        input_folder = self.selected_folder.get()
        if not input_folder or not os.path.isdir(input_folder):
            messagebox.showerror("Error", "Please select a valid input folder.")
            return
        
        output_folder = self.output_folder.get() or input_folder
        if not os.path.isdir(output_folder):
            try:
                os.makedirs(output_folder, exist_ok=True)
            except Exception as e:
                messagebox.showerror("Error", f"Cannot create output folder: {e}")
                return
        
        pdf_files = [f for f in os.listdir(input_folder) if f.lower().endswith('.pdf')]
        if not pdf_files:
            messagebox.showwarning("No PDFs", "No PDF files found in the selected folder.")
            return
        
        self.is_processing = True
        self.results = []
        self.copied_count = 0
        self.tree.delete(*self.tree.get_children())
        
        self.extract_btn.configure(state='disabled')
        self.browse_input_btn.configure(state='disabled')
        self.browse_output_btn.configure(state='disabled')
        self.open_folder_btn.configure(state='disabled')
        self.export_preview_btn.configure(state='disabled')
        self.export_errors_btn.configure(state='disabled')
        
        self.progress_bar['maximum'] = len(pdf_files)
        self.progress_bar['value'] = 0
        
        naming = NamingScheme(self.naming_scheme.get())
        organize = self.organize_by_month.get()
        dry_run = self.dry_run_mode.get()
        
        thread = threading.Thread(
            target=self.process_files,
            args=(input_folder, output_folder, pdf_files, naming, organize, dry_run),
            daemon=True
        )
        thread.start()
    
    def process_files(self, input_folder: str, output_folder: str, pdf_files: List[str],
                      naming_scheme: NamingScheme, organize_by_month: bool, dry_run: bool):
        total = len(pdf_files)
        
        for idx, filename in enumerate(pdf_files, 1):
            filepath = os.path.join(input_folder, filename)
            
            self.status_queue.put(('progress', idx, total, filename))
            
            try:
                result = process_single_pdf(filepath, filename)
                
                proposed_name = generate_proposed_filename(filename, result.data, naming_scheme)
                target_path = generate_target_path(output_folder, proposed_name, result.data, organize_by_month)
                
                result.proposed_filename = proposed_name
                result.target_path = target_path
                
                if not dry_run and result.status != Status.SKIPPED:
                    copy_error = self.copy_file_safe(filepath, target_path)
                    if copy_error:
                        result.reason = f"{result.reason}; Copy failed: {copy_error}"
                    else:
                        self.copied_count += 1
                
            except Exception as e:
                result = ParseResult(
                    filename=filename,
                    status=Status.SKIPPED,
                    reason=f"Processing error: {str(e)[:80]}",
                    data={},
                    proposed_filename="",
                    target_path=""
                )
                self.log_error(output_folder, filename, e)
            
            self.results.append(result)
            self.status_queue.put(('file_done', result))
        
        self.save_outputs(output_folder, dry_run)
        self.status_queue.put(('complete', dry_run))
    
    def copy_file_safe(self, source: str, target: str) -> Optional[str]:
        """
        Safely copy a file to target path. Creates directories as needed.
        Returns error message on failure, None on success.
        """
        try:
            target_unique = get_unique_filepath(target)
            target_dir = os.path.dirname(target_unique)
            
            if not os.path.exists(target_dir):
                os.makedirs(target_dir, exist_ok=True)
            
            shutil.copy2(source, target_unique)
            return None
            
        except PermissionError:
            return "Permission denied"
        except OSError as e:
            return f"OS error: {str(e)[:50]}"
        except Exception as e:
            return f"Copy error: {str(e)[:50]}"
    
    def log_error(self, folder: str, filename: str, exception: Exception):
        log_path = os.path.join(folder, 'errors.log')
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        
        try:
            with open(log_path, 'a', encoding='utf-8') as f:
                f.write(f"\n{'='*60}\n")
                f.write(f"Timestamp: {timestamp}\n")
                f.write(f"File: {filename}\n")
                f.write(f"Error: {str(exception)}\n")
                f.write(f"Traceback:\n{traceback.format_exc()[:500]}\n")
        except:
            pass
    
    def save_outputs(self, folder: str, dry_run: bool):
        timestamp = datetime.now().strftime('%Y%m%d_%H%M')
        self.last_excel_path = ""
        
        data_rows = []
        for r in self.results:
            row = {
                'Original_Filename': r.filename,
                'Proposed_Filename': r.proposed_filename,
                'Target_Path': r.target_path,
                'Invoice_Number': r.data.get('invoice_number', ''),
                'Invoice_Date': r.data.get('date', ''),
                'Total_Amount': r.data.get('total_amount', ''),
                'Vendor': r.data.get('vendor', ''),
                'Status': r.status.value,
                'Details': r.reason
            }
            data_rows.append(row)
        
        if data_rows:
            df = pd.DataFrame(data_rows)
            if dry_run:
                excel_path = os.path.join(folder, f'Preview_{timestamp}.xlsx')
            else:
                excel_path = os.path.join(folder, f'Invoices_Extracted_{timestamp}.xlsx')
            try:
                df.to_excel(excel_path, index=False, engine='openpyxl')
                self.last_excel_path = excel_path
            except Exception:
                pass
        
        if dry_run:
            preview_path = os.path.join(folder, 'preview.csv')
            try:
                with open(preview_path, 'w', newline='', encoding='utf-8-sig') as f:
                    writer = csv.writer(f)
                    writer.writerow(['Original_Filename', 'Proposed_Filename', 'Target_Path', 'Status'])
                    for r in self.results:
                        writer.writerow([r.filename, r.proposed_filename, r.target_path, r.status.value])
            except Exception:
                pass
        
        skipped = [r for r in self.results if r.status == Status.SKIPPED]
        if skipped:
            csv_path = os.path.join(folder, 'skipped_files.csv')
            try:
                with open(csv_path, 'w', newline='', encoding='utf-8-sig') as f:
                    writer = csv.writer(f)
                    writer.writerow(['Filename', 'Reason'])
                    for r in skipped:
                        writer.writerow([r.filename, r.reason])
            except Exception:
                pass
    
    def check_queue(self):
        try:
            while True:
                msg = self.status_queue.get_nowait()
                
                if msg[0] == 'progress':
                    _, idx, total, filename = msg
                    self.progress_bar['value'] = idx
                    self.progress_label.configure(text=f"Processing {idx} of {total}: {filename[:40]}...")
                
                elif msg[0] == 'file_done':
                    result = msg[1]
                    tag = result.status.value.lower()
                    self.tree.insert('', 'end', values=(
                        result.filename[:30] + ('...' if len(result.filename) > 30 else ''),
                        result.proposed_filename[:35] + ('...' if len(result.proposed_filename) > 35 else ''),
                        result.data.get('invoice_number', '')[:15],
                        result.data.get('date', ''),
                        result.data.get('total_amount', ''),
                        result.status.value,
                        result.reason[:40] + ('...' if len(result.reason) > 40 else '')
                    ), tags=(tag,))
                    self.tree.yview_moveto(1.0)
                
                elif msg[0] == 'complete':
                    dry_run = msg[1]
                    self.finish_processing(dry_run)
        
        except queue.Empty:
            pass
        
        self.root.after(100, self.check_queue)
    
    def finish_processing(self, dry_run: bool):
        self.is_processing = False
        self.extract_btn.configure(state='normal')
        self.browse_input_btn.configure(state='normal')
        self.browse_output_btn.configure(state='normal')
        self.open_folder_btn.configure(state='normal')
        self.export_preview_btn.configure(state='normal')
        self.export_errors_btn.configure(state='normal')
        
        self.progress_label.configure(text="Complete!")
        
        ok_count = sum(1 for r in self.results if r.status == Status.OK)
        partial_count = sum(1 for r in self.results if r.status == Status.PARTIAL)
        skipped_count = sum(1 for r in self.results if r.status == Status.SKIPPED)
        total = len(self.results)
        
        if dry_run:
            summary = (
                f"DRY RUN: {total} files previewed   |   "
                f"OK: {ok_count}   |   "
                f"Partial: {partial_count}   |   "
                f"Skipped: {skipped_count}   |   "
                f"(No files copied)"
            )
            msg_title = "Dry Run Complete"
            msg_text = (
                f"Preview complete for {total} file(s).\n\n"
                f"OK: {ok_count}\n"
                f"Partial: {partial_count}\n"
                f"Skipped: {skipped_count}\n\n"
                f"Preview CSV saved. Disable 'Dry Run' and run again to copy files."
            )
        else:
            summary = (
                f"Processed: {total} files   |   "
                f"OK: {ok_count}   |   "
                f"Partial: {partial_count}   |   "
                f"Skipped: {skipped_count}   |   "
                f"Copied: {self.copied_count}"
            )
            msg_title = "Extraction Complete"
            msg_text = (
                f"Processed {total} file(s).\n\n"
                f"OK: {ok_count}\n"
                f"Partial: {partial_count}\n"
                f"Skipped: {skipped_count}\n"
                f"Files copied: {self.copied_count}\n\n"
                f"Excel file saved to output folder."
            )
        
        self.summary_text.set(summary)
        messagebox.showinfo(msg_title, msg_text)
        
        if self.open_after_extraction.get() and self.last_excel_path and os.path.exists(self.last_excel_path):
            try:
                os.startfile(self.last_excel_path)
            except:
                pass
    
    def open_output_folder(self):
        folder = self.output_folder.get() or self.selected_folder.get()
        if folder and os.path.isdir(folder):
            try:
                os.startfile(folder)
            except:
                pass
    
    def export_preview_csv(self):
        if not self.results:
            messagebox.showinfo("No Data", "No results to export. Run extraction first.")
            return
        
        folder = self.output_folder.get() or self.selected_folder.get()
        save_path = filedialog.asksaveasfilename(
            defaultextension='.csv',
            filetypes=[('CSV files', '*.csv')],
            initialfile='preview.csv',
            initialdir=folder
        )
        
        if save_path:
            try:
                with open(save_path, 'w', newline='', encoding='utf-8-sig') as f:
                    writer = csv.writer(f)
                    writer.writerow(['Original_Filename', 'Proposed_Filename', 'Target_Path',
                                    'Invoice_Number', 'Date', 'Total', 'Status', 'Details'])
                    for r in self.results:
                        writer.writerow([
                            r.filename, r.proposed_filename, r.target_path,
                            r.data.get('invoice_number', ''),
                            r.data.get('date', ''),
                            r.data.get('total_amount', ''),
                            r.status.value, r.reason
                        ])
                messagebox.showinfo("Saved", f"Preview CSV saved to:\n{save_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save: {e}")
    
    def export_error_report(self):
        if not self.results:
            messagebox.showinfo("No Data", "No results available.")
            return
        
        skipped = [r for r in self.results if r.status == Status.SKIPPED]
        if not skipped:
            messagebox.showinfo("No Errors", "No skipped files to report.")
            return
        
        folder = self.output_folder.get() or self.selected_folder.get()
        save_path = filedialog.asksaveasfilename(
            defaultextension='.csv',
            filetypes=[('CSV files', '*.csv')],
            initialfile='skipped_files.csv',
            initialdir=folder
        )
        
        if save_path:
            try:
                with open(save_path, 'w', newline='', encoding='utf-8-sig') as f:
                    writer = csv.writer(f)
                    writer.writerow(['Filename', 'Reason'])
                    for r in skipped:
                        writer.writerow([r.filename, r.reason])
                messagebox.showinfo("Saved", f"Error report saved to:\n{save_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save: {e}")


def main():
    try:
        splash = SplashScreen()
        splash.run()
    except:
        pass
    
    root = tk.Tk()
    
    if sys.platform == 'win32':
        try:
            from ctypes import windll
            windll.shcore.SetProcessDpiAwareness(1)
        except:
            pass
    
    try:
        icon_path = resource_path('assets/logo.ico')
        if os.path.exists(icon_path):
            root.iconbitmap(icon_path)
    except:
        pass
    
    app = InvFrogApp(root)
    root.mainloop()


if __name__ == '__main__':
    main()
