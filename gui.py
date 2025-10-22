#!/usr/bin/env python3
"""
BCC Bulk Mailer Data Cleaner - GUI Application

A simple drag-and-drop interface for cleaning Excel mailing lists.
Perfect for teams who don't want to use command-line tools.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from tkinterdnd2 import DND_FILES, TkinterDnD
import threading
import sys
from pathlib import Path
from datetime import datetime

# Import the cleaning functions
try:
    from clean_excel_data import clean_excel_file, ProcessingStats
    CLEANER_AVAILABLE = True
except ImportError:
    CLEANER_AVAILABLE = False


class BulkMailerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("BCC Bulk Mailer Data Cleaner")
        self.root.geometry("900x700")
        self.root.resizable(True, True)

        # Styling
        self.setup_styles()

        # Variables
        self.input_file = tk.StringVar()
        self.output_file = tk.StringVar()
        self.fuzzy_threshold = tk.IntVar(value=85)
        self.save_report = tk.BooleanVar(value=True)
        self.processing = False

        # Build UI
        self.create_widgets()

        # Check if cleaner is available
        if not CLEANER_AVAILABLE:
            messagebox.showerror(
                "Error",
                "Could not import clean_excel_data module.\n\n"
                "Make sure clean_excel_data.py is in the same folder as this GUI."
            )

    def setup_styles(self):
        """Setup custom styles for widgets"""
        style = ttk.Style()
        style.theme_use('clam')

        # Custom colors
        style.configure('Title.TLabel', font=('Arial', 16, 'bold'), foreground='#2c3e50')
        style.configure('Subtitle.TLabel', font=('Arial', 10), foreground='#7f8c8d')
        style.configure('Info.TLabel', font=('Arial', 9), foreground='#34495e')
        style.configure('Success.TLabel', font=('Arial', 10, 'bold'), foreground='#27ae60')
        style.configure('Error.TLabel', font=('Arial', 10, 'bold'), foreground='#e74c3c')

        style.configure('Process.TButton', font=('Arial', 11, 'bold'), padding=10)
        style.map('Process.TButton',
                 background=[('active', '#27ae60'), ('!active', '#2ecc71')],
                 foreground=[('active', 'white'), ('!active', 'white')])

    def create_widgets(self):
        """Create all GUI widgets"""

        # Header
        header_frame = ttk.Frame(self.root, padding="20")
        header_frame.pack(fill=tk.X)

        title_label = ttk.Label(
            header_frame,
            text="🎯 BCC Bulk Mailer Data Cleaner",
            style='Title.TLabel'
        )
        title_label.pack()

        subtitle_label = ttk.Label(
            header_frame,
            text="Clean and validate your mailing lists for BCC Bulk Mailer",
            style='Subtitle.TLabel'
        )
        subtitle_label.pack()

        # Separator
        ttk.Separator(self.root, orient='horizontal').pack(fill=tk.X, pady=10)

        # Drop Zone
        self.create_drop_zone()

        # File Selection
        self.create_file_selection()

        # Options
        self.create_options()

        # Process Button
        self.create_process_button()

        # Progress
        self.create_progress_area()

        # Output/Report Area
        self.create_output_area()

        # Footer
        self.create_footer()

    def create_drop_zone(self):
        """Create drag-and-drop zone"""
        drop_frame = ttk.LabelFrame(self.root, text="📂 Drop Excel File Here", padding="20")
        drop_frame.pack(fill=tk.X, padx=20, pady=10)

        self.drop_label = ttk.Label(
            drop_frame,
            text="Drag and drop an Excel file (.xlsx or .xls)\nor click 'Browse' below",
            style='Info.TLabel',
            justify=tk.CENTER
        )
        self.drop_label.pack(pady=30)

        # Enable drag and drop
        try:
            drop_frame.drop_target_register(DND_FILES)
            drop_frame.dnd_bind('<<Drop>>', self.on_drop)
        except:
            # If drag-drop not available, just use button
            pass

    def create_file_selection(self):
        """Create file selection controls"""
        file_frame = ttk.LabelFrame(self.root, text="Files", padding="15")
        file_frame.pack(fill=tk.X, padx=20, pady=5)

        # Input file
        input_row = ttk.Frame(file_frame)
        input_row.pack(fill=tk.X, pady=5)

        ttk.Label(input_row, text="Input:", width=10).pack(side=tk.LEFT)

        input_entry = ttk.Entry(input_row, textvariable=self.input_file, state='readonly')
        input_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

        browse_btn = ttk.Button(input_row, text="Browse...", command=self.browse_input)
        browse_btn.pack(side=tk.LEFT)

        # Output file
        output_row = ttk.Frame(file_frame)
        output_row.pack(fill=tk.X, pady=5)

        ttk.Label(output_row, text="Output:", width=10).pack(side=tk.LEFT)

        output_entry = ttk.Entry(output_row, textvariable=self.output_file)
        output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

        ttk.Label(output_row, text="(auto-generated)", style='Subtitle.TLabel').pack(side=tk.LEFT)

    def create_options(self):
        """Create options controls"""
        options_frame = ttk.LabelFrame(self.root, text="⚙️ Options", padding="15")
        options_frame.pack(fill=tk.X, padx=20, pady=5)

        # Fuzzy threshold
        fuzzy_row = ttk.Frame(options_frame)
        fuzzy_row.pack(fill=tk.X, pady=5)

        ttk.Label(fuzzy_row, text="Fuzzy Matching:", width=20).pack(side=tk.LEFT)

        fuzzy_scale = ttk.Scale(
            fuzzy_row,
            from_=70,
            to=100,
            orient=tk.HORIZONTAL,
            variable=self.fuzzy_threshold,
            command=self.on_threshold_change
        )
        fuzzy_scale.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

        self.threshold_label = ttk.Label(fuzzy_row, text="85%", width=8)
        self.threshold_label.pack(side=tk.LEFT)

        # Help text
        help_text = ttk.Label(
            options_frame,
            text="Higher = stricter (fewer duplicates caught) | Lower = lenient (more duplicates caught)",
            style='Subtitle.TLabel'
        )
        help_text.pack(pady=2)

        # Save report checkbox
        report_check = ttk.Checkbutton(
            options_frame,
            text="Save detailed report (.txt file)",
            variable=self.save_report
        )
        report_check.pack(anchor=tk.W, pady=5)

    def create_process_button(self):
        """Create process button"""
        button_frame = ttk.Frame(self.root, padding="10")
        button_frame.pack(fill=tk.X, padx=20)

        self.process_btn = ttk.Button(
            button_frame,
            text="🚀 Clean Data",
            command=self.process_file,
            style='Process.TButton'
        )
        self.process_btn.pack(fill=tk.X)

    def create_progress_area(self):
        """Create progress indicator"""
        progress_frame = ttk.Frame(self.root, padding="10")
        progress_frame.pack(fill=tk.X, padx=20)

        self.progress_bar = ttk.Progressbar(
            progress_frame,
            mode='indeterminate',
            length=300
        )
        self.progress_bar.pack(fill=tk.X)

        self.status_label = ttk.Label(progress_frame, text="Ready", style='Info.TLabel')
        self.status_label.pack(pady=5)

    def create_output_area(self):
        """Create output/report display area"""
        output_frame = ttk.LabelFrame(self.root, text="📊 Processing Report", padding="10")
        output_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        self.output_text = scrolledtext.ScrolledText(
            output_frame,
            wrap=tk.WORD,
            font=('Courier New', 9),
            height=15
        )
        self.output_text.pack(fill=tk.BOTH, expand=True)

        # Initial message
        self.output_text.insert(tk.END, "👋 Welcome to BCC Bulk Mailer Data Cleaner!\n\n")
        self.output_text.insert(tk.END, "Instructions:\n")
        self.output_text.insert(tk.END, "1. Drag & drop an Excel file or click Browse\n")
        self.output_text.insert(tk.END, "2. Adjust options if needed\n")
        self.output_text.insert(tk.END, "3. Click 'Clean Data'\n")
        self.output_text.insert(tk.END, "4. Review the report and import to BCC Bulk Mailer\n\n")
        self.output_text.insert(tk.END, "Features:\n")
        self.output_text.insert(tk.END, "✓ Removes exact and fuzzy duplicates\n")
        self.output_text.insert(tk.END, "✓ Validates addresses, emails, phone numbers\n")
        self.output_text.insert(tk.END, "✓ Detects business names automatically\n")
        self.output_text.insert(tk.END, "✓ Formats everything for BCC Bulk Mailer\n")
        self.output_text.config(state=tk.DISABLED)

    def create_footer(self):
        """Create footer"""
        footer_frame = ttk.Frame(self.root, padding="10")
        footer_frame.pack(fill=tk.X, side=tk.BOTTOM)

        ttk.Label(
            footer_frame,
            text="© 2025 BCC Bulk Mailer Data Cleaner | Powered by Claude Code",
            style='Subtitle.TLabel'
        ).pack()

    def on_drop(self, event):
        """Handle file drop"""
        files = self.root.tk.splitlist(event.data)
        if files:
            file_path = files[0].strip('{}')
            if file_path.lower().endswith(('.xlsx', '.xls')):
                self.input_file.set(file_path)
                self.auto_set_output()
                self.update_drop_label(file_path)
            else:
                messagebox.showerror("Invalid File", "Please drop an Excel file (.xlsx or .xls)")

    def browse_input(self):
        """Browse for input file"""
        filename = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[
                ("Excel files", "*.xlsx *.xls"),
                ("All files", "*.*")
            ]
        )
        if filename:
            self.input_file.set(filename)
            self.auto_set_output()
            self.update_drop_label(filename)

    def auto_set_output(self):
        """Automatically set output filename"""
        if self.input_file.get():
            input_path = Path(self.input_file.get())
            output_path = input_path.stem + "_BCC_cleaned.csv"
            self.output_file.set(str(input_path.parent / output_path))

    def update_drop_label(self, filename):
        """Update drop zone label with filename"""
        file_name = Path(filename).name
        self.drop_label.config(text=f"✓ Selected: {file_name}")

    def on_threshold_change(self, value):
        """Update threshold label when slider moves"""
        threshold = int(float(value))
        self.threshold_label.config(text=f"{threshold}%")

    def log_output(self, message, clear=False):
        """Add message to output area"""
        self.output_text.config(state=tk.NORMAL)
        if clear:
            self.output_text.delete(1.0, tk.END)
        self.output_text.insert(tk.END, message + "\n")
        self.output_text.see(tk.END)
        self.output_text.config(state=tk.DISABLED)
        self.root.update()

    def update_status(self, message):
        """Update status label"""
        self.status_label.config(text=message)
        self.root.update()

    def process_file(self):
        """Process the selected file"""
        if self.processing:
            messagebox.showwarning("Processing", "Already processing a file. Please wait.")
            return

        if not self.input_file.get():
            messagebox.showerror("No File", "Please select an Excel file first.")
            return

        if not CLEANER_AVAILABLE:
            messagebox.showerror("Error", "Cleaning module not available.")
            return

        # Run in thread to avoid freezing UI
        thread = threading.Thread(target=self.run_cleaning, daemon=True)
        thread.start()

    def run_cleaning(self):
        """Run the cleaning process (in separate thread)"""
        self.processing = True

        # Update UI
        self.root.after(0, lambda: self.process_btn.config(state=tk.DISABLED))
        self.root.after(0, lambda: self.progress_bar.start(10))
        self.root.after(0, lambda: self.log_output("", clear=True))
        self.root.after(0, lambda: self.log_output("=" * 70))
        self.root.after(0, lambda: self.log_output("BCC BULK MAILER DATA CLEANING"))
        self.root.after(0, lambda: self.log_output("=" * 70))
        self.root.after(0, lambda: self.update_status("Processing..."))

        try:
            input_file = self.input_file.get()
            output_file = self.output_file.get() if self.output_file.get() else None
            threshold = self.fuzzy_threshold.get()
            save_report = self.save_report.get()

            self.root.after(0, lambda: self.log_output(f"\nInput:  {input_file}"))
            self.root.after(0, lambda: self.log_output(f"Output: {output_file}"))
            self.root.after(0, lambda: self.log_output(f"Fuzzy Threshold: {threshold}%"))
            self.root.after(0, lambda: self.log_output("\nProcessing..."))

            # Redirect stdout to capture print statements
            import io
            from contextlib import redirect_stdout

            output_buffer = io.StringIO()

            with redirect_stdout(output_buffer):
                result_file = clean_excel_file(
                    input_file,
                    output_file,
                    fuzzy_threshold=threshold,
                    save_report=save_report
                )

            # Get the output
            output_text = output_buffer.getvalue()

            # Display in GUI
            self.root.after(0, lambda: self.log_output("\n" + output_text))

            # Success
            self.root.after(0, lambda: self.update_status("✓ Complete!"))
            self.root.after(0, lambda: messagebox.showinfo(
                "Success!",
                f"Data cleaned successfully!\n\n"
                f"Output saved to:\n{result_file}\n\n"
                f"{'Report saved to: ' + str(Path(result_file).with_suffix('.txt')) if save_report else ''}\n\n"
                f"You can now import the CSV file into BCC Bulk Mailer."
            ))

        except Exception as e:
            error_msg = f"Error: {str(e)}"
            self.root.after(0, lambda: self.log_output(f"\n❌ {error_msg}"))
            self.root.after(0, lambda: self.update_status("Error occurred"))
            self.root.after(0, lambda: messagebox.showerror("Error", error_msg))

            import traceback
            self.root.after(0, lambda: self.log_output("\n" + traceback.format_exc()))

        finally:
            self.processing = False
            self.root.after(0, lambda: self.progress_bar.stop())
            self.root.after(0, lambda: self.process_btn.config(state=tk.NORMAL))


def main():
    """Main entry point"""
    try:
        # Try to use TkinterDnD for drag-and-drop
        root = TkinterDnD.Tk()
    except:
        # Fall back to regular Tkinter if drag-drop not available
        root = tk.Tk()
        print("Note: Drag-and-drop not available. Please install tkinterdnd2:")
        print("pip install tkinterdnd2")

    app = BulkMailerGUI(root)
    root.mainloop()


if __name__ == '__main__':
    main()
