#!/usr/bin/env python3
"""
BCC Bulk Mailer Data Cleaner - Interactive Chatbot GUI

An interactive, conversational interface that asks questions and confirms
before processing. Perfect for handling files with different formats.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except:
    TkinterDnD = None
import pandas as pd
import threading
from pathlib import Path
from datetime import datetime

# Import the cleaning functions
try:
    from clean_excel_data import clean_excel_file, parse_combined_address_regex, is_business_name
    CLEANER_AVAILABLE = True
except ImportError:
    CLEANER_AVAILABLE = False


class ChatMessage:
    """Represents a message in the chat"""
    def __init__(self, sender, text, widget_type=None, widget_data=None):
        self.sender = sender  # 'bot' or 'user'
        self.text = text
        self.widget_type = widget_type  # 'buttons', 'dropdown', 'preview', etc.
        self.widget_data = widget_data
        self.timestamp = datetime.now()


class InteractiveCleaner:
    def __init__(self, root):
        self.root = root
        self.root.title("BCC Bulk Mailer - Interactive Data Cleaner")
        self.root.geometry("1000x800")

        # State
        self.current_file = None
        self.df = None
        self.column_mapping = {}
        self.state = 'initial'  # initial, file_loaded, columns_mapped, ready
        self.fuzzy_threshold = 85

        # UI Setup
        self.setup_ui()
        self.show_welcome()

    def setup_ui(self):
        """Create the chat-like interface"""
        # Main container
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Header
        header = ttk.Frame(main_frame)
        header.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(
            header,
            text="🤖 BCC Bulk Mailer Interactive Cleaner",
            font=('Arial', 14, 'bold')
        ).pack()

        ttk.Label(
            header,
            text="I'll help you clean your mailing list step by step",
            font=('Arial', 9)
        ).pack()

        # Chat area (scrollable)
        chat_frame = ttk.Frame(main_frame)
        chat_frame.pack(fill=tk.BOTH, expand=True)

        # Scrollbar
        scrollbar = ttk.Scrollbar(chat_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Canvas for chat messages
        self.chat_canvas = tk.Canvas(
            chat_frame,
            bg='#f5f5f5',
            highlightthickness=0,
            yscrollcommand=scrollbar.set
        )
        self.chat_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.chat_canvas.yview)

        # Frame inside canvas to hold messages
        self.messages_frame = ttk.Frame(self.chat_canvas)
        self.canvas_frame = self.chat_canvas.create_window(
            (0, 0),
            window=self.messages_frame,
            anchor='nw'
        )

        # Bind resize
        self.messages_frame.bind('<Configure>', self._on_frame_configure)
        self.chat_canvas.bind('<Configure>', self._on_canvas_configure)

        # Input area
        input_frame = ttk.Frame(main_frame)
        input_frame.pack(fill=tk.X, pady=(10, 0))

        self.input_entry = ttk.Entry(input_frame, font=('Arial', 10))
        self.input_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        self.input_entry.bind('<Return>', lambda e: self.handle_text_input())

        ttk.Button(
            input_frame,
            text="Send",
            command=self.handle_text_input
        ).pack(side=tk.LEFT)

        ttk.Button(
            input_frame,
            text="📂 Browse File",
            command=self.browse_file
        ).pack(side=tk.LEFT, padx=(5, 0))

    def _on_frame_configure(self, event=None):
        """Reset scroll region when frame size changes"""
        self.chat_canvas.configure(scrollregion=self.chat_canvas.bbox('all'))

    def _on_canvas_configure(self, event):
        """Adjust frame width when canvas is resized"""
        self.chat_canvas.itemconfig(self.canvas_frame, width=event.width)

    def scroll_to_bottom(self):
        """Scroll chat to bottom"""
        self.root.update_idletasks()
        self.chat_canvas.yview_moveto(1.0)

    def add_bot_message(self, text):
        """Add a bot message to the chat"""
        msg_frame = ttk.Frame(self.messages_frame)
        msg_frame.pack(fill=tk.X, pady=5, padx=10)

        # Bot icon
        icon_label = ttk.Label(msg_frame, text="🤖", font=('Arial', 14))
        icon_label.pack(side=tk.LEFT, anchor='n', padx=(0, 10))

        # Message bubble
        bubble = tk.Frame(msg_frame, bg='#e3f2fd', relief=tk.RAISED, borderwidth=1)
        bubble.pack(side=tk.LEFT, fill=tk.X, expand=True)

        msg_label = tk.Label(
            bubble,
            text=text,
            bg='#e3f2fd',
            fg='#000',
            font=('Arial', 10),
            wraplength=600,
            justify=tk.LEFT,
            padx=15,
            pady=10
        )
        msg_label.pack()

        self.scroll_to_bottom()

    def add_user_message(self, text):
        """Add a user message to the chat"""
        msg_frame = ttk.Frame(self.messages_frame)
        msg_frame.pack(fill=tk.X, pady=5, padx=10)

        # Message bubble (right-aligned)
        bubble = tk.Frame(msg_frame, bg='#c8e6c9', relief=tk.RAISED, borderwidth=1)
        bubble.pack(side=tk.RIGHT, fill=tk.X)

        msg_label = tk.Label(
            bubble,
            text=text,
            bg='#c8e6c9',
            fg='#000',
            font=('Arial', 10),
            wraplength=600,
            justify=tk.LEFT,
            padx=15,
            pady=10
        )
        msg_label.pack()

        # User icon
        icon_label = ttk.Label(msg_frame, text="👤", font=('Arial', 14))
        icon_label.pack(side=tk.RIGHT, anchor='n', padx=(10, 0))

        self.scroll_to_bottom()

    def add_button_group(self, buttons):
        """Add a group of buttons (for user to click)"""
        btn_frame = ttk.Frame(self.messages_frame)
        btn_frame.pack(fill=tk.X, pady=5, padx=(60, 10))

        for text, command in buttons:
            btn = ttk.Button(btn_frame, text=text, command=command, width=20)
            btn.pack(side=tk.LEFT, padx=5)

        self.scroll_to_bottom()

    def add_dropdown_selector(self, label, options, callback):
        """Add a dropdown selector"""
        selector_frame = ttk.Frame(self.messages_frame)
        selector_frame.pack(fill=tk.X, pady=5, padx=(60, 10))

        ttk.Label(selector_frame, text=label, font=('Arial', 10)).pack(side=tk.LEFT, padx=(0, 10))

        var = tk.StringVar()
        dropdown = ttk.Combobox(selector_frame, textvariable=var, values=options, state='readonly', width=30)
        dropdown.pack(side=tk.LEFT, padx=5)
        dropdown.current(0)

        ttk.Button(
            selector_frame,
            text="✓ Confirm",
            command=lambda: callback(var.get())
        ).pack(side=tk.LEFT, padx=5)

        self.scroll_to_bottom()

    def add_data_preview(self, df, title="Preview"):
        """Add a data preview table"""
        preview_frame = ttk.LabelFrame(self.messages_frame, text=title, padding="10")
        preview_frame.pack(fill=tk.X, pady=5, padx=(60, 10))

        # Create scrolled text for preview
        preview_text = scrolledtext.ScrolledText(
            preview_frame,
            wrap=tk.NONE,
            font=('Courier New', 9),
            height=8,
            width=80
        )
        preview_text.pack(fill=tk.BOTH, expand=True)

        # Show first 5 rows
        preview_text.insert(tk.END, df.head(5).to_string(index=False))
        preview_text.config(state=tk.DISABLED)

        self.scroll_to_bottom()

    def show_welcome(self):
        """Show welcome message"""
        self.add_bot_message(
            "👋 Hi! I'm your interactive data cleaning assistant.\n\n"
            "I'll help you prepare your Excel file for BCC Bulk Mailer.\n\n"
            "Just upload your file (drag & drop or click 'Browse File') and I'll guide you through the process!"
        )

    def browse_file(self):
        """Browse for file"""
        filename = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            self.load_file(filename)

    def handle_text_input(self):
        """Handle text input from user"""
        text = self.input_entry.get().strip()
        if not text:
            return

        self.add_user_message(text)
        self.input_entry.delete(0, tk.END)

        # Process based on state
        # (For now, just acknowledge)
        self.add_bot_message(f"I heard: '{text}'. Let me think about that... 🤔")

    def load_file(self, filepath):
        """Load and analyze the Excel file"""
        self.add_user_message(f"📂 Loading: {Path(filepath).name}")

        try:
            self.df = pd.read_excel(filepath)
            self.current_file = filepath

            self.add_bot_message(
                f"✓ Great! I loaded your file.\n\n"
                f"📊 **File Info:**\n"
                f"   • Rows: {len(self.df)}\n"
                f"   • Columns: {len(self.df.columns)}\n\n"
                f"Let me show you what I found..."
            )

            self.root.after(1000, self.analyze_columns)

        except Exception as e:
            self.add_bot_message(f"❌ Oops! I couldn't load that file.\n\nError: {str(e)}")

    def analyze_columns(self):
        """Analyze columns and ask user for mapping"""
        columns = list(self.df.columns)

        self.add_bot_message(
            f"📋 **I detected these columns:**\n\n" +
            "\n".join([f"   • {col}" for col in columns]) +
            "\n\nLet me show you a preview of your data:"
        )

        self.add_data_preview(self.df, "Your Data (first 5 rows)")

        self.root.after(1000, lambda: self.ask_column_mapping(columns))

    def ask_column_mapping(self, columns):
        """Ask user to map columns"""
        self.add_bot_message(
            "🤔 **Let's figure out what each column contains.**\n\n"
            "I'll ask you a few questions to make sure I process this correctly."
        )

        # Check for common patterns
        has_full_address = any('full' in str(c).lower() and 'address' in str(c).lower() for c in columns)
        has_address_name = any('address' in str(c).lower() and 'name' in str(c).lower() for c in columns)

        if has_full_address and has_address_name:
            self.add_bot_message(
                "💡 I notice you have 'Address Name' and 'Full Address' columns.\n\n"
                "This looks like a **property management** or **condo association** format!\n\n"
                "Should I:\n"
                "• Use 'Address Name' for the person/business name?\n"
                "• Parse 'Full Address' to extract street, city, state, ZIP?"
            )

            self.add_button_group([
                ("Yes, that's correct!", lambda: self.confirm_property_format()),
                ("No, let me map manually", lambda: self.manual_column_mapping())
            ])
        else:
            self.add_bot_message(
                "I'll need your help identifying which columns contain what information.\n\n"
                "Let's go through them one by one..."
            )
            self.manual_column_mapping()

    def confirm_property_format(self):
        """User confirmed property management format"""
        self.add_user_message("Yes, that's correct!")

        self.add_bot_message(
            "Perfect! 👍\n\n"
            "I'll process it like this:\n\n"
            "• **Address Name** → Person/Business name (auto-detects which)\n"
            "• **Full Address** → Splits into street, city, state, ZIP\n"
            "• **Address** (unit numbers) → Goes to Address2 field\n\n"
            "I'll also:\n"
            "• Remove duplicates (exact + fuzzy matching)\n"
            "• Validate ZIP codes\n"
            "• Normalize state abbreviations\n"
            "• Detect business names (LLC, Inc, etc.)\n\n"
            "Want to see a preview of how the first few rows will look?"
        )

        self.add_button_group([
            ("Yes, show me preview", lambda: self.show_preview()),
            ("Skip preview, just process", lambda: self.start_processing())
        ])

    def manual_column_mapping(self):
        """Let user manually map columns"""
        self.add_user_message("Let me map manually")

        columns = list(self.df.columns)

        self.add_bot_message(
            "No problem! Let's map your columns.\n\n"
            "**Which column has the NAMES (person or business)?**"
        )

        self.add_dropdown_selector(
            "Name column:",
            ["(skip)"] + columns,
            lambda val: self.save_name_column(val, columns)
        )

    def save_name_column(self, column, all_columns):
        """Save name column and ask for address"""
        if column != "(skip)":
            self.column_mapping['name'] = column
            self.add_user_message(f"Name column: {column}")

        self.add_bot_message(
            "Great! Now, **which column has the complete ADDRESS?**\n"
            "(This should have street, city, state, ZIP - all in one)"
        )

        self.add_dropdown_selector(
            "Address column:",
            ["(skip)"] + all_columns,
            lambda val: self.save_address_column(val, all_columns)
        )

    def save_address_column(self, column, all_columns):
        """Save address column and continue"""
        if column != "(skip)":
            self.column_mapping['address'] = column
            self.add_user_message(f"Address column: {column}")

        self.add_bot_message(
            "Perfect! I have enough to get started.\n\n"
            f"**Mapping Summary:**\n"
            f"   • Names: {self.column_mapping.get('name', '(not set)')}\n"
            f"   • Addresses: {self.column_mapping.get('address', '(not set)')}\n\n"
            "Ready to see a preview?"
        )

        self.add_button_group([
            ("Yes, show preview", lambda: self.show_preview()),
            ("Skip preview, process now", lambda: self.start_processing())
        ])

    def show_preview(self):
        """Show preview of cleaned data"""
        self.add_user_message("Show me a preview")

        self.add_bot_message(
            "Generating preview... this will take a moment."
        )

        # Process in thread
        threading.Thread(target=self._generate_preview, daemon=True).start()

    def _generate_preview(self):
        """Generate preview (in thread)"""
        try:
            # Create temp output
            import tempfile
            import os

            temp_output = tempfile.mktemp(suffix='.csv')

            # Process file
            clean_excel_file(self.current_file, temp_output, fuzzy_threshold=self.fuzzy_threshold, save_report=False)

            # Read result
            result_df = pd.read_csv(temp_output)

            # Clean up
            try:
                os.remove(temp_output)
            except:
                pass

            # Show preview
            self.root.after(0, lambda: self._show_preview_result(result_df))

        except Exception as e:
            self.root.after(0, lambda: self.add_bot_message(f"❌ Preview failed: {str(e)}"))

    def _show_preview_result(self, result_df):
        """Show the preview result"""
        self.add_bot_message(
            f"✨ **Here's how your data will look after cleaning:**\n\n"
            f"   • {len(result_df)} rows (after removing duplicates)\n"
            f"   • Ready for BCC Bulk Mailer import"
        )

        self.add_data_preview(result_df, "Cleaned Data Preview (first 5 rows)")

        self.add_bot_message(
            "**Look good?**\n\n"
            "If everything looks correct, click 'Process Full File'.\n"
            "If something looks wrong, click 'Start Over'."
        )

        self.add_button_group([
            ("✓ Process Full File", lambda: self.start_processing()),
            ("← Start Over", lambda: self.restart())
        ])

    def start_processing(self):
        """Start the full processing"""
        self.add_user_message("Process the full file")

        self.add_bot_message(
            "🚀 **Processing your file...**\n\n"
            "This may take a minute depending on file size."
        )

        # Process in thread
        threading.Thread(target=self._process_file, daemon=True).start()

    def _process_file(self):
        """Process file in background thread"""
        try:
            output_file = str(Path(self.current_file).parent / (Path(self.current_file).stem + "_BCC_cleaned.csv"))

            clean_excel_file(self.current_file, output_file, fuzzy_threshold=self.fuzzy_threshold, save_report=True)

            self.root.after(0, lambda: self._processing_complete(output_file))

        except Exception as e:
            self.root.after(0, lambda: self.add_bot_message(f"❌ Processing failed: {str(e)}"))

    def _processing_complete(self, output_file):
        """Processing completed successfully"""
        self.add_bot_message(
            f"✅ **Done!**\n\n"
            f"Your cleaned file is ready:\n"
            f"   📄 {Path(output_file).name}\n\n"
            f"**What I did:**\n"
            f"   • Parsed addresses into separate fields\n"
            f"   • Removed duplicate records\n"
            f"   • Validated ZIP codes and emails\n"
            f"   • Formatted phone numbers\n"
            f"   • Detected business names\n\n"
            f"**Next steps:**\n"
            f"   1. Check the report file (.txt) for details\n"
            f"   2. Import the CSV into BCC Bulk Mailer\n"
            f"   3. Run CASS validation in BCC"
        )

        self.add_button_group([
            ("Open Output Folder", lambda: self._open_folder(output_file)),
            ("Process Another File", lambda: self.restart())
        ])

    def _open_folder(self, filepath):
        """Open folder containing output file"""
        import os
        import subprocess
        import platform

        folder = str(Path(filepath).parent)

        if platform.system() == 'Windows':
            os.startfile(folder)
        elif platform.system() == 'Darwin':  # macOS
            subprocess.Popen(['open', folder])
        else:  # Linux
            subprocess.Popen(['xdg-open', folder])

        self.add_user_message("Opened output folder")

    def restart(self):
        """Restart the process"""
        self.add_user_message("Start over")

        # Clear state
        self.current_file = None
        self.df = None
        self.column_mapping = {}
        self.state = 'initial'

        self.add_bot_message(
            "🔄 **Starting fresh!**\n\n"
            "Upload a new file when you're ready."
        )


def main():
    """Main entry point"""
    try:
        if TkinterDnD:
            root = TkinterDnD.Tk()
        else:
            root = tk.Tk()
    except:
        root = tk.Tk()

    app = InteractiveCleaner(root)
    root.mainloop()


if __name__ == '__main__':
    main()
