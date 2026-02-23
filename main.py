"""
CSV Email Sender - Tkinter Desktop Application
A simple desktop app for sending bulk emails via SMTP
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import threading
import time
from csv_parser import CSVParser, generate_template_csv
from email_sender import EmailBatchSender
from config import (
    SMTP_PRESETS, DEFAULT_DELAY_MS, DEFAULT_RANDOMIZE_PERCENT,
    DEFAULT_DELAY_PRESETS, WINDOW_TITLE, WINDOW_WIDTH, WINDOW_HEIGHT, LOG_FONT
)


class EmailSenderApp:
    """Main Tkinter application"""
    
    def __init__(self, root):
        self.root = root
        self.root.title(WINDOW_TITLE)
        self.root.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}")
        self.root.resizable(False, False)
        
        # State variables
        self.csv_file_path = tk.StringVar()
        self.attachment_files = []
        self.smtp_preset = tk.StringVar(value="Gmail")
        self.smtp_host = tk.StringVar(value=SMTP_PRESETS["Gmail"]["host"])
        self.smtp_port = tk.IntVar(value=SMTP_PRESETS["Gmail"]["port"])
        self.smtp_email = tk.StringVar()
        self.smtp_password = tk.StringVar()
        self.use_tls = tk.BooleanVar(value=True)
        self.delay_ms = tk.IntVar(value=DEFAULT_DELAY_MS)
        self.randomize_percent = tk.IntVar(value=DEFAULT_RANDOMIZE_PERCENT)
        self.default_subject = tk.StringVar()
        self.default_body = tk.StringVar()
        self.cc_recipients = tk.StringVar()
        self.bcc_recipients = tk.StringVar()
        
        # Sending state
        self.batch_sender = EmailBatchSender()
        self.sending_start_time = None
        self.progress_update_job = None
        self.is_sending = False
        
        # CSV data
        self.csv_data = []
        
        self._setup_ui()
        self._bind_preset_change()

        # Initial log message
        self._log("Ready. Upload a CSV file and configure SMTP settings to begin.", "info")
    
    def _setup_ui(self):
        """Setup UI layout"""
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        row = 0
        
        # Section 1: Upload Files
        row = self._create_file_section(main_frame, row)
        
        # Section 2: SMTP Configuration
        row = self._create_smtp_section(main_frame, row)
        
        # Section 3: Configure Sending
        row = self._create_sending_config_section(main_frame, row)
        
        # Section 4: Send Emails
        row = self._create_send_section(main_frame, row)
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
    
    def _create_file_section(self, parent, start_row):
        """Create file upload section"""
        frame = ttk.LabelFrame(parent, text="Section 1: Upload Files", padding="10")
        frame.grid(row=start_row, column=0, sticky=(tk.W, tk.E), pady=5)
        frame.columnconfigure(1, weight=1)
        
        row = 0
        
        # CSV File
        ttk.Label(frame, text="Data File (CSV/Excel):").grid(row=row, column=0, sticky=tk.W, pady=2)
        csv_entry = ttk.Entry(frame, textvariable=self.csv_file_path, width=50)
        csv_entry.grid(row=row, column=1, sticky=(tk.W, tk.E), padx=5, pady=2)
        ttk.Button(frame, text="Browse...", command=self._browse_csv).grid(row=row, column=2, pady=2)
        row += 1
        
        # Attachment Files - with listbox
        ttk.Label(frame, text="Attachment Files (optional):").grid(row=row, column=0, sticky=tk.NW, pady=2)
        
        # Attachment list frame
        attachment_frame = ttk.Frame(frame)
        attachment_frame.grid(row=row, column=1, sticky=(tk.W, tk.E), padx=5, pady=2)
        attachment_frame.columnconfigure(0, weight=1)
        
        # Listbox for attachments
        self.attachment_listbox = tk.Listbox(attachment_frame, height=4, exportselection=False)
        self.attachment_listbox.grid(row=0, column=0, sticky=(tk.W, tk.E))
        
        # Scrollbar for listbox
        attachment_scrollbar = ttk.Scrollbar(attachment_frame, orient=tk.VERTICAL, command=self.attachment_listbox.yview)
        attachment_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.attachment_listbox.config(yscrollcommand=attachment_scrollbar.set)
        
        # Button frame
        button_frame = ttk.Frame(frame)
        button_frame.grid(row=row, column=2, rowspan=2, padx=5)
        
        ttk.Button(button_frame, text="Add...", command=self._browse_attachments).pack(fill=tk.X, pady=2)
        ttk.Button(button_frame, text="Remove", command=self._remove_attachment).pack(fill=tk.X, pady=2)
        ttk.Button(button_frame, text="Clear All", command=self._clear_attachments).pack(fill=tk.X, pady=2)
        row += 1
        
        # Help text for attachments
        help_text = "ℹ️ CSV attachments override these. CSV: attachment1.pdf, attachment2.pdf"
        ttk.Label(frame, text=help_text, foreground="gray", font=("Arial", 8)).grid(
            row=row, column=1, sticky=tk.W, padx=5, pady=(0, 5))
        row += 1
        
        # CSV Template
        ttk.Button(frame, text="Download CSV Template", command=self._download_template).grid(row=row, column=0, columnspan=2, sticky=tk.W, pady=5)
        
        return start_row + 1
    
    def _create_smtp_section(self, parent, start_row):
        """Create SMTP configuration section"""
        frame = ttk.LabelFrame(parent, text="Section 2: SMTP Configuration", padding="10")
        frame.grid(row=start_row, column=0, sticky=(tk.W, tk.E), pady=5)
        frame.columnconfigure(1, weight=1)
        
        row = 0
        
        # Presets
        ttk.Label(frame, text="Presets:").grid(row=row, column=0, sticky=tk.W, pady=2)
        preset_frame = ttk.Frame(frame)
        preset_frame.grid(row=row, column=1, sticky=tk.W, pady=2)
        
        for preset_name in SMTP_PRESETS.keys():
            ttk.Radiobutton(
                preset_frame,
                text=preset_name,
                variable=self.smtp_preset,
                value=preset_name,
                command=self._on_preset_change
            ).pack(side=tk.LEFT, padx=5)
        row += 1
        
        # SMTP Server
        ttk.Label(frame, text="SMTP Server:").grid(row=row, column=0, sticky=tk.W, pady=2)
        ttk.Entry(frame, textvariable=self.smtp_host, width=30).grid(row=row, column=1, sticky=tk.W, pady=2)
        row += 1
        
        # Port
        ttk.Label(frame, text="Port:").grid(row=row, column=0, sticky=tk.W, pady=2)
        ttk.Entry(frame, textvariable=self.smtp_port, width=10).grid(row=row, column=1, sticky=tk.W, pady=2)
        ttk.Checkbutton(frame, text="Use TLS", variable=self.use_tls).grid(row=row, column=1, sticky=tk.W, padx=(100, 0))
        row += 1
        
        # Email
        ttk.Label(frame, text="Email:").grid(row=row, column=0, sticky=tk.W, pady=2)
        ttk.Entry(frame, textvariable=self.smtp_email, width=30).grid(row=row, column=1, sticky=(tk.W, tk.E), pady=2)
        row += 1
        
        # Password
        ttk.Label(frame, text="Password:").grid(row=row, column=0, sticky=tk.W, pady=2)
        ttk.Entry(frame, textvariable=self.smtp_password, show="•", width=30).grid(row=row, column=1, sticky=(tk.W, tk.E), pady=2)
        
        # App password help
        self.help_label = tk.Label(frame, text=SMTP_PRESETS["Gmail"]["help_text"], fg="#D32F2F", cursor="hand2", font=("Arial", 9))
        self.help_label.grid(row=row, column=2, padx=5, pady=2)
        self.help_label.bind("<Button-1>", self._open_help_url)
        
        return start_row + 1
    
    def _create_sending_config_section(self, parent, start_row):
        """Create sending configuration section"""
        frame = ttk.LabelFrame(parent, text="Section 3: Configure Sending", padding="10")
        frame.grid(row=start_row, column=0, sticky=(tk.W, tk.E), pady=5)
        frame.columnconfigure(1, weight=1)
        
        row = 0
        
        # Delay
        ttk.Label(frame, text="Delay:").grid(row=row, column=0, sticky=tk.W, pady=2)
        delay_frame = ttk.Frame(frame)
        delay_frame.grid(row=row, column=1, sticky=tk.W, pady=2)
        
        ttk.Entry(delay_frame, textvariable=self.delay_ms, width=8).pack(side=tk.LEFT)
        ttk.Label(delay_frame, text="ms").pack(side=tk.LEFT, padx=5)
        
        for label, value in DEFAULT_DELAY_PRESETS.items():
            ttk.Button(delay_frame, text=label, width=3, command=lambda v=value: self.delay_ms.set(v)).pack(side=tk.LEFT, padx=2)
        
        ttk.Checkbutton(
            frame,
            text=f"Randomize ±{DEFAULT_RANDOMIZE_PERCENT}%",
            variable=self.randomize_percent
        ).grid(row=row, column=2, sticky=tk.W, padx=10)
        row += 1
        
        # Default Subject
        ttk.Label(frame, text="Default Subject:").grid(row=row, column=0, sticky=tk.W, pady=2)
        ttk.Entry(frame, textvariable=self.default_subject, width=50).grid(row=row, column=1, columnspan=3, sticky=(tk.W, tk.E), pady=2)
        row += 1
        
        # Default Body
        ttk.Label(frame, text="Default Body:").grid(row=row, column=0, sticky=tk.NW, pady=2)
        body_text = tk.Text(frame, width=60, height=3)
        body_text.grid(row=row, column=1, columnspan=3, sticky=(tk.W, tk.E), pady=2)
        self.body_text_widget = body_text
        row += 1
        
        # CC
        ttk.Label(frame, text="CC:").grid(row=row, column=0, sticky=tk.W, pady=2)
        ttk.Entry(frame, textvariable=self.cc_recipients, width=50).grid(row=row, column=1, columnspan=3, sticky=(tk.W, tk.E), pady=2)
        row += 1
        
        # BCC
        ttk.Label(frame, text="BCC:").grid(row=row, column=0, sticky=tk.W, pady=2)
        ttk.Entry(frame, textvariable=self.bcc_recipients, width=50).grid(row=row, column=1, columnspan=3, sticky=(tk.W, tk.E), pady=2)
        
        return start_row + 1
    
    def _create_send_section(self, parent, start_row):
        """Create send control section"""
        frame = ttk.LabelFrame(parent, text="Section 4: Send Emails", padding="10")
        frame.grid(row=start_row, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        frame.columnconfigure(1, weight=1)
        parent.rowconfigure(start_row, weight=1)
        
        row = 0
        
        # Status indicator
        self.status_frame = ttk.Frame(frame)
        self.status_frame.grid(row=row, column=0, columnspan=4, sticky=tk.W, pady=5)
        
        self.status_indicator = tk.Label(
            self.status_frame,
            text="●",
            font=("Arial", 16),
            fg="gray"
        )
        self.status_indicator.pack(side=tk.LEFT, padx=5)
        
        self.status_text = tk.Label(
            self.status_frame,
            text="Ready to send",
            font=("Arial", 10)
        )
        self.status_text.pack(side=tk.LEFT, padx=5)
        row += 1
        
        # Buttons
        button_frame = ttk.Frame(frame)
        button_frame.grid(row=row, column=0, columnspan=4, sticky=tk.W, pady=5)
        
        self.start_button = ttk.Button(button_frame, text="▶ Start Sending", command=self._start_sending)
        self.start_button.pack(side=tk.LEFT, padx=5)
        
        self.abort_button = ttk.Button(button_frame, text="■ Abort", command=self._abort_sending, state=tk.DISABLED)
        self.abort_button.pack(side=tk.LEFT, padx=5)
        row += 1
        
        # Progress bar
        ttk.Label(frame, text="Progress:").grid(row=row, column=0, sticky=tk.W, pady=2)
        self.progress_var = tk.DoubleVar(value=0)
        self.progress_bar = ttk.Progressbar(frame, variable=self.progress_var, maximum=100, length=400)
        self.progress_bar.grid(row=row, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=2, padx=5)
        self.progress_label = ttk.Label(frame, text="0%")
        self.progress_label.grid(row=row, column=3, pady=2)
        row += 1
        
        # Status
        self.stats_label = ttk.Label(frame, text="Sent: 0  Failed: 0  Time: 00:00:00")
        self.stats_label.grid(row=row, column=0, columnspan=4, sticky=tk.W, pady=5)
        row += 1
        
        # Progress Log
        ttk.Label(frame, text="Progress Log:").grid(row=row, column=0, sticky=tk.NW, pady=2)
        row += 1
        
        self.log_widget = scrolledtext.ScrolledText(frame, width=70, height=15, font=LOG_FONT)
        self.log_widget.grid(row=row, column=0, columnspan=4, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        frame.rowconfigure(row, weight=1)
        
        # Configure log tags
        self.log_widget.tag_config("success", foreground="green")
        self.log_widget.tag_config("error", foreground="red")
        self.log_widget.tag_config("info", foreground="blue")
        self.log_widget.tag_config("warning", foreground="orange")
        
        return start_row + 1
    
    def _bind_preset_change(self):
        """Bind preset change handler"""
        self.smtp_preset.trace_add('write', lambda *args: self._on_preset_change())
    
    def _on_preset_change(self, event=None):
        """Handle SMTP preset change"""
        preset = self.smtp_preset.get()
        if preset in SMTP_PRESETS:
            config = SMTP_PRESETS[preset]
            self.smtp_host.set(config["host"])
            self.smtp_port.set(config["port"])
            self.use_tls.set(config["use_tls"])
            
            # Update help label
            if config["help_text"]:
                self.help_label.config(text=config["help_text"])
                self.help_label.bind("<Button-1>", self._open_help_url)
            else:
                self.help_label.config(text="", cursor="")
                self.help_label.unbind("<Button-1>")
    
    def _open_help_url(self, event):
        """Open help URL in browser"""
        preset = self.smtp_preset.get()
        if preset in SMTP_PRESETS and SMTP_PRESETS[preset]["help_url"]:
            import webbrowser
            webbrowser.open(SMTP_PRESETS[preset]["help_url"])
    
    def _browse_csv(self):
        """Browse for CSV/Excel file"""
        file_path = filedialog.askopenfilename(
            title="Select Data File",
            filetypes=[
                ("All supported", "*.csv *.xlsx *.xls"),
                ("CSV files", "*.csv"),
                ("Excel files", "*.xlsx *.xls"),
                ("All files", "*.*")
            ]
        )
        if file_path:
            self.csv_file_path.set(file_path)
            self._parse_csv()
    
    def _parse_csv(self):
        """Parse the CSV file"""
        file_path = self.csv_file_path.get()
        if not file_path:
            return
        
        parser = CSVParser()
        if parser.parse_file(file_path):
            self.csv_data = parser.get_data()
            self._log(f"✓ Loaded {parser.get_row_count()} recipients from {os.path.basename(file_path)}", "success")
        else:
            self.csv_data = []
            messagebox.showerror("Parse Error", parser.get_error())
    
    def _browse_attachments(self):
        """Browse for attachment files"""
        files = filedialog.askopenfilenames(
            title="Select Attachment Files",
            filetypes=[
                ("All files", "*.*")
            ]
        )
        if files:
            for file_path in files:
                if file_path not in self.attachment_files:
                    self.attachment_files.append(file_path)
                    self.attachment_listbox.insert(tk.END, os.path.basename(file_path))
    
    def _remove_attachment(self):
        """Remove selected attachment"""
        selection = self.attachment_listbox.curselection()
        if selection:
            index = selection[0]
            self.attachment_listbox.delete(index)
            del self.attachment_files[index]
    
    def _clear_attachments(self):
        """Clear all attachments"""
        self.attachment_files = []
        self.attachment_listbox.delete(0, tk.END)
    
    def _download_template(self):
        """Download CSV template"""
        file_path = filedialog.asksaveasfilename(
            title="Save Template",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")]
        )
        if file_path:
            if generate_template_csv(file_path):
                messagebox.showinfo("Template Created", f"Template saved to {file_path}")
            else:
                messagebox.showerror("Error", "Failed to create template")
    
    def _start_sending(self):
        """Start sending emails"""
        # Validate inputs
        if not self.csv_file_path.get():
            messagebox.showerror("Validation Error", "Please select a data file")
            return
        
        if not self.smtp_email.get():
            messagebox.showerror("Validation Error", "Please enter your email address")
            return
        
        if not self.smtp_password.get():
            messagebox.showerror("Validation Error", "Please enter your password")
            return
        
        if not self.smtp_host.get():
            messagebox.showerror("Validation Error", "Please enter SMTP server address")
            return
        
        if not self.csv_data:
            self._parse_csv()
            if not self.csv_data:
                messagebox.showerror("Validation Error", "No valid data found in CSV file")
                return
        
        # Update UI to show sending state
        self._set_sending_state(True)
        
        # Clear log and reset progress
        self.log_widget.delete(1.0, tk.END)
        self.progress_var.set(0)
        self.progress_label.config(text="0%")
        self.sending_start_time = time.time()
        
        # Get default body from text widget
        default_body = self.body_text_widget.get(1.0, tk.END).strip()
        
        # Start sending in background
        self.batch_sender.start_sending(
            host=self.smtp_host.get(),
            port=self.smtp_port.get(),
            email=self.smtp_email.get(),
            password=self.smtp_password.get(),
            use_tls=self.use_tls.get(),
            data=self.csv_data,
            default_subject=self.default_subject.get(),
            default_body=default_body,
            cc=self.cc_recipients.get() or None,
            bcc=self.bcc_recipients.get() or None,
            attachment_files=self.attachment_files,
            delay_ms=self.delay_ms.get(),
            randomize_percent=self.randomize_percent.get()
        )
        
        # Start progress updates
        self._update_progress()
    
    def _abort_sending(self):
        """Abort sending"""
        self.batch_sender.abort()
        self._log("⚠ Aborting send operation...", "warning")
        self.status_text.config(text="Aborting...")
    
    def _set_sending_state(self, sending):
        """Update UI state based on sending status"""
        self.is_sending = sending
        
        if sending:
            # Disable start button, enable abort
            self.start_button.config(state=tk.DISABLED)
            self.abort_button.config(state=tk.NORMAL)
            
            # Update status indicator
            self.status_indicator.config(fg="#00A4EF")  # Blue for running
            self.status_text.config(text="⏳ Sending in progress...")
            
            # Disable CSV file selection
            # (optional - you can also disable other controls)
        else:
            # Enable start button, disable abort
            self.start_button.config(state=tk.NORMAL)
            self.abort_button.config(state=tk.DISABLED)
            
            # Update status indicator
            self.status_indicator.config(fg="gray")
            self.status_text.config(text="Ready to send")
    
    def _update_progress(self):
        """Update progress from queue"""
        update = self.batch_sender.get_progress()
        
        while update:
            status = update['status']
            current = update['current']
            total = update['total']
            message = update['message']
            
            # Update progress bar
            if total > 0:
                percent = int(current / total * 100)
                self.progress_var.set(percent)
                self.progress_label.config(text=f"{percent}%")
            
            # Log message
            if status == 'sent':
                self._log(message, "success")
            elif status in ('error', 'failed'):
                self._log(message, "error")
            elif status == 'aborted':
                self._log(message, "warning")
            elif status == 'complete':
                self._log(message, "success")
            else:
                self._log(message, "info")
            
            # Update stats
            elapsed = time.time() - self.sending_start_time if self.sending_start_time else 0
            elapsed_str = self._format_time(elapsed)
            
            if status == 'complete':
                self.stats_label.config(text=f"{message}  Time: {elapsed_str}")
                self._set_sending_state(False)
            elif status in ('error', 'aborted'):
                self._set_sending_state(False)
            else:
                # Count sent and failed from log
                sent_count = self._count_log_tags("success")
                failed_count = self._count_log_tags("error")
                self.stats_label.config(
                    text=f"Sent: {sent_count}  Failed: {failed_count}  Time: {elapsed_str}"
                )
            
            # Get next update
            update = self.batch_sender.get_progress()
        
        # Schedule next update
        if self.batch_sender.is_sending():
            self.progress_update_job = self.root.after(100, self._update_progress)
        else:
            self._set_sending_state(False)
    
    def _count_log_tags(self, tag):
        """Count occurrences of a tag in the log"""
        # This is a simple count - for accurate count, you'd need to track separately
        count = 0
        try:
            # Get all text
            text = self.log_widget.get(1.0, tk.END)
            # Count occurrences of messages with the tag (simplified)
            for line in text.split('\n'):
                if '✓ Sent' in text[:text.index(line) + len(line)]:
                    count += 1
        except:
            pass
        return count
    
    def _log(self, message: str, tag=None):
        """Add message to log"""
        self.log_widget.insert(tk.END, message + "\n", tag)
        self.log_widget.see(tk.END)
    
    def _format_time(self, seconds: float) -> str:
        """Format time as HH:MM:SS"""
        hours = int(seconds // 3600)
        minutes = int((seconds % 3600) // 60)
        secs = int(seconds % 60)
        return f"{hours:02d}:{minutes:02d}:{secs:02d}"


def main():
    """Main entry point"""
    root = tk.Tk()
    app = EmailSenderApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
