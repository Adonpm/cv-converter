"""
Progress bar component
"""
import tkinter as tk
from tkinter import ttk

class ProgressDialog:
    def __init__(self, parent, title="Processing..."):
        self.parent = parent
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(title)
        self.dialog.geometry("400x150")
        self.dialog.resizable(False, False)

        # Configure dialog behavior
        self.dialog.transient(parent) 
        self.dialog.grab_set()

        self.cancelled = False

        self.create_widgets()

    def create_widgets(self):
        """Create progress dialog widgets"""
        main_frame = ttk.Frame(self.dialog, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Status label
        self.status_label = ttk.Label(main_frame, text="Initializing...", font=("Arial", 10))
        self.status_label.pack(pady=(0, 10))

        # Progress bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            main_frame,
            variable=self.progress_var,
            maximum=100,
            length=300
        )
        self.progress_bar.pack(pady=(0,10))

        # Progress text
        self.progress_text = ttk.Label(
            main_frame,
            text="0%",
            font=("Arial", 9)
        )
        self.progress_text.pack(pady=(0, 10))

        # Cancel button
        self.cancel_btn = ttk.Button(
            main_frame,
            text="Cancel",
            command=self.cancel
        )
        self.cancel_btn.pack()

    def update_progress(self, current, total, status="Processing..."):
        """Update progress"""
        if total>0:
            percentage = (current/total) * 100
            self.progress_var.set(percentage)
            self.progress_text.config(text=f"{percentage:.1f}% ({current}/{total})")
            self.status_label.config(text=status)
            self.dialog.update_idletasks()

    def cancel(self):
        """Cancel the operation"""
        self.cancelled = True
        self.dialog.destroy()

    def close(self):
        """Close the progress dialog"""
        self.dialog.destroy()
