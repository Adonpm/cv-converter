"""
CV conversion view
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path

class ConversionView:
    def __init__(self, parent_frame, controller):
        self.parent_frame = parent_frame
        self.controller = controller
        self.input_files = []
        self.create_widgets()

    def create_widgets(self):
        """Create conversion view widgets"""
        # Main container
        main_frame = ttk.Frame(self.parent_frame)
        main_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        main_frame.grid_columnconfigure(1, weight=1)

        # Input files section
        input_label = ttk.Label(main_frame, text="Input Files:", font=("Arial", 12, "bold"))
        input_label.grid(row=0, column=0, sticky="w", pady=(0,5))

        files_frame = ttk.Frame(main_frame)
        files_frame.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0,10))
        files_frame.grid_columnconfigure(0, weight=1)

        # Files listbox
        self.file_listbox = tk.Listbox(files_frame, height=6)
        self.file_listbox.grid(row=0, column=0, sticky="ew")

        # Files buttons
        files_btn_frame = ttk.Frame(files_frame)
        files_btn_frame.grid(row=0, column=1, sticky="ns", padx=(5, 0))

        add_file_button = ttk.Button(files_btn_frame, text="Add files", command=self.add_files)
        add_file_button.grid(row=0, column=0, pady=2, sticky="ew")

        remove_file_button = ttk.Button(files_btn_frame, text="Remove", command=self.remove_files)
        remove_file_button.grid(row=1, column=0, pady=2, sticky="ew")

        clearall_file_button = ttk.Button(files_btn_frame, text="Clear All", command=self.clear_files)
        clearall_file_button.grid(row=2, column=0, pady=2, sticky="ew")

        # Template selection
        template_label = ttk.Label(main_frame, text="Template:", font=("Arial", 12, "bold"))
        template_label.grid(row=2, column=0, sticky="w", pady=(10, 5))

        self.template_var = tk.StringVar(value = self.controller.get("default_template"))
        self.template_combo = ttk.Combobox(
            main_frame,
            textvariable=self.template_var,
            state="readonly"
        )
        self.template_combo.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(0,10))

        # Formatting options
        format_frame = ttk.LabelFrame(main_frame, text="Formatting options", padding=10)
        format_frame.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(0,10))
        format_frame.grid_columnconfigure(1, weight=1)

        # Font family
        ttk.Label(format_frame, text="Font Family:").grid(row=0, column=0, sticky="w", pady=2)
        self.font_family_var = tk.StringVar(value=self.controller.get("font_family"))
        font_combo = ttk.Combobox(
            format_frame, 
            textvariable=self.font_family_var,
            values=["Arial", "Calibri", "Times New Roman", "Helvetica"],
            state="readonly"
        )
        font_combo.grid(row=0, column=1, sticky="ew", padx=(10, 0), pady=2)

        # Font size
        ttk.Label(format_frame, text="Font Size:").grid(row=1, column=0, sticky="w", pady=2)
        self.font_size_var = tk.StringVar(value=self.controller.get("font_size"))
        size_combo = ttk.Combobox(
            format_frame,
            textvariable=self.font_size_var,
            values=["8", "9", "10", "11", "12", "14", "16"],
            state="readonly"
        )
        size_combo.grid(row=1, column=1, sticky="ew", padx=(10, 0), pady=2)

        # Client and opportunity
        client_frame = ttk.LabelFrame(main_frame, text="Bid Information", padding=10)
        client_frame.grid(row=5, column=0, columnspan=2, sticky="ew", pady=(0, 10))
        client_frame.grid_columnconfigure(1, weight=1)
        
        ttk.Label(client_frame, text="Client:").grid(row=0, column=0, sticky="w", pady=2)
        self.client_var = tk.StringVar(value=self.controller.get("client_name"))
        client_entry = ttk.Entry(client_frame, textvariable=self.client_var)
        client_entry.grid(row=0, column=1, sticky="ew", padx=(10, 0), pady=2)
        
        ttk.Label(client_frame, text="Opportunity:").grid(row=1, column=0, sticky="w", pady=2)
        self.opportunity_var = tk.StringVar(value=self.controller.get("opportunity_name"))
        opportunity_entry = ttk.Entry(client_frame, textvariable=self.opportunity_var)
        opportunity_entry.grid(row=1, column=1, sticky="ew", padx=(10, 0), pady=2)

        # Convert button
        convert_btn = ttk.Button(
            main_frame,
            text="Convert CVs",
            command=self.convert_cvs,
            style="Accent.TButton"
        )
        convert_btn.grid(row=6, column=0, columnspan=2, pady=20)

        # Load templates
        self.load_templates()

    def add_files(self):
        """Add input files"""
        file_types = [("Word Documents", "*.docx"), ("All Files", "*.*")]
        files = filedialog.askopenfilenames(
            title="Select Whoz CV files",
            filetypes=file_types
        )

        for file in files:
            if file not in self.input_files:
                self.input_files.append(file)
                self.file_listbox.insert(tk.END, Path(file).name)

    def remove_files(self):
        """Remove selected files"""
        selection = self.file_listbox.curselection()
        if selection:
            for index in reversed(selection):
                self.file_listbox.delete(index)
                del self.input_files[index]

    def clear_files(self):
        """Clear all files"""
        self.file_listbox.delete(0, tk.END)
        self.input_files.clear()

    def load_templates(self):
        """Load available templates"""
        templates = self.controller.get_templates()
        self.template_combo['values'] = templates
        if templates:
            self.template_combo.set(templates[0])

    def convert_cvs(self):
        """Start CV conversion process"""
        if not self.input_files:
            messagebox.showwarning("No Files", "Please select at least one input file.")
            return
        
        if not self.template_var.get():
            messagebox.showwarning("No Template", "Please select a template.")
            return
        
        # Prepare conversion settings
        settings = {
            "template": self.template_var.get(),
            "font_family": self.font_family_var.get(),
            "font_size": int(self.font_size_var.get()),
            "client_name": self.client_var.get(),
            "opportunity_name": self.opportunity_var.get()
        }
        
        # Start conversion
        self.controller.convert_cvs(self.input_files, settings)
