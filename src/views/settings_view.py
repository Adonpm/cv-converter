"""
Settings configuration view
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path

class SettingsView:
    def __init__(self, parent_frame, controller):
        self.parent_frame = parent_frame
        self.controller = controller
        self.create_widgets()
        self.load_current_settings()

    def create_widgets(self):
        """Create settings view widgets"""
        # Main container with scrollbar
        canvas = tk.Canvas(self.parent_frame)
        scrollbar = ttk.Scrollbar(self.parent_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "Configure",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0,0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        main_frame = ttk.Frame(scrollable_frame, padding=20)
        main_frame.pack(fill="both", expand=True)

        # Title
        title_label = ttk.Label(
            main_frame, 
            text="Application Settings",
            font=("Arial", 16, "bold")
        )
        title_label.pack(anchor="w", pady=(0, 20))

        # Default template settings
        template_frame = ttk.LabelFrame(
            main_frame,
            text="Default Template Settings",
            padding=15
            )
        template_frame.pack(fill="x", pady=(0,15))
        template_frame.grid_columnconfigure(1, weight=1)

        template_frame_label = ttk.Label(template_frame, text="Default template:")
        template_frame_label.grid(row=0, column=0, sticky="w", pady=5)
        self.default_template_var = tk.StringVar()
        self.default_template_combo = ttk.Combobox(
            template_frame,
            textvariable=self.default_template_var,
            state="readonly"
        )
        self.default_template_combo.grid(row=0, column=1, sticky="ew", padx=(10,0), pady=5)

        # Font settings
        font_frame = ttk.LabelFrame(
            main_frame,
            text="Default Font Settings",
            padding=15
        )
        font_frame.pack(fill="x", pady=(0,15))
        font_frame.grid_columnconfigure(1, weight=1)

        font_family_label = ttk.Label(font_frame, text="Font Family:")
        font_family_label.grid(row=0, column=0, sticky="w", pady=5)
        self.font_family_var = tk.StringVar()
        font_family_combo = ttk.Combobox(
            font_frame,
            textvariable=self.font_family_var,
            values=["Arial", "Calibri", "Times New Roman", "Helvetica", "Tahoma"],
            state="readonly"
        ) 
        font_family_combo.grid(row=0, column=1, sticky="ew", padx=(10, 0), pady=5)

        font_size_label = ttk.Label(font_frame, text="Font Size:")
        font_size_label.grid(row=1, column=0, sticky="w", pady=5)
        self.font_size_var = tk.StringVar()
        font_size_combo = ttk.Combobox(
            font_frame,
            textvariable=self.font_size_var,
            values=["8", "9", "10", "11", "12", "14", "16", "18"],
            state="readonly"
        )
        font_size_combo.grid(row=1, column=1, sticky="ew", padx=(10,0), pady=5)

        # Directory Settings
        dir_frame = ttk.LabelFrame(main_frame, text="Directory Settings", padding=15)
        dir_frame.pack(fill="x", pady=(0, 15))
        dir_frame.grid_columnconfigure(1, weight=1)

        output_dir_label = ttk.Label(dir_frame, text="Output Directory:")
        output_dir_label.grid(row=0, column=0, sticky="w", pady=5)

        dir_inner_frame = ttk.Frame(dir_frame)
        dir_inner_frame.grid(row=0, column=1, sticky="ew", padx=(10,0), pady=5)
        dir_inner_frame.grid_columnconfigure(0, weight=1)

        self.output_dir_var = tk.StringVar()
        output_dir_entry = ttk.Entry(dir_inner_frame, textvariable=self.output_dir_var)
        output_dir_entry.grid(row=0, column=0, sticky="ew")

        browse_btn = ttk.Button(
            dir_inner_frame,
            text="Browse",
            command=self.browse_output_directory
        )
        browse_btn.grid(row=0, column=1, padx=(5,0))

        # Default Bid Information
        bid_frame = ttk.LabelFrame(main_frame, text="Default Bid Information", padding=15)
        bid_frame.pack(fill="x", pady=(0, 15))
        bid_frame.grid_columnconfigure(1, weight=1)
        
        ttk.Label(bid_frame, text="Default Client:").grid(
            row=0, column=0, sticky="w", pady=5
        )
        self.default_client_var = tk.StringVar()
        client_entry = ttk.Entry(bid_frame, textvariable=self.default_client_var)
        client_entry.grid(row=0, column=1, sticky="ew", padx=(10, 0), pady=5)
        
        ttk.Label(bid_frame, text="Default Opportunity:").grid(
            row=1, column=0, sticky="w", pady=5
        )
        self.default_opportunity_var = tk.StringVar()
        opportunity_entry = ttk.Entry(bid_frame, textvariable=self.default_opportunity_var)
        opportunity_entry.grid(row=1, column=1, sticky="ew", padx=(10, 0), pady=5)

        # Template Management
        template_mgmt_frame = ttk.LabelFrame(main_frame, text="Template Management", padding=15)
        template_mgmt_frame.pack(fill="x", pady=(0, 15))
        
        ttk.Label(
            template_mgmt_frame,
            text="Add new templates by copying .docx files to the templates folder:",
            wraplength=500
        ).pack(anchor="w", pady=(0, 10))
        
        templates_path_frame = ttk.Frame(template_mgmt_frame)
        templates_path_frame.pack(fill="x")
        templates_path_frame.grid_columnconfigure(0, weight=1)
        
        self.templates_path_var = tk.StringVar()
        templates_path_entry = ttk.Entry(
            templates_path_frame,
            textvariable=self.templates_path_var,
            state="readonly"
        )
        templates_path_entry.grid(row=0, column=0, sticky="ew")

        open_templates_btn = ttk.Button(
            templates_path_frame,
            text="Open Folder",
            command=self.open_templates_folder
        )
        open_templates_btn.grid(row=0, column=1, padx=(5, 0))
        
        refresh_templates_btn = ttk.Button(
            templates_path_frame,
            text="Refresh",
            command=self.refresh_templates
        )
        refresh_templates_btn.grid(row=0, column=2, padx=(5, 0))

        # Action buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill="x", pady=(20, 0))

        ttk.Button(
            button_frame,
            text="Save Settings",
            command=self.save_settings,
            style="Accent.TButton"
        ).pack(side="left", padx=(0, 10))
        
        ttk.Button(
            button_frame,
            text="Reset to Defaults",
            command=self.reset_to_defaults
        ).pack(side="left", padx=(0, 10))
        
        ttk.Button(
            button_frame,
            text="Cancel",
            command=self.cancel_changes
        ).pack(side="left")



    def load_current_settings(self):
        """Load current settings from the controller"""
        # Load templates
        templates = self.controller.get_templates()
        self.default_template_combo['values'] = templates

        # Set current values
        self.default_template_var.set(self.controller.get_setting("default_template", ""))
        self.font_family_var.set(self.controller.get_setting("font_family", "Arial"))
        self.font_size_var.set(str(self.controller.get_setting("font_size", 11)))
        self.output_dir_var.set(self.controller.get_setting("output_directory", ""))
        self.default_client_var.set(self.controller.get_setting("client_name", ""))
        self.default_opportunity_var.set(self.controller.get_setting("opportunity_name", ""))

        # Set templates path
        templates_dir = self.controller.get_templates_directory()
        self.templates_path_var.set(str(templates_dir))

    def browse_output_directory(self):
        """Browse for output directory"""
        directory = filedialog.askdirectory(
            title="Select Output Directory",
            initialdir=self.output_dir_var.get()
        )
        if directory:
            self.output_dir_var.set(directory)

    def open_templates_folder(self):
        """Open the templates directory in the file explorer"""
        import subprocess
        import sys

        templates_dir = self.controller.get_templates_directory()

        try:
            if sys.platform.startswith("darwin"):  # macOS
                subprocess.run(['open', str(templates_dir)])
            elif sys.platform.startswith("win"):  # Windows
                subprocess.run(['explorer', str(templates_dir)])
            else:  # Linux
                subprocess.run(['xdg-open', str(templates_dir)])
        except Exception as e:
            messagebox.showerror("Error", f"Could not open templates folder: {e}")

    def refresh_templates(self):
        """Refresh the templates list"""
        self.controller.refresh_templates()
        templates = self.controller.get_templates()
        self.default_template_combo['values'] = templates
        messagebox.showinfo("Templates Refreshed", f"Found {len(templates)} templates.")

    def save_settings(self):
        """Save current settings"""
        try:
            # Validate font size
            font_size = int(self.font_size_var.get())
            if font_size < 6 or font_size > 72:
                raise ValueError("Font size must be between 6 and 72.")
            
            # Validate output directory
            output_dir = Path(self.output_dir_var.get())
            if not output_dir.exists():
                output_dir.mkdir(parents=True, exist_ok=True)

            # Prepare settings
            settings = {
                "default_template": self.default_template_var.get(),
                "font_family": self.font_family_var.get(),
                "font_size": font_size,
                "output_directory": self.output_dir_var.get(),
                "client_name": self.default_client_var.get(),
                "opportunity_name": self.default_opportunity_var.get()
            }

            # Save settings
            self.controller.save_settings(settings)
            messagebox.showinfo("Settings Saved", "Your settings have been saved successfully.")

        except ValueError as e:
            messagebox.showerror("Invalid Input", str(e))

        except Exception as e:
            messagebox.showerror("Error", f"Could not save settings: {e}")

    def reset_to_defaults(self):
        """Reset settings to default values"""
        if messagebox.askyesno("Reset Settings", "Are you sure you want to reset settings to default values?"):
            self.controller.reset_settings()
            self.load_current_settings()
            messagebox.showinfo("Settings Reset", "Settings have been reset to default values.")

    def cancel_changes(self):
        """Cancel changes and reload original settings"""
        self.load_current_settings()
