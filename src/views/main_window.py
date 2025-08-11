"""
Main application window
"""
import tkinter as tk
from tkinter import ttk

class MainWindow:
    def __init__(self, root, controller):
        self.root = root
        self.controller = controller
        self.setup_window()
        self.create_widgets()
        
    def setup_window(self):
        '''Configure main window'''
        self.root.title("CV Converter")
        self.root.geometry("800x600")
        self.root.minsize(600, 400)

        # Configure grid weights
        self.root.grid_rowconfigure(1, weight=1)
        self.root.grid_columnconfigure(0, weight=1)

    def create_widgets(self):
        '''Create main window widgets'''
        # Header frame
        self.header_frame = ttk.Frame(self.root)
        self.header_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=5)

        # Title
        title_label = ttk.Label(
            self.header_frame,
            text="CV Converter",
            font=("Arial", 16, "bold")
        )
        title_label.grid(row=0, column=0, sticky="w")

        # Navigation buttons
        nav_frame = ttk.Frame(self.header_frame)
        nav_frame.grid(row=0, column=1, sticky="e")

        self.convert_btn = ttk.Button(
            nav_frame,
            text="Convert CVs",
            command=self.show_conversion_view
        )
        self.convert_btn.grid(row=0, column=0, padx=5)

        self.settings_btn = ttk.Button(
            nav_frame,
            text="Settings",
            command=self.show_settings_view
        )
        self.settings_btn.grid(row=0, column=1, padx=5)

        # Configure header grid
        self.header_frame.grid_columnconfigure(1, weight=1)

        # Main content frame
        self.content_frame = ttk.Frame(self.root)
        self.content_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=5)
        self.content_frame.grid_rowconfigure(0, weight=1)
        self.content_frame.grid_columnconfigure(0, weight=1)

        # Welcome message
        self.show_welcome()

    def show_welcome(self):
        '''Show welcome message'''
        self.clear_content()

        welcome_frame = ttk.Frame(self.content_frame)
        welcome_frame.grid(row=0, column=0, sticky="nsew")
        welcome_frame.grid_rowconfigure(0, weight=1)
        welcome_frame.grid_columnconfigure(0, weight=1)

        # Main container for all content
        content_frame = ttk.Frame(welcome_frame)
        content_frame.grid(row=0, column=0)

        # App icon/logo (using Unicode)
        logo_label = ttk.Label(
            content_frame,
            text="📄",  # Document conversion icon
            font=("Arial", 85)
        )
        logo_label.grid(row=0, column=0, pady=(0, 20))

        # Welcome title
        title_label = ttk.Label(
            content_frame,
            text="Welcome to CV Converter",
            font=("Arial", 18, "bold")
        )
        title_label.grid(row=1, column=0, pady=(0, 10))

        # Description with icons
        desc_label = ttk.Label(
            content_frame,
            text="🔄 Click 'Convert CVs' to start converting your CVs\n\n"
                "⚙️ Click 'Settings' to configure your preferences",
            font=("Arial", 12),
            justify="center"
        )
        desc_label.grid(row=2, column=0, pady=(0, 20))

        # Feature highlights
        features_frame = ttk.Frame(content_frame)
        features_frame.grid(row=3, column=0, pady=10)

        ttk.Label(features_frame, text="✨ Multiple template support", 
                font=("Arial", 10)).grid(row=0, column=0, padx=10)
        ttk.Label(features_frame, text="🎨 Customizable formatting", 
                font=("Arial", 10)).grid(row=0, column=1, padx=10)
        ttk.Label(features_frame, text="📁 Batch processing", 
                font=("Arial", 10)).grid(row=0, column=2, padx=10)

    def clear_content(self):
        '''Clear the content frame'''
        for widget in self.content_frame.winfo_children():
            widget.destroy()

    def show_conversion_view(self):
        '''Show conversion view'''
        self.controller.show_conversion_view()

    def show_settings_view(self):
        '''Show settings view'''
        self.controller.show_settings_view()    

    def get_content_frame(self):
        '''Get the main content frame for other views'''
        return self.content_frame
