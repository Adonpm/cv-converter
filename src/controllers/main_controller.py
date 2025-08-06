"""
Main application controller
"""
import tkinter as tk
from tkinter import ttk
from models.settings_model import SettingsModel
from views.main_window import MainWindow
from views.conversion_view import ConversionView
from controllers.conversion_controller import ConversionController

class MainController:
    def __init__(self, root):
        self.root = root
        self.settings_model = SettingsModel()
        self.current_view = None

        # Initialize controllers
        self.conversion_controller = ConversionController(self.settings_model, self)

        # Initialize main window
        self.main_window = MainWindow(root, self)

        # Load window geometry
        geometry = self.settings_model.get("window_geometry", "800x600")
        self.root.geometry(geometry)

        # Bind window close event
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def show_conversion_view(self):
        '''Show CV conversion view'''
        self.main_window.clear_content()

        # Placeholder for conversion view
        content_frame = self.main_window.get_content_frame()
        self.current_view = ConversionView(content_frame, self.conversion_controller)

    def show_settings_view(self):
        '''Show settings view'''
        self.main_window.clear_content()

        # Placeholder for settings view
        content_frame = self.main_window.get_content_frame()
        placeholder = ttk.Label(
            content_frame,
            text="Settings View\n(To be implemented)",
            font=("Arial", 14),
            justify="center"
        )
        placeholder.grid(row=0, column=0)

    def on_closing(self):
        '''Handle application closing'''
        # Save window geometry
        self.settings_model.set("window_geometry", self.root.geometry())
        self.settings_model.save_settings()
        self.root.destroy()

    # Pass through methods for controllers
    def get_templates(self):
        return self.conversion_controller.get_templates()
    
    def convert_cvs(self, input_files, settings):
        return self.conversion_controller.convert_cvs(input_files, settings)
