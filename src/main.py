'''
CV Converter Application Entry Point
'''
import tkinter as tk
from controllers.main_controller import MainController
from utils.config import Config

def main():
    # Ensure required directories exist
    Config.ensure_directories()

    # Create main window
    root = tk.Tk()

    # Initialize app
    app = MainController(root)

    # Start the main event loop
    root.mainloop()

if __name__ == "__main__":
    main()