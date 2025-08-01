"""
Settings model for managing user preferences
"""
import json
from pathlib import Path
from utils.config import Config

class SettingsModel:
    def __init__(self):
        self.settings = self._load_default_settings()
        self.load_settings()

    def _load_default_settings(self):
        """Load default settings from the configuration"""
        return {
            "font_family": Config.DEFAULT_FONT_FAMILY,
            "font_size": Config.DEFAULT_FONT_SIZE,
            "default_template": "",
            "output_directory": str(Config.DEFAULT_OUTPUT_DIR),
            "client_name": "",
            "opportunity_name": "",
            "last_input_directory": str(Path.home()),
            "window_geometry": "800x600"
        }
    
    def load_settings(self):
        """Load settings from file"""
        settings_file = Config.get_settings_file()
        if settings_file.exists():
            try:
                with open(settings_file, 'r') as f:
                    saved_settings = json.load(f)
                    self.settings.update(saved_settings)
            except Exception as e:
                print(f"Error loading settings: {e}")

    def save_settings(self):
        """Save current settings to file"""
        Config.ensure_directories()
        try:
            with open(Config.get_settings_file(), 'w') as f:
                json.dump(self.settings, f, indent=2)
        except Exception as e:
            print(f"Error saving settings: {e}")

    def get(self, key, default=None):
        """Get a setting value"""
        return self.settings.get(key, default)
    
    def set(self, key, value):
        '''Set a setting value'''
        self.settings[key] = value

    def update(self, new_settings):
        '''Update multiple settings'''
        self.settings.update(new_settings)