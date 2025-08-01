'''
Application configuration and constants
'''
import os
import json
from pathlib import Path

class Config:
    # Application settings
    APP_NAME = "CV Converter"
    APP_VERSION = "1.0.0"

    # Paths
    BASE_DIR = Path(__file__).parent.parent.parent
    TEMPLATES_DIR = BASE_DIR / "templates"
    CONFIG_DIR = BASE_DIR / "config"
    ASSETS_DIR = BASE_DIR / "assets"

    # Default settings
    DEFAULT_FONT_FAMILY = "Arial"
    DEFAULT_FONT_SIZE = 11
    DEFAULT_OUTPUT_DIR = BASE_DIR / "output"

    # File types
    SUPPORTED_INPUT_FORMATS = [".docx"]
    TEMPLATE_FORMATS = [".docx"]

    @classmethod
    def ensure_directories(cls):
        """Create necessary directories if they don't exist"""
        for directory in [cls.TEMPLATES_DIR, cls.CONFIG_DIR, cls.DEFAULT_OUTPUT_DIR]:
            directory.mkdir(exist_ok=True)

    @classmethod
    def get_settings_file(cls):
        """Get path to settings file"""
        return cls.CONFIG_DIR / "settings.json"
    