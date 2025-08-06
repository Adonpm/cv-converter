"""
Settings controller for managing application settings
"""
from models.template_manager import TemplateManager
from utils.config import Config

class SettingsController:
    def __init__(self, settings_model):
        self.settings_model = settings_model
        self.template_manager = TemplateManager()

    def get_templates(self):
        """Get list of available templates"""
        return self.template_manager.get_template_list()
    
    def get_templates_directory(self):
        """Get templates directory path"""
        return Config.TEMPLATES_DIR
    
    def refresh_templates(self):
        """Refresh template list"""
        self.template_manager.load_templates()
    
    def get_setting(self, key, default=None):
        """Get a setting value"""
        return self.settings_model.get(key, default)
    
    def save_settings(self, new_settings):
        """Save settings"""
        self.settings_model.update(new_settings)
        self.settings_model.save_settings()
    
    def reset_settings(self):
        """Reset settings to defaults"""
        self.settings_model.settings = self.settings_model._load_default_settings()
        self.settings_model.save_settings()
    