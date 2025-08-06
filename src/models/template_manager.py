"""
Template manager for handling CV templates
"""
from docx import Document
from pathlib import Path
from utils.config import Config
import shutil

class TemplateManager:
    def __init__(self):
        self.templates = {}
        self.load_templates()

    def load_templates(self):
        """Load available templates"""
        self.templates = {}
        templates_dir = Config.TEMPLATES_DIR

        if templates_dir.exists():
            for template_file in templates_dir.glob("*.docx"):
                template_name = template_file.stem
                self.templates[template_name] = str(template_file)

    def get_template_list(self):
        """Get list of available templates"""
        return list(self.templates.keys())
    
    def get_template_path(self, template_name):
        """Get the file path of a specific template"""
        return self.templates.get(template_name)
    
    def apply_template(self, template_name, cv_data, output_path, formatting_options=None, client_name=None, opportunity_name=None):
        """Apply a template to the provided CV data"""
        template_path = self.get_template_path(template_name)
        if not template_path:
            raise ValueError(f"Template '{template_name}' not found.")
        
        try:
            # Load template
            doc = Document(template_path)

            # Apply CV data to the template
            self._populate_template(doc, cv_data, formatting_options or {}, client_name, opportunity_name)

            # Save output
            doc.save(output_path)
            return True
        
        except Exception as e:
            raise Exception(f"Error applying template: {str(e)}")

    def _populate_template(self, doc, cv_data, formatting_options, client_name=None, opportunity_name=None):
        """Populate template with CV data"""
        # Replace placeholders in paragraphs
        for paragraph in doc.paragraphs:
            self._replace_placeholders(paragraph, cv_data)
        
        # Replace placeholders in tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        self._replace_placeholders(paragraph, cv_data)

        # Update headers and footers
        for section in doc.sections:
            self._update_header_footer(section, cv_data, formatting_options, client_name, opportunity_name)

    def _replace_placeholders(self, paragraph, cv_data):
        """Replace placeholders in paragraph text"""
        text = paragraph.text

        # Define placeholder mappings
        placeholders = {
            "{{NAME}}": cv_data.get("name", ""),
            "{{TITLE}}": cv_data.get("current_position", ""),
            "{{SUMMARY}}": cv_data.get("summary", ""),
            "{{EXPERIENCE}}": cv_data.get("experience", ""), # Need to fix
            "{{EDUCATION}}": "\n".join(cv_data.get("education", [])),
            "{{SKILLS}}": "\n".join(cv_data.get("skills", [])), # Need to fix
            "{{CERTIFICATIONS}}": "\n".join(cv_data.get("certifications", [])),
            "{{PROJECTS}}": "\n".join(cv_data.get("professional_experience", []))
        }

        # Replace placeholders
        for placeholder, value in placeholders.items():
            if placeholder in text:
                # Clear existing text
                paragraph.clear()
                # Add new text
                run = paragraph.add_run(text.replace(placeholder, value))

    def _update_header_footer(self, section, cv_data, formatting_options, client_name=None, opportunity_name=None):
        """Update header and footer with optional client/opportunity info from tkinter interface"""
        client = client_name if client_name is not None else ""
        opportunity = opportunity_name if opportunity_name is not None else ""
        # Update footer
        if section.footer:
            for paragraph in section.footer.paragraphs:
                text = paragraph.text
                if "{{CLIENT}}" in text:
                    paragraph.clear()
                    paragraph.add_run(text.replace("{{CLIENT}}", client))
                if "{{OPPORTUNITY}}" in text:
                    paragraph.clear()
                    paragraph.add_run(text.replace("{{OPPORTUNITY}}", opportunity))
