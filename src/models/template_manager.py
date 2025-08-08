"""
Template manager for handling CV templates
"""
from docx import Document
from pathlib import Path
from utils.config import Config
import shutil
from docx.shared import Pt

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
    
    def apply_template(self, template_name, cv_data, output_path, formatting_options=None):
        """Apply a template to the provided CV data"""
        template_path = self.get_template_path(template_name)
        if not template_path:
            raise ValueError(f"Template '{template_name}' not found.")
        
        try:
            # Load template
            doc = Document(template_path)

            # Apply CV data to the template
            self._populate_template(doc, cv_data, formatting_options or {})

            # Save output
            doc.save(output_path)
            return True
        
        except Exception as e:
            raise Exception(f"Error applying template: {str(e)}")

    def _populate_template(self, doc, cv_data, formatting_options):
        """Populate template with CV data"""
        # Replace placeholders in paragraphs
        for paragraph in doc.paragraphs:
            self._replace_placeholders(paragraph, cv_data, formatting_options)
        
        # Replace placeholders in tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        self._replace_placeholders(paragraph, cv_data, formatting_options)

        # Update headers and footers
        for section in doc.sections:
            self._update_header_footer(section, cv_data, formatting_options)

    def _replace_placeholders(self, paragraph, cv_data, formatting_options):
        """Replace placeholders in paragraph text"""
        text = paragraph.text

        # Define placeholder mappings
        placeholders = {
            "{{NAME}}": cv_data.get("name", "") or "",
            "{{TITLE}}": cv_data.get("current_position", "") or "",
            "{{SUMMARY}}": cv_data.get("summary", "") or "",
            "{{YEARS OF EXPERIENCE}}": cv_data.get("years_experience", "") or "", 
            "{{EDUCATION}}": self._format_education(cv_data.get("education", [])),
            ########################################################
            "{{SKILLS}}": self._format_skills(cv_data.get("skills", [])), 
            ########################################################
            "{{MEMBERSHIPS}}": self._format_memberships(cv_data.get("professional_memberships", [])),
            "{{CERTIFICATIONS}}": self._format_certifications(cv_data.get("certifications", [])),
            "{{PROJECTS}}": self._format_professional_experience(cv_data.get("professional_experience", []))
        }

        # Replace placeholders
        for placeholder, value in placeholders.items():
            if placeholder in text:
                # Ensure value is a string, not None
                safe_value = str(value) if value is not None else ""

                # Clear existing text
                paragraph.clear()
                # Add new text
                run = paragraph.add_run(text.replace(placeholder, safe_value))

                ############################################################
                # Apply font formatting to all placeholders
                '''
                if formatting_options.get("font_family"):
                    run.font.name = formatting_options["font_family"]
                if formatting_options.get("font_size"):
                    run.font.size = Pt(formatting_options["font_size"])
                '''
                ############################################################

    def _format_education(self, education_list):
        """Format education entries"""
        if not education_list:
            return ""
        return "\n".join(f"• {edu}" for edu in education_list if edu)

    ########################################################
    def _format_skills(self, skills_list):
        """Format skills entries"""
        pass
    ########################################################

    def _format_certifications(self, certifications_list):
        """Format certifications entries"""
        if not certifications_list:
            return ""
        
        formatted = []
        for cert in certifications_list:
            if isinstance(cert, dict):
                cert_text = []
                if cert.get("programme"):
                    cert_text.append(cert["programme"])
                if cert.get("entity"):
                    cert_text.append(f"({cert['entity']})")
                if cert.get("date"):
                    cert_text.append(f"- {cert['date']}")
                formatted.append(" ".join(cert_text))
            else:
                formatted.append(str(cert))
        
        return "\n".join([f"{cert}" for cert in formatted])
    
    def _format_memberships(self, memberships_list):
        """Format professional membership entries"""
        if not memberships_list:
            return ""
        
        formatted = []
        for membership in memberships_list:
            if isinstance(membership, dict):
                membership_text = []
                if membership.get("programme"):
                    membership_text.append(membership["programme"])
                if membership.get("date"):
                    membership_text.append(f"- {membership['date']}")
                formatted.append(" ".join(membership_text))
            else:
                formatted.append(str(membership))
        
        return "\n".join([f"{membership}" for membership in formatted])

    def _format_professional_experience(self, experience_list):
        """Format professional experience entries"""
        if not experience_list:
            return ""
        
        formatted = []
        for exp in experience_list:
            exp_text = []
            
            # Add role and company
            if exp.get("experience_header") and exp.get("client"):
                exp_text.append(f"{exp['experience_header']} - {exp['client']}")
            elif exp.get("experience_header"):
                exp_text.append(exp['experience_header'])
            elif exp.get("client"):
                exp_text.append(exp['client'])
            
            # Add duration
            if exp.get("duration_period"):
                exp_text.append(f"({exp['duration_period']})")
            
            # Add location
            #if exp.get("location"):
                #exp_text.append(f"Location: {exp['location']}")
            
            # Add description
            if exp.get("expert_mission_description"):
                description = exp['expert_mission_description']
                exp_text.append(f"\n{description}")
            
            # Add technical expertise
            #if exp.get("technical_expertise"):
                #exp_text.append(f"\nTechnical expertise: {exp['technical_expertise']}")
            
            formatted.append("\n".join(exp_text))
        
        return "\n\n".join([f"• {exp}" for exp in formatted]) 

    def _update_header_footer(self, section, cv_data, formatting_options):
        """Update header and footer with optional client/opportunity info from tkinter interface"""
        client = formatting_options.get("client_name", "")
        opportunity = formatting_options.get("opportunity_name", "")
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
