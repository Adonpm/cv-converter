"""
CV Parser for extracting content from Whoz template CVs
"""
from docx import Document
import re
from pathlib import Path

class CVParser:
    def __init__(self):
        self.extracted_data = {}

    def parse_cv(self, file_path):
        """Parse a Whoz CV and extract structured data"""
        try:
            doc = Document(file_path)
            self.extracted_data = {
                "name": "",
                "current_position": "",
                "years_experience": "",
                "skills": [], # Needs to address
                "nationality": "",
                "date_of_birth": "",
                "education": [],
                "certifications": [],
                "professional_memberships": [],
                "software": [],
                "key_capabilities": "",
                "languages": [],
                "professional_experience": [],
                "contact": {}, # Needs to address
                "summary": "",
                "publications_or_awards": []
            }

            self._extract_content_from_tables(doc)
            return self.extracted_data

        except Exception as e:
            raise Exception(f"Error parsing CV: {str(e)}")

    def _extract_content_from_tables(self, doc):
        """Extract content from document tables"""
        for table_num, table in enumerate(doc.tables):
            self._process_table(table, table_num)

    def _process_table(self, table, table_num):
        """Process individual table based on its content"""
        for row_num, row in enumerate(table.rows):
            cells = []
            for cell in row.cells:
                text = cell.text.strip()
                if text:
                    cells.append(text)
                
            if not cells:
                continue

            # Process 1st table - Basic Information
            if table_num == 0:
                self._extract_basic_info(cells, row_num)

            # Process other tables based on section headers
            elif cells:
                potential_header = cells[0].upper()

                # Check if this is actually a header (no dates/numbers)
                if any(char.isdigit() for char in potential_header):
                    section_header = "PROFESSIONAL EXPERIENCE"
                else:
                    section_header = potential_header

                if "EDUCATION" in section_header:
                    self._extract_education(cells)
                elif "CERTIFICATION" in section_header:
                    self._extract_certifications(cells)
                elif "PROFESSIONAL MEMBERSHIPS" in section_header or "ACCREDITATIONS" in section_header:
                    self._extract_memberships(cells)
                elif "SOFTWARE" in section_header:
                    self._extract_software(cells)
                elif "KEY CAPABILITIES" in section_header:
                    self._extract_key_capabilities(cells)
                elif "LANGUAGES" in section_header:
                    self._extract_languages(cells)
                elif "PROFESSIONAL EXPERIENCE" in section_header:
                    self._extract_professional_experience(cells, row_num)
                elif "PUBLICATIONS" in section_header or "PROFESSIONAL AWARDS" in section_header:
                    self._extract_publications_professional_awards(cells)

    def _extract_basic_info(self, cells, row_num):
        """Extract basic information from first table"""
        if row_num == 0 and len(cells) >= 1:
            # Extract name from "LAST NAME First name: xxx yy"
            name_text = cells[0]
            if ":" in name_text:
                self.extracted_data["name"] = name_text.split(":")[-1].strip()

            # Extract nationality if present
            if len(cells) >= 3 and "Nationality" in cells[2]:
                nationality_text = cells[2]
                if ":" in nationality_text:
                    nationality = nationality_text.split(":")[-1].strip()
                    if nationality != "(To be completed)":
                        self.extracted_data["nationality"] = nationality

        elif row_num == 1 and len(cells) >= 1:
            # Extract current postion
            position_text = cells[0]
            if "Current position:" in position_text:
                self.extracted_data["current_position"] = position_text.split(":")[-1].strip()

            # Extract date of birth if present
            if len(cells) >= 3 and "Date of birth" in cells[2]:
                dob_text = cells[2]
                if ":" in dob_text:
                    dob = dob_text.split(":")[-1].strip()
                    if not dob.startswith("(**/**/****)"):
                        self.extracted_data["date_of_birth"] = dob
            
        elif row_num == 2 and len(cells) >= 1:
            # Extract years of experience
            experience_text = cells[0]
            if "Years of experience:" in experience_text:
                self.extracted_data["years_experience"] =  experience_text.split(":")[-1].strip()
    
    ###################################################################
    ################ FIX THIS LATER ###################################
    ###################################################################
    def _extract_education(self, cells): 
        """Extract education information"""
        if len(cells) >= 2:
            education_text = cells[1].strip()
            if education_text and education_text != "(To be completed)":
                # Split by year patterns or newlines
                education_items = re.split(r'\d{4}', education_text)
                for item in education_items:
                    item = item.strip()
                    self.extracted_data["education"].append(item)    

    def _extract_certifications(self, cells):
        """Extract certification information"""
        if len(cells) >= 2:
            cert_text = cells[1].strip()
            if cert_text and cert_text != "(To be completed)":
                # Parse certification details
                cert_info = {}
                lines = cert_text.split("\n")
                for line in lines:
                    line = line.strip()
                    if "Obtention date:" in line:
                        cert_info["date"] = line.replace("Obtention date:", "").strip()
                    elif "Delivering entity:" in line:
                        cert_info["entity"] = line.replace("Delivering entity:", "").strip()
                    elif "Programme:" in line:
                        cert_info["programme"] = line.replace("Programme:", "").strip()
                    
                if cert_info:
                    self.extracted_data["certifications"].append(cert_info)
                
    def _extract_memberships(self, cells):
        """Extract professional memberships"""
        if len(cells) >= 2:
            membership_text = cells[1].strip()
            if membership_text and membership_text != "(To be completed)":
                # Parse membership details
                membership_info = {}
                lines = membership_text.split('\n')
                for line in lines:
                    line = line.strip()
                    if "Obtention date:" in line:
                        membership_info["date"] = line.replace("Obtention date:", "").strip()
                    elif "Programme:" in line:
                        membership_info["programme"] = line.replace("Programme:", "").strip()
                
                if membership_info:
                    self.extracted_data["professional_memberships"].append(membership_info)

    def _extract_software(self, cells):
        """Extract software information"""
        if len(cells) >= 2:
            software_text = cells[1].strip()
            if software_text and software_text != "(To be completed)":
                self.extracted_data["software"].append(software_text)

    def _extract_key_capabilities(self, cells):
        """Extract key capabilities"""
        if len(cells) >= 2:
            capabilities_text = cells[1].strip()
            if capabilities_text and capabilities_text != "(To be completed)":
                self.extracted_data["key_capabilities"] = capabilities_text
                
                # Also extract summary from presentation section
                if "Presentation:" in capabilities_text:
                    presentation_part = capabilities_text.split("Presentation:")[-1].strip()
                    self.extracted_data["summary"] = presentation_part

    def _extract_languages(self, cells):
        """Extract language information"""
        if len(cells) >= 2:
            language_text = cells[1].strip()
            if language_text:
                # Parse language entries (e.g., "English: Native – C2")
                languages = language_text.split(',')
                for lang in languages:
                    lang = lang.strip()
                    if ':' in lang:
                        lang_name = lang.split(':')[0].strip()
                        lang_level = lang.split(':')[1].strip()
                        self.extracted_data["languages"].append({
                            "language": lang_name,
                            "level": lang_level
                        })

    def _extract_professional_experience(self, cells, row_num):
        """Extract professional experience"""
        # Skip header row
        if row_num == 0:
            return

        if len(cells) >= 2:
            # First cell contains duration and dates
            duration_text = cells[0].strip()
            # Second cell contains job details
            job_text = cells[1].strip()

            if duration_text and job_text and job_text != "(To be completed)":
                experience_entry = {
                    "duration": duration_text,
                    "details": job_text
                }
                # Parse duration for start/end dates
                if "From" in duration_text and "Duration" in duration_text:
                    parts = re.split(r'From|Duration:', duration_text)
                    if len(parts) >= 2:
                        experience_entry["start_date"] = parts[1].strip()
                    if len(parts) >= 3:
                        experience_entry["duration_period"] = parts[2].strip()
                
                # Parse job details for company, role, location, etc.
                self._parse_job_details(experience_entry, job_text)
                
                self.extracted_data["professional_experience"].append(experience_entry)

    def _parse_job_details(self, experience_entry, job_text):
        """Parse detailed job information"""
        if "Project title" in job_text:
            if "Location" in job_text and "Client name" in job_text and "Detailed project description" in job_text and "Detailed description of the expert mission (on this project)" in job_text and "Technical expertise" in job_text:
                raw_parts = re.split(r'Project title:|Location:|Client name:|Detailed project description:|Detailed description of the expert mission \(on this project\):|Technical expertise:', job_text)
                parts = [part for part in raw_parts if part is not None]
                experience_entry["experience_header"] = parts[0].strip()
                experience_entry["project_title"] = parts[1].strip()
                experience_entry["location"] = parts[2].strip()
                experience_entry["client"] = parts[3].strip()
                experience_entry["project_description"] = parts[4].strip()
                experience_entry["expert_mission_description"] = parts[5].strip()
                experience_entry["technical_expertise"] = parts[6].strip()
        else:
            if "Location" in job_text and "Description" in job_text:
                parts = re.split(r'Location:|Description', job_text)
                lines = parts[0].split('\n')
                if len(lines) >= 2:
                    experience_entry["employee_company"] = lines[0].strip()
                    experience_entry["employee_position"] = lines[1].strip()
                    experience_entry["employee_location"] = parts[1].strip()
                    experience_entry["employee_description"] = parts[2].strip()
                else:
                    experience_entry["employee_company"] = parts[0].strip()
                    experience_entry["employee_location"] = parts[1].strip()
                    experience_entry["employee_description"] = parts[2].strip()

    def _extract_publications_professional_awards(self, cells):
        """Extract Publications or Awards"""
        if len(cells) >= 2:
            publications_professional_awards_text = cells[1].strip()
            if publications_professional_awards_text and publications_professional_awards_text != "(To be completed)":
                self.extracted_data["publications_or_awards"].append(publications_professional_awards_text)
