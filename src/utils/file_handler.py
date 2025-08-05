"""
File handling utilities
"""
from pathlib import Path
import shutil
from datetime import datetime

class FileHandler:
    @staticmethod
    def create_output_directory(base_dir, batch_name=None):
        """Create output directory for converted files"""
        if not batch_name:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            batch_name = f"cv_conversion_{timestamp}"

        output_dir = Path(base_dir) / batch_name
        output_dir.mkdir(parents=True, exist_ok=True)
        return output_dir
    
    @staticmethod
    def generate_output_filename(input_file, template_name, output_dir):
        """Generate output filename for converted CV"""
        input_path = Path(input_file)
        base_name = input_path.stem
        
        output_filename = f"{base_name}_{template_name}_converted.docx"
        return output_dir / output_filename
    
    @staticmethod
    def validate_input_file(file_path):
        """Validate input file"""
        path = Path(file_path)

        if not path.exists():
            return False, "File does not exist"
        
        if path.suffix.lower() != ".docx":
            return False, "Only .docx files are supported"
        
        if path.stat().st_size == 0:
            return False, "File is empty"
        
        return True, "Valid"
    
    @staticmethod
    def backup_file(file_path, backup_dir):
        """Create backup of original file"""
        source = Path(file_path)
        backup_path = backup_dir / source.name
        shutil.copy2(source, backup_path)
        return backup_path
