"""
Conversion controller for managing CV conversion process
"""
from models.cv_parser import CVParser
from models.template_manager import TemplateManager
from utils.file_handler import FileHandler
from pathlib import Path
import threading
from tkinter import messagebox

class ConversionController:
    def __init__(self, settings_model):
        self.settings_model = settings_model
        self.cv_parser = CVParser()
        self.template_manager = TemplateManager()
        self.file_handler = FileHandler()

    def get_templates(self):
        """Get list of available templates"""
        return self.template_manager.get_template()
    
    def convert_cvs(self, input_files, settings):
        """Convert multiple CVs"""
        def conversion_thread():
            try:
                # Create output directoty
                output_dir = self.file_handler.create_output_directory(self.settings_model.get("output_directory"))
                
                successful_conversions = 0
                failed_conversions = []

                for input_file in input_files:
                    try:
                        # Validate input file
                        is_valid, message = self.file_handler.validate_input_file(input_file)
                        if not is_valid:
                            failed_conversions.append(f"{Path(input_file).name}: {message}")
                            continue

                        # Parse CV
                        cv_data = self.cv_parser.parse_cv(input_file)

                        # Generate output filename
                        output_file = self.file_handler.generate_output_filename(input_file, settings['template'], output_dir)

                        # Apply template
                        formatting_options = {
                            "font_family": settings['font_family'],
                            "font_size": settings["font_size"],
                            "client_name": settings["client_name"],
                            "opportunity_name": settings["opportunity_name"]
                        }

                        self.template_manager.apply_template(
                            settings['template'],
                            cv_data,
                            output_file,
                            formatting_options
                            ###################################
                            # Add logic to handle client and opportunity names from tkinter
                            ####################################
                        )

                        successful_conversions += 1

                    except Exception as e:
                        failed_conversions.append(f"{Path(input_file).name}: {str(e)}")

                # Show results
                self._show_conversion_results(successful_conversions, failed_conversions, output_dir)

            except Exception as e:
                messagebox.showerror("Conversion Error", f"An error occurred during conversion: {str(e)}")

        # Run conversion in a separate thread to avoid blocking the UI
        thread = threading.Thread(target=conversion_thread)
        thread.daemon = True
        thread.start()
                
    def _show_conversion_results(self, successful, failed, output_dir):
            """Show conversion results"""
            message = f"Conversion completed!\n\n"
            message += f"Successfully converted: {successful} files\n"

            if failed:
                message += f"Failed conversions: {len(failed)} files\n\n"
                message += "Failed files:\n"
                for failure in failed[:5]:  # Show only first 5 failures
                    message += f"- {failure}\n"
                if len(failed) > 5:
                    message += f"... and {len(failed) - 5} more files failed.\n"

            message += f"\nOutput directory: {output_dir}"

            if failed:
                messagebox.showwarning("Conversion Results", message)
            else:
                messagebox.showinfo("Conversion Complete", message)
