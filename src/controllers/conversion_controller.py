"""
Conversion controller for managing CV conversion process with progress tracking
"""
from models.cv_parser import CVParser
from models.template_manager import TemplateManager
from utils.file_handler import FileHandler
from pathlib import Path
import threading, time
from tkinter import messagebox
from views.components.progress_bar import ProgressDialog

class ConversionController:
    def __init__(self, settings_model, main_window):
        self.settings_model = settings_model
        self.main_window = main_window
        self.cv_parser = CVParser()
        self.template_manager = TemplateManager()
        self.file_handler = FileHandler()
        self.progress_dialog = None

    def get_templates(self):
        """Get list of available templates"""
        return self.template_manager.get_template_list()
    
    def convert_cvs(self, input_files, settings):
        """Convert multiple CVs"""
        if not input_files:
            return
        
        # Show progress dialog
        self.progress_dialog = ProgressDialog(
            self.main_window.root,
            "Converting CVs.."
        )

        def conversion_thread():
            try:
                # Create output directoty
                output_dir = self.file_handler.create_output_directory(self.settings_model.get("output_directory"))
                
                successful_conversions = 0
                failed_conversions = []
                total_files = len(input_files)

                # Create backup directory
                backup_dir = output_dir / "backups"
                backup_dir.mkdir(exist_ok=True)

                for i, input_file in enumerate(input_files):
                    if self.progress_dialog.cancelled:
                        break

                    try:
                        # Update progress
                        filename = Path(input_file).name
                        status = f"Processing {filename}"
                        self.progress_dialog.update_progress(i, total_files, status)

                        # Validate input file
                        is_valid, message = self.file_handler.validate_input_file(input_file)
                        if not is_valid:
                            failed_conversions.append(f"{Path(input_file).name}: {message}")
                            continue

                        # Create backup
                        self.file_handler.backup_file(input_file, backup_dir)

                        # Parse CV
                        self.progress_dialog.update_progress(i, total_files, f"Parsing {filename}...")
                        cv_data = self.cv_parser.parse_cv(input_file)

                        # Generate output filename
                        output_file = self.file_handler.generate_output_filename(input_file, settings['template'], output_dir)

                        # Apply template
                        self.progress_dialog.update_progress(i, total_files, f"Applying template to {filename}...")
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

                        # Pause to show progress
                        time.sleep(0.1)

                    except Exception as e:
                        failed_conversions.append(f"{Path(input_file).name}: {str(e)}")
                
                # Final progress update
                if not self.progress_dialog.cancelled:
                    self.progress_dialog.update_progress(total_files, total_files, "Conversion complete!")
                    time.sleep(0.5)
                
                # Close progress dialog
                self.progress_dialog.close()

                # Show results (only if not cancelled)
                if not self.progress_dialog.cancelled:
                    self._show_conversion_results(successful_conversions, failed_conversions, output_dir)

            except Exception as e:
                messagebox.showerror("Conversion Error", f"An error occurred during conversion: {str(e)}")

        # Run conversion in a separate thread to avoid blocking the UI
        thread = threading.Thread(target=conversion_thread)
        thread.daemon = True
        thread.start()
                
    def _show_conversion_results(self, successful, failed, output_dir):
            """Show conversion results"""
            title = "Conversion Results"

            message = f"Batch Conversion completed!\n\n"
            message += f"✓ Successfully converted: {successful} files\n"

            if failed:
                message += f"✗ Failed conversions: {len(failed)} files\n\n"
                message += "Failed files:\n"
                for failure in failed[:5]:  # Show only first 5 failures
                    message += f"- {failure}\n"
                if len(failed) > 5:
                    message += f"... and {len(failed) - 5} more files failed.\n"

            message += f"\n📁 Output directory: {output_dir}\n"
            message += f"📁 Backups created in: {output_dir}/backups"

            if failed and successful == 0:
                messagebox.showerror(title, message)
            elif failed:
                messagebox.showwarning(title, message)
            else:
                messagebox.showinfo(title, message)

            # Update settings with last used values
            self.settings_model.set("last_input_directory", str(Path(output_dir).parent))
            self.settings_model.save_settings()
