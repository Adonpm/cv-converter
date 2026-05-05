# CV Converter Application

A professional Python GUI application built with Tkinter using Model-View-Controller (MVC) architecture for converting Whoz CV documents into client-specific templates.

## Features

- **Clean MVC Architecture**: Separates presentation, business logic, and data management
- **User-Friendly Interface**: Professional GUI with responsive layout
- **Settings Persistence**: Automatically saves window geometry and user preferences
- **Template Management**: Support for multiple CV templates and formats
- **Font Customisation**: User-selectable font family and size applied across all converted output
- **Batch Processing**: Convert multiple CVs in one run with progress tracking


## Project Structure

```
cv-converter/
├── src/
│   ├── main.py                          # Application entry point
│   ├── models/
│   │   ├── settings_model.py            # Settings management
│   │   ├── cv_parser.py                 # Document parsing logic
│   │   └── template_manager.py          # Template handling
│   ├── views/
│   │   ├── main_window.py               # Main GUI window
│   │   ├── conversion_view.py           # CV conversion interface
│   │   ├── settings_view.py             # Settings configuration
│   │   └── components/
│   │       └── progress_bar.py          # Progress dialog component
│   ├── controllers/
│   │   ├── main_controller.py           # Main application controller
│   │   ├── conversion_controller.py     # CV conversion logic
│   │   └── settings_controller.py       # Settings management
│   └── utils/
│       ├── config.py                    # Application configuration
│       └── file_handler.py              # File operations
├── sample_templates/                    # Sample CV templates for reference
│   ├── CV_Template_v1.docx
│   └── CV_Template_v2.docx
├── templates/                           # User's working templates (gitignored)
├── config/                              # Configuration files (gitignored)
├── output/                              # Converted CV output (gitignored)
├── build.py                             # PyInstaller build script
├── requirements.txt                     # Python dependencies
└── README.md
```


## Installation

1. **Clone the repository:**

```bash
git clone https://github.com/Adonpm/cv-converter.git
cd cv-converter
```

2. **Install dependencies:**

```bash
pip install -r requirements.txt
```

3. **Run the application:**

```bash
python src/main.py
```


## Requirements

- Python 3.7+
- tkinter (usually included with Python)
- pathlib (included with Python 3.4+)


## Usage

1. **Launch the application** by running `python src/main.py`
2. **Convert CVs**: Click "Convert CVs" to access the conversion interface
3. **Settings**: Click "Settings" to configure application preferences
4. **Window preferences** are automatically saved and restored between sessions

### Templates

Sample templates are provided in the `sample_templates/` folder. Copy them into the `templates/` folder to get started. Templates use placeholders to map CV data automatically — see the supported placeholders below.

### Supported Placeholders

| Placeholder | Description |
|---|---|
| `{{NAME}}` | Candidate full name (uppercased) |
| `{{TITLE}}` | Current job title |
| `{{SUMMARY}}` | Professional summary |
| `{{HOME}}` | Home location (nationality) |
| `{{YEARS OF EXPERIENCE}}` | Total years of experience |
| `{{EXPERIENCE SUMMARY}}` | Summary list of past employers and dates |
| `{{EDUCATION}}` | Education entries |
| `{{SKILLS}}` | Top skills extracted from experience |
| `{{MEMBERSHIPS}}` | Professional memberships |
| `{{CERTIFICATIONS}}` | Certifications |
| `{{PROJECTS}}` | Full professional experience with headers and descriptions |
| `{{CLIENT}}` | Client name (from Bid Information) |
| `{{OPPORTUNITY}}` | Opportunity name (from Bid Information) |


## Configuration

The application uses a JSON-based configuration system:

- **Settings file**: `config/settings.json`
- **Templates directory**: `templates/`
- **Output directory**: `output/`
- **Default settings** are automatically created on first run


## Building the Executable

To build a standalone Windows executable:

```bash
python build.py
```

This uses PyInstaller with `--onedir` mode and outputs to the `dist/` folder. The distributable consists of two items that must be kept together:

- `CV-Converter.exe` — the application executable
- `_internal/` — required dependencies folder (do not delete or move)

Zip both items together when sharing with end users.


## Development

### Key Components

- **MainController**: Manages application flow and view switching
- **MainWindow**: Primary GUI window with header and content areas
- **SettingsModel**: Handles user preferences and configuration persistence
- **CVParser**: Extracts structured data from Whoz-format CV `.docx` files
- **TemplateManager**: Populates templates with extracted CV data, handles floating text boxes and styled paragraph insertion
- **Config**: Centralized application configuration and constants


### Architecture Pattern

The application follows MVC architecture:

- **Models**: Data management and business logic
- **Views**: GUI components and user interaction
- **Controllers**: Application flow and coordination between models and views

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Submit a pull request

## Support

For issues and questions, please create an issue in the GitHub repository.