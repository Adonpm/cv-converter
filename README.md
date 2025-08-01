# CV Converter Application

A professional Python GUI application built with Tkinter using Model-View-Controller (MVC) architecture for converting CV documents between different templates.

## Features

- **Clean MVC Architecture**: Separates presentation, business logic, and data management
- **User-Friendly Interface**: Professional GUI with responsive layout
- **Settings Persistence**: Automatically saves window geometry and user preferences
- **Template Management**: Support for multiple CV templates and formats
- **Cross-Platform**: Works on Windows, macOS, and Linux


## Project Structure

```
cv-converter/
├── src/
│   ├── main.py                     # Application entry point
│   ├── models/
│   │   ├── settings_model.py       # Settings management
│   │   ├── cv_parser.py           # Document parsing logic
│   │   └── template_manager.py     # Template handling
│   ├── views/
│   │   ├── main_window.py         # Main GUI window
│   │   ├── conversion_view.py     # CV conversion interface
│   │   └── settings_view.py       # Settings configuration
│   ├── controllers/
│   │   ├── main_controller.py     # Main application controller
│   │   ├── conversion_controller.py # CV conversion logic
│   │   └── settings_controller.py # Settings management
│   └── utils/
│       ├── config.py              # Application configuration
│       └── file_handler.py        # File operations
├── templates/                     # CV template files
├── config/                        # Configuration files
├── requirements.txt              # Python dependencies
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

## Configuration

The application uses a JSON-based configuration system:

- **Settings file**: `config/settings.json`
- **Templates directory**: `templates/`
- **Default settings** are automatically created on first run


## Development

### Key Components

- **MainController**: Manages application flow and view switching
- **MainWindow**: Primary GUI window with header and content areas
- **SettingsModel**: Handles user preferences and configuration persistence
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
