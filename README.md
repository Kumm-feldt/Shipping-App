
# Guatemala Marítima, S.A. - Vessel Documentation Manager

## Table of Contents
- [Introduction](#introduction)
- [Features](#features)
- [Installation](#installation)
- [Usage](#usage)
- [Dependencies](#dependencies)
- [Configuration](#configuration)
- [Troubleshooting](#troubleshooting)
- [Contributors](#contributors)
- [License](#license)

---

## Introduction
The **Guatemala Marítima, S.A. Vessel Documentation Manager** is a desktop application built using Python and Tkinter. The tool streamlines the creation and organization of maritime documentation for new vessels. It automates folder creation, document preparation, and file management based on user input.

---

## Features
- **GUI for Document Management**: User-friendly interface to manage vessel-related documents.
- **Automated Folder Creation**: Creates a main directory and subdirectories for each vessel.
- **Custom Document Generation**: Populates Word documents using a template with vessel details.
- **Integration with Excel and Word Files**: Copies and renames existing Excel and Word templates for each vessel.
- **Multi-Language Labels**: Supports Spanish and English text fields.

---

## Installation
1. Clone the repository:
   ```bash
   git clone <repository_url>
   cd <repository_name>
   ```
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Place the required fonts and image files (`gt.png`, `title.png`, `flecha.png`, etc.) in the project directory.

---

## Usage
1. Run the application:
   ```bash
   python index.py
   ```
2. Enter the vessel's name to create a new record.
3. Fill in the required form fields:
   - Agent name
   - Vessel details (IMO, flag, dimensions, etc.)
   - Arrival and departure details
4. Automatically generate folders and pre-filled documents by clicking "Save."

### Directory Structure
The application creates the following folder structure for each vessel:
- `BLs/`
- `BOL. PAGO PRECALCULO/`
- `CARTAS DE CORRECCION/`
- `CARTAS DE DESPACHO/`
- `FACTURAS EPQ/`
- `SCANS DE BLs/`
- `SOL. ACTIVIDAD PERMITIDA/`
- `JUST. SOBRANTES & FALTANTES/`
- `DOCS. CAPITAN/`
- `REC. REINTEGRO/`

### Supported File Formats
- `.xlsx`
- `.docx`

---

## Dependencies
The project uses the following Python libraries:
- `tkinter` for GUI development.
- `Pillow` for image handling.
- `pathlib` for path operations.
- `docxtpl` for populating Word document templates.
- `pyglet` for font management.
- `shutil` and `os` for file and directory operations.

Install all dependencies using:
```bash
pip install -r requirements.txt
```

---

## Configuration
- Update the paths for images, templates, and fonts as per your system.
- Place templates (`solicitud-zarpe.docx`, `cartas-de-flete.docx`, etc.) in the root directory.
- Fonts (`Black.ttf`, `Medium.ttf`, etc.) must be added to the project folder.

---

## Troubleshooting
- **Error: Missing Template Files**  
  Ensure that required `.docx` and `.xlsx` files are in the project directory.
- **Font Issues**  
  Verify that custom fonts are correctly installed and loaded.
- **Permission Issues**  
  Run the application with sufficient permissions to create folders and write files.

---

## Contributors
- **Primary Developer**: Guatemala Marítima, S.A.

---

## License
This project is licensed under the [MIT License](LICENSE).

---

```

### Notes
- The `README.md` assumes you have the necessary images, templates, and fonts in the correct paths.
- If you want me to add anything specific (like a sample output or advanced usage), let me know!