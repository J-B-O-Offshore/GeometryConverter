# GeometryConverter

The **GeometryConverter** is an Excel- and Python-based application for creating, saving, loading, and assembling **MP, TP, and Tower geometries** from a centrally stored SQL database.  
It supports overlap detection, automatic skirt extraction, and exporting assembled structures in multiple formats.

## Features
- Create, save, and load geometries directly from Excel or Python.
- Assemble MP, TP, and Tower structures from a central SQL database.
- Detect overlaps between components.
- Extract skirts automatically.
- Export assembled structures to:
  - **JBOOST**
  - **WLGEN**
  - **Bladed**
  - **Genie / Sesam**

## Requirements
- Python (installed and added to the system `PATH`)
- Correctly configured paths in **GlobalConfig**:
  - `python_scripts_path`
  - `python path`
- Required Python packages installed via:
  ```bash
  install_requirements
