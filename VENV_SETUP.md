# Virtual Environment Setup

This project now includes a Python virtual environment to manage dependencies.

## Quick Start

### Option 1: Use the activation script
```bash
source activate_venv.sh
```

### Option 2: Manual activation
```bash
source venv/bin/activate
```

## Running the Syllabus Generator

Once the virtual environment is activated:
```bash
python syllabus_generator.py
```

## Dependencies

The following packages are installed in the virtual environment:
- `python-dateutil` - For date calculations and Easter date computation
- `prettytable` - For formatted table output in the console
- `ics` - For generating iCal calendar files
- `pytz` - For timezone handling

## Deactivating the Virtual Environment

To deactivate the virtual environment:
```bash
deactivate
```

## Installing Additional Dependencies

If you need to add new dependencies:
1. Activate the virtual environment
2. Install the package: `pip install package_name`
3. Update requirements.txt: `pip freeze > requirements.txt`

## Fresh Setup

If you need to recreate the virtual environment:
```bash
rm -rf venv
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```
