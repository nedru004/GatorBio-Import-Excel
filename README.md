# GatorBio Excel to ASY Converter

This Python program converts a GatorBio Assay Form Excel file (`.xlsx` or `.xlsm`) into a `.asy` file that can be imported into the GatorBio BLI machine software.

## Features

- **Simple Table-Based Reading**: Reads data from sequential table rows
- **Sample Information**: Extracts sample IDs, types, concentrations, and molecular weights
- **Probe Information**: Extracts probe information for Max Plate configuration
- **Assay Steps Configuration**: Parses user-defined assay loops and steps
- **GUI Support**: Interactive file selection dialogs
- **Complete Configuration**: Generates all required parameters for the GatorBio BLI machine

## Requirements

- Python 3.6 or higher
- openpyxl (for reading Excel files)
- tkinter (usually included with Python, for GUI dialogs)

## Installation

1. Create a virtual environment (recommended):
```bash
python3 -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

## Usage

### GUI Mode (Interactive)

Run the script without any arguments to open file selection dialogs:

```bash
python excel_to_asy.py
```

This will:
1. Open a dialog to select your Excel file (.xlsx or .xlsm)
2. Open a dialog to choose where to save the .asy file
3. Show a success message when complete

### Command Line Mode

```bash
python excel_to_asy.py "GatorBio Assay Form.xlsm" -o "output.asy"
```

### Command Line Arguments

- `excel_file`: (Required) Path to the Excel file (.xlsx or .xlsm)
- `-o, --output`: Output .asy file path (default: same name as Excel file)

### Example

```bash
python excel_to_asy.py "GatorBio Assay Form.xlsm" -o "MyAssay.asy"
```

## Excel File Format

The script expects an Excel file with three sheets:

### 1. PreExperiment Sheet
Contains key-value pairs for configuration parameters:
- Assay name, description, user information
- Machine settings (temperatures, speeds, times)
- Shaker settings

### 2. Experiment Sheet
Contains two tables:

**SampleInfo Table (Columns C-H, Rows 14-110):**
- Row 14: Header row
- Rows 15-110: Data rows (one per well, up to 96 wells)
- Column C: SampleID
- Column D: Type (Buffer, Load, Sample, Regeneration, Neutralization, etc.)
- Column E: Concentration
- Column F: Molecular Weight
- Column G: Molar Concentration
- Column H: Information

**ProbeInfo (Max Plate) Table (Columns Q-V, Rows 14-110):**
- Row 14: Header row
- Rows 15-110: Data rows (one per well, up to 96 wells)
- Column Q: SampleID/Name
- Column R: Type (Probe, Buffer, Sample, Load, Regeneration, Neutralization, etc.)
- Column S: Concentration
- Column T: Molecular Weight
- Column U: Molar Concentration
- Column V: Information

### 3. Assay Sheet
Contains user-defined assay loops and steps:
- Loop column: Defines which loop each step belongs to
- Probe column: Probe column number
- Plate and Column: Plate number and column within plate
- Time (Sec) and Speed (rpm): Step parameters
- StepType: Baseline, Loading, Association, Dissociation, etc.

## Sample Types

The script supports the following sample types:
- `0`: Empty/Blank
- `1`: Sample/Analyte (Probe in ProbeInfo)
- `2`: Background (Sample in ProbeInfo)
- `4`: Buffer (default)
- `5`: Load
- `6`: Regeneration
- `7`: Neutralization

## Output Format

The generated `.asy` file contains:

1. **BasicInformation Section**: Contains PreExperiment and Experiment JSON configurations
2. **PreExperiment**: Assay parameters, machine settings, shaker settings, etc.
3. **Experiment**: 
   - SampleInfo: Array of 96 sample definitions (read from C14-H110)
   - ProbeInfo: Array of 96 probe definitions (read from Q14-V110)
   - Regeneration settings (rs)
   - Flow settings (fs)
   - listLoopStep: User-defined assay step loops

## Notes

- The script reads tables sequentially: row 15 maps to index 0, row 16 to index 1, etc.
- Molar concentrations are automatically calculated if not provided (from concentration and molecular weight)
- All concentrations are assumed to be in µg/mL, molecular weights in kDa
- Types 6 (Regeneration) and 7 (Neutralization) automatically get SampleID set to "N/A"
- The generated .asy file is compatible with GatorBio software version 2.17.7.0416
