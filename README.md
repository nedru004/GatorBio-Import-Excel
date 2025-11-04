# GatorBio Excel to ASY Converter

This Python program converts a GatorBio Assay Form Excel file (`.xlsx`) into a `.asy` file that can be imported into the GatorBio BLI machine software.

## Features

- **Plate Layout Parsing**: Reads 96-well plate layouts from the Excel file
- **Sample Information**: Extracts sample IDs, concentrations, and molecular weights
- **Assay Steps Configuration**: Parses assay steps (Capture, Association, Baseline, Dissociation, Regeneration)
- **Complete Configuration**: Generates all required parameters for the GatorBio BLI machine

## Requirements

- Python 3.6 or higher
- openpyxl (for reading Excel files)

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

### Basic Usage

```bash
python excel_to_asy.py "GatorBio Assay Form.xlsx"
```

This will create a `.asy` file with the same name as the Excel file.

### Advanced Usage

```bash
python excel_to_asy.py "GatorBio Assay Form.xlsx" -o "output.asy" -u "Your Name" -d "Assay Description"
```

### Command Line Arguments

- `excel_file`: (Required) Path to the Excel file (.xlsx)
- `-o, --output`: Output .asy file path (default: same name as Excel file)
- `-u, --user`: User name for the assay (default: "User")
- `-d, --description`: Assay description (default: "")

### Example

```bash
python excel_to_asy.py "GatorBio Assay Form.xlsx" -u "John Doe" -d "EPO Binding Assay"
```

## Excel File Format

The script expects the Excel file to have:

1. **96 Well Plate Layout**: 
   - A header row with "96 Well Plate"
   - Column numbers (1-12) and row letters (A-H)
   - Sample information can be entered in cells corresponding to well positions

2. **Sample Information** (optional):
   - Column B: Well positions (e.g., "A1", "B1")
   - Column C: Sample IDs
   - Column D: Concentrations (µg/mL)
   - Column E: Molecular Weights (kDa)

3. **Assay Steps** (optional):
   - A header row with "Assay Steps"
   - Step definitions with:
     - Step type (Capture/Baseline=0, Association=1, Dissociation=2, Regeneration=3)
     - Sample column/well position
     - Speed (rpm)
     - Time (seconds)

If assay steps are not found in the Excel file, the script will use default steps:
- Capture (120s)
- Association (120s)
- Baseline (60s)
- Dissociation (200s)
- Regeneration (400s)

## Sample Types

The script supports the following sample types:
- `0`: Empty
- `1`: Analyte
- `2`: Background control
- `4`: Positive control (blank) - default
- `5`: Regenerant
- `6`: Negative control
- `7`: Reference

## Output Format

The generated `.asy` file contains:

1. **BasicInformation Section**: Contains PreExperiment and Experiment JSON configurations
2. **PreExperiment**: Assay parameters, machine settings, shaker settings, etc.
3. **Experiment**: 
   - SampleInfo: Array of 96 sample definitions
   - ProbeInfo: Array of 96 probe definitions
   - Regeneration settings (rs)
   - Flow settings (fs)
   - listLoopStep: Assay step definitions

## Troubleshooting

If the script cannot find sample information or assay steps in the Excel file, it will:
- Create empty/default entries for samples
- Use default assay steps

You can examine the Excel file structure using:
```bash
python examine_excel.py "GatorBio Assay Form.xlsx"
```

## Notes

- The script automatically calculates molar concentrations from concentration and molecular weight values
- Well positions are converted from Excel format (A1-H12) to column numbers (1-96)
- All concentrations are assumed to be in µg/mL, molecular weights in kDa
- The generated .asy file is compatible with GatorBio software version 2.17.7.0416
