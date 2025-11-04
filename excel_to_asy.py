#!/usr/bin/env python3
"""
GatorBio Excel to ASY Converter

This script reads a GatorBio Assay Form Excel file and converts it to a .asy file
that can be imported into the GatorBio BLI machine software.
"""

import json
import sys
import argparse
from datetime import datetime
from pathlib import Path

try:
    import openpyxl
    from openpyxl import load_workbook
except ImportError:
    print("Error: openpyxl is required. Install it with: pip install openpyxl")
    sys.exit(1)


class SampleType:
    """Sample type constants from the .asy file"""
    EMPTY = 0
    ANALYTE = 1  # Analyte samples
    BACKGROUND = 2  # Background control
    POSITIVE = 4  # Positive control (blank)
    REGENERANT = 5  # Regenerant
    NEGATIVE = 6  # Negative control
    REFERENCE = 7  # Reference


def parse_well_position(well_str):
    """
    Convert Excel well position (e.g., 'A1', 'B12') to column number (1-96).
    A1 = 1, A2 = 2, ..., A12 = 12, B1 = 13, ..., H12 = 96
    """
    if not well_str or well_str == '':
        return None
    
    well_str = str(well_str).strip().upper()
    if len(well_str) < 2:
        return None
    
    # Extract letter and number
    letter = well_str[0]
    try:
        number = int(well_str[1:])
    except ValueError:
        return None
    
    # Convert letter to row (A=0, B=1, ..., H=7)
    row = ord(letter) - ord('A')
    if row < 0 or row > 7:
        return None
    
    # Convert to 1-based column number (A1=1, A2=2, ..., H12=96)
    if number < 1 or number > 12:
        return None
    
    column = row * 12 + number
    return column


def map_sample_type_label_to_code(label):
    """
    Map sample type labels from Excel to numeric codes.
    Labels: Buffer, Load, Sample, Regeneration, Neutralization, Probe, Background, Negative, Reference
    """
    if not label:
        return SampleType.POSITIVE  # Default to empty/positive control
    
    label_lower = str(label).strip().lower()
    
    if "buffer" in label_lower or "baseline" in label_lower:
        return SampleType.POSITIVE  # 4
    elif "load" in label_lower:
        return SampleType.REGENERANT  # 5 - Load is typically regenerant
    elif "sample" in label_lower or "analyte" in label_lower:
        return SampleType.ANALYTE  # 1
    elif "regeneration" in label_lower or "regen" in label_lower:
        return SampleType.REGENERANT  # 5
    elif "neutralization" in label_lower:
        return SampleType.REGENERANT  # 5
    elif "probe" in label_lower:
        return SampleType.ANALYTE  # 1 - Probe is treated as analyte in ProbeInfo
    elif "background" in label_lower:
        return SampleType.BACKGROUND  # 2
    elif "negative" in label_lower:
        return SampleType.NEGATIVE  # 6
    elif "reference" in label_lower:
        return SampleType.REFERENCE  # 7
    else:
        return SampleType.POSITIVE  # Default


def read_excel_file(excel_path):
    """Read and parse the Excel file with multiple sheets"""
    wb = load_workbook(excel_path, data_only=True)
    
    # Get all sheets
    sheets = {}
    for sheet_name in wb.sheetnames:
        sheets[sheet_name] = wb[sheet_name]
    
    print(f"Excel file has {len(wb.sheetnames)} sheet(s): {', '.join(wb.sheetnames)}")
    
    return sheets


def parse_pre_experiment_sheet(ws):
    """Parse the PreExperiment sheet and return a dictionary of configuration values"""
    config = {}
    
    # Default values
    defaults = {
        "AssayDescription": "",
        "AssayUser": "User",
        "CreationTime": datetime.now().strftime("%m-%d-%Y %H:%M:%S"),
        "ModificationTime": datetime.now().strftime("%m-%d-%Y %H:%M:%S"),
        "PreAssayShakerASpeed": 1000,
        "PreAssayShakerBSpeed": 1000,
        "PreAssayTime": 300,
        "GapTime": 200,
        "AssayShakerATemperature": 25,
        "AssayShakerBTemperature": 25,
        "bRegeneration": True,
        "bRegenerationStart": False,
        "bPlateAFlat": True,
        "AssayTemperature": 25,
    }
    
    config.update(defaults)
    
    # Parse key-value pairs from the sheet
    for row in ws.iter_rows(values_only=True):
        if row[0] and isinstance(row[0], str):
            key = row[0].strip()
            value = row[1] if len(row) > 1 else None
            
            # Map Excel keys to config keys
            if "Assay Description" in key:
                config["AssayDescription"] = str(value) if value else ""
            elif "Assay User" in key:
                config["AssayUser"] = str(value) if value else "User"
            elif "Creation Time" in key:
                config["CreationTime"] = str(value) if value else config["CreationTime"]
            elif "Modification Time" in key:
                config["ModificationTime"] = str(value) if value else config["ModificationTime"]
            elif "RegenerationStart" in key:
                config["bRegenerationStart"] = bool(value) if value is not None else False
            elif "Regeneration" in key and "Start" not in key:
                config["bRegeneration"] = bool(value) if value is not None else True
            elif "PlateAFlat" in key or "Plate A Flat" in key:
                config["bPlateAFlat"] = bool(value) if value is not None else True
            elif "PreAssayTime" in key:
                config["PreAssayTime"] = int(value) if value is not None else 300
            elif "PreAssayShakerASpeed" in key:
                config["PreAssayShakerASpeed"] = int(value) if value is not None else 1000
            elif "PreAssayShakerBSpeed" in key:
                config["PreAssayShakerBSpeed"] = int(value) if value is not None else 1000
            elif "Assay Temperature" in key:
                temp = int(value) if value is not None else 25
                config["AssayShakerATemperature"] = temp
                config["AssayShakerBTemperature"] = temp
                config["AssayTemperature"] = temp
    
    return config


def parse_plate_layout(ws):
    """
    Parse the Experiment sheet to extract sample and probe information.
    The sheet has:
    - Rows 2-11: Grid layout showing sample types
    - Row 14: Header row with columns: Well, Type, Sample Name, Conc., MW, M Conc., Information
    - Rows 15+: Detailed sample information
    """
    samples = []
    probe_info = []
    
    # Create empty arrays for 96 wells
    for i in range(96):
        samples.append({
            "Type": SampleType.POSITIVE,  # Default to empty/positive control
            "Concentration": -1.0,
            "MolecularWeight": -1.0,
            "MolarConcentration": -1.0,
            "Information": "",
            "SampleID": ""
        })
        probe_info.append({
            "Type": SampleType.ANALYTE,  # Default probe type
            "Concentration": -1.0,
            "MolecularWeight": -1.0,
            "MolarConcentration": -1.0,
            "Information": "",
            "SampleID": "Probe"  # Default probe ID
        })
    
    # Find the detailed sample information section (starting around row 14)
    header_row = None
    for row_idx, row in enumerate(ws.iter_rows(values_only=True), 1):
        if row[1] and isinstance(row[1], str) and "96 Well Plate" in row[1] and "Type" in str(row[2]):
            header_row = row_idx
            break
    
    if header_row:
        # Parse detailed sample information starting from header_row + 1
        # Columns: B=Well, C=Type/SampleID, D=Sample Name/Type, E=Conc, F=MW, G=M Conc, H=Information
        for row_idx in range(header_row + 1, ws.max_row + 1):
            row = list(ws.iter_rows(min_row=row_idx, max_row=row_idx, values_only=True))[0]
            
            # Check if this is a data row (has a well position)
            well_cell = row[1] if len(row) > 1 else None  # Column B
            if not well_cell:
                continue
            
            well_pos = parse_well_position(str(well_cell))
            if not well_pos:
                continue
            
            well_idx = well_pos - 1  # Convert to 0-based index
            if well_idx < 0 or well_idx >= 96:
                continue
            
            # Parse sample information
            # Column C (index 2): Could be type code or sample ID like "EPO-R"
            type_or_id = row[2] if len(row) > 2 else None
            
            # Column D (index 3): Sample name/type label (Buffer, Load, Sample, etc.)
            sample_name_type = row[3] if len(row) > 3 else None
            
            # Column E (index 4): Concentration
            conc = row[4] if len(row) > 4 else None
            
            # Column F (index 5): Molecular Weight
            mw = row[5] if len(row) > 5 else None
            
            # Column G (index 6): Molar Concentration
            m_conc = row[6] if len(row) > 6 else None
            
            # Column H (index 7): Information
            info = row[7] if len(row) > 7 else None
            
            # Determine sample type and ID
            sample_type = SampleType.POSITIVE
            sample_id = ""
            
            if sample_name_type:
                # Use the sample name/type label to determine type
                sample_type = map_sample_type_label_to_code(sample_name_type)
                sample_id = str(sample_name_type).strip()
            
            # If type_or_id looks like a sample ID (e.g., "EPO-R"), use it
            if type_or_id and isinstance(type_or_id, str) and type_or_id.strip():
                sample_id = str(type_or_id).strip()
            
            # Parse concentrations
            concentration = -1.0
            molecular_weight = -1.0
            molar_concentration = -1.0
            
            if conc:
                try:
                    concentration = float(conc)
                except (ValueError, TypeError):
                    pass
            
            if mw:
                try:
                    molecular_weight = float(mw)
                except (ValueError, TypeError):
                    pass
            
            if m_conc:
                try:
                    molar_concentration = float(m_conc)
                except (ValueError, TypeError):
                    pass
            elif concentration != -1.0 and molecular_weight != -1.0 and molecular_weight > 0:
                # Calculate molar concentration if not provided
                molar_concentration = (concentration / molecular_weight) * 1000  # Convert to µM
            
            # Update sample info
            samples[well_idx] = {
                "Type": sample_type,
                "Concentration": concentration,
                "MolecularWeight": molecular_weight,
                "MolarConcentration": molar_concentration,
                "Information": str(info).strip() if info else "",
                "SampleID": sample_id
            }
    
    # Also parse the grid layout (rows 2-11) to get probe information
    # Look for "Max Plate" which contains probe information
    for row_idx in range(2, 12):
        row = list(ws.iter_rows(min_row=row_idx, max_row=row_idx, values_only=True))[0]
        
        # Find the "Max Plate" section (starts around column Q, index 16)
        if len(row) > 16:
            row_letter_cell = row[16]  # Column Q, should be A-H
            if row_letter_cell and isinstance(row_letter_cell, str) and len(row_letter_cell) == 1:
                row_letter = row_letter_cell.strip().upper()
                if row_letter >= 'A' and row_letter <= 'H':
                    row_num = ord(row_letter) - ord('A')
                    
                    # Check columns 1-12 (starting at column R, index 17)
                    for col_num in range(1, 13):
                        cell = row[16 + col_num] if len(row) > 16 + col_num else None
                        if cell:
                            cell_str = str(cell).strip()
                            if "Probe" in cell_str:
                                well_idx = row_num * 12 + col_num - 1
                                if well_idx >= 0 and well_idx < 96:
                                    probe_info[well_idx]["SampleID"] = "Probe"
    
    return samples, probe_info


def parse_assay_sheet(ws):
    """
    Parse the Assay sheet to extract regeneration settings and assay steps.
    Returns: (regeneration_settings, assay_steps)
    """
    # Default regeneration settings
    rs = {
        "name": None,
        "repeat": 3,
        "RTime": 5,
        "RSpeed": 1000,
        "NTime": 5,
        "NSpeed": 1000
    }
    
    # Default flow settings
    fs = {
        "name": None,
        "repeatRP": 5,
        "ReaTime": 5,
        "ReaSpeed": 1000,
        "Rea2Time": 5,
        "Rea2Speed": 1000,
        "PriTime": 5,
        "PriSpeed": 1000,
        "CapTime": 120,
        "CapSpeed": 1000,
        "W1Time": 10,
        "W1Speed": 1000,
        "W2Time": 10,
        "W2Speed": 1000,
        "W3Time": 10,
        "W3Speed": 1000,
        "W4Time": 10,
        "W4Speed": 1000,
        "W5Time": 10,
        "W5Speed": 1000,
        "W6Time": 10,
        "W6Speed": 1000
    }
    
    # Parse regeneration settings
    in_rs_section = False
    for row in ws.iter_rows(values_only=True):
        if row[0] and isinstance(row[0], str):
            key = str(row[0]).strip().lower()
            value = row[1] if len(row) > 1 else None
            
            if "regeneration step" in key or "rs" in key:
                in_rs_section = True
                continue
            
            if in_rs_section and value is not None:
                if "repeat" in key:
                    try:
                        rs["repeat"] = int(float(value))
                    except (ValueError, TypeError):
                        pass
                elif "rtime" in key:
                    try:
                        rs["RTime"] = int(float(value))
                    except (ValueError, TypeError):
                        pass
                elif "rspeed" in key:
                    try:
                        rs["RSpeed"] = int(float(value))
                    except (ValueError, TypeError):
                        pass
                elif "ntime" in key:
                    try:
                        rs["NTime"] = int(float(value))
                    except (ValueError, TypeError):
                        pass
                elif "nspeed" in key:
                    try:
                        rs["NSpeed"] = int(float(value))
                    except (ValueError, TypeError):
                        pass
    
    # Parse assay steps
    steps = []
    current_loop = []
    in_steps_section = False
    header_found = False
    
    for row in ws.iter_rows(values_only=True):
        if row[0] and isinstance(row[0], str):
            if "experiment step" in str(row[0]).lower() or "step" in str(row[0]).lower():
                if "steptype" in str(row).lower() or "steptype" in str(row).lower():
                    header_found = True
                    in_steps_section = True
                    continue
        
        if in_steps_section and header_found and row[0] is not None:
            # Parse step row: Step, Plate, Column, Sec, rpm, StepType
            try:
                step_num = int(float(row[0])) if row[0] else None
                plate = int(float(row[1])) if len(row) > 1 and row[1] else 1
                column = int(float(row[2])) if len(row) > 2 and row[2] else 1
                time = int(float(str(row[3]))) if len(row) > 3 and row[3] else 120
                speed = int(float(row[4])) if len(row) > 4 and row[4] else 1000
                step_type_str = str(row[5]).strip().lower() if len(row) > 5 and row[5] else ""
            except (ValueError, TypeError):
                continue
            
            # Map step type string to numeric code
            step_type = None
            if "baseline" in step_type_str or "capture" in step_type_str:
                step_type = 0
            elif "loading" in step_type_str or "load" in step_type_str:
                step_type = 1  # Loading is like association
            elif "assoc" in step_type_str or "association" in step_type_str:
                step_type = 1
            elif "dissoc" in step_type_str or "dissociation" in step_type_str:
                step_type = 2
            elif "regen" in step_type_str or "regeneration" in step_type_str:
                step_type = 3
            else:
                continue  # Skip unknown step types
            
            # Create step
            step = {
                "type": step_type,
                "isMultiStep": False,
                "isSubElementExist": False,
                "listProbeColumn": [plate],  # Use plate number as probe column
                "listStep": [{"SampleColumn": column, "Speed": speed, "Time": time}]
            }
            
            # Determine if this starts a new loop
            # A new loop typically starts with Baseline/Capture steps
            if step_type == 0 and current_loop:
                # Save previous loop and start new one
                steps.append(current_loop)
                current_loop = [step]
            else:
                current_loop.append(step)
    
    # Add final loop
    if current_loop:
        steps.append(current_loop)
    
    return rs, fs, steps if steps else None


def create_pre_experiment_config(pre_exp_data):
    """Create the PreExperiment configuration JSON from parsed data"""
    now = datetime.now()
    time_str = now.strftime("%m-%d-%Y %H:%M:%S")
    
    config = {
        "AssayDescription": pre_exp_data.get("AssayDescription", ""),
        "AssayUser": pre_exp_data.get("AssayUser", "User"),
        "CreationTime": pre_exp_data.get("CreationTime", time_str),
        "ModificationTime": pre_exp_data.get("ModificationTime", time_str),
        "StartExperimentTime": None,
        "EndExperimentTime": None,
        "AnotherSavePath": "",
        "AssayType": 2,
        "PreAssayShakerASpeed": pre_exp_data.get("PreAssayShakerASpeed", 1000),
        "PreAssayShakerBSpeed": pre_exp_data.get("PreAssayShakerBSpeed", 1000),
        "PreAssayTime": pre_exp_data.get("PreAssayTime", 300),
        "GapTime": 200,
        "AssayShakerATemperature": pre_exp_data.get("AssayShakerATemperature", 25),
        "AssayShakerBTemperature": pre_exp_data.get("AssayShakerBTemperature", 25),
        "MachineRealType": "GatorPrime",
        "idleShakerATemperature": 30,
        "idleShakerBTemperature": 30,
        "PlateAType": 0,
        "bPlateAFlat": pre_exp_data.get("bPlateAFlat", True),
        "bRegeneration": pre_exp_data.get("bRegeneration", True),
        "bRegenerationStart": pre_exp_data.get("bRegenerationStart", False),
        "RegenerationNum": 99999,
        "IsOpenRNAfterAssay": {str(i): True for i in range(7)},
        "IsOpenRNBeforeAssay": {str(i): True for i in range(7)},
        "RegenerationMode": 0,
        "ShakerASpeedDeviationLow": 0,
        "ShakerBSpeedDeviationLow": 0,
        "ShakerASpeedDeviationHigh": 0,
        "ShakerBSpeedDeviationHigh": 0,
        "ShakerATempDeviation": 0.0,
        "ShakerBTempDeviation": 0.0,
        "SpectrometerTempDeviation": 0.0,
        "ParentName": None,
        "ResultName": None,
        "SoftWareVersion": "2.17.7.0416",
        "SeriesNo": None,
        "ErrorChannelList": None,
        "ErrorPreAssayList": None,
        "ErrorSampleList": None,
        "ErrorRegenerationList": None,
        "AnalysisSettingList": [],
        "ReportSettingList": [],
        "PlateColumns": [12, 12],
        "threshold": "Infinity",
        "bsingle": False,
        "bThresh": False,
        "strFileID": None
    }
    
    return config


def create_experiment_config(samples, probe_info, rs, fs, assay_steps=None):
    """Create the Experiment configuration JSON"""
    
    # Default assay steps if none provided
    if assay_steps is None:
        assay_steps = [[
            {
                "type": 0,  # Capture
                "isMultiStep": False,
                "isSubElementExist": False,
                "listProbeColumn": [1],
                "listStep": [{"SampleColumn": 1, "Speed": 1000, "Time": 120}]
            },
            {
                "type": 1,  # Association
                "isMultiStep": False,
                "isSubElementExist": False,
                "listProbeColumn": [1],
                "listStep": [{"SampleColumn": 2, "Speed": 1000, "Time": 120}]
            },
            {
                "type": 0,  # Baseline
                "isMultiStep": False,
                "isSubElementExist": False,
                "listProbeColumn": [1],
                "listStep": [{"SampleColumn": 3, "Speed": 1000, "Time": 60}]
            },
            {
                "type": 2,  # Dissociation
                "isMultiStep": False,
                "isSubElementExist": False,
                "listProbeColumn": [1],
                "listStep": [{"SampleColumn": 4, "Speed": 1000, "Time": 200}]
            },
            {
                "type": 3,  # Regeneration
                "isMultiStep": False,
                "isSubElementExist": False,
                "listProbeColumn": [1],
                "listStep": [{"SampleColumn": 5, "Speed": 1000, "Time": 400}]
            }
        ]]
    
    config = {
        "SampleInfo": samples,
        "ProbeInfo": probe_info,
        "rs": rs,
        "fs": fs,
        "listLoopStep": assay_steps,
        "ErrorMessage": None,
        "WarningMessage": "",
        "ConcentrationUnit": 0,
        "MolarConcentrationUnit": 0
    }
    
    return config


def generate_asy_file(excel_path, output_path=None):
    """Main function to convert Excel file to .asy file"""
    excel_path = Path(excel_path)
    
    if not excel_path.exists():
        raise FileNotFoundError(f"Excel file not found: {excel_path}")
    
    # Read Excel file
    print(f"Reading Excel file: {excel_path}")
    sheets = read_excel_file(excel_path)
    
    # Parse PreExperiment sheet
    pre_exp_config_data = {}
    if "PreExperiment" in sheets:
        print("Parsing PreExperiment sheet...")
        pre_exp_config_data = parse_pre_experiment_sheet(sheets["PreExperiment"])
    else:
        print("Warning: PreExperiment sheet not found, using defaults")
    
    # Parse Experiment sheet
    samples = []
    probe_info = []
    if "Experiment" in sheets:
        print("Parsing Experiment sheet...")
        samples, probe_info = parse_plate_layout(sheets["Experiment"])
        print(f"Parsed {len([s for s in samples if s['SampleID']])} samples")
    else:
        print("Warning: Experiment sheet not found, creating empty samples")
        samples = [{"Type": 4, "Concentration": -1.0, "MolecularWeight": -1.0, 
                   "MolarConcentration": -1.0, "Information": "", "SampleID": ""} for _ in range(96)]
        probe_info = [{"Type": 1, "Concentration": -1.0, "MolecularWeight": -1.0, 
                      "MolarConcentration": -1.0, "Information": "", "SampleID": ""} for _ in range(96)]
    
    # Parse Assay sheet
    rs = {"name": None, "repeat": 3, "RTime": 5, "RSpeed": 1000, "NTime": 5, "NSpeed": 1000}
    fs = {"name": None, "repeatRP": 5, "ReaTime": 5, "ReaSpeed": 1000, "Rea2Time": 5, "Rea2Speed": 1000,
          "PriTime": 5, "PriSpeed": 1000, "CapTime": 120, "CapSpeed": 1000,
          "W1Time": 10, "W1Speed": 1000, "W2Time": 10, "W2Speed": 1000, "W3Time": 10, "W3Speed": 1000,
          "W4Time": 10, "W4Speed": 1000, "W5Time": 10, "W5Speed": 1000, "W6Time": 10, "W6Speed": 1000}
    assay_steps = None
    
    if "Assay" in sheets:
        print("Parsing Assay sheet...")
        rs, fs, assay_steps = parse_assay_sheet(sheets["Assay"])
        if assay_steps:
            print(f"Found {len(assay_steps)} assay loop(s) with {sum(len(loop) for loop in assay_steps)} total steps")
        else:
            print("No assay steps found in Assay sheet, using defaults")
    else:
        print("Warning: Assay sheet not found, using defaults")
    
    # Create configurations
    pre_exp_config = create_pre_experiment_config(pre_exp_config_data)
    exp_config = create_experiment_config(samples, probe_info, rs, fs, assay_steps)
    
    # Generate output filename
    if output_path is None:
        output_path = excel_path.parent / f"{excel_path.stem}.asy"
    else:
        output_path = Path(output_path)
    
    # Write .asy file
    print(f"Generating .asy file: {output_path}")
    with open(output_path, 'w') as f:
        f.write("[BasicInformation]\n")
        f.write(f'PreExperiment = "{json.dumps(pre_exp_config, separators=(",", ":"))}"\n')
        f.write(f'Experiment = "{json.dumps(exp_config, separators=(",", ":"))}"\n')
    
    print(f"Successfully generated: {output_path}")
    return output_path


def main():
    parser = argparse.ArgumentParser(
        description="Convert GatorBio Assay Form Excel file to .asy file"
    )
    parser.add_argument(
        "excel_file",
        help="Path to the Excel file (.xlsx)"
    )
    parser.add_argument(
        "-o", "--output",
        help="Output .asy file path (default: same name as Excel file)"
    )
    
    args = parser.parse_args()
    
    try:
        generate_asy_file(
            args.excel_file,
            args.output
        )
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
