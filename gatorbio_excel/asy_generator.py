from __future__ import annotations

import json
from pathlib import Path

from .excel_parser import (
    create_experiment_config,
    create_pre_experiment_config,
    parse_assay_sheet,
    parse_plate_layout,
    parse_pre_experiment_sheet,
    read_excel_file,
)


def generate_asy_file(excel_path: Path | str, output_path: Path | str | None = None) -> Path:
    """Main function to convert Excel file to .asy file."""

    excel_path = Path(excel_path)
    if not excel_path.exists():
        raise FileNotFoundError(f"Excel file not found: {excel_path}")

    print(f"Reading Excel file: {excel_path}")
    sheets = read_excel_file(excel_path)

    pre_exp_config_data = {}
    if "PreExperiment" in sheets:
        print("Parsing PreExperiment sheet...")
        pre_exp_config_data = parse_pre_experiment_sheet(sheets["PreExperiment"])
    else:
        print("Warning: PreExperiment sheet not found, using defaults")

    samples = []
    probe_info = []
    if "Experiment" in sheets:
        print("Parsing Experiment sheet...")
        samples, probe_info = parse_plate_layout(sheets["Experiment"])
        print(f"Parsed {len([s for s in samples if s['SampleID']])} samples")

        # Order: Type, MolarConcentration, Information, SampleID, MolecularWeight, Concentration
        core_sample_keys = [
            "Type",
            "MolarConcentration",
            "Information",
            "SampleID",
            "MolecularWeight",
            "Concentration",
        ]
        core_probe_keys = [
            "Type",
            "MolarConcentration",
            "Information",
            "SampleID",
            "MolecularWeight",
            "Concentration",
        ]
        samples_for_asy = [{k: sample.get(k) for k in core_sample_keys} for sample in samples]
        probe_info_for_asy = [{k: probe.get(k) for k in core_probe_keys} for probe in probe_info]
    else:
        print("Warning: Experiment sheet not found, creating empty samples")
        samples_for_asy = [
            {
                "Type": 4,
                "MolarConcentration": -1.0,
                "Information": "",
                "SampleID": "",
                "MolecularWeight": -1.0,
                "Concentration": -1.0,
            }
            for _ in range(96)
        ]
        probe_info_for_asy = [
            {
                "Type": 1,
                "MolarConcentration": -1.0,
                "Information": "",
                "SampleID": "",
                "MolecularWeight": -1.0,
                "Concentration": -1.0,
            }
            for _ in range(96)
        ]

    rs = {"name": None, "repeat": 3, "RTime": 5, "RSpeed": 1000, "NTime": 5, "NSpeed": 1000}
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
        "W6Speed": 1000,
    }
    assay_steps = None

    if "Assay" in sheets:
        print("Parsing Assay sheet...")
        rs, fs, assay_steps = parse_assay_sheet(sheets["Assay"])
        if assay_steps:
            print(
                f"Found {len(assay_steps)} assay loop(s) with "
                f"{sum(len(loop) for loop in assay_steps)} total steps"
            )
        else:
            print("No assay steps found in Assay sheet, using defaults")
    else:
        print("Warning: Assay sheet not found, using defaults")

    # Get number of loops from assay_steps
    print(f"assay_steps: {assay_steps}")
    num_loops = len(assay_steps)
    pre_exp_config = create_pre_experiment_config(pre_exp_config_data, num_loops)
    exp_config = create_experiment_config(samples_for_asy, probe_info_for_asy, rs, fs, assay_steps)

    if output_path is None:
        output_path = excel_path.parent / f"{excel_path.stem}.asy"
    else:
        output_path = Path(output_path)

    print(f"Generating .asy file: {output_path}")
    with open(output_path, "w") as fp:
        fp.write("[BasicInformation]\n")
        fp.write(f'PreExperiment = "{json.dumps(pre_exp_config, separators=(",", ":"))}"\n')
        fp.write(f'Experiment = "{json.dumps(exp_config, separators=(",", ":"))}"\n')

    print(f"Successfully generated: {output_path}")
    return output_path


__all__ = ["generate_asy_file"]

