from __future__ import annotations

import json
from collections import defaultdict
from pathlib import Path
from typing import Dict, List, Tuple
from math import ceil
from datetime import datetime

from .common import (
    SampleType,
    sanitize_identifier,
    map_assay_code_to_label,
    map_max_plate_code_to_label,
)
from .excel_parser import parse_plate_layout, read_excel_file


def generate_liquid_handler_notebook(
    excel_path: Path | str,
    output_path: Path | str | None = None,
    buffer_source: str = "tube",
) -> Path:
    """Generate a Jupyter notebook for liquid handling using pylabrobot.
    
    Args:
        excel_path: Path to the Excel file to process
        output_path: Path for the output notebook (default: same name as Excel with .ipynb)
        buffer_source: Buffer source type - "tube" for 50mL tube or "trough" for 60mL trough (default: "tube")
    """

    excel_path = Path(excel_path)
    if not excel_path.exists():
        raise FileNotFoundError(f"Excel file not found: {excel_path}")

    print(f"Reading Excel file: {excel_path}")
    sheets = read_excel_file(excel_path)

    if "Experiment" not in sheets:
        raise ValueError("Experiment sheet not found in Excel file")

    assay_plate, max_plate = parse_plate_layout(sheets["Experiment"])

    initial_concentrations: Dict[str, float] = {}

    ASSAY_DILUTION_VOLUME_UL = 300
    MAX_DILUTION_VOLUME_UL = 300
    ASSAY_FINAL_VOLUME_UL = 200
    MAX_FINAL_VOLUME_UL = 200

    buffer_resource_name = "stock_buffer"

    ASSAY_SHARED_RESOURCE_MAP = {
        "buffer": "stock_buffer",
        "regeneration": "stock_regeneration",
        "neutralization": "stock_neutralization",
        "activation": "stock_buffer",
        "quench": "stock_buffer",
        "wash": "stock_buffer",
    }

    MAX_SHARED_RESOURCE_MAP = {
        "probe": "stock_buffer",
        "buffer": "stock_buffer",
        "regeneration": "stock_regeneration",
        "neutralization": "stock_neutralization",
        "activation": "stock_buffer",
        "quench": "stock_buffer",
        "wash": "stock_buffer",
    }

    combined_samples: Dict[str, dict] = {}

    def register_entry(entry, well_idx, plate_kind):
        sample_id = str(entry.get("SampleID") or "").strip()
        concentration = entry.get("Concentration", -1.0)
        type_code = entry.get("Type", None)
        
        # Convert numeric type code to string label
        if plate_kind == "assay":
            sample_type = map_assay_code_to_label(type_code) if type_code is not None else "buffer"
        else:
            sample_type = map_max_plate_code_to_label(type_code) if type_code is not None else "buffer"

        if (not sample_id or sample_id.upper() == "N/A") and sample_type == "sample":
            return

        record = combined_samples.setdefault(
            sample_id,
            {
                "occurrences": [],
                "max_conc": -1.0,
                "has_assay": False,
                "has_max": False,
                "levels": [],
            },
        )

        if plate_kind == "assay":
            record["has_assay"] = True
        else:
            record["has_max"] = True

        record["occurrences"].append(
            {
                "plate": plate_kind,
                "well_idx": well_idx,
                "concentration": concentration,
                "type": sample_type,  # Now using string label instead of numeric code
            }
        )

        if concentration is not None and concentration > 0:
            record["max_conc"] = max(record["max_conc"], concentration)

    for well_idx, sample in enumerate(assay_plate):
        register_entry(sample, well_idx, "assay")

    for well_idx, entry in enumerate(max_plate):
        register_entry(entry, well_idx, "max")

    for sample_id, record in combined_samples.items():
        buckets = {}
        for occ in record["occurrences"]:
            conc = occ.get("concentration", -1.0)
            conc_value = float(conc) if conc is not None and conc > 0 else 0.0
            key = round(conc_value, 6)
            bucket = buckets.setdefault(key, {"value": conc_value, "occurrences": []})
            bucket["occurrences"].append(occ)

        record["levels"] = sorted(
            buckets.values(),
            key=lambda b: b["value"],
            reverse=True,
        )

        if record["max_conc"] <= 0:
            record["max_conc"] = 1.0

    combined_sample_ids = sorted(combined_samples.keys())
    assay_ids = [sid for sid in combined_sample_ids if combined_samples[sid]["has_assay"]]
    max_ids = [sid for sid in combined_sample_ids if combined_samples[sid]["has_max"]]

    row_usage = [1] * 8  # next available column per row (1-indexed)
    max_column_used = 1
    MAX_DILUTION_PLATES = 2
    MAX_COLUMNS_PER_ROW = MAX_DILUTION_PLATES * 12
    current_row_pointer = 0
    for sample_id in combined_sample_ids:
        if sample_id == "":
            continue
        record = combined_samples[sample_id]
        total_occurrences = len(record["occurrences"])
        wells_needed = max(total_occurrences, 1)
        assigned = False
        for attempt in range(8):
            row_idx = (current_row_pointer + attempt) % 8
            start_col_candidate = row_usage[row_idx]
            end_col = start_col_candidate + wells_needed - 1
            if end_col > MAX_COLUMNS_PER_ROW:
                continue
            record["row_letter"] = chr(ord("A") + row_idx)
            record["start_col"] = start_col_candidate
            record["next_offset"] = {"assay": 0, "max": 0}
            row_usage[row_idx] = end_col + 1
            max_column_used = max(max_column_used, end_col)
            current_row_pointer = (row_idx + 1) % 8
            assigned = True
            break

        if not assigned:
            raise ValueError(
                "Unable to assign dilutions without overlapping wells. Reduce the number of dilutions "
                "or extend the deck configuration."
            )

    max_column_used = max(1, max_column_used)
    plates_needed = max(1, ceil(max_column_used / 12))
    if plates_needed > 2:
        raise ValueError(
            "More than two dilution plates are required. Reduce the number of dilution wells or update the deck layout."
        )

    dilution_plate_names = ["dilution_plate"]
    if plates_needed >= 2:
        dilution_plate_names.append("dilution_plate_2")

    column_sequences = {
        "assay": defaultdict(list),
        "max": defaultdict(list),
    }

    stock_resource_map = {sample_id: None for sample_id in combined_sample_ids}

    init_conc_vars = []
    for sample_id in assay_ids:
        record = combined_samples[sample_id]
        safe_name = sanitize_identifier(sample_id, prefix="S")
        var_name = safe_name.upper()
        default_conc = initial_concentrations.get(sample_id, record["max_conc"])
        if default_conc is None or default_conc <= 0:
            default_conc = record["max_conc"]
        init_conc_vars.append(
            f"INITIAL_CONC_{var_name} = {default_conc}  # µg/mL - Stock concentration for {sample_id}"
        )

    init_max_conc_vars = []
    for sample_id in max_ids:
        record = combined_samples[sample_id]
        safe_name = sanitize_identifier(sample_id, prefix="P")
        var_name = f"MAX_{safe_name.upper()}"
        default_conc = record["max_conc"]
        if safe_name.upper() != "PROBE":
            init_max_conc_vars.append(
                f"INITIAL_CONC_{var_name} = {default_conc}  # µg/mL - Stock concentration for Max Plate {sample_id}"
            )

    sample_loading_lines = ["## Sample Stock Placement\n"]
    manual_positions = set()  # Use single set to prevent duplicates across assay and max plates

    resource_volume_uL = defaultdict(float)
    for res in {"stock_buffer", "stock_regeneration", "stock_neutralization"}:
        resource_volume_uL[res] = 0.0
    for res in {val for val in stock_resource_map.values() if val}:
        resource_volume_uL.setdefault(res, 0.0)

    tube_resources = sorted(resource_volume_uL.keys())

    def well_position_sort_key(position: str) -> tuple[int, int]:
        row_letter = position[0].upper()
        column_number = int(position[1:])
        return (column_number, ord(row_letter))

    assay_shared_wells = defaultdict(list)
    for idx, entry in enumerate(assay_plate):
        type_code = entry.get("Type", SampleType.Assay.EMPTY)
        entry_type = map_assay_code_to_label(type_code)
        resource = ASSAY_SHARED_RESOURCE_MAP.get(entry_type)
        if resource:
            well_position = entry.get("WellPosition")
            if not well_position:
                raise ValueError(f"Missing WellPosition for assay sample index {idx}")
            assay_shared_wells[resource].append(well_position)
            resource_volume_uL[resource] += ASSAY_FINAL_VOLUME_UL

    max_shared_wells = defaultdict(list)
    for idx, entry in enumerate(max_plate):
        type_code = entry.get("Type", SampleType.MaxPlate.EMPTY)
        entry_type = map_max_plate_code_to_label(type_code)
        resource = MAX_SHARED_RESOURCE_MAP.get(entry_type)
        if resource:
            well_position = entry.get("WellPosition")
            if not well_position:
                raise ValueError(f"Missing WellPosition for Max Plate index {idx}")
            max_shared_wells[resource].append(well_position)
            resource_volume_uL[resource] += MAX_FINAL_VOLUME_UL

    notebook = {
        "cells": [],
        "metadata": {
            "kernelspec": {
                "display_name": "Python 3",
                "language": "python",
                "name": "python3",
            },
            "language_info": {"name": "python", "version": "3.8.0"},
        },
        "nbformat": 4,
        "nbformat_minor": 4,
    }

    notebook["cells"].append(
        {
            "cell_type": "markdown",
            "metadata": {},
            "source": [
                "# Liquid Handler Program for GatorBio Assay\n",
                f"Generated from: {excel_path.name}\n",
                f"Generated on: {datetime.now().strftime('%Y-%B-%d %H:%M:%S')}\n",
                "This notebook prepares assay and Max Plate (ProbeInfo) dilutions using pylabrobot.",
            ],
        }
    )

    notebook["cells"].append(
        {
            "cell_type": "markdown",
            "metadata": {},
            "source": ["## Setup Machine"],
        }
    )

    notebook["cells"].append(
        {
            "cell_type": "code",
            "execution_count": None,
            "metadata": {},
            "source": [
                "%load_ext autoreload\n",
                "%autoreload 2\n",
                "from pylabrobot.liquid_handling import LiquidHandler\n",
                "from pylabrobot.liquid_handling.backends import STARBackend\n",
                "from pylabrobot.resources import Deck, Coordinate\n",
                "from pylabrobot.liquid_handling import Strictness, set_strictness\n",
                "from pylabrobot.resources.hamilton import STARDeck\n",
                "from pylabrobot.liquid_handling.utils import get_wide_single_resource_liquid_op_offsets\n"
                "import time\n",
                "\n",
                'lh = LiquidHandler(backend=STARBackend(read_timeout=600), deck=STARDeck(core_grippers="1000uL-5mL-on-waste"))\n',
                "await lh.setup(skip_iswap=True)\n",
                "set_strictness(Strictness.STRICT)",
            ],
        }
    )

    notebook["cells"].append(
        {
            "cell_type": "markdown",
            "metadata": {},
            "source": ["## Set Variables\n", "Change these values for your experiment."],
        }
    )

    variables_code = [
        "# Assay plate parameters\n",
        "FINAL_VOLUME = 200  # µL transferred to assay 96-well plate\n",
        "DILUTION_VOLUME = 300  # µL per dilution in deep well plate\n",
        "\n",
        "# Max Plate parameters\n",
        "MAX_PLATE_FINAL_VOLUME = 200  # µL transferred to Max Plate\n",
        "MAX_PLATE_DILUTION_VOLUME = 300  # µL per dilution in Max Plate deep well\n",
        "\n",
        "# Carrier Rail Assignment\n",
        f"BUFFER_SOURCE = '{buffer_source}'\n",
        "PLATE_CARRIER_RAIL = 19\n",
        "TUBE_CARRIER_RAIL = 35\n",
        "TIP_CARRIER_RAIL = 7\n",
        "\n",
        "# Mixing parameters\n",
        "MIX_CYCLES = 10\n",
        "MIX_VOLUME = 200  # µL per mix\n",
        "MIX_FLOW_RATE = 100  # µL/s\n",
        "CHANGE_TIPS_BETWEEN_DILUTIONS = False\n",
        "\n",
    ]

    notebook["cells"].append(
        {
            "cell_type": "code",
            "execution_count": None,
            "metadata": {},
            "source": variables_code,
        }
    )

    summary_cell_index = len(notebook["cells"])
    notebook["cells"].append(
        {
            "cell_type": "markdown",
            "metadata": {},
            "source": ["## Stock Volume Requirements\n", "Volumes will be calculated below.\n"],
        }
    )

    sample_instructions_cell_index = len(notebook["cells"])
    notebook["cells"].append(
        {
            "cell_type": "markdown",
            "metadata": {},
            "source": ["## Sample Stock Placement\n", "Instructions will be generated below.\n"],
        }
    )

    dilution_plate_names_repr = repr(dilution_plate_names)
    deck_code = [
        "from pylabrobot.resources import (\n",
        "    TIP_CAR_480_A00,\n",
        "    PLT_CAR_L5AC_A00,\n",
        "    Cor_96_wellplate_360ul_Fb,\n",
        "    Cor_96_wellplate_2mL_Vb,\n",
        "    hamilton_96_tiprack_1000uL,\n",
        "    Tube_CAR_32_A00,\n",
        "    hamilton_tube_carrier_12_b00,\n",
        "    Cor_Falcon_tube_50mL_Vb,\n",
        "    Trough_CAR_5R60_A00,\n",
        "    hamilton_1_trough_60ml_Vb,\n",
        ")\n",
        "from pylabrobot.liquid_handling.standard import Mix\n",
        "\n",
        'lh.deck.get_resource("trash_core96").location = Coordinate(-260, 106, 216.4)\n',
        "\n",
        "# Tips\n",
        'tip_car = TIP_CAR_480_A00(name="tip_carrier")\n',
        'tip_car[0] = hamilton_96_tiprack_1000uL(name="main_tips")\n',
        'tip_car[1] = hamilton_96_tiprack_1000uL(name="spare_tips")\n',
        "lh.deck.assign_child_resource(tip_car, rails=TIP_CARRIER_RAIL)\n",
        "\n",
        "TIP_WELL_ORDER = [\n",
        "    'A1', 'B1', 'C1', 'D1', 'E1', 'F1', 'G1', 'H1',\n",
        "    'A2', 'B2', 'C2', 'D2', 'E2', 'F2', 'G2', 'H2',\n",
        "    'A3', 'B3', 'C3', 'D3', 'E3', 'F3', 'G3', 'H3',\n",
        "    'A4', 'B4', 'C4', 'D4', 'E4', 'F4', 'G4', 'H4',\n",
        "    'A5', 'B5', 'C5', 'D5', 'E5', 'F5', 'G5', 'H5',\n",
        "    'A6', 'B6', 'C6', 'D6', 'E6', 'F6', 'G6', 'H6',\n",
        "    'A7', 'B7', 'C7', 'D7', 'E7', 'F7', 'G7', 'H7',\n",
        "    'A8', 'B8', 'C8', 'D8', 'E8', 'F8', 'G8', 'H8',\n",
        "    'A9', 'B9', 'C9', 'D9', 'E9', 'F9', 'G9', 'H9',\n",
        "    'A10', 'B10', 'C10', 'D10', 'E10', 'F10', 'G10', 'H10',\n",
        "    'A11', 'B11', 'C11', 'D11', 'E11', 'F11', 'G11', 'H11',\n",
        "    'A12', 'B12', 'C12', 'D12', 'E12', 'F12', 'G12', 'H12',\n",
        "]\n",
        "TIP_POSITIONS = [(\"main_tips\", pos) for pos in TIP_WELL_ORDER] + [(\"spare_tips\", pos) for pos in TIP_WELL_ORDER]\n",
        "tip_position_index = 0\n",
        "\n",
        "# Plates\n",
        f"dilution_plate_names = {dilution_plate_names_repr}\n",
        'plt_car = PLT_CAR_L5AC_A00(name="plate_carrier")\n',
        'plt_car[0] = Cor_96_wellplate_360ul_Fb(name="final_plate")  # Assay plate\n',
        'plt_car[1] = Cor_96_wellplate_360ul_Fb(name="max_plate_final")  # Max Plate\n',
        'plt_car[2] = Cor_96_wellplate_2mL_Vb(name=dilution_plate_names[0])  # Shared dilutions\n',
        "if len(dilution_plate_names) > 1:\n",
        '    plt_car[3] = Cor_96_wellplate_2mL_Vb(name=dilution_plate_names[1])  # Additional dilutions\n',
        "lh.deck.assign_child_resource(plt_car, rails=PLATE_CARRIER_RAIL)\n",
        "\n",
        f"# Stock containers\n",
        f"stock_resources = {tube_resources}\n",
        "# Separate buffer from other resources if using trough\n",
        'if BUFFER_SOURCE == "trough":\n',
        '    trough_car = Trough_CAR_5R60_A00(name="trough_carrier")\n',
        '    for i, resource_name in enumerate(stock_resources):\n',
        '        if i >= 6:\n',
        '            print(f"Warning: Not enough trough positions for {resource_name}")\n',
        '            continue\n',
        '        trough_car[i] = hamilton_1_trough_60ml_Vb(name=resource_name)\n',
        "    lh.deck.assign_child_resource(trough_car, rails=TUBE_CARRIER_RAIL)\n",
        "else:\n",
        '    tube_car = hamilton_tube_carrier_12_b00(name="tube_carrier")\n',
        "    for i, resource_name in enumerate(stock_resources):\n",
        "        if i >= 12:\n",
        '            print(f"Warning: Not enough tube positions for {resource_name}")\n',
        "            continue\n",
        "        tube_car[i] = Cor_Falcon_tube_50mL_Vb(name=resource_name)\n",
        "    lh.deck.assign_child_resource(tube_car, rails=TUBE_CARRIER_RAIL)\n",
        "\n",
        "lh.summary()",
    ]

    notebook["cells"].append(
        {
            "cell_type": "code",
            "execution_count": None,
            "metadata": {},
            "source": deck_code,
        }
    )

    notebook["cells"].append(
        {
            "cell_type": "markdown",
            "metadata": {},
            "source": ["## Helper Functions"],
        }
    )

    helper_functions = [
        "def next_tip_positions(count):\n",
        "    global tip_position_index\n",
        "    if tip_position_index + count > len(TIP_POSITIONS):\n",
        "        raise RuntimeError('Not enough tips available for this run.')\n",
        "    positions = TIP_POSITIONS[tip_position_index:tip_position_index + count]\n",
        "    tip_position_index += count\n",
        "    resources = []\n",
        "    for rack_name, position in positions:\n",
        "        resources.append(lh.deck.get_resource(rack_name)[position][0])\n",
        "    return resources\n",
        "\n",
        "def group_wells_by_column(wells):\n",
        "    columns = {}\n",
        "    for well in wells:\n",
        "        column = ''.join([ch for ch in well if ch.isdigit()])\n",
        "        columns.setdefault(column, []).append(well)\n",
        "    for column in columns:\n",
        "        columns[column].sort()\n",
        "    return columns\n",
        "\n",
        "async def multi_channel_transfer_from_tube(tube_resource_name, plate_name, column_transfers, BUFFER_SOURCE, keep_tips=False):\n",
        "    resource_container = lh.deck.get_resource(tube_resource_name)\n",
        "    plate = lh.deck.get_resource(plate_name)\n",
        "    MAX_TIP_VOLUME = 1000\n",
        "    channel_set = set()\n",
        "    per_channel_total = [0]*8\n",
        "    for column_number, volumes in column_transfers:\n",
        "        channel_indices = [idx for idx, vol in enumerate(volumes) if vol and vol > 0]\n",
        "        channel_set.update(channel_indices)\n",
        "        for idx in channel_indices:\n",
        "            per_channel_total[idx] += volumes[idx]\n",
        "    # Add extra volume to each channel total since plunger movement is non-linear\n"
        "    per_channel_total = [total + 20 for total in per_channel_total]\n",
        "    # Check tip presence from machine\n",
        "    tip_presence = await lh.backend.request_tip_presence()\n",
        "    channels_with_tips = [i for i, has_tip in enumerate(tip_presence) if has_tip == 1]\n",
        "    # Check if we need to discard tips (if required channels don't match existing tips)\n",
        "    if not set(channel_set).issubset(set(channels_with_tips)) or not keep_tips:\n",
        "        if channels_with_tips:\n",
        "           await lh.discard_tips()\n",
        "        tip_resources = next_tip_positions(len(channel_set))\n",
        "        await lh.pick_up_tips(tip_resources, use_channels=list(channel_set))\n",
        "    if BUFFER_SOURCE != 'trough':\n",
        "        for idx in channel_set:\n",
        "            await lh.aspirate([resource_container], vols=[per_channel_total[idx]], use_channels=[idx])\n",
        "    for column_number, volumes in column_transfers:\n",
        "        channels_to_use = []\n",
        "        volumes_filtered = []\n",
        "        for idx, volume in enumerate(volumes):\n",
        "            if volume and volume > 0:\n",
        "                channels_to_use.append(idx)\n",
        "                volumes_filtered.append(volume)\n",
        "        target_positions = [f\"{chr(ord('A') + channel)}{column_number}\" for channel in channels_to_use]\n",
        "        targets_all = [plate[pos][0] for pos in target_positions]\n",
        "        if BUFFER_SOURCE == 'trough':\n",
        "            wide_offsets = [x + Coordinate(0,0,1) for x in get_wide_single_resource_liquid_op_offsets(resource_container,num_channels=8)]\n",
        "            await lh.aspirate([resource_container]*len(channels_to_use), vols=volumes_filtered, use_channels=channels_to_use, offsets=[x for i,x in enumerate(wide_offsets) if i in channels_to_use], spread='custom')\n",
        "        await lh.dispense(targets_all, vols=volumes_filtered, use_channels=channels_to_use, offsets=[Coordinate(0, 0, 5)] * len(channels_to_use))\n",
        "\n",
        "async def multi_channel_serial_dilution(source_plate_name, target_plate_name, source_column, target_column, transfer_volumes, keep_tips=False):\n",
        "    channel_indices = [idx for idx, vol in enumerate(transfer_volumes) if vol and vol > 0]\n",
        "    source_plate = lh.deck.get_resource(source_plate_name)\n",
        "    target_plate = lh.deck.get_resource(target_plate_name)\n",
        "    source_positions = [f\"{chr(ord('A') + idx)}{source_column}\" for idx in channel_indices]\n",
        "    target_positions = [f\"{chr(ord('A') + idx)}{target_column}\" for idx in channel_indices]\n",
        "    # Check tip presence from machine\n",
        "    tip_presence = await lh.backend.request_tip_presence()\n",
        "    channels_with_tips = [i for i, has_tip in enumerate(tip_presence) if has_tip == 1]\n",
        "    # Check if we need to discard tips (if required channels don't match existing tips)\n",
        "    if not set(channel_indices).issubset(set(channels_with_tips)) or not keep_tips:\n",
        "        if channels_with_tips:\n",
        "           await lh.discard_tips()\n",
        "        tip_resources = next_tip_positions(len(channel_indices))\n",
        "        await lh.pick_up_tips(tip_resources, use_channels=list(channel_indices))\n",
        "    source_containers = [source_plate[pos][0] for pos in source_positions]\n",
        "    transfer_list = [transfer_volumes[idx] for idx in channel_indices]\n",
        "    await lh.aspirate(source_containers, vols=transfer_list, use_channels=channel_indices)\n",
        "    target_containers = [target_plate[pos][0] for pos in target_positions]\n",
        "    mix_args = None\n",
        "    if MIX_CYCLES and MIX_CYCLES > 0:\n",
        "        mix_args = [Mix(volume=MIX_VOLUME, repetitions=MIX_CYCLES, flow_rate=MIX_FLOW_RATE) for idx in channel_indices]\n",
        "    await lh.dispense(target_containers, vols=transfer_list, use_channels=channel_indices, mix=mix_args)\n",
        "\n",
        "async def transfer_to_final_plate(target_plate_name, entries, volume):\n",
        '    """Transfer liquid from source plate to target plate"""\n',
        "    # Sort by target well position: column first, then row (top to bottom)\n",
        "    def well_sort_key(entry):\n",
        "        target_well = entry[2]\n",
        "        column_number = int(target_well[1:])\n",
        "        row_letter = target_well[0].upper()\n",
        "        return (column_number, ord(row_letter))\n",
        "    ordered = sorted(entries, key=well_sort_key)\n",
        "    if not ordered:\n",
        "        return\n",
        "    target_plate = lh.deck.get_resource(target_plate_name)\n",
        "    source_containers = [lh.deck.get_resource(source_plate_name)[source_well][0] for source_plate_name, source_well, _ in ordered]\n",
        "    target_containers = [target_plate[target_well][0] for _, _, target_well in ordered]\n",
        "    tips = next_tip_positions(len(ordered))\n",
        "    await lh.pick_up_tips(tips)\n",
        "    if len(set([x[1] for x in ordered])) != len(ordered):\n",
        "        for i, source_well in enumerate(source_containers):\n",
        "            await lh.aspirate([source_well], vols=[volume], use_channels=[i], blow_out_air_volume=[10])\n",
        "    else:\n",
        "        await lh.aspirate(\n",
        "            source_containers,\n",
        "            vols=[volume] * len(ordered),\n",
        "            use_channels=list(range(len(ordered))),\n",
        "            blow_out_air_volume=[20] * len(ordered),\n",
        "        )\n",
        "    await lh.dispense(\n",
        "        target_containers,\n",
        "        vols=[volume] * len(ordered),\n",
        "        use_channels=list(range(len(ordered))),\n",
        "        blow_out_air_volume=[25] * len(ordered),\n",
        "        offsets=[Coordinate(0, 0, 5)] * len(ordered),\n",
        "    )\n",
        "    await lh.discard_tips()\n",
    ]

    notebook["cells"].append(
        {
            "cell_type": "code",
            "execution_count": None,
            "metadata": {},
            "source": helper_functions,
        }
    )

    notebook["cells"].append(
        {
            "cell_type": "markdown",
            "metadata": {},
            "source": ["## Execute Liquid Handling"],
        }
    )


    main_code = ['print("Starting liquid handling...")\n']

    main_code.append("\n")

    sample_dilution_map = {}
    max_dilution_map = {}
    sample_prefill_commands: list[str] = []
    buffer_prefill_plan: Dict[str, Dict[int, Dict[int, float]]] = defaultdict(lambda: defaultdict(dict))
    dilution_plan: Dict[
        Tuple[str, str, int, int], Dict[int, Dict[str, object]]
    ] = defaultdict(dict)
    dilution_logs: list[str] = []
    occupied_dilution_positions: Dict[str, Dict[str, str]] = defaultdict(dict)

    def process_sequence(sample_id, record, stock_resource, sequence, plate_kind):
        if not sequence:
            return

        row_letter = record["row_letter"]
        start_col = record["start_col"]
        offsets = record.setdefault("next_offset", {"assay": 0, "max": 0})
        
        # For combined sequences, use the maximum offset to ensure we don't overlap
        # Use the larger of the two offsets as the base
        offset_base = max(offsets["assay"], offsets["max"])
        
        # Use assay volume for dilution calculations (they should be the same anyway)
        final_volume_ul = ASSAY_DILUTION_VOLUME_UL

        def target_details_for_offset(offset):
            absolute_column = start_col + offset
            plate_index = (absolute_column - 1) // 12
            if plate_index >= len(dilution_plate_names):
                raise ValueError(
                    f"Insufficient dilution plates allocated for sample '{sample_id}'."
                )
            plate_resource_name = dilution_plate_names[plate_index]
            column_within_plate = ((absolute_column - 1) % 12) + 1
            target_well = f"{row_letter}{column_within_plate}"
            return plate_resource_name, target_well

        def concentration_value(occ):
            conc = occ.get("concentration")
            return conc if conc and conc > 0 else None

        def compute_factor(numerator, denominator):
            if numerator and denominator and numerator > 0 and denominator > 0:
                ratio = numerator / denominator
                if ratio > 1:
                    return ratio
            return None

        plan = []
        for idx, occ in enumerate(sequence):
            curr_conc = concentration_value(occ)
            prev_conc = concentration_value(sequence[idx - 1]) if idx > 0 else None
            next_conc = concentration_value(sequence[idx + 1]) if idx + 1 < len(sequence) else None

            factor_in = compute_factor(prev_conc, curr_conc)
            transfer_in = final_volume_ul / factor_in if factor_in and factor_in > 1 else 0.0

            factor_out = compute_factor(curr_conc, next_conc)
            transfer_out = final_volume_ul / factor_out if factor_out and factor_out > 1 else 0.0

            # Calculate base_total: for consistency, use the same calculation for all steps
            # This ensures buffer_prefill is the same when dilution factors are the same
            if idx == len(sequence) - 1:
                # Last step: use same base_total as intermediate steps for consistency
                # Calculate what transfer_out would be if there was a next step with same factor
                if factor_in and factor_in > 1:
                    # Assume same dilution factor continues (for consistency)
                    base_total = final_volume_ul + (final_volume_ul / factor_in)
                else:
                    # No factor available, just use final_volume_ul
                    base_total = final_volume_ul
            else:
                # All other steps: base_total includes reserve for next dilution
                base_total = final_volume_ul + transfer_out
            
            # Calculate buffer_prefill: volume of buffer to add before transfer_in
            # For first step: no buffer needed (stock is loaded directly)
            # For all other steps: buffer_prefill = base_total - transfer_in
            if idx == 0:
                buffer_prefill = 0.0
            else:
                buffer_prefill = base_total - transfer_in
                if buffer_prefill < 0:
                    buffer_prefill = 0.0

            plan.append(
                {
                    "occ": occ,
                    "transfer_in": transfer_in,
                    "transfer_out": transfer_out,
                    "base_total": base_total,
                    "buffer_prefill": buffer_prefill,
                }
            )

        unique_concs = {
            occ.get("concentration")
            for occ in sequence
            if occ.get("concentration") and occ.get("concentration") > 0
        }
        use_single_well = len(unique_concs) <= 1

        for offset_increment, entry in enumerate(plan):
            occ = entry["occ"]
            offset = offset_base if use_single_well else offset_base + offset_increment
            plate_resource_name, target_well = target_details_for_offset(offset)
            well_idx = occ["well_idx"]
            conc = occ.get("concentration", -1.0)
            conc_display = conc if conc and conc > 0 else 0.0

            existing_owner = occupied_dilution_positions[plate_resource_name].get(target_well)
            if existing_owner and existing_owner != sample_id:
                raise ValueError(
                    f"Dilution well conflict: {target_well} on {plate_resource_name} already assigned to {existing_owner}."
                )
            occupied_dilution_positions[plate_resource_name][target_well] = sample_id

            # Map to the appropriate final plate based on the occurrence's plate type
            occ_plate_kind = occ.get("plate")
            if occ_plate_kind == "assay":
                sample_dilution_map[(sample_id, well_idx)] = (plate_resource_name, target_well)
            elif occ_plate_kind == "max":
                max_dilution_map[(sample_id, well_idx)] = (plate_resource_name, target_well)

            liquid_desc = describe_liquid(occ_plate_kind, occ.get("type"))
            transfer_in = entry["transfer_in"]
            transfer_out = entry["transfer_out"]
            base_total = entry["base_total"]
            buffer_prefill = entry["buffer_prefill"]

            if use_single_well and offset_increment > 0:
                dilution_logs.append(
                    f"# {sample_id} reuses {target_well}; no additional buffer or stock required"
                )
                continue

            if offset_increment == 0:
                load_volume = base_total * len(plan) if use_single_well else base_total
                if stock_resource:
                    column_number = int(target_well[1:])
                    row_index = ord(target_well[0]) - ord("A")
                    volumes = [0.0] * 8
                    volumes[row_index] = round(load_volume, 1)
                    transfers_json = json.dumps([[column_number, volumes]])
                    sample_prefill_commands.append(
                        f"await multi_channel_transfer_from_tube('{stock_resource}', '{plate_resource_name}', {transfers_json}, BUFFER_SOURCE)\n"
                    )
                    message = f"Loaded {sample_id} {liquid_desc} into {target_well} ({load_volume:.1f} µL)"
                    if transfer_out > 0 and not use_single_well:
                        message += f"; reserving {transfer_out:.1f} µL for next dilution"
                else:
                    # Use the occurrence's plate kind for the key (occ_plate_kind already defined above)
                    key = (sample_id, well_idx, occ_plate_kind)
                    if key not in manual_positions:
                        manual_positions.add(key)
                        manual_key = "assay" if occ_plate_kind == "assay" else "max"
                        total_volume_mL = load_volume / 1000
                        note = " (includes reserve for next dilution)" if transfer_out > 0 and not use_single_well else ""
                        sample_loading_lines.append(
                            f"- `{sample_id}` → `{plate_resource_name}['{target_well}']` ({total_volume_mL:.2f} mL){note}\n"
                        )
                    message = f"Ensure {sample_id} {liquid_desc} ({load_volume:.1f} µL) is pre-loaded into {plate_resource_name} well {target_well}"
                sample_prefill_commands.append(f"# {message}\n" if not stock_resource else f"print(\"{message}\")\n")
                continue

            if buffer_prefill > 0:
                resource_volume_uL[buffer_resource_name] += buffer_prefill
                column_number = int(target_well[1:])
                row_idx = ord(target_well[0]) - ord("A")
                column_map = buffer_prefill_plan[plate_resource_name]
                row_map = column_map.setdefault(column_number, {})
                row_map[row_idx] = row_map.get(row_idx, 0.0) + buffer_prefill

            if transfer_in > 0:
                source_offset = offset - 1 if not use_single_well else offset
                source_plate_resource_name, source_well = target_details_for_offset(source_offset)
                source_column_number = int(source_well[1:])
                row_idx = ord(target_well[0]) - ord("A")
                key = (
                    source_plate_resource_name,
                    plate_resource_name,
                    source_column_number,
                    column_number,
                )
                row_entries = dilution_plan.setdefault(key, {})
                row_entries[row_idx] = {
                    "transfer_volume": transfer_in,
                    "total_volume": base_total,
                    "sample_id": sample_id,
                    "conc_display": conc_display,
                }
            else:
                dilution_logs.append(
                    f"# No dilution transfer required for {target_well}; volume maintained at {base_total:.1f} µL"
                )

        # Update both offsets to the same value since we're using combined sequences
        new_offset = offset_base + (1 if use_single_well else len(plan))
        record["next_offset"]["assay"] = new_offset
        record["next_offset"]["max"] = new_offset

    def describe_liquid(plate_kind, type_label):
        """Describe liquid type using string label instead of numeric code."""
        if plate_kind == "assay":
            if type_label == "sample":
                return "sample stock"
            if type_label == "load":
                return "load stock"
            return "buffer"
        else:
            if type_label == "sample":
                return "sample stock"
            if type_label == "load":
                return "load stock"
            return "buffer"

    def column_transfers_for_wells(wells, per_well_volume, volume_limit=1000.0):
        column_map: Dict[int, List[float]] = {}
        for well in sorted((w for w in wells if w), key=well_position_sort_key):
            row_idx = ord(well[0]) - ord("A")
            column_number = int(well[1:])
            column_map.setdefault(column_number, [0.0] * 8)[row_idx] = per_well_volume
        transfers: List[List[List[float]]] = []
        current_batch: List[List[float]] = []
        per_channel_totals = [0.0] * 8
        for column in sorted(column_map.keys()):
            volumes = [round(vol, 1) for vol in column_map[column]]
            exceeds = any(per_channel_totals[idx] + volumes[idx] > volume_limit for idx in range(8))
            if exceeds and current_batch:
                transfers.append(current_batch)
                current_batch = []
                per_channel_totals = [0.0] * 8
            for idx in range(8):
                per_channel_totals[idx] += volumes[idx]
            current_batch.append([column, volumes])
        if current_batch:
            transfers.append(current_batch)
        return transfers

    def format_transfer_batch(entries):
        # Sort by target well position: column first, then row (top to bottom)
        ordered = sorted(entries, key=lambda e: well_position_sort_key(e[2]))
        return "[" + ", ".join(
            f"('{source_plate}', '{source_well}', '{target_well}')" for source_plate, source_well, target_well in ordered
        ) + "]"

    for sample_id in combined_sample_ids:
        if sample_id == "":
            continue
        record = combined_samples[sample_id]
        stock_resource = stock_resource_map[sample_id]

        ordered_occurrences = []
        for level in record["levels"]:
            ordered_occurrences.extend(level["occurrences"])
        if len(ordered_occurrences) < len(record["occurrences"]):
            for occ in record["occurrences"]:
                if occ not in ordered_occurrences:
                    ordered_occurrences.append(occ)

        # Combine all occurrences across both plates and sort by concentration (descending)
        # This ensures all dilutions are calculated together for the complete concentration series
        combined_sequence = ordered_occurrences.copy()
        combined_sequence.sort(key=lambda occ: occ.get("concentration", -1.0), reverse=True)
        
        # Only process if there are occurrences to dilute
        if combined_sequence:
            # Use the first plate_kind found for the column assignment (doesn't matter which)
            first_plate_kind = combined_sequence[0]["plate"]
            column_sequences[first_plate_kind][record["start_col"]].append(
                (sample_id, record, stock_resource, combined_sequence)
            )

    for plate_kind in ("assay", "max"):
        columns = column_sequences[plate_kind]
        if not columns:
            continue
        label = "Assay" if plate_kind == "assay" else "Max"
        for column_idx in sorted(columns.keys()):
            dilution_logs.append(f"print(\"Processing {label} column {column_idx}\")")
            for sample_id, record, stock_resource, sequence in columns[column_idx]:
                process_sequence(sample_id, record, stock_resource, sequence, plate_kind)

    if sample_prefill_commands:
        main_code.append('print("Starting sample stock transfers...")\n')
        main_code.extend(sample_prefill_commands)
        main_code.append("\n")
    else:
        main_code.append("# No sample stock transfers required\n\n")

    buffer_batches: list[tuple[str, list[list[object]]]] = []
    if buffer_prefill_plan:
        for plate_resource_name, column_map in sorted(buffer_prefill_plan.items()):
            column_transfers = []
            per_channel_totals = [0.0] * 8
            current_batch: list[list[object]] = []
            for column_number in sorted(column_map.keys()):
                row_map = column_map[column_number]
                volumes_list = [round(float(row_map.get(row_idx, 0.0)), 1) for row_idx in range(8)]
                exceeds_limit = False
                for idx in range(8):
                    if per_channel_totals[idx] + volumes_list[idx] > 1000:
                        exceeds_limit = True
                        break
                if exceeds_limit and current_batch:
                    buffer_batches.append((plate_resource_name, current_batch))
                    current_batch = []
                    per_channel_totals = [0.0] * 8
                for idx in range(8):
                    per_channel_totals[idx] += volumes_list[idx]
                current_batch.append([column_number, volumes_list])
            if current_batch:
                buffer_batches.append((plate_resource_name, current_batch))

    buffer_follow_up = []
    for plate_name, wells, volume in (
        ("final_plate", assay_shared_wells.get("stock_buffer", []), ASSAY_FINAL_VOLUME_UL),
        ("max_plate_final", max_shared_wells.get("stock_buffer", []), MAX_FINAL_VOLUME_UL),
    ):
        for batch in column_transfers_for_wells(wells, volume):
            buffer_follow_up.append((plate_name, batch))

    shared_resources = []
    for resource_map, plate_name, volume in (
        (assay_shared_wells, "final_plate", ASSAY_FINAL_VOLUME_UL),
        (max_shared_wells, "max_plate_final", MAX_FINAL_VOLUME_UL),
    ):
        for resource in sorted(resource_map.keys()):
            wells = [w for w in resource_map[resource] if w]
            if not wells:
                continue
            batches = column_transfers_for_wells(wells, volume)
            shared_resources.extend((resource, plate_name, wells, batch) for batch in batches)

    buffer_total_calls = len(buffer_batches) + len(buffer_follow_up)
    buffer_calls_done = 0

    if buffer_batches:
        main_code.append('print("Prefilling buffer into dilution plate columns...")\n')
        for batch_plate, column_transfers in buffer_batches:
            keep_tips = buffer_calls_done > 0
            transfers_json = json.dumps(column_transfers)
            main_code.append(
                f"await multi_channel_transfer_from_tube('{buffer_resource_name}', '{batch_plate}', {transfers_json}, BUFFER_SOURCE, keep_tips={str(keep_tips)})\n"
            )
            buffer_calls_done += 1
        main_code.append("\n")
    else:
        if buffer_total_calls == 0:
            main_code.append("# No buffer prefill required\n\n")
        else:
            main_code.append("# No dilution-plate buffer prefill required\n\n")

    if buffer_follow_up:
        for plate_name, transfers in buffer_follow_up:
            keep_tips = buffer_calls_done > 0
            main_code.append(f'print("Loading stock_buffer into {plate_name}...")\n')
            transfers_json = json.dumps(transfers)
            main_code.append(
                f"await multi_channel_transfer_from_tube('{buffer_resource_name}', '{plate_name}', {transfers_json}, BUFFER_SOURCE, keep_tips={str(keep_tips)})\n"
            )
            buffer_calls_done += 1
            wells_loaded = sum(1 for _, vols in transfers for vol in vols if vol > 0)
            main_code.append(f'print("stock_buffer: loaded {wells_loaded} wells on {plate_name}")\n')
        main_code.append("\n")

    def load_shared_resource(resource_name, transfers_final_batches, transfers_max_batches):
        tasks = []
        if transfers_final_batches:
            tasks.extend(("final_plate", batch) for batch in transfers_final_batches)
        if transfers_max_batches:
            tasks.extend(("max_plate_final", batch) for batch in transfers_max_batches)
        if not tasks:
            return
        for idx, (plate, batch) in enumerate(tasks):
            keep = idx < len(tasks) - 1
            main_code.append(f'print("Loading {resource_name} into {plate}...")\n')
            transfers_json = json.dumps(batch)
            main_code.append(
                f"await multi_channel_transfer_from_tube('{resource_name}', '{plate}', {transfers_json}, BUFFER_SOURCE, keep_tips={str(keep)})\n"
            )
            wells_loaded = sum(sum(1 for vol in volumes if vol > 0) for _, volumes in batch)
            main_code.append(f'print("{resource_name}: loaded {wells_loaded} wells on {plate}")\n')
        main_code.append("\n")

    load_shared_resource(
        "stock_neutralization",
        column_transfers_for_wells(assay_shared_wells.get("stock_neutralization", []), ASSAY_FINAL_VOLUME_UL),
        column_transfers_for_wells(max_shared_wells.get("stock_neutralization", []), MAX_FINAL_VOLUME_UL),
    )
    load_shared_resource(
        "stock_regeneration",
        column_transfers_for_wells(assay_shared_wells.get("stock_regeneration", []), ASSAY_FINAL_VOLUME_UL),
        column_transfers_for_wells(max_shared_wells.get("stock_regeneration", []), MAX_FINAL_VOLUME_UL),
    )

    if dilution_plan:
        main_code.append('print("Starting serial dilutions...")\n')
        track_labels = []
        for (
            source_plate_resource_name,
            target_plate_resource_name,
            source_column_number,
            target_column_number,
        ), row_map in sorted(dilution_plan.items()):
            transfer_volumes = [round(float(row_map[row_idx]["transfer_volume"]), 1) if row_idx in row_map else 0.0 for row_idx in range(8)]
            sample_labels = [row_map[row_idx]["sample_id"] if row_idx in row_map else "" for row_idx in range(8)]
            if track_labels == sample_labels:
                keep_tips = True
            else:
                keep_tips = False
            track_labels = sample_labels
            transfer_json = json.dumps(transfer_volumes)
            main_code.append(
                f"await multi_channel_serial_dilution('{source_plate_resource_name}', '{target_plate_resource_name}', {source_column_number}, {target_column_number}, {transfer_json}, keep_tips={str(keep_tips)})\n"
            )
        main_code.append(
            "await lh.discard_tips()\n\n"
        )
    else:
        main_code.append("# No serial dilutions required\n\n")

    volume_resources = sorted(resource_volume_uL.keys())
    volume_summary_lines = [
        "## Stock Volume Requirements\n",
        "Estimated minimum volumes to load in each 50 mL tube (add extra for dead volume):\n",
    ]
    for resource in volume_resources:
        volume_uL = resource_volume_uL.get(resource, 0.0)
        volume_summary_lines.append(f"- `{resource}`: {volume_uL/1000:.2f} mL (≈ {volume_uL:.0f} µL)\n")
    if buffer_resource_name in volume_resources:
        volume_summary_lines.append(
            "\nProbe wells buffer usage is included in the `stock_buffer` total.\n"
        )

    notebook["cells"][summary_cell_index]["source"] = volume_summary_lines

    if len(sample_loading_lines) == 1:
        sample_loading_lines.append(
            "Robot will automatically load assay and Max Plate wells from stock tubes. No manual loading required.\n"
        )

    notebook["cells"][sample_instructions_cell_index]["source"] = sample_loading_lines

    if sample_dilution_map:
        main_code.append('print("Transferring assay dilutions to final plate...")\n')
        # Collect all entries and sort by target well position
        all_entries: List[tuple[str, str, str]] = []
        for (sample_id, well_idx), (source_plate, source_well) in sample_dilution_map.items():
            final_well = assay_plate[well_idx].get("WellPosition")
            if not final_well:
                raise ValueError(f"Missing WellPosition for assay well index {well_idx}")
            all_entries.append((source_plate, source_well, final_well))
        
        # Sort by target well position (column number, then row letter)
        all_entries.sort(key=lambda e: well_position_sort_key(e[2]))
        
        # Batch into groups of 8
        batch_entries: List[tuple[str, str, str]] = []
        for entry in all_entries:
            batch_entries.append(entry)
            if len(batch_entries) == 8:
                batch_literal = format_transfer_batch(batch_entries)
                main_code.append(
                    f"await transfer_to_final_plate('final_plate', {batch_literal}, FINAL_VOLUME)\n"
                )
                batch_entries = []
        if batch_entries:
            batch_literal = format_transfer_batch(batch_entries)
            main_code.append(
                f"await transfer_to_final_plate('final_plate', {batch_literal}, FINAL_VOLUME)\n"
            )
        main_code.append("\n")

    if max_dilution_map:
        main_code.append('\nprint("Transferring Max Plate dilutions to final plate...")\n')
        # Collect all entries and sort by target well position
        all_entries: List[tuple[str, str, str]] = []
        for (sample_id, well_idx), (source_plate, source_well) in max_dilution_map.items():
            final_well = max_plate[well_idx].get("WellPosition")
            if not final_well:
                raise ValueError(f"Missing WellPosition for Max Plate well index {well_idx}")
            all_entries.append((source_plate, source_well, final_well))
        
        # Sort by target well position (column number, then row letter)
        all_entries.sort(key=lambda e: well_position_sort_key(e[2]))
        
        # Batch into groups of 8
        batch_entries: List[tuple[str, str, str]] = []
        for entry in all_entries:
            batch_entries.append(entry)
            if len(batch_entries) == 8:
                batch_literal = format_transfer_batch(batch_entries)
                main_code.append(
                    f"await transfer_to_final_plate('max_plate_final', {batch_literal}, MAX_PLATE_FINAL_VOLUME)\n"
                )
                batch_entries = []
        if batch_entries:
            batch_literal = format_transfer_batch(batch_entries)
            main_code.append(
                f"await transfer_to_final_plate('max_plate_final', {batch_literal}, MAX_PLATE_FINAL_VOLUME)\n"
            )
        main_code.append("\n")

    main_code.append('\nprint("Liquid handling complete!")')

    notebook["cells"].append(
        {
            "cell_type": "code",
            "execution_count": None,
            "metadata": {},
            "source": main_code,
        }
    )

    if output_path is None:
        output_path = excel_path.parent / f"{excel_path.stem}_liquid_handler.ipynb"
    else:
        output_path = Path(output_path)

    print(f"Generating notebook: {output_path}")
    with open(output_path, "w", encoding="utf-8") as fh:
        json.dump(notebook, fh, indent=2, ensure_ascii=False)

    print(f"Successfully generated: {output_path}")
    return output_path


__all__ = ["generate_liquid_handler_notebook"]

