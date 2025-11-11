from __future__ import annotations

import json
from collections import defaultdict
from pathlib import Path
from typing import Dict, Optional, List, Tuple
from math import ceil, isclose
from datetime import datetime

from .common import SampleType, sanitize_identifier
from .excel_parser import parse_plate_layout, read_excel_file


def generate_liquid_handler_notebook(
    excel_path: Path | str,
    output_path: Path | str | None = None,
    initial_concentrations: Optional[Dict[str, float]] = None,
) -> Path:
    """Generate a Jupyter notebook for liquid handling using pylabrobot."""

    excel_path = Path(excel_path)
    if not excel_path.exists():
        raise FileNotFoundError(f"Excel file not found: {excel_path}")

    print(f"Reading Excel file: {excel_path}")
    sheets = read_excel_file(excel_path)

    if "Experiment" not in sheets:
        raise ValueError("Experiment sheet not found in Excel file")

    samples, probe_info = parse_plate_layout(sheets["Experiment"])

    if initial_concentrations is None:
        initial_concentrations = {}

    ASSAY_DILUTION_VOLUME_UL = 300
    MAX_DILUTION_VOLUME_UL = 300
    ASSAY_FINAL_VOLUME_UL = 200
    MAX_FINAL_VOLUME_UL = 200

    buffer_resource_name = "stock_buffer"

    ASSAY_SHARED_RESOURCE_MAP = {
        SampleType.Assay.Buffer: "stock_buffer",
        SampleType.Assay.Regeneration: "stock_regeneration",
        SampleType.Assay.Neutralization: "stock_neutralization",
        SampleType.Assay.Activation: "stock_buffer",
        SampleType.Assay.Quench: "stock_buffer",
        SampleType.Assay.Wash: "stock_buffer",
    }

    MAX_SHARED_RESOURCE_MAP = {
        SampleType.MaxPlate.Probe: "stock_buffer",
        SampleType.MaxPlate.Buffer: "stock_buffer",
        SampleType.MaxPlate.Regeneration: "stock_regeneration",
        SampleType.MaxPlate.Neutralization: "stock_neutralization",
        SampleType.MaxPlate.Activation: "stock_buffer",
        SampleType.MaxPlate.Quench: "stock_buffer",
        SampleType.MaxPlate.Wash: "stock_buffer",
    }

    combined_liquids: Dict[str, dict] = {}

    def register_entry(entry, well_idx, plate_kind):
        sample_id = str(entry.get("SampleID") or "").strip()
        concentration = entry.get("Concentration", -1.0)
        sample_type = entry.get("Type", SampleType.Assay.Buffer if plate_kind == "assay" else SampleType.MaxPlate.Buffer)

        shared_map = ASSAY_SHARED_RESOURCE_MAP if plate_kind == "assay" else MAX_SHARED_RESOURCE_MAP
        if sample_type in shared_map:
            return

        if not sample_id or sample_id.upper() == "N/A":
            return

        record = combined_liquids.setdefault(
            sample_id,
            {
                "occurrences": [],
                "max_conc": -1.0,
                "has_assay": False,
                "has_max": False,
                "levels": [],
                "avg_dilution_factor": None,
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
                "type": sample_type,
            }
        )

        if concentration is not None and concentration > 0:
            record["max_conc"] = max(record["max_conc"], concentration)

    for well_idx, sample in enumerate(samples):
        register_entry(sample, well_idx, "assay")

    for well_idx, probe in enumerate(probe_info):
        register_entry(probe, well_idx, "max")

    for sample_id, record in combined_liquids.items():
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

        positive_levels = [level["value"] for level in record["levels"] if level["value"] > 0]
        dilution_factors = []
        for idx_level in range(1, len(positive_levels)):
            prev_val = positive_levels[idx_level - 1]
            curr_val = positive_levels[idx_level]
            if prev_val > 0 and curr_val > 0:
                dilution_factors.append(prev_val / curr_val)

        if dilution_factors:
            record["avg_dilution_factor"] = sum(dilution_factors) / len(dilution_factors)

        if record["max_conc"] <= 0:
            record["max_conc"] = 1.0

    combined_sample_ids = sorted(combined_liquids.keys())
    assay_ids = [sid for sid in combined_sample_ids if combined_liquids[sid]["has_assay"]]
    max_ids = [sid for sid in combined_sample_ids if combined_liquids[sid]["has_max"]]

    row_usage = [1] * 8  # next available column per row (1-indexed)
    max_column_used = 1
    MAX_DILUTION_PLATES = 2
    MAX_COLUMNS_PER_ROW = MAX_DILUTION_PLATES * 12
    current_row_pointer = 0

    for sample_id in combined_sample_ids:
        record = combined_liquids[sample_id]
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
        record = combined_liquids[sample_id]
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
        record = combined_liquids[sample_id]
        safe_name = sanitize_identifier(sample_id, prefix="P")
        var_name = f"MAX_{safe_name.upper()}"
        default_conc = record["max_conc"]
        if safe_name.upper() != "PROBE":
            init_max_conc_vars.append(
                f"INITIAL_CONC_{var_name} = {default_conc}  # µg/mL - Stock concentration for Max Plate {sample_id}"
            )

    sample_loading_lines = ["## Sample Stock Placement\n"]
    manual_positions = {"assay": set(), "max": set()}

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
    for idx, sample in enumerate(samples):
        sample_type = sample.get("Type", SampleType.Assay.EMPTY)
        resource = ASSAY_SHARED_RESOURCE_MAP.get(sample_type)
        if resource:
            well_position = sample.get("WellPosition")
            if not well_position:
                raise ValueError(f"Missing WellPosition for assay sample index {idx}")
            assay_shared_wells[resource].append(well_position)
            resource_volume_uL[resource] += ASSAY_FINAL_VOLUME_UL

    max_shared_wells = defaultdict(list)
    for idx, probe in enumerate(probe_info):
        probe_type = probe.get("Type", SampleType.MaxPlate.EMPTY)
        resource = MAX_SHARED_RESOURCE_MAP.get(probe_type)
        if resource:
            well_position = probe.get("WellPosition")
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
                f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n",
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
        "# Mixing and flow parameters\n",
        "MIX_CYCLES = 3\n",
        "FLOW_RATE = 100  # µL/sec\n",
        "CHANGE_TIPS_BETWEEN_DILUTIONS = True\n",
        "\n",
        "# Assay initial stock concentrations (µg/mL)\n",
    ] + (init_conc_vars if init_conc_vars else ["# None detected\n"])

    variables_code += ["\n", "# Max Plate initial stock concentrations (µg/mL)\n"] + (
        init_max_conc_vars if init_max_conc_vars else ["# None detected\n"]
    )

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
        "    nest_96_wellplate_2mL_Vb,\n",
        "    hamilton_96_tiprack_1000uL,\n",
        "    Tube_CAR_32_A00,\n",
        "    hamilton_tube_carrier_12_b00,\n",
        "    Cor_Falcon_tube_50mL_Vb,\n",
        ")\n",
        "from pylabrobot.liquid_handling.standard import Mix\n",
        "\n",
        'lh.deck.get_resource("trash_core96").location = Coordinate(-260, 106, 216.4)\n',
        "\n",
        "# Tips\n",
        'tip_car = TIP_CAR_480_A00(name="tip_carrier")\n',
        'tip_car[0] = hamilton_96_tiprack_1000uL(name="main_tips")\n',
        "lh.deck.assign_child_resource(tip_car, rails=7)\n",
        "\n",
        "TIP_POSITIONS = [\n",
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
        "tip_position_index = 0\n",
        "\n",
        "# Plates\n",
        f"dilution_plate_names = {dilution_plate_names_repr}\n",
        'plt_car = PLT_CAR_L5AC_A00(name="plate_carrier")\n',
        'plt_car[0] = Cor_96_wellplate_360ul_Fb(name="final_plate")  # Assay plate\n',
        'plt_car[1] = nest_96_wellplate_2mL_Vb(name=dilution_plate_names[0])  # Shared dilutions\n',
        'plt_car[2] = Cor_96_wellplate_360ul_Fb(name="max_plate_final")  # Max Plate\n',
        "if len(dilution_plate_names) > 1:\n",
        '    plt_car[3] = nest_96_wellplate_2mL_Vb(name=dilution_plate_names[1])  # Additional dilutions\n',
        "lh.deck.assign_child_resource(plt_car, rails=1)\n",
        "\n",
        "# 50 mL stock tubes\n",
        f"stock_resources = {tube_resources}\n",
        'tube_car = hamilton_tube_carrier_12_b00(name="tube_carrier")\n',
        "for i, resource_name in enumerate(stock_resources):\n",
        "    if i >= 12:\n",
        '        print(f"Warning: Not enough tube positions for {resource_name}")\n',
        "        continue\n",
        "    tube_car[11 - i] = Cor_Falcon_tube_50mL_Vb(name=resource_name)\n",
        "lh.deck.assign_child_resource(tube_car, rails=35)\n",
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
        "    return positions\n",
        "\n",
        "def get_tip_resources(positions):\n",
        "    return [lh.deck.get_resource('main_tips')[pos] for pos in positions]\n",
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
        "async def serial_dilution(source_plate_name, target_plate_name, source_well, target_well, transfer_volume, final_volume, mix_cycles=MIX_CYCLES):\n",
        '    """Transfer an aliquot from source to target and mix for serial dilution."""\n',
        "    source_plate = lh.deck.get_resource(source_plate_name)\n",
        "    target_plate = lh.deck.get_resource(target_plate_name)\n",
        '    tip_positions = next_tip_positions(1)\n',
        "    tips = get_tip_resources(tip_positions)\n",
        "    await lh.pick_up_tips(tips)\n",
        "    await lh.aspirate([source_plate[source_well]], vols=[transfer_volume], flow_rates=[FLOW_RATE])\n",
        "    await lh.dispense([target_plate[target_well]], vols=[transfer_volume], flow_rates=[FLOW_RATE])\n",
        "    for _ in range(mix_cycles):\n",
        "        mix_volume = min(final_volume * 0.8, final_volume)\n",
        "        await lh.aspirate([target_plate[target_well]], vols=[mix_volume], flow_rates=[FLOW_RATE])\n",
        "        await lh.dispense([target_plate[target_well]], vols=[mix_volume], flow_rates=[FLOW_RATE], offsets=[Coordinate(0, 0, 5)])\n",
        "    await lh.drop_tips([lh.deck.get_resource('trash')]*len(tips))\n",
        "\n",
        "async def transfer_from_tube_to_plate(tube_resource_name, plate_name, target_well, volume):\n",
        '    """Transfer liquid from a 50 mL tube to a plate well"""\n',
        '    tube_resource = lh.deck.get_resource(tube_resource_name)\n',
        "    plate = lh.deck.get_resource(plate_name)\n",
        "    tip_positions = next_tip_positions(1)\n",
        "    tips = get_tip_resources(tip_positions)\n",
        "    await lh.pick_up_tips(tips)\n",
        "    await lh.aspirate([tube_resource], vols=[volume], flow_rates=[FLOW_RATE])\n",
        "    await lh.dispense([plate[target_well]], vols=[volume], flow_rates=[FLOW_RATE])\n",
        "    await lh.drop_tips([lh.deck.get_resource('trash')]*len(tips))\n",
        "\n",
        "async def load_plate_columns_from_stock(stock_resource_name, plate_name, wells, volume_per_well):\n",
        "    if not wells:\n",
        "        return\n",
        "    wells_by_column = group_wells_by_column(wells)\n",
        "    for column in sorted(wells_by_column.keys(), key=lambda c: int(c)):\n",
        "        column_wells = wells_by_column[column]\n",
        "        if len(column_wells) == 8:\n",
        "            volumes = [volume_per_well] * 8\n",
        "            await multi_channel_transfer_from_tube(stock_resource_name, plate_name, int(column), volumes)\n",
        "        else:\n",
        "            for well in column_wells:\n",
        "                await transfer_from_tube_to_plate(stock_resource_name, plate_name, well, volume_per_well)\n",
        "\n",
        "def build_channel_indices(volumes):\n",
        "    return [idx for idx, vol in enumerate(volumes) if vol and vol > 0]\n",
        "\n",
        "async def multi_channel_transfer_from_tube(tube_resource_name, plate_name, column_number, volumes, change_tips=True):\n",
        "    channel_indices = build_channel_indices(volumes)\n",
        "    if not channel_indices:\n",
        "        return\n",
        "    tube = lh.deck.get_resource(tube_resource_name)\n",
        "    plate = lh.deck.get_resource(plate_name)\n",
        "    tip_positions = next_tip_positions(len(channel_indices))\n",
        "    tip_resources = get_tip_resources(tip_positions)\n",
        "    await lh.pick_up_tips(tip_resources)\n",
        "    source_containers = [tube for _ in channel_indices]\n",
        "    volumes_list = [volumes[idx] for idx in channel_indices]\n",
        "    await lh.aspirate(source_containers, vols=volumes_list, use_channels=channel_indices, flow_rates=[FLOW_RATE] * len(channel_indices))\n",
        "    target_positions = [f\"{chr(ord('A') + idx)}{column_number}\" for idx in channel_indices]\n",
        "    targets = [plate[pos] for pos in target_positions]\n",
        "    await lh.dispense(targets, vols=volumes_list, use_channels=channel_indices, flow_rates=[FLOW_RATE] * len(channel_indices))\n",
        "    await lh.drop_tips([lh.deck.get_resource('trash')]*len(channel_indices))\n",
        "\n",
        "async def multi_channel_serial_dilution(source_plate_name, target_plate_name, source_column, target_column, transfer_volumes, total_volumes, sample_labels, change_tips, reuse_state):\n",
        "    channel_indices = build_channel_indices(transfer_volumes)\n",
        "    if not channel_indices:\n",
        "        return\n",
        "    source_plate = lh.deck.get_resource(source_plate_name)\n",
        "    target_plate = lh.deck.get_resource(target_plate_name)\n",
        "    source_positions = [f\"{chr(ord('A') + idx)}{source_column}\" for idx in channel_indices]\n",
        "    target_positions = [f\"{chr(ord('A') + idx)}{target_column}\" for idx in channel_indices]\n",
        "    if change_tips or not reuse_state.get('loaded'):\n",
        "        tip_positions = next_tip_positions(len(channel_indices))\n",
        "        tip_resources = get_tip_resources(tip_positions)\n",
        "        await lh.pick_up_tips(tip_resources)\n",
        "        reuse_state['loaded'] = True\n",
        "        reuse_state['tip_resources'] = tip_resources\n",
        "    else:\n",
        "        tip_resources = reuse_state['tip_resources']\n",
        "    source_containers = [source_plate[pos] for pos in source_positions]\n",
        "    transfer_list = [transfer_volumes[idx] for idx in channel_indices]\n",
        "    await lh.aspirate(source_containers, vols=transfer_list, use_channels=channel_indices, flow_rates=[FLOW_RATE] * len(channel_indices))\n",
        "    target_containers = [target_plate[pos] for pos in target_positions]\n",
        "    mix_args = None\n",
        "    if MIX_CYCLES and MIX_CYCLES > 0:\n",
        "        mix_args = []\n",
        "        for idx, channel_idx in enumerate(channel_indices):\n",
        "            channel_total = total_volumes[channel_idx] or transfer_list[idx]\n",
        "            mix_volume = channel_total * 0.8 if channel_total else transfer_list[idx]\n",
        "            mix_args.append(Mix(volume=mix_volume, repetitions=MIX_CYCLES, flow_rate=FLOW_RATE))\n",
        "    await lh.dispense(target_containers, vols=transfer_list, use_channels=channel_indices, flow_rates=[FLOW_RATE] * len(channel_indices), mix=mix_args)\n",
        "    for idx, channel_idx in enumerate(channel_indices):\n",
        "        label = sample_labels[channel_idx]\n",
        "        if label:\n",
        "            print(f\"{label}: column {source_column}->{target_column} dilution complete\")\n",
        "    if change_tips:\n",
        "        await lh.drop_tips([lh.deck.get_resource('trash')]*len(channel_indices))\n",
        "        reuse_state['loaded'] = False\n",
        "        reuse_state['tip_resources'] = []\n",
        "    else:\n",
        "        reuse_state['tip_resources'] = tip_resources\n",
        "\n",
        "async def transfer_to_final_plate(source_plate_name, target_plate_name, source_well, target_well, volume):\n",
        '    """Transfer liquid from source plate to target plate"""\n',
        "    source_plate = lh.deck.get_resource(source_plate_name)\n",
        "    target_plate = lh.deck.get_resource(target_plate_name)\n",
        "    tip_positions = next_tip_positions(1)\n",
        "    tips = get_tip_resources(tip_positions)\n",
        "    await lh.pick_up_tips(tips)\n",
        "    await lh.aspirate([source_plate[source_well]], vols=[volume], flow_rates=[FLOW_RATE])\n",
        "    await lh.dispense([target_plate[target_well]], vols=[volume], flow_rates=[FLOW_RATE])\n",
        "    await lh.drop_tips([lh.deck.get_resource('trash')]*len(tips))\n",
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

    has_assay_liquids = bool(assay_ids)
    has_max_liquids = bool(max_ids)

    main_code = ['print("Starting liquid handling...")\n']
    if has_assay_liquids:
        main_code.append('print("Preparing assay plate dilutions...")\n')
    else:
        main_code.append("# No assay plate dilutions detected\n")

    if has_max_liquids:
        main_code.append('\nprint("Preparing Max Plate dilutions...")\n')
    else:
        main_code.append("\n# No Max Plate liquids detected\n")

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
        offset_base = offsets[plate_kind]
        final_volume_ul = ASSAY_DILUTION_VOLUME_UL if plate_kind == "assay" else MAX_DILUTION_VOLUME_UL
        map_dict = sample_dilution_map if plate_kind == "assay" else max_dilution_map
        manual_key = "assay" if plate_kind == "assay" else "max"
        dict_name = "dilution_well_map_samples" if plate_kind == "assay" else "dilution_well_map_max"

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

        avg_factor = record.get("avg_dilution_factor")
        plan = []
        for idx, occ in enumerate(sequence):
            curr_conc = concentration_value(occ)
            prev_conc = concentration_value(sequence[idx - 1]) if idx > 0 else None
            next_conc = concentration_value(sequence[idx + 1]) if idx + 1 < len(sequence) else None

            factor_in = compute_factor(prev_conc, curr_conc)
            used_avg_in = False
            if factor_in is None and avg_factor and avg_factor > 1:
                factor_in = avg_factor
                used_avg_in = True
            transfer_in = final_volume_ul / factor_in if factor_in and factor_in > 1 else 0.0

            factor_out = compute_factor(curr_conc, next_conc)
            used_avg_out = False
            if factor_out is None and avg_factor and avg_factor > 1:
                factor_out = avg_factor
                used_avg_out = True
            transfer_out = final_volume_ul / factor_out if factor_out and factor_out > 1 else 0.0

            base_total = final_volume_ul + transfer_out
            buffer_prefill = base_total - transfer_in if idx > 0 else 0.0
            if buffer_prefill < 0:
                buffer_prefill = 0.0

            plan.append(
                {
                    "occ": occ,
                    "transfer_in": transfer_in,
                    "transfer_out": transfer_out,
                    "base_total": base_total,
                    "buffer_prefill": buffer_prefill,
                    "used_avg_in": used_avg_in,
                    "used_avg_out": used_avg_out,
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

            map_dict[(sample_id, well_idx)] = (plate_resource_name, target_well)

            liquid_desc = describe_liquid(plate_kind, occ.get("type"))
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
                    sample_prefill_commands.append(
                        f"await transfer_from_tube_to_plate('{stock_resource}', '{plate_resource_name}', '{target_well}', {load_volume:.1f})\n"
                    )
                    message = f"Loaded {sample_id} {liquid_desc} into {target_well} ({load_volume:.1f} µL)"
                    if transfer_out > 0 and not use_single_well:
                        message += f"; reserving {transfer_out:.1f} µL for next dilution"
                else:
                    key = (sample_id, well_idx)
                    if key not in manual_positions[manual_key]:
                        manual_positions[manual_key].add(key)
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
                    "used_avg": entry["used_avg_in"] and avg_factor and avg_factor > 1,
                }
            else:
                dilution_logs.append(
                    f"# No dilution transfer required for {target_well}; volume maintained at {base_total:.1f} µL"
                )

        record["next_offset"][plate_kind] = offset_base + (1 if use_single_well else len(plan))

    def describe_liquid(plate_kind, type_code):
        if plate_kind == "assay":
            if type_code == SampleType.Assay.Sample:
                return "sample stock"
            if type_code == SampleType.Assay.Load:
                return "load stock"
            return "buffer"
        else:
            if type_code == SampleType.MaxPlate.Sample:
                return "sample stock"
            if type_code == SampleType.MaxPlate.Load:
                return "load stock"
            if type_code == SampleType.MaxPlate.Probe:
                return "probe stock"
            return "buffer"

    for sample_id in combined_sample_ids:
        record = combined_liquids[sample_id]
        stock_resource = stock_resource_map[sample_id]

        ordered_occurrences = []
        for level in record["levels"]:
            ordered_occurrences.extend(level["occurrences"])
        if len(ordered_occurrences) < len(record["occurrences"]):
            for occ in record["occurrences"]:
                if occ not in ordered_occurrences:
                    ordered_occurrences.append(occ)

        assay_sequence = [occ for occ in ordered_occurrences if occ["plate"] == "assay"]
        max_sequence = [occ for occ in ordered_occurrences if occ["plate"] == "max"]

        if assay_sequence:
            column_sequences["assay"][record["start_col"]].append(
                (sample_id, record, stock_resource, assay_sequence)
            )
        if max_sequence:
            column_sequences["max"][record["start_col"]].append(
                (sample_id, record, stock_resource, max_sequence)
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

    if buffer_prefill_plan:
        main_code.append('print("Prefilling buffer into dilution plate columns...")\n')
        for plate_resource_name, column_map in sorted(buffer_prefill_plan.items()):
            for column_number in sorted(column_map.keys()):
                row_map = column_map[column_number]
                volumes_list = [round(float(row_map.get(row_idx, 0.0)), 1) for row_idx in range(8)]
                volumes_json = json.dumps(volumes_list)
                main_code.append(
                    f"await multi_channel_transfer_from_tube('{buffer_resource_name}', '{plate_resource_name}', {column_number}, {volumes_json})\n"
                )
        main_code.append("\n")
    else:
        main_code.append("# No buffer prefill required\n\n")

    if dilution_plan:
        main_code.append('print("Starting serial dilutions...")\n')
        main_code.append('tip_reuse_state = {"loaded": False}\n')
        for (
            source_plate_resource_name,
            target_plate_resource_name,
            source_column_number,
            target_column_number,
        ), row_map in sorted(dilution_plan.items()):
            transfer_volumes = [round(float(row_map[row_idx]["transfer_volume"]), 1) if row_idx in row_map else 0.0 for row_idx in range(8)]
            total_volumes = [round(float(row_map[row_idx]["total_volume"]), 1) if row_idx in row_map else 0.0 for row_idx in range(8)]
            sample_labels = [row_map[row_idx]["sample_id"] if row_idx in row_map else "" for row_idx in range(8)]
            transfer_json = json.dumps(transfer_volumes)
            total_json = json.dumps(total_volumes)
            labels_json = json.dumps(sample_labels)
            main_code.append(
                f"await multi_channel_serial_dilution('{source_plate_resource_name}', '{target_plate_resource_name}', {source_column_number}, {target_column_number}, {transfer_json}, {total_json}, {labels_json}, CHANGE_TIPS_BETWEEN_DILUTIONS, tip_reuse_state)\n"
            )
        main_code.append(
            "if not CHANGE_TIPS_BETWEEN_DILUTIONS and tip_reuse_state['loaded']:\n    await lh.drop_tips([lh.deck.get_resource('trash')] * len(tip_reuse_state['tip_resources']))\n    tip_reuse_state['loaded'] = False\n    tip_reuse_state['tip_resources'] = []\n\n"
        )
        main_code.append("\n")
    else:
        main_code.append("# No serial dilutions required\n\n")

    for log_line in dilution_logs:
        main_code.append(log_line + "\n")

    if assay_shared_wells:
        main_code.append('print("Loading shared reagents into assay final plate...")\n')
        for resource in sorted(assay_shared_wells.keys()):
            wells = sorted(set(assay_shared_wells[resource]), key=well_position_sort_key)
            wells_literal = json.dumps(wells)
            main_code.append(
                f"await load_plate_columns_from_stock('{resource}', 'final_plate', {wells_literal}, FINAL_VOLUME)\n"
            )
            main_code.append(f'print("{resource}: loaded {len(wells)} assay wells")\n')
        main_code.append("\n")

    if max_shared_wells:
        main_code.append('print("Loading shared reagents into Max Plate...")\n')
        for resource in sorted(max_shared_wells.keys()):
            wells = sorted(set(max_shared_wells[resource]), key=well_position_sort_key)
            wells_literal = json.dumps(wells)
            main_code.append(
                f"await load_plate_columns_from_stock('{resource}', 'max_plate_final', {wells_literal}, MAX_PLATE_FINAL_VOLUME)\n"
            )
            main_code.append(f'print("{resource}: loaded {len(wells)} Max Plate wells")\n')
        main_code.append("\n")

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
        for (sample_id, well_idx), (source_plate, source_well) in sorted(
            sample_dilution_map.items(), key=lambda x: (x[0][0], x[0][1])
        ):
            final_well = samples[well_idx].get("WellPosition")
            if not final_well:
                raise ValueError(f"Missing WellPosition for assay well index {well_idx}")
            main_code.append(
                f"await transfer_to_final_plate('{source_plate}', 'final_plate', '{source_well}', '{final_well}', FINAL_VOLUME)  # {sample_id}\n"
            )

    if max_dilution_map:
        main_code.append('\nprint("Transferring Max Plate dilutions to final plate...")\n')
        for (sample_id, well_idx), (source_plate, source_well) in sorted(
            max_dilution_map.items(), key=lambda x: (x[0][0], x[0][1])
        ):
            final_well = probe_info[well_idx].get("WellPosition")
            if not final_well:
                raise ValueError(f"Missing WellPosition for Max Plate well index {well_idx}")
            main_code.append(
                f"await transfer_to_final_plate('{source_plate}', 'max_plate_final', '{source_well}', '{final_well}', MAX_PLATE_FINAL_VOLUME)  # {sample_id}\n"
            )

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

