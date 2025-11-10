from __future__ import annotations

"""
Shared helpers and constants used across the GatorBio Excel conversion tools.
"""

class SampleType:
    """Sample type constants for both assay and max plates."""

    class Assay:
        EMPTY = 0
        Sample = 1  # Analyte samples
        Buffer = 4  # Wash/Baseline Buffer
        Load = 5  # Load on probe (ligand)
        Regeneration = 6  # Regenerate probes
        Neutralization = 7  # Neutralize probes
        Activation = 8  # Activation
        Quench = 9  # Quench
        Wash = 10  # Wash

    class MaxPlate:
        EMPTY = 0
        Probe = 1  # Probe sensor
        Sample = 2  # Analyte samples
        Buffer = 5  # Wash/Baseline Buffer
        Load = 6  # Load on probe (ligand)
        Regeneration = 7  # Regenerate probes
        Neutralization = 8  # Neutralize probes
        Activation = 9  # Activation
        Quench = 10  # Quench
        Wash = 11  # Wash


_ASSAY_LABEL_MAP = {
    "": SampleType.Assay.Buffer,
    "empty": SampleType.Assay.EMPTY,
    "buffer": SampleType.Assay.Buffer,
    "baseline": SampleType.Assay.Buffer,
    "sample": SampleType.Assay.Sample,
    "analyte": SampleType.Assay.Sample,
    "load": SampleType.Assay.Load,
    "probe": SampleType.Assay.Load,
    "regeneration": SampleType.Assay.Regeneration,
    "regen": SampleType.Assay.Regeneration,
    "neutralization": SampleType.Assay.Neutralization,
    "activation": SampleType.Assay.Activation,
    "quench": SampleType.Assay.Quench,
    "wash": SampleType.Assay.Wash,
    "background": SampleType.Assay.Buffer,
    "negative": SampleType.Assay.Regeneration,
    "reference": SampleType.Assay.Neutralization,
}

_MAX_LABEL_MAP = {
    "": SampleType.MaxPlate.Buffer,
    "empty": SampleType.MaxPlate.EMPTY,
    "probe": SampleType.MaxPlate.Probe,
    "sensor": SampleType.MaxPlate.Probe,
    "sample": SampleType.MaxPlate.Sample,
    "analyte": SampleType.MaxPlate.Sample,
    "buffer": SampleType.MaxPlate.Buffer,
    "baseline": SampleType.MaxPlate.Buffer,
    "load": SampleType.MaxPlate.Load,
    "regeneration": SampleType.MaxPlate.Regeneration,
    "regen": SampleType.MaxPlate.Regeneration,
    "neutralization": SampleType.MaxPlate.Neutralization,
    "activation": SampleType.MaxPlate.Activation,
    "quench": SampleType.MaxPlate.Quench,
    "wash": SampleType.MaxPlate.Wash,
}


def map_assay_label_to_code(label: object) -> int:
    if label is None:
        return SampleType.Assay.Buffer
    label_str = str(label).strip().lower()
    for key, value in _ASSAY_LABEL_MAP.items():
        if key and key in label_str:
            return value
    return _ASSAY_LABEL_MAP.get(label_str, SampleType.Assay.Buffer)


def map_max_plate_label_to_code(label: object) -> int:
    if label is None:
        return SampleType.MaxPlate.Buffer
    label_str = str(label).strip().lower()
    for key, value in _MAX_LABEL_MAP.items():
        if key and key in label_str:
            return value
    return _MAX_LABEL_MAP.get(label_str, SampleType.MaxPlate.Buffer)


def sanitize_identifier(name: object, prefix: str = "ID") -> str:
    """Sanitize a string to be used in variable/resource names and identifiers."""

    if not name:
        return prefix

    safe = "".join(ch if str(ch).isalnum() else "_" for ch in str(name).strip())
    while "__" in safe:
        safe = safe.replace("__", "_")
    safe = safe.strip("_")
    if not safe:
        safe = prefix
    if safe[0].isdigit():
        safe = f"{prefix}_{safe}"
    return safe


__all__ = [
    "SampleType",
    "map_assay_label_to_code",
    "map_max_plate_label_to_code",
    "sanitize_identifier",
]

