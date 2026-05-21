"""
Per-sheet configuration for PPT table generation.

Add sheet names here (lowercase) to override default table widths,
margins, or column sizing behavior.

Usage in slides.py:
    from ppt_generation.sheet_config import get_sheet_config
    config = get_sheet_config(sheet_name)
    table_width_emu = config["table_width_emu"] or <default calculation>
"""

from pptx.util import Inches


# ─── Default table layout (used for all sheets unless overridden) ───
DEFAULT_CONFIG = {
    "left_margin_in": 0.7,
    "right_margin_in": 0.7,       # LEFT_MARGIN + extra_right_margin combined
    "table_width_in": None,       # None = auto-calculate from slide width - margins
    "max_col_width_in": 3.0,      # Max width for any single column
    "compact_row_height_in": 0.18,
    "fixed_col_widths": {         # Fixed widths for known columns
        "S.No": 0.48,
        "ID": 0.4,
        "RAG": 0.45,
    },
}


# ─── Per-sheet overrides ───
# Keys must be LOWERCASE and stripped.
# Only include fields you want to override; rest falls back to DEFAULT_CONFIG.

SHEET_OVERRIDES = {

    # FS sheets — tighter table, no RAG legend
    "fs": {
        "table_width_in": 11.5,
        "left_margin_in": 0.5,
        "right_margin_in": 0.5,
        "max_col_width_in": 2.5,
        "skip_rag_legend": True,
    },

    # Phase sheets — same treatment
    "phase": {
        "table_width_in": 11.5,
        "left_margin_in": 0.5,
        "right_margin_in": 0.5,
        "max_col_width_in": 2.5,
        "skip_rag_legend": True,
    },

    # Assumptions — compact rows
    "assumptions": {
        "compact_row_height_in": 0.20,
    },

    # Add more overrides as needed:
    # "risks": { "table_width_in": 11.5 },
    # "issues": { "table_width_in": 11.5 },
}


def get_sheet_config(sheet_name: str) -> dict:
    """
    Get merged config for a sheet. Checks exact match first,
    then checks if the sheet name STARTS WITH any override key
    (so "FS - Sprint 1" matches the "fs" config).

    Returns a dict with all keys from DEFAULT_CONFIG, with overrides applied.
    Also includes pre-computed EMU values for convenience.
    """
    key = sheet_name.lower().strip()

    # Find matching override: exact match first, then prefix match
    overrides = {}
    if key in SHEET_OVERRIDES:
        overrides = SHEET_OVERRIDES[key]
    else:
        for prefix, cfg in SHEET_OVERRIDES.items():
            if key.startswith(prefix):
                overrides = cfg
                break

    # Merge with defaults
    config = {**DEFAULT_CONFIG, **overrides}

    # Merge fixed_col_widths (don't fully replace, merge)
    config["fixed_col_widths"] = {
        **DEFAULT_CONFIG.get("fixed_col_widths", {}),
        **overrides.get("fixed_col_widths", {}),
    }

    # Pre-compute EMU values
    config["left_margin_emu"] = int(Inches(config["left_margin_in"]))
    config["right_margin_emu"] = int(Inches(config["right_margin_in"]))
    config["compact_row_height_emu"] = int(Inches(config["compact_row_height_in"]))

    if config["table_width_in"]:
        config["table_width_emu"] = int(Inches(config["table_width_in"]))
    else:
        config["table_width_emu"] = None  # Caller should calculate from slide width

    config["fixed_col_widths_emu"] = {
        k: Inches(v) for k, v in config["fixed_col_widths"].items()
    }

    config["max_col_width_emu"] = int(Inches(config["max_col_width_in"]))

    return config
