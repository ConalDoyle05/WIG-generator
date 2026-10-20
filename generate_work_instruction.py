"""
generate_work_instruction.py

Strict-convention Manufacturing Work Instruction generator.

Purpose:
    Creates a manufacturing work instruction (WI) PowerPoint from one assembly
    Excel workbook, catalog part images, final assembly image, assembly process
    images, testing images, and packaging images.

Expected project structure (each root is overridable with a CLI flag; see
Run, below):
    Project Folder/
        generate_work_instruction.py
        Templates/XX-XXX-XX REV XX.pptx
        Part_Catalog_Images/
        Design Folder/<prefix>/<assembly number>/<assembly number>.xlsx
        Design Folder/<prefix>/<assembly number>/Assembly/
        Design Folder/<prefix>/<assembly number>/Testing/      optional
        Design Folder/<prefix>/<assembly number>/Packaging/    optional
        Design Folder/<prefix>/<assembly number>/CANCELLED/    created automatically

Image conventions:
    Part catalog images, final assembly images, and all process-step images may
    be .png, .jpg, or .jpeg. The file name must still match the part number,
    assembly number, or step number exactly.

Revision handling:
    On a revision, superseded WI PowerPoints are moved into the CANCELLED
    folder after the new deck is saved, so the assembly folder always holds
    exactly one current work instruction.

Excel conventions:
    Required sheets: Title Info, Parts, Equipment, Assembly.
    Optional sheets: Testing, Packaging.
    Columns are read by POSITION, not by header text, so the header row can be
    named however the operator prefers. Row 1 is always treated as the header
    row and data always starts on row 2. Column order, left to right:
        Title Info: WI # | WI Title | REV # | PRODUCTION CELL | FORMER WI# |
                    ASSEMBLY TIME | Parent Assemblies
        Parts:      Part # | BOMNAME | Quantity | Part Name
        Equipment:  Materials | Tools
        Assembly / Testing / Packaging: Step # | Description
    Rows with BOMNAME = MFG BOM are treated as included subassemblies and
    placed on the cover page through template placeholders.

Adding a process section:
    Assembly, Testing, and Packaging are driven by the PROCESS_SECTIONS table.
    Adding another section is a single entry there plus a template slide.

Install:
    pip install -r requirements.txt        (or: pip install .  to get the `wig` command)

Run:
    python generate_work_instruction.py                      (asks for everything)
    python generate_work_instruction.py --assembly 12-345-67 \\
        --design-folder "D:/Shared/Design Folder" \\
        --part-images-folder "D:/Shared/Part_Catalog_Images" \\
        --template-path "D:/Shared/Templates/XX-XXX-XX REV XX.pptx"

    Add --initials CD --welding N --steps-per-slide 4 --initial-release (or
    --revision D.2 --change-description "...") to skip the prompts, several
    assembly numbers (or --assembly-file) to build in one run, --pdf to export
    PDFs, --config to describe a template with a different slide order, and
    --init DIR [--sample] to create a starter project. See --help and the README.
"""

from __future__ import annotations
import argparse
import copy
import dataclasses
import json
import math
import os
import re
import shutil
import subprocess
import sys
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Dict, Iterator, List, Optional, Tuple
from openpyxl import load_workbook
from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

# =============================================================================
# 1. PROJECT PATHS AND FILE CONVENTIONS
# =============================================================================

SCRIPT_DIR = Path(__file__).resolve().parent
DESIGN_FOLDER = SCRIPT_DIR / "Design Folder"
PART_IMAGE_FOLDER = SCRIPT_DIR / "Part_Catalog_Images"
TEMPLATE_PATH = SCRIPT_DIR / "Templates" / "XX-XXX-XX REV XX.pptx"
EXCEL_EXTENSION = ".xlsx"
# Accepted image formats for part catalog, final assembly, and process images.
# The first entry is used when reporting an expected-but-missing image.
IMAGE_EXTENSIONS = [".png", ".jpg", ".jpeg"]
IMAGE_EXTENSION = IMAGE_EXTENSIONS[0]
# Superseded work instructions are moved here on a revision. The folder is
# created alongside Assembly, Testing, and Packaging inside the assembly folder.
CANCELLED_FOLDER_NAME = "CANCELLED"

# =============================================================================
# 2. EXCEL FORMAT REQUIREMENTS
# =============================================================================

SHEET_TITLE_INFO = "Title Info"
SHEET_PARTS = "Parts"
SHEET_EQUIPMENT = "Equipment"
SHEET_ASSEMBLY_STEPS = "Assembly"
SHEET_TESTING = "Testing"
SHEET_PACKAGING = "Packaging"
REQUIRED_BASE_SHEETS = [
    SHEET_TITLE_INFO,
    SHEET_PARTS,
    SHEET_EQUIPMENT,
    SHEET_ASSEMBLY_STEPS,
]
REQUIRED_TITLE_FIELDS = [
    "WI #",
    "WI Title",
    "REV #",
    "PRODUCTION CELL",
    "FORMER WI#",
    "ASSEMBLY TIME",
]
# Fixed Title Info layout: values are always read from row 2.
TITLE_INFO_COLUMNS = {
    "WI #": 1,
    "WI Title": 2,
    "REV #": 3,
    "PRODUCTION CELL": 4,
    "FORMER WI#": 5,
    "ASSEMBLY TIME": 6,
    "Parent Assemblies": 7,
}

PART_NUMBER_COL = "Part #"
BOMNAME_COL = "BOMNAME"
QUANTITY_COL = "Quantity"
PART_NAME_COL = "Part Name"
MATERIALS_COL = "Materials"
TOOLS_COL = "Tools"
STEP_NUMBER_COL = "Step #"
STEP_DESCRIPTION_COL = "Description"

# Fixed column layouts. Data sheets are read by column POSITION, not by header
# text, because the header row exists for the operator filling in the workbook
# and its wording is not reliable ("BOM NAME" vs "BOMNAME", "Qty" vs
# "Quantity", and so on). Row 1 is always the header row; data starts on row 2.
# If a sheet's column order ever changes, update the position here.
PARTS_COLUMNS = {
    PART_NUMBER_COL: 1,
    BOMNAME_COL: 2,
    QUANTITY_COL: 3,
    PART_NAME_COL: 4,
}
EQUIPMENT_COLUMNS = {
    MATERIALS_COL: 1,
    TOOLS_COL: 2,
}
PROCESS_COLUMNS = {
    STEP_NUMBER_COL: 1,
    STEP_DESCRIPTION_COL: 2,
}

# Header spellings that are recognized during the optional layout sanity check.
# These are never used to locate a column. They only let the script warn when a
# recognized header appears in an unexpected position, which usually means the
# sheet's columns were reordered rather than merely renamed.
COLUMN_ALIASES = {
    PART_NUMBER_COL: {"PART", "PARTNO", "PARTNUM", "PARTNUMBER", "PN"},
    BOMNAME_COL: {"BOM", "BOMNAME", "BOMTYPE"},
    QUANTITY_COL: {"QTY", "QUANTITY"},
    PART_NAME_COL: {"PARTNAME", "PARTDESCRIPTION"},
    MATERIALS_COL: {"MATERIAL", "MATERIALS"},
    TOOLS_COL: {"TOOL", "TOOLS", "TOOLING"},
    STEP_NUMBER_COL: {"STEP", "STEPNO", "STEPNUM", "STEPNUMBER"},
    STEP_DESCRIPTION_COL: {"DESC", "DESCRIPTION", "INSTRUCTION", "INSTRUCTIONS"},
}

# =============================================================================
# 3. POWERPOINT TEMPLATE SETTINGS
# =============================================================================

TITLE_SLIDE_INDEX = 0
REVISION_LOG_SLIDE_INDEX = 1
# Revision log table layout: DWG REV | WI REV | Change | By | Date.
REVISION_LOG_COLUMNS = 5
WELDING_SLIDE_INDEX = 3
PARTS_TEMPLATE_SLIDE_INDEX = 4
EQUIPMENT_SLIDE_INDEX = 5
PART_CARD_POSITIONS = [
    (0.45, 1.00), (5.05, 1.00),
    (0.45, 2.25), (5.05, 2.25),
    (0.45, 3.50), (5.05, 3.50),
    (0.45, 4.75), (5.05, 4.75),
    (0.45, 6.00), (5.05, 6.00),
]
# Derived so the two can never disagree: one card position is one part slot.
PARTS_PER_SLIDE = len(PART_CARD_POSITIONS)
TITLE_IMAGE_BOX = (5.05, 1.75, 4.30, 5.00)
PROCESS_TITLE_BOX = (2.30, 0.20, 5.40, 0.70)

# Numbered badge drawn on the top-left corner of each step picture, keyed by steps
# per slide: (diameter in inches, font size). Smaller grids get smaller badges.
STEP_BADGE = {2: (0.50, 18), 4: (0.42, 16), 8: (0.34, 12)}
STEP_BADGE_INSET = 0.05
STEP_BADGE_NAME = "Step Badge"
PART_IMAGE_W = 0.95
PART_IMAGE_H = 0.95
PART_TEXT_W = 3.45
PART_TEXT_H = 1.08

# Maximum description length that fits the standard process-step text box.
# This threshold is based on a known 153-character instruction that fits
# without changing the font size or box dimensions.
STANDARD_DESCRIPTION_CHAR_LIMIT = 153

# Long description boxes always keep their original top position and grow
# downward only. They stop at whichever comes first: the box on the row below
# them, or this limit near the bottom of the 7.5-inch slide.
LONG_DESCRIPTION_BOTTOM_LIMIT = 7.30

# Clearance kept between a grown description box and the box beneath it.
LONG_DESCRIPTION_ROW_GAP = 0.05

# Image box and description box for each step position, plus the description
# font size, keyed by the number of steps placed on one slide.
PROCESS_LAYOUTS = {
    2: ([((0.45, 1.05, 4.45, 4.85), (0.45, 6.05, 4.45, 0.85)),
         ((5.10, 1.05, 4.45, 4.85), (5.10, 6.05, 4.45, 0.85))], 14),
    4: ([((0.45, 1.05, 4.45, 2.35), (0.45, 3.50, 4.45, 0.55)),
         ((5.10, 1.05, 4.45, 2.35), (5.10, 3.50, 4.45, 0.55)),
         ((0.45, 4.20, 4.45, 2.35), (0.45, 6.65, 4.45, 0.55)),
         ((5.10, 4.20, 4.45, 2.35), (5.10, 6.65, 4.45, 0.55))], 10),
    8: ([((0.25, 1.00, 2.25, 2.05), (0.25, 3.10, 2.25, 0.50)),
         ((2.65, 1.00, 2.25, 2.05), (2.65, 3.10, 2.25, 0.50)),
         ((5.05, 1.00, 2.25, 2.05), (5.05, 3.10, 2.25, 0.50)),
         ((7.45, 1.00, 2.25, 2.05), (7.45, 3.10, 2.25, 0.50)),
         ((0.25, 4.00, 2.25, 2.05), (0.25, 6.10, 2.25, 0.50)),
         ((2.65, 4.00, 2.25, 2.05), (2.65, 6.10, 2.25, 0.50)),
         ((5.05, 4.00, 2.25, 2.05), (5.05, 6.10, 2.25, 0.50)),
         ((7.45, 4.00, 2.25, 2.05), (7.45, 6.10, 2.25, 0.50))], 8),
}

# =============================================================================
# 4. DATA CONTAINERS AND PROCESS SECTION DEFINITIONS
# =============================================================================

@dataclass(frozen=True)
class SectionSpec:
    """Defines one process section end to end.

    A section owns an Excel sheet, an image subfolder, a slide title prefix,
    and the index of its template slide. Everything the generator does for
    Assembly, Testing, and Packaging is driven from this one definition, so a
    new section needs one entry here and one template slide, not edits in six
    separate places.

    has_step_placeholders is deliberately separate from required. It describes
    the TEMPLATE SLIDE (does it carry {Assembly Step picture}-style markers that
    must be stripped by name, or is it a plain slide whose text is cleared
    wholesale?), while required describes the WORKBOOK (must the sheet exist and
    have rows?). Today only Assembly is both, but they are unrelated facts, and
    fusing them would silently mangle the next required section that is added.
    """
    sheet: str
    folder: str
    title_prefix: str
    template_index: int
    required: bool
    has_step_placeholders: bool

# Order matters: sections are built in the order listed, and each one's slide
# index is offset by how many slides the earlier sections expanded to.
PROCESS_SECTIONS = [
    SectionSpec(SHEET_ASSEMBLY_STEPS, "Assembly", "Assembly Step", 6,
                required=True, has_step_placeholders=True),
    SectionSpec(SHEET_TESTING, "Testing", "Testing", 8,
                required=False, has_step_placeholders=False),
    SectionSpec(SHEET_PACKAGING, "Packaging", "Packaging", 9,
                required=False, has_step_placeholders=False),
]

# Placeholder markers stripped by name from a template slide that has them.
STEP_PLACEHOLDER_MARKERS = [
    "{Assembly Step picture}",
    "{Step Description}",
    "Step #{##}",
    "Step #{###}",
]

@dataclass
class PartRow:
    """Stores one part row from the Parts sheet."""
    part_number: str
    bomname: str
    quantity: str
    part_name: str

@dataclass
class ProcessRow:
    """Stores one row from Assembly, Testing, or Packaging process sheets."""
    step_number: str
    description: str

@dataclass
class WorkInstructionData:
    """Holds all validated data required to generate one work instruction deck."""
    assembly_number: str
    reviser_initials: str
    uses_welding: bool
    change_description: str
    is_revision: bool
    new_revision: Optional[str]
    title_info: Dict[str, str]
    parts: List[PartRow]
    included_subassemblies: List[PartRow]
    materials: List[str]
    tools: List[str]
    # Process steps keyed by sheet name, one entry per PROCESS_SECTIONS spec.
    process_steps: Dict[str, List[ProcessRow]]

    def steps_for(self, spec: SectionSpec) -> List[ProcessRow]:
        """Returns the step rows read for one process section."""
        return self.process_steps.get(spec.sheet, [])

@dataclass
class ArchiveResult:
    """Records which superseded work instructions were moved to CANCELLED."""
    moved: List[Path] = field(default_factory=list)
    failed: List[Tuple[Path, str]] = field(default_factory=list)

# =============================================================================
# 5. BASIC CLEANUP, VALIDATION, AND FORMATTING HELPERS
# =============================================================================

def clean_cell(value) -> str:
    """Converts an Excel cell value into clean text used throughout the generator."""
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%m/%d/%Y")
    text = str(value).strip()
    if text.endswith(".0") and text.replace(".0", "").isdigit():
        text = text[:-2]
    return text

def normalize_number(value) -> str:
    """Normalizes assembly, part, and step numbers for file and folder matching."""
    return clean_cell(value).upper().replace(" ", "")

def required_path(path: Path, description: str, errors: List[str], *, folder: bool = False) -> None:
    """Adds a validation error when a required file or folder is missing."""
    exists = path.exists() and (path.is_dir() if folder else path.is_file())
    if not exists:
        kind = "folder" if folder else "file"
        errors.append(f"Missing {description}:\n  Expected {kind}: {path}")

def raise_validation_errors(errors: List[str]) -> None:
    """Stops generation and prints all validation errors together."""
    if not errors:
        return
    message = "\n\n".join(errors)
    raise FileNotFoundError(
        "The work instruction could not be created because validation failed.\n\n"
        f"{message}\n\n"
        "Fix the item(s) above and run the script again."
    )

def normalize_header(value) -> str:
    """Reduces a header cell to letters and digits only for loose comparison."""
    return "".join(character for character in clean_cell(value).upper() if character.isalnum())

def resolve_columns(ws, sheet_name: str, layout: Dict[str, int]) -> Dict[str, int]:
    """Returns the fixed column index for each field on a data sheet.

    Columns are located by position, so the header row may be worded however
    the operator prefers. Header text is still inspected, but only as a sanity
    check: if a recognized header sits in an unexpected column, the sheet was
    probably reordered, and a warning is printed so the mismatch is visible
    instead of silently producing wrong data.
    """
    required_columns = max(layout.values())
    if ws.max_column < required_columns:
        raise ValueError(
            f"Sheet '{sheet_name}' has {ws.max_column} column(s) but needs at least "
            f"{required_columns}.\n"
            "Expected column order: "
            + " | ".join(field_name for field_name, _ in sorted(layout.items(), key=lambda item: item[1]))
        )

    headers = {
        column: normalize_header(ws.cell(row=1, column=column).value)
        for column in range(1, ws.max_column + 1)
    }
    for field_name, column in layout.items():
        aliases = COLUMN_ALIASES.get(field_name, set())
        matches = [
            other_column
            for other_column, header in headers.items()
            if header and header in aliases
        ]
        if matches and column not in matches:
            print(
                f"Warning: Sheet '{sheet_name}' column {column} is read as '{field_name}', "
                f"but a header matching '{field_name}' was found in column {matches[0]}. "
                "Columns are read by position, so check that the column order is correct."
            )
    return dict(layout)

def iter_sheet_rows(ws, cols: Dict[str, int]) -> Iterator[Tuple[int, Dict[str, str]]]:
    """Yields (row number, cleaned field values) for every non-blank data row.

    Data always starts on row 2 because row 1 is the operator's header row.
    Rows where every mapped column is empty are skipped, so trailing blank
    rows in a workbook do not produce validation errors.
    """
    for row in range(2, ws.max_row + 1):
        values = {
            field_name: clean_cell(ws.cell(row=row, column=column).value)
            for field_name, column in cols.items()
        }
        if any(values.values()):
            yield row, values

def first_blank_field(fields: Tuple[Tuple[str, str], ...]) -> Optional[str]:
    """Returns the label of the first blank field, or None when all are filled."""
    return next((label for label, value in fields if not value), None)

def format_parent_list(parent_text: str) -> str:
    """Formats parent assemblies from comma- or semicolon-separated text."""
    parent_text = clean_cell(parent_text)
    if not parent_text:
        return "None"
    parents = [
        item.strip()
        for item in parent_text.replace(",", ";").split(";")
        if item.strip()
    ]
    return ", ".join(parents) if parents else "None"

def format_included_subassemblies(data: WorkInstructionData) -> str:
    """Formats included subassemblies from BOMNAME = MFG BOM rows."""
    if not data.included_subassemblies:
        return "None"
    return ", ".join(part.part_number for part in data.included_subassemblies)

# =============================================================================
# 6. PATH BUILDERS
# =============================================================================

def assembly_base_folder(assembly_number: str) -> Path:
    """Returns the design-folder location for one assembly number."""
    assembly_number = normalize_number(assembly_number)
    prefix_folder = assembly_number.split("-")[0]
    return DESIGN_FOLDER / prefix_folder / assembly_number

def assembly_excel_path(assembly_number: str) -> Path:
    """Builds the expected Excel workbook path for the assembly."""
    assembly_number = normalize_number(assembly_number)
    return assembly_base_folder(assembly_number) / f"{assembly_number}{EXCEL_EXTENSION}"

def cancelled_folder(assembly_number: str) -> Path:
    """Returns the CANCELLED folder that holds superseded work instructions."""
    return assembly_base_folder(assembly_number) / CANCELLED_FOLDER_NAME

def image_path_from_folder(folder: Path, item_number: str) -> Path:
    """Finds an image by number, accepting PNG, JPG, or JPEG.

    The exact name is tried first. If nothing matches, the folder is scanned
    case-insensitively, so a file saved as PART.JPG or Part.Png is still found
    on Linux and macOS, where the earlier exact check would have missed it.
    Windows finds it either way. Returns the expected PNG path when no file
    exists, so callers can report a clear missing-image error.
    """
    item_number = normalize_number(item_number)
    for extension in IMAGE_EXTENSIONS:
        image_path = folder / f"{item_number}{extension}"
        if image_path.exists():
            return image_path

    if folder.is_dir():
        wanted = item_number.casefold()
        for candidate in sorted(folder.iterdir()):
            if (candidate.is_file()
                    and candidate.stem.casefold() == wanted
                    and candidate.suffix.casefold() in IMAGE_EXTENSIONS):
                return candidate

    return folder / f"{item_number}{IMAGE_EXTENSION}"

def catalog_image_path(item_number: str) -> Path:
    """Finds a shared catalog image, used for both parts and final assemblies."""
    return image_path_from_folder(PART_IMAGE_FOLDER, item_number)

def section_image_folder(spec: SectionSpec, assembly_number: str) -> Path:
    """Returns the image folder for one process section."""
    return assembly_base_folder(assembly_number) / spec.folder

def section_image_path(spec: SectionSpec, assembly_number: str, step_number: str) -> Path:
    """Finds the step image for one process section."""
    return image_path_from_folder(section_image_folder(spec, assembly_number), step_number)

REVISION_IN_FILENAME = re.compile(r"REV\s*([A-Za-z0-9.]+)")

def revision_sort_key(path: Path) -> Tuple[str, int, str]:
    """Sorts WI files by revision so REV D.10 ranks above REV D.9."""
    match = REVISION_IN_FILENAME.search(path.stem)
    if not match:
        return ("", -1, path.name)
    drawing_rev, mwi_revision = split_revision(match.group(1))
    numeric_revision = int(mwi_revision) if mwi_revision.isdigit() else -1
    return (drawing_rev.upper(), numeric_revision, path.name)

def existing_work_instructions(assembly_number: str) -> List[Path]:
    """Lists WI PowerPoints in the assembly folder, oldest revision first.

    Only the top level of the assembly folder is searched, so decks already
    moved into CANCELLED are ignored.
    """
    folder = assembly_base_folder(assembly_number)
    if not folder.is_dir():
        return []
    return sorted(folder.glob(f"*{assembly_number}*REV*.pptx"), key=revision_sort_key)

def find_existing_wi(assembly_number: str) -> Optional[Path]:
    """Finds the latest existing WI PowerPoint for revision-log preservation."""
    matches = existing_work_instructions(assembly_number)
    return matches[-1] if matches else None

# =============================================================================
# 7. EXCEL READING FUNCTIONS
# =============================================================================

def require_sheets(wb) -> None:
    """Verifies that the workbook contains all required base sheets."""
    missing = [sheet for sheet in REQUIRED_BASE_SHEETS if sheet not in wb.sheetnames]
    if missing:
        raise ValueError(
            f"Workbook is missing required sheet(s): {', '.join(missing)}\n"
            f"Found sheets: {', '.join(wb.sheetnames)}"
        )

def read_title_info(wb, errors: List[str]) -> Dict[str, str]:
    """Reads cover-page metadata from fixed columns in row 2.

    Blank required fields are collected into the shared error list rather than
    raised immediately, so a missing title value is reported alongside missing
    parts, steps, and images instead of hiding them.
    """
    ws = wb[SHEET_TITLE_INFO]
    title_info = {
        field_name: clean_cell(ws.cell(row=2, column=column).value)
        for field_name, column in TITLE_INFO_COLUMNS.items()
    }
    missing_fields = [
        field_name
        for field_name in REQUIRED_TITLE_FIELDS
        if not title_info.get(field_name)
    ]
    if missing_fields:
        errors.append(
            "The following required Title Info values are blank:\n\n"
            + "\n".join(missing_fields)
        )
    return title_info

def read_parts(wb, errors: List[str]) -> Tuple[List[PartRow], List[PartRow]]:
    """Reads parts from the fixed Part #/BOMNAME/Quantity/Part Name column positions."""
    ws = wb[SHEET_PARTS]
    cols = resolve_columns(ws, SHEET_PARTS, PARTS_COLUMNS)

    parts: List[PartRow] = []
    included_subassemblies: List[PartRow] = []
    missing_part_images: List[str] = []

    for row, values in iter_sheet_rows(ws, cols):
        part_number = normalize_number(values[PART_NUMBER_COL])
        quantity = values[QUANTITY_COL]
        part_name = values[PART_NAME_COL]
        bomname = values[BOMNAME_COL].upper()

        blank_field = first_blank_field((
            ("Part Number", part_number),
            ("Quantity", quantity),
            ("Part Name", part_name),
        ))
        if blank_field:
            part_reference = f" for part {part_number}" if part_number else ""
            errors.append(
                f"Sheet '{SHEET_PARTS}', row {row}: {blank_field} is blank{part_reference}."
            )
            continue

        if not catalog_image_path(part_number).exists():
            missing_part_images.append(part_number)

        part = PartRow(
            part_number=part_number,
            bomname=bomname,
            quantity=quantity,
            part_name=part_name,
        )
        parts.append(part)
        if bomname == "MFG BOM":
            included_subassemblies.append(part)

    if not parts:
        errors.append(f"Sheet '{SHEET_PARTS}' has no valid part rows.")
    if missing_part_images:
        errors.append(
            f"Missing Part Images ({len(missing_part_images)}):\n\n"
            + "\n".join(sorted(missing_part_images))
        )
    return parts, included_subassemblies

def read_equipment(wb) -> Tuple[List[str], List[str]]:
    """Reads materials and tools from the fixed Equipment column positions."""
    ws = wb[SHEET_EQUIPMENT]
    cols = resolve_columns(ws, SHEET_EQUIPMENT, EQUIPMENT_COLUMNS)
    materials: List[str] = []
    tools: List[str] = []
    for _, values in iter_sheet_rows(ws, cols):
        if values[MATERIALS_COL]:
            materials.append(values[MATERIALS_COL])
        if values[TOOLS_COL]:
            tools.append(values[TOOLS_COL])
    return materials, tools

def read_process_sheet(
    wb,
    spec: SectionSpec,
    assembly_number: str,
    errors: List[str],
) -> List[ProcessRow]:
    """Reads one process section's sheet and validates its step images."""
    if spec.sheet not in wb.sheetnames:
        if spec.required:
            errors.append(f"Workbook is missing required sheet: {spec.sheet}")
        return []

    ws = wb[spec.sheet]
    cols = resolve_columns(ws, spec.sheet, PROCESS_COLUMNS)
    rows: List[ProcessRow] = []
    missing_step_images: List[str] = []

    for row, values in iter_sheet_rows(ws, cols):
        step_number = normalize_number(values[STEP_NUMBER_COL])
        description = values[STEP_DESCRIPTION_COL]

        blank_field = first_blank_field((
            (STEP_NUMBER_COL, step_number),
            (STEP_DESCRIPTION_COL, description),
        ))
        if blank_field:
            step_reference = f" for step {step_number}" if step_number else ""
            errors.append(
                f"Sheet '{spec.sheet}', row {row}: {blank_field} is blank{step_reference}."
            )
            continue

        if not section_image_path(spec, assembly_number, step_number).exists():
            missing_step_images.append(step_number)
        rows.append(ProcessRow(step_number=step_number, description=description))

    if spec.required and not rows:
        errors.append(f"Sheet '{spec.sheet}' has no valid rows.")
    if missing_step_images:
        errors.append(
            f"Missing {spec.sheet} Images ({len(missing_step_images)}):\n\n"
            + "\n".join(sorted(missing_step_images))
        )
    return rows

def require_section_folder(spec: SectionSpec, assembly_number: str, errors: List[str]) -> None:
    """Adds a validation error when a process section's image folder is missing."""
    required_path(
        section_image_folder(spec, assembly_number),
        f"{spec.folder.lower()} image folder for {assembly_number}",
        errors,
        folder=True,
    )

def validate_project_paths(assembly_number: str, errors: List[str]) -> None:
    """Checks every folder and file that must exist before the workbook is opened."""
    required_path(DESIGN_FOLDER, "design folder", errors, folder=True)
    required_path(PART_IMAGE_FOLDER, "part catalog image folder", errors, folder=True)
    required_path(TEMPLATE_PATH.parent, "PowerPoint template folder", errors, folder=True)
    required_path(assembly_base_folder(assembly_number), f"assembly folder for {assembly_number}", errors, folder=True)
    for spec in PROCESS_SECTIONS:
        if spec.required:
            require_section_folder(spec, assembly_number, errors)
    required_path(assembly_excel_path(assembly_number), "assembly Excel workbook", errors)
    required_path(TEMPLATE_PATH, "PowerPoint template", errors)
    required_path(catalog_image_path(assembly_number), f"final assembly image for assembly {assembly_number}", errors)

def read_and_validate_data(
    assembly_number: str,
    reviser_initials: str,
    uses_welding: bool,
    change_description: str,
    is_revision: bool,
    new_revision: Optional[str],
) -> WorkInstructionData:
    """Validates paths, reads Excel data, and returns a complete data object."""
    assembly_number = normalize_number(assembly_number)
    errors: List[str] = []

    validate_project_paths(assembly_number, errors)
    raise_validation_errors(errors)

    wb = load_workbook(assembly_excel_path(assembly_number), data_only=True)
    require_sheets(wb)

    title_info = read_title_info(wb, errors)
    title_info["DATE"] = datetime.now().strftime("%m/%d/%Y")
    if new_revision:
        title_info["REV #"] = new_revision

    parts, included_subassemblies = read_parts(wb, errors)
    materials, tools = read_equipment(wb)

    process_steps: Dict[str, List[ProcessRow]] = {}
    for spec in PROCESS_SECTIONS:
        process_steps[spec.sheet] = read_process_sheet(wb, spec, assembly_number, errors)

    # An optional section only needs its image folder once it actually has steps.
    for spec in PROCESS_SECTIONS:
        if not spec.required and process_steps[spec.sheet]:
            require_section_folder(spec, assembly_number, errors)
    raise_validation_errors(errors)

    return WorkInstructionData(
        assembly_number=assembly_number,
        reviser_initials=reviser_initials,
        uses_welding=uses_welding,
        change_description=change_description,
        is_revision=is_revision,
        new_revision=new_revision,
        title_info=title_info,
        parts=parts,
        included_subassemblies=included_subassemblies,
        materials=materials,
        tools=tools,
        process_steps=process_steps,
    )

# =============================================================================
# 8. POWERPOINT UTILITY FUNCTIONS
# =============================================================================

def split_revision(revision: str) -> Tuple[str, str]:
    """Splits a revision like D.1 into drawing revision D and WI revision 1."""
    revision = clean_cell(revision)
    if "." not in revision:
        return revision, ""
    drawing_rev, mwi_revision = revision.split(".", 1)
    return drawing_rev, mwi_revision

def find_table(slide, description: str):
    """Returns the first table on a slide, or raises when the slide has none."""
    for shape in slide.shapes:
        if shape.has_table:
            return shape.table
    raise ValueError(f"{description} does not contain a table.")

def read_existing_revision_log(existing_pptx: Path) -> List[List[str]]:
    """Reads previous revision-log rows from an existing WI PowerPoint."""
    prs = Presentation(existing_pptx)
    table = find_table(
        prs.slides[REVISION_LOG_SLIDE_INDEX],
        "Existing WI revision log slide",
    )
    column_count = min(REVISION_LOG_COLUMNS, len(table.columns))
    rows = []
    for row_index in range(1, len(table.rows)):
        values = [clean_cell(table.cell(row_index, col_index).text) for col_index in range(column_count)]
        if any(values):
            rows.append(values)
    return rows

def delete_slide(prs: Presentation, slide_index: int) -> None:
    """Deletes a slide from the presentation by index."""
    slide_id_list = prs.slides._sldIdLst
    slides = list(slide_id_list)
    slide_id_list.remove(slides[slide_index])

BLANK_LAYOUT_INDEX = 6

def blank_slide_layout(prs: Presentation):
    """Returns the emptiest slide layout to host a duplicated slide.

    Standard Office templates put the Blank layout at index 6, which is used
    when present. Templates built from a custom master may order layouts
    differently or have fewer of them, so fall back to whichever layout carries
    the fewest placeholders rather than raising IndexError.
    """
    layouts = list(prs.slide_layouts)
    if not layouts:
        raise ValueError("The PowerPoint template contains no slide layouts.")
    if len(layouts) > BLANK_LAYOUT_INDEX:
        return layouts[BLANK_LAYOUT_INDEX]
    return min(layouts, key=lambda layout: len(layout.placeholders))

def duplicate_slide(prs: Presentation, slide_index: int):
    """Duplicates a template slide by copying its shapes."""
    source = prs.slides[slide_index]
    copied = prs.slides.add_slide(blank_slide_layout(prs))
    for shape in source.shapes:
        copied.shapes._spTree.insert_element_before(copy.deepcopy(shape.element), "p:extLst")
    return copied

def move_slide(prs: Presentation, old_index: int, new_index: int) -> None:
    """Moves a slide from one index position to another."""
    xml_slides = prs.slides._sldIdLst
    slides = list(xml_slides)
    slide = slides[old_index]
    xml_slides.remove(slide)
    xml_slides.insert(new_index, slide)

def expand_template_slide(prs: Presentation, template_index: int, slide_count: int) -> None:
    """Duplicates a template slide until the section has slide_count pages."""
    for _ in range(slide_count - 1):
        duplicate_slide(prs, template_index)
        move_slide(prs, len(prs.slides) - 1, template_index + 1)

def remove_shape(shape) -> None:
    """Removes one shape from a slide."""
    shape.element.getparent().remove(shape.element)

def remove_shapes(slide, *, pictures: bool = False, text: bool = False) -> None:
    """Removes pictures and/or text shapes from a slide."""
    for shape in list(slide.shapes):
        if (pictures and shape.shape_type == MSO_SHAPE_TYPE.PICTURE) or (text and shape.has_text_frame):
            remove_shape(shape)

def remove_template_text(slide, markers: List[str]) -> None:
    """Removes template placeholder text shapes from a slide."""
    for shape in list(slide.shapes):
        if shape.has_text_frame and any(marker in shape.text for marker in markers):
            remove_shape(shape)

def replace_text_everywhere(slide, replacements: Dict[str, str]) -> None:
    """Replaces placeholders while preserving formatting."""
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                for old, new in replacements.items():
                    if old in run.text:
                        run.text = run.text.replace(old, str(new))
                # Gray out "None"
                if run.text.strip() == "None":
                    run.font.color.rgb = RGBColor(128, 128, 128)
                    run.font.bold = True

def title_replacements(data: WorkInstructionData) -> Dict[str, str]:
    """Builds all title-slide placeholder replacement values."""
    return {
        "{Assembly Number}": data.assembly_number,
        "{TITLE}": data.title_info["WI Title"],
        "{REV}": data.title_info["REV #"],
        "{DATE}": data.title_info["DATE"],
        "{CELL}": data.title_info["PRODUCTION CELL"],
        "{FORMER}": data.title_info["FORMER WI#"],
        "{TIME}": data.title_info["ASSEMBLY TIME"],
        "{PARENT_ASSEMBLIES}": format_parent_list(data.title_info.get("Parent Assemblies", "")),
        "{INCLUDED_SUBASSEMBLIES}": format_included_subassemblies(data),
    }

def add_textbox(
    slide,
    text: str,
    x: float,
    y: float,
    w: float,
    h: float,
    font_size: int = 14,
    bold: bool = False,
    align_center: bool = False,
    red_border: bool = False,
    border_weight: float = 2.0,
) -> None:
    """Adds a formatted textbox to a slide."""
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    text_frame = box.text_frame
    text_frame.clear()
    text_frame.word_wrap = True
    paragraph = text_frame.paragraphs[0]
    # Assigning multi-line text splits it into one run per line (joined by
    # <a:br>), so every run must be formatted, not just the first, or lines
    # after the first fall back to the theme's default run size.
    paragraph.text = text
    if align_center:
        paragraph.alignment = PP_ALIGN.CENTER
    for run in paragraph.runs:
        run.font.size = Pt(font_size)
        run.font.bold = bold
        run.font.color.rgb = RGBColor(0, 0, 0)
    if red_border:
        box.line.color.rgb = RGBColor(255, 0, 0)
        box.line.width = Pt(border_weight)

# Standard border weight for the red content boxes used on every content slide.
CONTENT_BORDER_WEIGHT = 3

def add_bordered_box(slide, text: str, box: Tuple[float, float, float, float],
                     font_size: int, bold: bool = False) -> None:
    """Adds the standard centered, red-bordered content box used on content slides."""
    add_textbox(
        slide,
        text,
        *box,
        font_size=font_size,
        bold=bold,
        align_center=True,
        red_border=True,
        border_weight=CONTENT_BORDER_WEIGHT,
    )

def process_slide_title(spec: SectionSpec, items: List[ProcessRow]) -> str:
    """Titles a process slide by the steps it holds: 'Assembly Steps 1-4', 'Testing Step 3'."""
    # "Assembly Step" already names a step; "Testing" and "Packaging" need the word added.
    singular = spec.title_prefix if spec.title_prefix.endswith("Step") else f"{spec.title_prefix} Step"
    first, last = items[0].step_number, items[-1].step_number
    if len(items) == 1 or first == last:
        return f"{singular} {first}"
    return f"{singular}s {first}\u2013{last}"

def add_step_badge(slide, step_number: str, picture, items_per_slide: int) -> None:
    """Draws the step number in a red circle on the top-left corner of a step picture."""
    diameter, font_size = STEP_BADGE[items_per_slide]
    badge = slide.shapes.add_shape(
        MSO_SHAPE.OVAL,
        picture.left + Inches(STEP_BADGE_INSET),
        picture.top + Inches(STEP_BADGE_INSET),
        Inches(diameter),
        Inches(diameter),
    )
    # The name keeps renumber_all_slides from mistaking a badge for the page number.
    badge.name = f"{STEP_BADGE_NAME} {step_number}"
    badge.fill.solid()
    badge.fill.fore_color.rgb = RGBColor(255, 0, 0)
    badge.line.fill.background()
    frame = badge.text_frame
    frame.margin_left = frame.margin_right = frame.margin_top = frame.margin_bottom = 0
    frame.text = str(step_number)
    paragraph = frame.paragraphs[0]
    paragraph.alignment = PP_ALIGN.CENTER
    for run in paragraph.runs:
        run.font.size = Pt(font_size)
        run.font.bold = True
        run.font.color.rgb = RGBColor(255, 255, 255)

def add_picture_fit(slide, image_path: Path, x: float, y: float, w: float, h: float):
    """Adds an image centered inside a target box while preserving aspect ratio.

    Returns the picture shape, whose position is where the image really landed
    (it is smaller than the box when the aspect ratios differ).
    """
    try:
        with Image.open(image_path) as img:
            img_w, img_h = img.size
    except OSError as error:
        raise ValueError(
            f"Could not read image file:\n  {image_path}\n\n"
            "The file exists but is not a readable PNG, JPG, or JPEG. It may be "
            "corrupt, still uploading, or saved with the wrong extension.\n"
            f"Underlying error: {error}"
        )
    if not img_w or not img_h:
        raise ValueError(f"Image has no usable dimensions:\n  {image_path}")

    box_ratio = w / h
    img_ratio = img_w / img_h
    if img_ratio > box_ratio:
        final_w = w
        final_h = w / img_ratio
    else:
        final_h = h
        final_w = h * img_ratio
    final_x = x + (w - final_w) / 2
    final_y = y + (h - final_h) / 2
    return slide.shapes.add_picture(
        str(image_path),
        Inches(final_x),
        Inches(final_y),
        width=Inches(final_w),
        height=Inches(final_h),
    )

SLIDE_NUMBER_BOX = (9.10, 7.05, 0.75, 0.30)
# A shape is only treated as the existing slide-number placeholder if it sits in
# the footer band. Without this, a content shape whose text happens to be digits
# (a quantity, a short step description) could be overwritten with a page number.
SLIDE_NUMBER_MIN_TOP = 6.50

def is_slide_number_shape(shape) -> bool:
    """Reports whether a shape looks like the template's slide-number placeholder."""
    if not shape.has_text_frame or shape.name.startswith(STEP_BADGE_NAME):
        return False
    text = shape.text.strip()
    return text.isdigit() and len(text) <= 3

def find_slide_number_shape(slide):
    """Finds the shape holding the slide number, preferring one in the footer band."""
    candidates = [shape for shape in slide.shapes if is_slide_number_shape(shape)]
    if not candidates:
        return None
    footer_minimum = Inches(SLIDE_NUMBER_MIN_TOP)
    in_footer = [
        shape for shape in candidates
        if shape.top is not None and shape.top >= footer_minimum
    ]
    return in_footer[0] if in_footer else candidates[0]

def renumber_all_slides(prs: Presentation) -> None:
    """Adds or updates slide numbers on every slide."""
    x, y, w, h = SLIDE_NUMBER_BOX
    for number, slide in enumerate(prs.slides, start=1):
        shape = find_slide_number_shape(slide)
        if shape is None:
            add_textbox(slide, str(number), x, y, w, h, font_size=14, align_center=True)
            continue
        shape.text = str(number)
        shape.left = Inches(x)
        shape.top = Inches(y)
        shape.width = Inches(w)
        shape.height = Inches(h)
        for paragraph in shape.text_frame.paragraphs:
            paragraph.alignment = PP_ALIGN.CENTER
            for run in paragraph.runs:
                run.font.size = Pt(14)
                run.font.bold = False

def set_parts_title(slide, title: str) -> None:
    """Updates the parts-list title or creates one if the template title is missing."""
    for shape in slide.shapes:
        if shape.has_text_frame and "Parts List" in shape.text:
            shape.text = title
            return
    add_textbox(slide, title, 0.55, 0.40, 8.0, 0.45, font_size=24, bold=True)

# =============================================================================
# 9. SLIDE POPULATION FUNCTIONS
# =============================================================================

def populate_title_slide(slide, data: WorkInstructionData) -> None:
    """Populates the title slide with metadata and the final assembly image."""
    replace_text_everywhere(slide, title_replacements(data))
    add_picture_fit(slide, catalog_image_path(data.assembly_number), *TITLE_IMAGE_BOX)

def populate_revision_log_slide(slide, data: WorkInstructionData) -> None:
    """Populates the revision log table while preserving previous revisions."""
    table = find_table(slide, "Revision Log slide")
    existing_rows: List[List[str]] = []
    if data.is_revision:
        existing_wi = find_existing_wi(data.assembly_number)
        if existing_wi is None:
            raise FileNotFoundError(
                f"Cannot create revision {data.title_info['REV #']} for assembly {data.assembly_number}.\n\n"
                "The previous revision of this WI could not be found.\n"
                "A revision must be created from the existing WI so the revision log history can be preserved.\n\n"
                f"Expected location:\n  {assembly_base_folder(data.assembly_number)}\n\n"
                f"Expected file pattern:\n  *{data.assembly_number}*REV*.pptx\n\n"
                f"Superseded revisions are archived in the {CANCELLED_FOLDER_NAME} folder and are "
                "not searched, because a revision must be built from the current WI.\n"
                "Place the current REV PowerPoint in the assembly folder and run the script again."
            )
        existing_rows = read_existing_revision_log(existing_wi)
    drawing_rev, mwi_revision = split_revision(data.title_info["REV #"])
    new_row = [
        drawing_rev,
        mwi_revision,
        data.change_description,
        data.reviser_initials,
        data.title_info["DATE"],
    ]
    all_rows = existing_rows + [new_row]
    # Check capacity before writing anything so the table is never left half filled.
    available_rows = len(table.rows) - 1
    if len(all_rows) > available_rows:
        raise ValueError(
            f"Revision log table is full: {len(all_rows)} revision(s) to write but only "
            f"{available_rows} data row(s) in the template.\n\n"
            f"Add at least {len(all_rows) - available_rows} more row(s) to the revision log "
            "table in the PowerPoint template."
        )
    column_count = min(REVISION_LOG_COLUMNS, len(table.columns))
    for row_index, row_values in enumerate(all_rows, start=1):
        for col_index, value in enumerate(row_values[:column_count]):
            cell = table.cell(row_index, col_index)
            cell.text = str(value)
            for paragraph in cell.text_frame.paragraphs:
                paragraph.alignment = PP_ALIGN.CENTER
                for run in paragraph.runs:
                    run.font.size = Pt(12)
                    run.font.bold = False
                    run.font.color.rgb = RGBColor(0, 0, 0)

def populate_parts_slide(slide, parts: List[PartRow], slide_number_label: str) -> None:
    """Populates one parts-list slide with up to ten parts."""
    # A duplicated template slide's picture shapes keep XML that references
    # image relationships from the original slide's .rels part, which the
    # duplicate does not have. Any picture already on the template (e.g. a
    # placeholder graphic) must be stripped before part-card images are
    # added, or duplicated slides show a broken or wrong image.
    remove_shapes(slide, pictures=True)
    remove_template_text(
        slide,
        ["{Part Number}", "{BOMNAME}", "{Quantity}", "{Description}", "{Part Name}"],
    )
    set_parts_title(slide, f"Parts List - {slide_number_label}")
    for index, part in enumerate(parts[:PARTS_PER_SLIDE]):
        x, y = PART_CARD_POSITIONS[index]
        add_picture_fit(slide, catalog_image_path(part.part_number), x, y, PART_IMAGE_W, PART_IMAGE_H)
        add_bordered_box(
            slide,
            f"{part.part_number}  |  QTY: {part.quantity}\n{part.part_name}",
            (x + 1.05, y, PART_TEXT_W, PART_TEXT_H),
            font_size=13,
        )

def populate_equipment_slide(slide, materials: List[str], tools: List[str]) -> None:
    """Populates the equipment slide with materials and tools."""
    remove_template_text(slide, ["*Materials*", "{Material", "*Tools*", "{Tool"])
    for heading, items, box in (
        ("Materials", materials, (0.90, 1.55, 3.85, 4.65)),
        ("Tools", tools, (5.10, 1.55, 3.85, 4.65)),
    ):
        body = "\n".join(f"- {item}" for item in items) if items else "- N/A"
        add_bordered_box(slide, f"{heading}\n{body}", box, font_size=16)

def boxes_share_a_column(first, second) -> bool:
    """Reports whether two boxes overlap horizontally, so one can sit under the other."""
    left, _, width, _ = first
    other_left, _, other_width, _ = second
    return other_left < left + width and left < other_left + other_width

def description_bottom_limit(layout_boxes, index: int) -> float:
    """Returns how far down one slot's description box is allowed to grow.

    A grown box must not run into whatever sits beneath it. This scans the
    layout for any box lower down that shares horizontal space with this slot
    and stops short of the highest one, falling back to the slide's bottom
    limit when nothing is below. Deriving the stop point from the layout means
    the 2, 4, and 8-up grids all stay correct without per-layout special cases.
    """
    desc_box = layout_boxes[index][1]
    _, top, _, _ = desc_box
    limits = [LONG_DESCRIPTION_BOTTOM_LIMIT]
    for image_box, other_desc_box in layout_boxes:
        for other_box in (image_box, other_desc_box):
            _, other_top, _, _ = other_box
            if other_top > top and boxes_share_a_column(desc_box, other_box):
                limits.append(other_top - LONG_DESCRIPTION_ROW_GAP)
    return min(limits)

def adjust_description_box(desc_box, description: str, bottom_limit: float):
    """Returns either the normal box or a taller long-description box.

    The top-left position and width never change. Descriptions up to
    153 characters use the standard box. Longer descriptions grow downward
    only, keeping the same width so they cannot overlap the neighboring
    column, and stopping at bottom_limit so they cannot overlap the row below.
    """
    x, y, w, h = desc_box
    if len(description) <= STANDARD_DESCRIPTION_CHAR_LIMIT:
        return desc_box
    desired_height = h + 0.85
    # Keep the original top edge fixed and grow downward only.
    # Never shrink below the template height, even in a tight slot.
    max_height = max(h, bottom_limit - y)
    return (x, y, w, min(desired_height, max_height))

def populate_process_slide(
    slide,
    spec: SectionSpec,
    assembly_number: str,
    items: List[ProcessRow],
    slide_number: int,
    items_per_slide: int,
) -> None:
    """Populates one process slide with assembly, testing, or packaging steps."""
    remove_shapes(slide, pictures=True)
    if spec.has_step_placeholders:
        remove_template_text(slide, STEP_PLACEHOLDER_MARKERS)
    else:
        remove_shapes(slide, text=True)

    add_bordered_box(
        slide,
        process_slide_title(spec, items),
        PROCESS_TITLE_BOX,
        font_size=30,
        bold=True,
    )

    layout_boxes, desc_font_size = PROCESS_LAYOUTS[items_per_slide]
    for index, item in enumerate(items):
        image_box, desc_box = layout_boxes[index]
        picture = add_picture_fit(
            slide,
            section_image_path(spec, assembly_number, item.step_number),
            *image_box,
        )
        add_step_badge(slide, item.step_number, picture, items_per_slide)
        add_bordered_box(
            slide,
            item.description,
            adjust_description_box(
                desc_box,
                item.description,
                description_bottom_limit(layout_boxes, index),
            ),
            font_size=desc_font_size,
            bold=True,
        )

# =============================================================================
# 10. DECK BUILDING LOGIC
# =============================================================================

def populate_process_section(
    prs: Presentation,
    spec: SectionSpec,
    template_index: int,
    items: List[ProcessRow],
    assembly_number: str,
    items_per_slide: int,
) -> int:
    """Duplicates and fills all slides for one process section.

    Returns the number of slides the section occupies. A section with no steps
    has its template slide deleted and returns 0.
    """
    if not items:
        delete_slide(prs, template_index)
        return 0
    slide_count = math.ceil(len(items) / items_per_slide)
    expand_template_slide(prs, template_index, slide_count)
    for slide_offset in range(slide_count):
        start = slide_offset * items_per_slide
        populate_process_slide(
            prs.slides[template_index + slide_offset],
            spec,
            assembly_number,
            items[start:start + items_per_slide],
            slide_offset + 1,
            items_per_slide,
        )
    return slide_count

def populate_parts_section(prs: Presentation, parts: List[PartRow]) -> int:
    """Duplicates and fills all parts-list slides. Returns the slide count."""
    slide_count = max(1, math.ceil(len(parts) / PARTS_PER_SLIDE))
    expand_template_slide(prs, PARTS_TEMPLATE_SLIDE_INDEX, slide_count)
    for slide_offset in range(slide_count):
        start = slide_offset * PARTS_PER_SLIDE
        label = f"{slide_offset + 1} of {slide_count}" if slide_count > 1 else "1"
        populate_parts_slide(
            prs.slides[PARTS_TEMPLATE_SLIDE_INDEX + slide_offset],
            parts[start:start + PARTS_PER_SLIDE],
            label,
        )
    return slide_count

def delete_blank_optional_template_slides(prs: Presentation) -> None:
    """Removes optional template pages that still contain the blank-data marker.

    The slides remain part of the master PowerPoint template. They are only
    removed from the generated work instruction when no content has replaced
    the template's 'Leave page blank if no data' instruction.
    """
    blank_marker = "leave page blank if no data"
    optional_titles = {"cycle time", "size and weight"}

    for slide_index in range(len(prs.slides) - 1, -1, -1):
        slide = prs.slides[slide_index]
        combined = "\n".join(
            clean_cell(shape.text)
            for shape in slide.shapes
            if shape.has_text_frame and clean_cell(shape.text)
        ).lower()

        has_blank_marker = blank_marker in combined
        is_optional_static_slide = any(title in combined for title in optional_titles)

        if has_blank_marker and is_optional_static_slide:
            delete_slide(prs, slide_index)

def unique_destination(folder: Path, file_name: str) -> Path:
    """Returns a non-colliding path inside a folder by adding a numeric suffix."""
    destination = folder / file_name
    if not destination.exists():
        return destination
    counter = 2
    while True:
        candidate = folder / f"{destination.stem} ({counter}){destination.suffix}"
        if not candidate.exists():
            return candidate
        counter += 1

def archive_superseded_work_instructions(assembly_number: str, current_output: Path) -> ArchiveResult:
    """Moves older WI PowerPoints into the CANCELLED folder.

    Runs only after the new deck has been saved, so the previous revision stays
    available while the revision log is being read. Everything in the assembly
    folder except the newly created file is moved, which leaves exactly one
    current work instruction in the main folder.
    """
    result = ArchiveResult()
    current_output = current_output.resolve()
    superseded = [
        path
        for path in existing_work_instructions(assembly_number)
        if path.resolve() != current_output
    ]
    if not superseded:
        return result

    destination_folder = cancelled_folder(assembly_number)
    try:
        destination_folder.mkdir(parents=True, exist_ok=True)
    except OSError as error:
        result.failed = [(path, str(error)) for path in superseded]
        return result

    for path in superseded:
        destination = unique_destination(destination_folder, path.name)
        try:
            shutil.move(str(path), str(destination))
        except (OSError, PermissionError) as error:
            result.failed.append((path, str(error)))
        else:
            result.moved.append(destination)
    return result

def save_presentation(prs: Presentation, output_path: Path) -> None:
    """Saves the deck, refusing to overwrite an existing work instruction."""
    if output_path.exists():
        raise FileExistsError(
            f"Output file already exists:\n{output_path}\n\n"
            "Rename, move, or close the existing file before generating a new one."
        )
    try:
        prs.save(output_path)
    except PermissionError:
        raise PermissionError(
            f"Cannot save output file:\n{output_path}\n\n"
            "Make sure the PowerPoint is not open and that you have permission to save here."
        )

def build_work_instruction(data: WorkInstructionData, items_per_slide: int) -> Tuple[Path, ArchiveResult]:
    """Builds and saves the complete work instruction PowerPoint."""
    prs = Presentation(TEMPLATE_PATH)
    last_needed = max([TITLE_SLIDE_INDEX, REVISION_LOG_SLIDE_INDEX, WELDING_SLIDE_INDEX,
                       PARTS_TEMPLATE_SLIDE_INDEX, EQUIPMENT_SLIDE_INDEX,
                       *(spec.template_index for spec in PROCESS_SECTIONS)])
    if len(prs.slides) <= last_needed:
        raise ValueError(
            f"The template has {len(prs.slides)} slides, but the slide layout expects at least {last_needed + 1}.\n"
            f"Template: {TEMPLATE_PATH}\n"
            "Check the slide order in the README, or pass --config with your template's slide positions."
        )
    welding_slide_id = prs.slides[WELDING_SLIDE_INDEX].slide_id
    populate_title_slide(prs.slides[TITLE_SLIDE_INDEX], data)
    populate_revision_log_slide(prs.slides[REVISION_LOG_SLIDE_INDEX], data)

    part_slide_count = populate_parts_section(prs, data.parts)

    # Every section after the parts list shifts down by the number of extra
    # slides the earlier sections added. A section that was deleted because it
    # had no steps contributes -1, which pulls the later sections back up.
    slide_offset = part_slide_count - 1
    populate_equipment_slide(
        prs.slides[EQUIPMENT_SLIDE_INDEX + slide_offset],
        data.materials,
        data.tools,
    )
    for spec in PROCESS_SECTIONS:
        section_slide_count = populate_process_section(
            prs,
            spec,
            spec.template_index + slide_offset,
            data.steps_for(spec),
            data.assembly_number,
            items_per_slide,
        )
        slide_offset += section_slide_count - 1

    # Cycle Time and Size and Weight are optional template pages.
    # If their blank-data marker is still present, omit them from the generated deck.
    delete_blank_optional_template_slides(prs)

    # The welding slide is found by identity, not by its template index: a blank
    # optional page (Cycle Time) ahead of it is deleted above and would shift it,
    # so a fixed index would delete the wrong slide.
    if not data.uses_welding:
        for index, slide in enumerate(prs.slides):
            if slide.slide_id == welding_slide_id:
                delete_slide(prs, index)
                break
    renumber_all_slides(prs)

    output_path = assembly_base_folder(data.assembly_number) / f"{data.assembly_number} REV {data.title_info['REV #']}.pptx"
    save_presentation(prs, output_path)

    # Only a revision supersedes an existing deck. An initial release has
    # nothing to cancel, so the CANCELLED folder is not created.
    archive = ArchiveResult()
    if data.is_revision:
        archive = archive_superseded_work_instructions(data.assembly_number, output_path)
    return output_path, archive

# =============================================================================
# 11. COMMAND-LINE ENTRY POINT
# =============================================================================

# An operator answers several prompts in a row. Re-asking a mistyped answer is
# far better than discarding the ones already given, but the retry count is
# bounded so a piped or redirected run cannot loop forever.
MAX_INPUT_ATTEMPTS = 3

def prompt_until_valid(prompt: str, is_valid, error_message: str, *, upper: bool = False) -> str:
    """Asks until the answer passes is_valid, then returns the cleaned answer."""
    for attempt in range(1, MAX_INPUT_ATTEMPTS + 1):
        try:
            answer = clean_cell(input(prompt))
        except EOFError:
            raise ValueError(error_message)
        if upper:
            answer = answer.upper()
        if is_valid(answer):
            return answer
        if attempt < MAX_INPUT_ATTEMPTS:
            print(f"  {error_message}")
    raise ValueError(error_message)

def get_required_text(prompt: str, error_message: str, *, upper: bool = False) -> str:
    """Asks for a value that simply cannot be blank."""
    return prompt_until_valid(prompt, lambda answer: bool(answer), error_message, upper=upper)

def get_choice(prompt: str, choices: List[str], error_message: str) -> str:
    """Asks for one of a fixed set of answers, compared case-insensitively."""
    return prompt_until_valid(
        prompt,
        lambda answer: answer in choices,
        error_message,
        upper=True,
    )

def get_yes_no(prompt: str) -> bool:
    """Prompts the user for Y/N input and returns True for Y and False for N."""
    return get_choice(prompt, ["Y", "N"], "Input must be Y or N.") == "Y"

def get_steps_per_slide() -> int:
    """Prompts the user for a supported number of process steps per slide."""
    valid = [str(option) for option in sorted(PROCESS_LAYOUTS)]
    answer = get_choice(
        f"How many process steps per slide? ({'/'.join(valid)}): ",
        valid,
        f"Steps per slide must be {', '.join(valid)}.",
    )
    return int(answer)

def get_release_inputs() -> Tuple[bool, Optional[str], str]:
    """Prompts the user for initial release or revision details."""
    release_type = get_choice(
        "Is this an Initial Release or Revision? (I/R): ",
        ["I", "R"],
        "Release type must be I or R.",
    )
    if release_type == "I":
        return False, None, "Initial Release"
    new_revision = get_required_text(
        "Enter NEW revision (Ex: D.2): ",
        "New revision cannot be blank.",
        upper=True,
    )
    change_description = get_required_text(
        "Enter revision change description: ",
        "Revision change description cannot be blank.",
    )
    return True, new_revision, change_description

def report_archive(archive: ArchiveResult) -> None:
    """Prints which superseded work instructions were archived, and which could not be."""
    if archive.moved:
        print(f"Moved to {CANCELLED_FOLDER_NAME} ({len(archive.moved)}):")
        for path in archive.moved:
            print(f"  {path.name}")
    if archive.failed:
        print(
            f"\nWARNING: {len(archive.failed)} superseded work instruction(s) could not be moved "
            f"to the {CANCELLED_FOLDER_NAME} folder."
        )
        print("Close the file(s) and move them manually so only the current revision remains:")
        for path, reason in archive.failed:
            print(f"  {path.name} - {reason}")

# -----------------------------------------------------------------------------
# Template layout config and PDF export
# -----------------------------------------------------------------------------

# Slide keys accepted in a --config file, in the order they must appear in the
# template. Sections after the parts list shift down as earlier sections grow,
# so the order is fixed even though the positions are configurable.
CONFIG_SLIDE_ORDER = ["welding", "parts", "equipment", "assembly", "testing", "packaging"]
CONFIG_SLIDE_KEYS = ["title", "revision_log"] + CONFIG_SLIDE_ORDER
CONFIG_PATH_KEYS = {
    "design_folder": "DESIGN_FOLDER",
    "part_images_folder": "PART_IMAGE_FOLDER",
    "template_path": "TEMPLATE_PATH",
}

def load_config(config_path: Path) -> None:
    """Applies a JSON config: template slide positions and default folder paths.

    Example:
        {"slides": {"title": 0, "revision_log": 1, "welding": 3, "parts": 4,
                    "equipment": 5, "assembly": 6, "testing": 8, "packaging": 9},
         "paths": {"template_path": "Templates/our template.pptx"}}

    Any key may be left out. Relative paths resolve against the config file.
    Command-line flags still override whatever the config says.
    """
    global TITLE_SLIDE_INDEX, REVISION_LOG_SLIDE_INDEX, WELDING_SLIDE_INDEX
    global PARTS_TEMPLATE_SLIDE_INDEX, EQUIPMENT_SLIDE_INDEX, PROCESS_SECTIONS
    global DESIGN_FOLDER, PART_IMAGE_FOLDER, TEMPLATE_PATH

    config_path = Path(config_path)
    try:
        config = json.loads(config_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise ValueError(f"Could not read config file {config_path}: {exc}")
    if not isinstance(config, dict):
        raise ValueError(f"Config file {config_path} must contain a JSON object.")
    unknown_top = set(config) - {"slides", "paths"}
    if unknown_top:
        raise ValueError(f"Unknown config key(s): {', '.join(sorted(unknown_top))}. Allowed: slides, paths.")

    slides = {
        "title": TITLE_SLIDE_INDEX,
        "revision_log": REVISION_LOG_SLIDE_INDEX,
        "welding": WELDING_SLIDE_INDEX,
        "parts": PARTS_TEMPLATE_SLIDE_INDEX,
        "equipment": EQUIPMENT_SLIDE_INDEX,
    }
    for spec in PROCESS_SECTIONS:
        slides[spec.sheet.lower()] = spec.template_index

    given = config.get("slides", {})
    if not isinstance(given, dict):
        raise ValueError('Config "slides" must be an object such as {"parts": 4}.')
    bad_keys = set(given) - set(CONFIG_SLIDE_KEYS)
    if bad_keys:
        raise ValueError(f"Unknown slide key(s): {', '.join(sorted(bad_keys))}. Allowed: {', '.join(CONFIG_SLIDE_KEYS)}.")
    for key, value in given.items():
        if not isinstance(value, int) or isinstance(value, bool) or value < 0:
            raise ValueError(f'Slide position for "{key}" must be a whole number 0 or higher, got {value!r}.')
        slides[key] = value

    if len(set(slides.values())) != len(slides):
        raise ValueError("Each template slide needs its own position; two entries share one.")
    ordered = [slides[key] for key in CONFIG_SLIDE_ORDER]
    if ordered != sorted(ordered):
        raise ValueError(
            "Template slides must appear in this order: "
            + ", ".join(CONFIG_SLIDE_ORDER)
            + ". Title and revision_log may sit anywhere, but they are filled in before the other pages "
            "are copied, so keep them ahead of the welding slide."
        )
    if max(slides["title"], slides["revision_log"]) > slides["welding"]:
        raise ValueError("The title and revision_log slides must come before the welding slide.")

    TITLE_SLIDE_INDEX = slides["title"]
    REVISION_LOG_SLIDE_INDEX = slides["revision_log"]
    WELDING_SLIDE_INDEX = slides["welding"]
    PARTS_TEMPLATE_SLIDE_INDEX = slides["parts"]
    EQUIPMENT_SLIDE_INDEX = slides["equipment"]
    PROCESS_SECTIONS = [
        dataclasses.replace(spec, template_index=slides[spec.sheet.lower()]) for spec in PROCESS_SECTIONS
    ]

    paths = config.get("paths", {})
    if not isinstance(paths, dict) or set(paths) - set(CONFIG_PATH_KEYS):
        raise ValueError(f"Config \"paths\" may only contain: {', '.join(CONFIG_PATH_KEYS)}.")
    for key, value in paths.items():
        resolved = (config_path.parent / value).resolve() if not Path(value).is_absolute() else Path(value)
        globals()[CONFIG_PATH_KEYS[key]] = resolved

def find_soffice() -> Optional[str]:
    """Locates LibreOffice (SOFFICE_BIN overrides the search)."""
    override = os.environ.get("SOFFICE_BIN")
    if override:
        return override if Path(override).exists() else None
    return shutil.which("soffice") or shutil.which("libreoffice")

def export_pdf(pptx_path: Path) -> Path:
    """Converts a finished deck to PDF next to it, using headless LibreOffice."""
    soffice = find_soffice()
    if soffice is None:
        raise RuntimeError(
            "PDF export needs LibreOffice, and it was not found on PATH.\n"
            "Install it, or set SOFFICE_BIN to the soffice executable, then run again."
        )
    result = subprocess.run(
        [soffice, "--headless", "--convert-to", "pdf", "--outdir", str(pptx_path.parent), str(pptx_path)],
        capture_output=True, text=True, timeout=300,
    )
    pdf_path = pptx_path.with_suffix(".pdf")
    if result.returncode != 0 or not pdf_path.exists():
        raise RuntimeError(f"LibreOffice could not convert {pptx_path.name}:\n{result.stderr.strip() or result.stdout.strip()}")
    return pdf_path

# -----------------------------------------------------------------------------
# Command-line interface
# -----------------------------------------------------------------------------

@dataclass
class Answers:
    """The operator's answers that apply to every assembly in a run."""
    initials: str
    uses_welding: bool
    items_per_slide: int
    is_revision: bool
    new_revision: Optional[str]
    change_description: str

def build_parser() -> argparse.ArgumentParser:
    layouts = [str(option) for option in sorted(PROCESS_LAYOUTS)]
    parser = argparse.ArgumentParser(
        prog="wig",
        description="Create work instruction PowerPoints from assembly workbooks and image folders. "
                    "Any answer not given as a flag is asked for interactively.",
    )
    where = parser.add_argument_group("what to build")
    where.add_argument("--assembly", nargs="+", metavar="NUMBER",
                       help="One or more assembly numbers, for example 12-345-67. Several numbers build in one run.")
    where.add_argument("--assembly-file", metavar="FILE",
                       help="Text file with one assembly number per line. Blank lines and lines starting with # are skipped.")
    where.add_argument("--init", metavar="DIR",
                       help="Create an empty project skeleton in DIR (folders, starter template, blank workbook) and exit.")
    where.add_argument("--sample", action="store_true",
                       help="With --init, also add one made-up assembly (12-345-67) so you can run it straight away.")

    ingest = parser.add_argument_group("import from OneNote")
    ingest.add_argument("--import-onenote", metavar="DOCX",
                        help="Import steps from a OneNote page exported as Word (a table with a photo and a "
                             "description per row) into --assembly: writes the numbered photos and the "
                             "Assembly sheet, then exits without building a deck.")
    ingest.add_argument("--section", choices=["assembly", "testing", "packaging"], default="assembly",
                        help="With --import-onenote, which part of the instruction the page holds (default assembly). "
                             "Export a separate page for Testing and for Packaging.")
    ingest.add_argument("--replace", action="store_true",
                        help="With --import-onenote, overwrite steps and photos that already exist.")
    ingest.add_argument("--max-px", type=int, default=1600,
                        help="With --import-onenote, shrink photos so the longest side is at most this many pixels (default 1600).")
    ingest.add_argument("--image-format", choices=["jpg", "png"], default="jpg",
                        help="With --import-onenote, format to save photos in (default jpg, which keeps decks small).")

    paths = parser.add_argument_group("where things are")
    paths.add_argument("--config", metavar="FILE", help="JSON file with template slide positions and default paths.")
    paths.add_argument("--design-folder", help="Overrides the default Design Folder root.")
    paths.add_argument("--part-images-folder", help="Overrides the default part-catalog image folder.")
    paths.add_argument("--template-path", help="Overrides the default PowerPoint template path.")

    answers = parser.add_argument_group("answers (skip the prompts)")
    answers.add_argument("--initials", help="Reviser initials for the revision log.")
    answers.add_argument("--welding", type=str.upper, choices=["Y", "N"], help="Keep the welding slide (Y) or remove it (N).")
    answers.add_argument("--steps-per-slide", choices=layouts, help="Process steps per slide.")
    release = answers.add_mutually_exclusive_group()
    release.add_argument("--initial-release", action="store_true", help="This is the first release of the instruction.")
    release.add_argument("--revision", metavar="REV", help="Revise the existing instruction to this revision, for example D.2.")
    answers.add_argument("--change-description", help="One-line description of the change (required with --revision).")

    output = parser.add_argument_group("output")
    output.add_argument("--pdf", action="store_true", help="Also export each finished deck to PDF (needs LibreOffice).")
    return parser

def read_assembly_file(path: Path) -> List[str]:
    """Reads assembly numbers from a text file, one per line."""
    try:
        lines = Path(path).read_text(encoding="utf-8").splitlines()
    except OSError as exc:
        raise ValueError(f"Could not read assembly file {path}: {exc}")
    return [line.strip() for line in lines if line.strip() and not line.strip().startswith("#")]

def collect_answers(args: argparse.Namespace) -> Answers:
    """Uses flags where given and asks for everything else."""
    if args.revision is not None and not clean_cell(args.change_description):
        raise ValueError("--revision needs --change-description.")
    if args.change_description and args.revision is None:
        raise ValueError("--change-description only applies together with --revision.")

    initials = clean_cell(args.initials).upper() if args.initials else get_required_text(
        "Enter reviser initials (Ex: CD): ", "Reviser initials cannot be blank.", upper=True)
    uses_welding = (args.welding == "Y") if args.welding else get_yes_no("Is welding used in this assembly? (Y/N): ")
    items_per_slide = int(args.steps_per_slide) if args.steps_per_slide else get_steps_per_slide()
    if args.initial_release:
        is_revision, new_revision, change = False, None, "Initial Release"
    elif args.revision is not None:
        new_revision = clean_cell(args.revision).upper()
        if not new_revision:
            raise ValueError("--revision cannot be blank.")
        is_revision, change = True, clean_cell(args.change_description)
    else:
        is_revision, new_revision, change = get_release_inputs()
    return Answers(initials, uses_welding, items_per_slide, is_revision, new_revision, change)

def check_assembly_setup(assembly_number: str) -> List[str]:
    """Returns setup problems for one assembly (empty when it is ready to build)."""
    errors: List[str] = []
    required_path(assembly_base_folder(assembly_number), f"assembly folder for {assembly_number}", errors, folder=True)
    required_path(assembly_excel_path(assembly_number), "assembly Excel workbook", errors)
    return errors

def build_one(assembly_number: str, answers: Answers, make_pdf: bool) -> None:
    """Reads, validates and builds one assembly, printing a summary."""
    print(f"\nAssembly number: {assembly_number}")
    print(f"Expected Excel: {assembly_excel_path(assembly_number)}")
    print(f"Expected final assembly image from part catalog: {catalog_image_path(assembly_number)}")
    for spec in PROCESS_SECTIONS:
        print(f"Expected {spec.folder.lower()} image folder: {section_image_folder(spec, assembly_number)}")

    data = read_and_validate_data(
        assembly_number, answers.initials, answers.uses_welding, answers.change_description,
        answers.is_revision, answers.new_revision,
    )
    output_path, archive = build_work_instruction(data, answers.items_per_slide)

    print("\nWork instruction created successfully.")
    print(f"Output file: {output_path}")
    report_archive(archive)
    print(f"Parts read: {len(data.parts)}")
    print(f"Parts slides created: {math.ceil(len(data.parts) / PARTS_PER_SLIDE)}")
    for spec in PROCESS_SECTIONS:
        steps = data.steps_for(spec)
        print(f"{spec.sheet} rows read: {len(steps)}")
        if spec.required:
            print(f"{spec.sheet} slides created: {math.ceil(len(steps) / answers.items_per_slide)}")
    if make_pdf:
        print(f"PDF: {export_pdf(output_path)}")

def import_steps_from_onenote(assembly_number: str, args: argparse.Namespace) -> None:
    """Runs the OneNote import for one assembly and prints what it did."""
    from wig_import import import_onenote
    spec = next(spec for spec in PROCESS_SECTIONS if spec.sheet.lower() == args.section)
    result = import_onenote(
        Path(args.import_onenote),
        assembly_base_folder(assembly_number),
        assembly_excel_path(assembly_number),
        replace=args.replace,
        max_px=args.max_px,
        image_format=args.image_format,
        sheet=spec.sheet,
        folder=spec.folder,
    )
    print(f"Imported {result.steps} {spec.sheet} step(s) for {assembly_number} from OneNote page"
          + (f' "{result.title}"' if result.title else "") + ".")
    print(f"Photos:   {section_image_folder(spec, assembly_number)}")
    print(f"Workbook: {assembly_excel_path(assembly_number)}"
          + ("  (created: fill in Title Info, Parts and Equipment)" if result.workbook_created else ""))
    if result.backup_path:
        print(f"Backup of the previous workbook: {result.backup_path}")
    if result.removed_images:
        print(f"Replaced {result.removed_images} existing step photo(s).")
    for warning in result.warnings:
        print(f"WARNING: {warning}")

def main(argv: Optional[List[str]] = None) -> int:
    """Runs the tool. Returns 0 when everything built, 1 otherwise."""
    global DESIGN_FOLDER, PART_IMAGE_FOLDER, TEMPLATE_PATH
    args = build_parser().parse_args(argv)

    try:
        if args.sample and not args.init:
            raise ValueError("--sample only applies together with --init.")
        if args.init:
            from wig_scaffold import scaffold_project
            made = scaffold_project(Path(args.init), sample=args.sample)
            print(f"Created project skeleton in {args.init}")
            print(f"  Template:  {made['template_path']}")
            print("  Workbook:  Templates/assembly-workbook-template.xlsx (copy it to start an assembly)")
            if args.sample:
                print(f"  Sample assembly ready: {made['assembly_dir'].name}")
            print("Next: restyle the template in PowerPoint (keep the slide order and {placeholders}),")
            print("then run:  wig --assembly <number> "
                  f'--design-folder "{made["design_folder"]}" --part-images-folder "{made["part_images_folder"]}" '
                  f'--template-path "{made["template_path"]}"')
            return 0

        if args.config:
            load_config(Path(args.config))
        if args.design_folder:
            DESIGN_FOLDER = Path(args.design_folder)
        if args.part_images_folder:
            PART_IMAGE_FOLDER = Path(args.part_images_folder)
        if args.template_path:
            TEMPLATE_PATH = Path(args.template_path)

        numbers: List[str] = list(args.assembly or [])
        if args.assembly_file:
            numbers += read_assembly_file(Path(args.assembly_file))
        if not numbers:
            numbers = [get_required_text("Enter assembly number (Ex: 12-345-67): ", "Assembly number cannot be blank.")]
        assemblies = list(dict.fromkeys(normalize_number(number) for number in numbers))
        if not all(assemblies):
            raise ValueError("Assembly number cannot be blank.")
        if args.import_onenote:
            if len(assemblies) != 1:
                raise ValueError("--import-onenote needs exactly one --assembly number.")
            import_steps_from_onenote(assemblies[0], args)
            return 0
        if args.pdf and find_soffice() is None:
            raise ValueError("--pdf needs LibreOffice, and it was not found. Install it or set SOFFICE_BIN.")

        # Fail before asking the remaining questions if an assembly is not set up.
        failures: Dict[str, str] = {}
        ready: List[str] = []
        for number in assemblies:
            problems = check_assembly_setup(number)
            if problems:
                failures[number] = "\n".join(problems)
            else:
                ready.append(number)
        if ready:
            answers = collect_answers(args)
            for number in ready:
                try:
                    build_one(number, answers, args.pdf)
                except (ValueError, OSError, RuntimeError) as exc:
                    failures[number] = str(exc)
    except (ValueError, OSError, RuntimeError) as exc:
        print(f"\nERROR: {exc}", file=sys.stderr)
        return 1

    if failures:
        print(f"\n{len(failures)} of {len(assemblies)} assembl{'y' if len(assemblies) == 1 else 'ies'} failed:", file=sys.stderr)
        for number, reason in failures.items():
            print(f"\n[{number}]\n{reason}", file=sys.stderr)
        return 1
    if len(assemblies) > 1:
        print(f"\nAll {len(assemblies)} assemblies built.")
    return 0

if __name__ == "__main__":
    sys.exit(main())
