"""Revision tracking: what each deck contained, and what changed since the last one.

Every real build saves a record next to its deck ("12-345-67 REV A.wig.json"):
each step's text, a fingerprint of its photo and its time, plus the parts,
materials and tools. When the next revision is built, the two records are
compared, so the deck can mark exactly what changed and the build can say so.

Comparing records rather than slides means a reworded step is caught even if
the slide looks similar, and a photo swap is caught even if the text didn't
change. A step that only moved (because one was inserted before it) is listed
but not marked: its content didn't change.

Pure functions on plain dicts; the generator builds the records.
"""

from __future__ import annotations

import json
import os
import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple

FORMAT_VERSION = 1
SUFFIX = ".wig.json"
CHANGED, ADDED = "changed", "added"
_DASHES = str.maketrans({"–": "-", "—": "-", "‑": "-", "‒": "-", "−": "-"})


def manifest_path(deck: Path) -> Path:
    """'12-345-67 REV A.pptx' -> '12-345-67 REV A.wig.json'."""
    return Path(deck).with_suffix(SUFFIX)


def load_manifest(path: Path) -> Optional[dict]:
    """The record, or None if there isn't one or it can't be read (never fatal: it only disables marks)."""
    try:
        data = json.loads(Path(path).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None
    return data if isinstance(data, dict) and data.get("format") == FORMAT_VERSION else None


def save_manifest(path: Path, manifest: dict) -> None:
    path = Path(path)
    handle, temp_name = tempfile.mkstemp(dir=str(path.parent), prefix=".wig-", suffix=".tmp")
    try:
        with os.fdopen(handle, "w", encoding="utf-8") as file:
            json.dump(manifest, file, indent=2, ensure_ascii=False)
            file.write("\n")
        os.replace(temp_name, path)
    except BaseException:
        if os.path.exists(temp_name):
            os.unlink(temp_name)
        raise


def part_key(number: str) -> str:
    return (number or "").translate(_DASHES).upper().replace(" ", "")


@dataclass
class Changes:
    previous_revision: str
    # (section, step number) -> CHANGED or ADDED: the steps to mark in the new deck.
    marks: Dict[Tuple[str, str], str] = field(default_factory=dict)
    # Part keys (see part_key) of parts that are new or whose quantity or name changed.
    part_marks: Set[str] = field(default_factory=set)
    lines: List[str] = field(default_factory=list)
    compact: List[str] = field(default_factory=list)

    @property
    def any(self) -> bool:
        return bool(self.lines)

    def compact_text(self, limit: int = 150) -> str:
        """A short change description for the revision log, e.g. 'Assembly 3, 5 changed; Assembly 6 added'."""
        text = "; ".join(self.compact) or "No content changes"
        return text if len(text) <= limit else text[: limit - 1].rstrip("; ,") + "…"


def _step_key(entry: dict) -> Tuple[str, Optional[str]]:
    return (entry.get("description", ""), entry.get("photo"))


def _compare_steps(section: str, old: List[dict], new: List[dict], changes: Changes) -> None:
    """Three passes, so inserting one step doesn't mark every step after it:
    1. unchanged in place, 2. same content at a new number (moved), 3. the rest: changed or added."""
    old_by_number = {e["step"]: e for e in old}
    consumed: Set[str] = set()
    settled: Set[str] = set()     # new step numbers already accounted for
    moved: List[Tuple[str, str]] = []
    changed: List[str] = []
    added: List[str] = []

    for entry in new:
        same = old_by_number.get(entry["step"])
        if same is not None and _step_key(same) == _step_key(entry):
            consumed.add(entry["step"])
            settled.add(entry["step"])

    for entry in new:
        if entry["step"] in settled:
            continue
        match = next((e["step"] for e in old if e["step"] not in consumed and _step_key(e) == _step_key(entry)), None)
        if match is not None:
            consumed.add(match)
            settled.add(entry["step"])
            moved.append((entry["step"], match))

    for entry in new:
        number = entry["step"]
        if number in settled:
            continue
        same = old_by_number.get(number)
        if same is not None and number not in consumed:
            consumed.add(number)
            what = [label for label, key in (("text", "description"), ("photo", "photo"))
                    if same.get(key) != entry.get(key)]
            changes.marks[(section, number)] = CHANGED
            changed.append(number)
            changes.lines.append(f"{section} {number}: {' and '.join(what) or 'content'} changed")
        else:
            changes.marks[(section, number)] = ADDED
            added.append(number)
            changes.lines.append(f"{section} {number}: added")

    for number, was in moved:
        changes.lines.append(f"{section} {number}: same step, renumbered from {was}")
    removed = [e for e in old if e["step"] not in consumed]
    for entry in removed:
        changes.lines.append(f"{section} {entry['step']} in REV {changes.previous_revision} removed: "
                             f"{entry.get('description', '')}")

    if changed:
        changes.compact.append(f"{section} {', '.join(changed)} changed")
    if added:
        changes.compact.append(f"{section} {', '.join(added)} added")
    if removed:
        changes.compact.append(f"REV {changes.previous_revision} {section} "
                               f"{', '.join(e['step'] for e in removed)} removed")


def _compare_list(label: str, old: List[str], new: List[str], changes: Changes) -> None:
    old_keys = {item.lower(): item for item in old}
    new_keys = {item.lower(): item for item in new}
    added = [new_keys[k] for k in new_keys if k not in old_keys]
    removed = [old_keys[k] for k in old_keys if k not in new_keys]
    if added:
        changes.lines.append(f"{label} added: {', '.join(added)}")
    if removed:
        changes.lines.append(f"{label} removed: {', '.join(removed)}")
    if added or removed:
        changes.compact.append(f"{label.lower()} updated")


def compare(previous: dict, current: dict) -> Changes:
    """What changed from the previous deck's record to the one about to be built."""
    changes = Changes(previous_revision=str(previous.get("revision", "?")))
    old_sections = previous.get("sections", {})
    new_sections = current.get("sections", {})
    for section in list(new_sections) + [s for s in old_sections if s not in new_sections]:
        _compare_steps(section, old_sections.get(section, []), new_sections.get(section, []), changes)

    old_parts = {part_key(p["number"]): p for p in previous.get("parts", [])}
    new_parts = {part_key(p["number"]): p for p in current.get("parts", [])}
    part_bits = []
    for key, part in new_parts.items():
        old = old_parts.get(key)
        if old is None:
            changes.part_marks.add(key)
            changes.lines.append(f"Part {part['number']} ({part.get('name', '')}) added")
            part_bits.append(f"+{part['number']}")
            continue
        if str(old.get("quantity")) != str(part.get("quantity")):
            changes.part_marks.add(key)
            changes.lines.append(f"Part {part['number']}: quantity {old.get('quantity')} → {part.get('quantity')}")
            part_bits.append(f"{part['number']} qty")
        if old.get("name") != part.get("name"):
            changes.part_marks.add(key)
            changes.lines.append(f"Part {part['number']}: name {old.get('name')} → {part.get('name')}")
            part_bits.append(f"{part['number']} name")
    for key, old in old_parts.items():
        if key not in new_parts:
            changes.lines.append(f"Part {old['number']} ({old.get('name', '')}) removed")
            part_bits.append(f"-{old['number']}")
    if part_bits:
        changes.compact.append("parts " + ", ".join(dict.fromkeys(part_bits)))

    _compare_list("Materials", previous.get("materials", []), current.get("materials", []), changes)
    _compare_list("Tools", previous.get("tools", []), current.get("tools", []), changes)
    return changes
