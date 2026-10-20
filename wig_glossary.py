"""The polish glossary: tool codes, shorthand terms, units, count words, safety words and settings.

The glossary is a JSON file you own. It belongs to your project folder, not to
this repository: once you start adding your own shop's shorthand, part names
and tool codes, it describes how your workplace works, so keep it private.
`wig --init` and the first `wig --polish` write a small generic starter.

    {
      "version": 1,
      "settings": {"part_format": "{name} ({number})", "require_part_number": "always"},
      "tools": {"W1/2": {"name": "1/2\\" wrench", "phrase": "a 1/2\\" wrench"}},
      "terms": {"tight": "tighten", "w": "with", "insp": "inspect"},
      "units": {"nm": "N·m", "in": "in"},
      "count_words": ["screw", "pin"],
      "safety_words": ["torque", "hot"],
      "words": ["fixture"]
    }

tools    code -> name and phrase. Matched ignoring case. The phrase is what the
         step says ("Tighten with a 1/2" wrench"); the name is for tool lists.
terms    shorthand -> replacement. Keys may be several words ("tight w"); the
         longest match wins. Matched ignoring case.
units    unit as typed -> how it is written. A number needs a unit, a count word
         or a part after it, otherwise the step is flagged.
count_words   nouns that make a bare number a count ("4 screws").
safety_words  any of these in a note forces review.
words    extra words to treat as correctly spelled, on top of the built-in list.

Nothing here guesses: a word that is not a term, tool, unit, part or known word
is flagged so you can add it.
"""

from __future__ import annotations

import json
import os
import re
import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Set

FORMAT_VERSION = 1
REQUIRE_MODES = ("always", "when_part_mentioned")
DEFAULT_SETTINGS = {
    "part_format": "{name} ({number})",
    "require_part_number": "always",
    # Looks like a part number: capitals and digits in dash-separated groups, with at
    # least one digit (P-004, 12-345-67). Used only to flag numbers missing from Parts.
    "part_number_pattern": r"(?=[A-Z0-9-]*\d)[A-Z0-9]+(?:-[A-Z0-9]+)+",
}
TOP_LEVEL_KEYS = {"version", "settings", "tools", "terms", "units", "count_words", "safety_words", "words"}

STARTER_GLOSSARY = {
    "version": FORMAT_VERSION,
    "settings": dict(DEFAULT_SETTINGS),
    "tools": {
        "W1/2": {"name": "1/2\" wrench", "phrase": "a 1/2\" wrench"},
        "PH2": {"name": "#2 Phillips screwdriver", "phrase": "a #2 Phillips screwdriver"},
        "HK4": {"name": "4 mm hex key", "phrase": "a 4 mm hex key"},
        "TQS": {"name": "torque screwdriver", "phrase": "the torque screwdriver"},
    },
    "terms": {
        "tight": "tighten", "w": "with", "w/": "with", "w/o": "without", "ins": "insert",
        "inst": "install", "pl": "place", "pos": "position", "rem": "remove", "insp": "inspect",
        "chk": "check", "tq": "torque", "lube": "lubricate", "conn": "connect",
        "disc": "disconnect", "assy": "assembly", "subassy": "subassembly", "qty": "quantity",
        "ea": "each", "approx": "approximately", "b/w": "between", "thru": "through",
        "til": "until", "&": "and", "rh": "right-hand", "lh": "left-hand", "cw": "clockwise",
        "ccw": "counterclockwise", "fwd": "forward", "btm": "bottom", "frt": "front",
        "hw": "hardware", "scr": "screw", "scrs": "screws", "wshr": "washer",
        "wshrs": "washers", "brkt": "bracket", "dwg": "drawing", "ref": "reference",
    },
    "units": {
        "in": "in", "inch": "in", "inches": "in", "\"": "\"", "mm": "mm", "cm": "cm",
        "nm": "N·m", "n-m": "N·m", "n·m": "N·m", "in-lb": "in-lb", "in-lbs": "in-lb",
        "inlb": "in-lb", "ft-lb": "ft-lb", "ft-lbs": "ft-lb", "ftlb": "ft-lb", "psi": "psi",
        "deg": "°", "°": "°", "sec": "s", "s": "s", "min": "min", "lb": "lb", "lbs": "lb",
        "kg": "kg", "%": "%",
    },
    "count_words": [
        "screw", "bolt", "nut", "washer", "pin", "rivet", "clip", "tab", "hole", "place",
        "location", "side", "time", "turn", "wire", "cable", "tie", "label", "part", "piece",
        "spacer", "standoff", "insert", "bracket", "plate", "drop", "layer", "strip",
    ],
    "safety_words": [
        "torque", "hot", "sharp", "weld", "caution", "warning", "danger", "pinch", "lockout",
        "voltage", "pressure", "chemical", "crush", "lift",
    ],
    "words": [],
}

# Correctly spelled words the polish step never flags. Deliberately general:
# common English plus shop vocabulary. Your own words go in the glossary's
# "words" list, and words from every step you approve are learned automatically.
BASE_WORDS = """
a about above across after again against align aligned all along also an and angle another any
apart apply are around as assemble assembled assembly at attach attached away back base be before
behind below bend bent between both bottom box bracket break bring by cable can cap carefully
center centered check clamp clean clear clip close closed connect connector correct correctly cover
crimp cut cycle debris do does down drawing drill drive each edge empty end engage engaged even
excess face fasten feed fill finger firm firmly first fit fitting fixture flat flush for force
from front full fully gap gaps gently glue go good grease groove guide half hand hardware has
have head hold hole holes horizontal housing if in inner insert inside install into is it its
jig keep key label latch layer leave left length level lever lift light line lined lock locked
log loose lower lubricate make mark match material mount mounted move must nearest next no not
note now nut of off on onto open opening or orient orientation other out outer outside over pack
package pair part parts pass past pattern pin place plate plug point position press pull push
put rail rear reference release remove repeat replace reset rest return right rod rotate route
run same screw screws seal seat seated secure set shaft side sides single slide slot slowly snap
snug so spacer spring square stand start step stop straight strip surface tab tap tape the then
there thread threaded threads through tie tight tighten tightened to too tool top toward towards
track tray turn twist two under until up upper use used using verify vertical visible wait wall
washer when where while wipe wire with without work yes zero
one three four five six seven eight nine ten
here this that these those them they it's itself its which what who why how than very more less
most least much many few some only just still already almost always never once twice again also
be been being was were will would should could may might can't don't do not none
new old small large big short long wide narrow thin thick high low deep near far inner outer
middle centre corner end ends both either neither each other same different correct wrong
make sure ensure confirm look looking see seen shown show hold holding keep keeping leave
take put set get give place placed lay lie rest sit fit fits fitted use uses needed need needs
by per via onto into upon within without against along across around behind beside beyond
first second third last final next previous following before after during while until then once
left right top bottom front back rear side sides up down in out on off over under
all any every no nothing everything something part parts piece pieces item items unit units
red black white blue green yellow orange grey gray silver clear
again align adjust area bag bin bolt bubble bushing carton case catch chamfer cord count damage
dent dowel ensure fastener flange gasket gear handle hinge hook joint knob lid loop magnet nozzle
o-ring panel pipe plastic plunger port pump rivet rubber scratch shim sleeve socket solder sticker
stud switch tube valve wheel wrench screwdriver hex phillips torque driver pliers cutter hammer
mallet gun drill bit file punch scale caliper gauge meter
any approximately clockwise counterclockwise each every forward left-hand right-hand
quantity subassembly between through without bottom front each check inspect install connect
disconnect position remove insert approximately reference drawing hardware
""".split()


@dataclass
class Tool:
    code: str
    name: str
    phrase: str


@dataclass
class Glossary:
    path: Optional[Path]
    settings: Dict[str, str]
    tools: Dict[str, Tool]            # keyed by code, lower case
    terms: Dict[str, str]             # keyed by shorthand, lower case, single spaces
    units: Dict[str, str]             # keyed by unit as typed, lower case
    count_words: Set[str]
    safety_words: Set[str]
    words: Set[str]
    max_term_words: int = 1
    part_number_re: re.Pattern = field(default=re.compile(DEFAULT_SETTINGS["part_number_pattern"]))


def _normalize_key(key: str) -> str:
    return re.sub(r"\s+", " ", key.strip()).lower()


def _reject_duplicates(pairs):
    seen = {}
    for key, value in pairs:
        if key in seen:
            raise ValueError(f'the key "{key}" appears twice in the same section')
        seen[key] = value
    return seen


def _parse(text: str, path: Optional[Path]) -> dict:
    where = f"{path}" if path else "glossary"
    try:
        data = json.loads(text, object_pairs_hook=_reject_duplicates)
    except json.JSONDecodeError as exc:
        raise ValueError(
            f"{where} is not valid JSON: {exc.msg} at line {exc.lineno}, column {exc.colno}.\n"
            "The usual causes are a comma after the last entry in a list or section, or a missing quote."
        )
    except ValueError as exc:
        raise ValueError(f"{where}: {exc}.")
    if not isinstance(data, dict):
        raise ValueError(f"{where} must contain a JSON object.")
    return data


def build_glossary(data: dict, path: Optional[Path] = None) -> Glossary:
    """Validates raw glossary data and returns a Glossary. Raises ValueError listing every problem."""
    problems: List[str] = []

    unknown = set(data) - TOP_LEVEL_KEYS
    if unknown:
        problems.append(f"Unknown section(s): {', '.join(sorted(unknown))}. Allowed: {', '.join(sorted(TOP_LEVEL_KEYS))}.")
    version = data.get("version", FORMAT_VERSION)
    if version != FORMAT_VERSION:
        problems.append(f"Unsupported glossary version {version!r}; this WIG reads version {FORMAT_VERSION}.")

    def section(name: str, kind: type, default):
        value = data.get(name, default)
        if not isinstance(value, kind):
            problems.append(f'"{name}" must be {"an object" if kind is dict else "a list"}.')
            return default
        return value

    settings = dict(DEFAULT_SETTINGS)
    raw_settings = section("settings", dict, {})
    for key, value in raw_settings.items():
        if key not in DEFAULT_SETTINGS:
            problems.append(f'Unknown setting "{key}". Allowed: {", ".join(DEFAULT_SETTINGS)}.')
        elif not isinstance(value, str):
            problems.append(f'Setting "{key}" must be text.')
        else:
            settings[key] = value
    if settings["require_part_number"] not in REQUIRE_MODES:
        problems.append(f'Setting "require_part_number" must be one of: {", ".join(REQUIRE_MODES)}.')
    if "{number}" not in settings["part_format"]:
        problems.append('Setting "part_format" must contain {number}, for example "{name} ({number})".')
    try:
        settings["part_format"].format(name="x", number="y")
    except (KeyError, IndexError, ValueError):
        problems.append('Setting "part_format" may only use {name} and {number}.')
    try:
        part_number_re = re.compile(settings["part_number_pattern"])
    except re.error as exc:
        problems.append(f'Setting "part_number_pattern" is not a valid pattern: {exc}.')
        part_number_re = re.compile(DEFAULT_SETTINGS["part_number_pattern"])

    terms: Dict[str, str] = {}
    for key, value in section("terms", dict, {}).items():
        norm = _normalize_key(key) if isinstance(key, str) else ""
        if not norm:
            problems.append("A term has an empty shorthand.")
        elif not isinstance(value, str) or not value.strip():
            problems.append(f'Term "{key}" must map to non-empty text.')
        elif norm in terms:
            problems.append(f'Term "{key}" is listed twice (ignoring case and spacing).')
        else:
            terms[norm] = value.strip()

    tools: Dict[str, Tool] = {}
    for code, entry in section("tools", dict, {}).items():
        lowered = code.strip().lower()
        if not lowered or re.search(r"\s", code.strip()):
            problems.append(f'Tool code "{code}" must be one word with no spaces.')
            continue
        if not isinstance(entry, dict) or not isinstance(entry.get("name"), str) or not entry["name"].strip():
            problems.append(f'Tool "{code}" needs a "name", for example {{"name": "1/2\\" wrench"}}.')
            continue
        extra = set(entry) - {"name", "phrase"}
        if extra:
            problems.append(f'Tool "{code}" has unknown field(s): {", ".join(sorted(extra))}. Allowed: name, phrase.')
        phrase = entry.get("phrase", entry["name"])
        if not isinstance(phrase, str) or not phrase.strip():
            problems.append(f'Tool "{code}" has an empty phrase.')
            continue
        if lowered in tools:
            problems.append(f'Tool code "{code}" is listed twice (ignoring case).')
            continue
        tools[lowered] = Tool(code.strip(), entry["name"].strip(), phrase.strip())

    units: Dict[str, str] = {}
    for key, value in section("units", dict, {}).items():
        norm = key.strip().lower() if isinstance(key, str) else ""
        if not norm or not isinstance(value, str) or not value.strip():
            problems.append(f'Unit "{key}" must map to non-empty text.')
        elif norm in units:
            problems.append(f'Unit "{key}" is listed twice (ignoring case).')
        else:
            units[norm] = value.strip()

    def word_list(name: str) -> Set[str]:
        values = section(name, list, [])
        result: Set[str] = set()
        for value in values:
            if not isinstance(value, str) or not value.strip():
                problems.append(f'"{name}" may only hold non-empty text.')
            else:
                result.add(value.strip().lower())
        return result

    count_words = word_list("count_words")
    safety_words = word_list("safety_words")
    words = word_list("words")

    # One shorthand must mean one thing, or the result depends on which rule ran first.
    for code in tools:
        if code in terms:
            problems.append(f'"{code}" is both a tool code and a term. Rename one.')
        if code in units:
            problems.append(f'"{code}" is both a tool code and a unit. Rename one.')
    for key in terms:
        if key in units:
            problems.append(f'"{key}" is both a term and a unit. Rename one.')

    if problems:
        where = f"{path}" if path else "The glossary"
        raise ValueError(f"{where} has {len(problems)} problem(s):\n  " + "\n  ".join(problems))

    return Glossary(
        path=path, settings=settings, tools=tools, terms=terms, units=units,
        count_words=count_words, safety_words=safety_words, words=words,
        max_term_words=max([len(key.split(" ")) for key in terms] or [1]),
        part_number_re=part_number_re,
    )


def load_glossary(path: Path) -> Glossary:
    path = Path(path)
    try:
        text = path.read_text(encoding="utf-8")
    except OSError as exc:
        raise ValueError(f"Could not read glossary {path}: {exc}")
    return build_glossary(_parse(text, path), path)


def write_starter_glossary(path: Path) -> bool:
    """Writes the generic starter glossary unless a file is already there. Returns True if written."""
    path = Path(path)
    if path.exists():
        return False
    _write_json(path, STARTER_GLOSSARY)
    return True


def _write_json(path: Path, data: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    text = json.dumps(data, indent=2, ensure_ascii=False) + "\n"
    # Write to a temporary file first, so a crash never leaves half a glossary behind.
    handle, temp_name = tempfile.mkstemp(dir=str(path.parent), prefix=".glossary-", suffix=".tmp")
    try:
        with os.fdopen(handle, "w", encoding="utf-8") as file:
            file.write(text)
        os.replace(temp_name, path)
    except BaseException:
        if os.path.exists(temp_name):
            os.unlink(temp_name)
        raise


def _edit(path: Path, change) -> Glossary:
    """Loads the raw JSON, applies change(data), validates the result, and only then saves it."""
    path = Path(path)
    if not path.exists():
        write_starter_glossary(path)
    data = _parse(path.read_text(encoding="utf-8"), path)
    change(data)
    glossary = build_glossary(data, path)
    _write_json(path, data)
    return glossary


def add_term(path: Path, shorthand: str, phrase: str) -> Glossary:
    norm = _normalize_key(shorthand)
    if not norm or not phrase.strip():
        raise ValueError("Both the shorthand and the phrase are needed.")

    def change(data):
        terms = data.setdefault("terms", {})
        for key, value in terms.items():
            if _normalize_key(key) == norm:
                if value == phrase.strip():
                    return
                raise ValueError(f'"{shorthand}" already means "{value}". Edit the glossary file to change it.')
        terms[norm] = phrase.strip()
    return _edit(path, change)


def add_tool(path: Path, code: str, name: str, phrase: Optional[str] = None) -> Glossary:
    code = code.strip()
    if not code or not name.strip():
        raise ValueError("A tool needs a code and a name.")

    def change(data):
        tools = data.setdefault("tools", {})
        for key, value in tools.items():
            if key.lower() == code.lower():
                raise ValueError(f'Tool code "{code}" already exists ({value.get("name", "?")}). Edit the glossary file to change it.')
        entry = {"name": name.strip()}
        if phrase and phrase.strip():
            entry["phrase"] = phrase.strip()
        tools[code] = entry
    return _edit(path, change)


def add_words(path: Path, words: Iterable[str]) -> Glossary:
    new = [w.strip().lower() for w in words if w and w.strip()]
    if not new:
        raise ValueError("Give at least one word.")

    def change(data):
        existing = data.setdefault("words", [])
        known = {w.lower() for w in existing if isinstance(w, str)}
        for word in new:
            if word not in known:
                existing.append(word)
                known.add(word)
    return _edit(path, change)


def describe(glossary: Glossary) -> str:
    return (f"{len(glossary.terms)} terms, {len(glossary.tools)} tools, {len(glossary.units)} units, "
            f"{len(glossary.count_words)} count words, {len(glossary.safety_words)} safety words, "
            f"{len(glossary.words)} extra words; part format \"{glossary.settings['part_format']}\", "
            f"part number required: {glossary.settings['require_part_number']}")
