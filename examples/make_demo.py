"""Builds a small synthetic project under ./demo and generates a work instruction from it.

Nothing here is real production data: the parts, steps and images are made up
so you can see the full input -> output flow without setting anything up.

    python examples/make_demo.py            # build the deck
    python examples/make_demo.py --pdf      # also export a PDF (needs LibreOffice)

Then open demo/Design Folder/12/12-345-67/12-345-67 REV A.pptx.
It is the same as running `wig --init demo --sample` followed by one `wig` command.
"""

from __future__ import annotations

import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import generate_work_instruction as wig  # noqa: E402
from wig_scaffold import SAMPLE_ASSEMBLY, scaffold_project  # noqa: E402

DEMO = ROOT / "demo"


def main() -> int:
    if DEMO.exists():
        shutil.rmtree(DEMO)
    made = scaffold_project(DEMO, sample=True)
    argv = [
        "--assembly", SAMPLE_ASSEMBLY,
        "--design-folder", str(made["design_folder"]),
        "--part-images-folder", str(made["part_images_folder"]),
        "--template-path", str(made["template_path"]),
        "--initials", "CD", "--welding", "N", "--steps-per-slide", "4", "--initial-release",
    ]
    if "--pdf" in sys.argv[1:]:
        argv.append("--pdf")
    return wig.main(argv)


if __name__ == "__main__":
    sys.exit(main())
