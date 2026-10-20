"""Tests for the command-line interface: flags, batch mode, --init, --config, --pdf."""

from __future__ import annotations

import io
import json
import os
import shutil
import sys
import tempfile
import unittest
from contextlib import redirect_stderr, redirect_stdout
from pathlib import Path
from unittest import mock

from pptx import Presentation

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(Path(__file__).resolve().parent))

import generate_work_instruction as wig  # noqa: E402
from docx_fixture import make_docx, png_bytes  # noqa: E402
from wig_scaffold import SAMPLE_ASSEMBLY, scaffold_project  # noqa: E402

GLOBALS = [
    "DESIGN_FOLDER", "PART_IMAGE_FOLDER", "TEMPLATE_PATH", "TITLE_SLIDE_INDEX",
    "REVISION_LOG_SLIDE_INDEX", "WELDING_SLIDE_INDEX", "PARTS_TEMPLATE_SLIDE_INDEX",
    "EQUIPMENT_SLIDE_INDEX", "PROCESS_SECTIONS",
]


class CliTestCase(unittest.TestCase):
    def setUp(self):
        saved = {name: getattr(wig, name) for name in GLOBALS}
        self.addCleanup(lambda: [setattr(wig, name, value) for name, value in saved.items()])
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.root = Path(tmp.name)
        self.made = scaffold_project(self.root / "project", sample=True)

    def run_cli(self, *argv, stdin: str | None = None):
        out, err = io.StringIO(), io.StringIO()
        with redirect_stdout(out), redirect_stderr(err):
            if stdin is None:
                code = wig.main(list(argv))
            else:
                with mock.patch("builtins.input", side_effect=stdin.splitlines()):
                    code = wig.main(list(argv))
        return code, out.getvalue(), err.getvalue()

    def paths(self):
        return [
            "--design-folder", str(self.made["design_folder"]),
            "--part-images-folder", str(self.made["part_images_folder"]),
            "--template-path", str(self.made["template_path"]),
        ]

    ANSWERS = ["--initials", "cd", "--welding", "n", "--steps-per-slide", "4", "--initial-release"]

    def deck(self, assembly=SAMPLE_ASSEMBLY, rev="A") -> Path:
        prefix = assembly.split("-")[0]
        return self.made["design_folder"] / prefix / assembly / f"{assembly} REV {rev}.pptx"

    def clone_sample(self, new_number: str) -> None:
        """Copies the sample assembly to another number so batch runs have two."""
        design, catalog = self.made["design_folder"], self.made["part_images_folder"]
        src = design / "12" / SAMPLE_ASSEMBLY
        dst = design / "12" / new_number
        shutil.copytree(src, dst)
        (dst / f"{SAMPLE_ASSEMBLY}.xlsx").rename(dst / f"{new_number}.xlsx")
        shutil.copy(catalog / f"{SAMPLE_ASSEMBLY}.png", catalog / f"{new_number}.png")


class FlagTests(CliTestCase):
    def test_flags_build_a_deck_without_prompts(self):
        code, out, err = self.run_cli("--assembly", SAMPLE_ASSEMBLY, *self.paths(), *self.ANSWERS)
        self.assertEqual(code, 0, err)
        self.assertTrue(self.deck().exists())
        self.assertIn("created successfully", out)
        # Welding "n", Cycle Time and Size and Weight pages blank: cover, log, parts, equipment, steps.
        self.assertEqual(len(Presentation(self.deck()).slides), 5)

    def test_missing_answers_are_prompted(self):
        code, _, err = self.run_cli(
            "--assembly", SAMPLE_ASSEMBLY, *self.paths(), stdin="CD\nN\n4\nI\n"
        )
        self.assertEqual(code, 0, err)
        self.assertTrue(self.deck().exists())

    def test_assembly_number_is_prompted_when_not_given(self):
        code, _, err = self.run_cli(*self.paths(), *self.ANSWERS, stdin=f"{SAMPLE_ASSEMBLY}\n")
        self.assertEqual(code, 0, err)
        self.assertTrue(self.deck().exists())

    def test_revision_by_flags_archives_the_old_deck(self):
        self.run_cli("--assembly", SAMPLE_ASSEMBLY, *self.paths(), *self.ANSWERS)
        code, out, err = self.run_cli(
            "--assembly", SAMPLE_ASSEMBLY, *self.paths(),
            "--initials", "CD", "--welding", "N", "--steps-per-slide", "4",
            "--revision", "b", "--change-description", "Changed torque note",
        )
        self.assertEqual(code, 0, err)
        self.assertTrue(self.deck(rev="B").exists())
        self.assertFalse(self.deck(rev="A").exists())
        self.assertIn("Moved to CANCELLED", out)

    def test_revision_needs_a_change_description(self):
        code, _, err = self.run_cli("--assembly", SAMPLE_ASSEMBLY, *self.paths(), "--revision", "B")
        self.assertEqual(code, 1)
        self.assertIn("--change-description", err)

    def test_missing_assembly_reports_an_error_and_fails(self):
        code, _, err = self.run_cli("--assembly", "99-999-99", *self.paths(), *self.ANSWERS)
        self.assertEqual(code, 1)
        self.assertIn("99-999-99", err)

    def test_invalid_flag_value_is_rejected_by_argparse(self):
        with redirect_stderr(io.StringIO()), self.assertRaises(SystemExit):
            wig.main(["--steps-per-slide", "3"])


class BatchTests(CliTestCase):
    def test_several_assemblies_build_in_one_run(self):
        self.clone_sample("12-345-68")
        code, out, err = self.run_cli(
            "--assembly", SAMPLE_ASSEMBLY, "12-345-68", *self.paths(), *self.ANSWERS
        )
        self.assertEqual(code, 0, err)
        self.assertTrue(self.deck().exists())
        self.assertTrue(self.deck("12-345-68").exists())
        self.assertIn("All 2 assemblies built", out)

    def test_one_bad_assembly_does_not_stop_the_others(self):
        code, _, err = self.run_cli(
            "--assembly", "99-999-99", SAMPLE_ASSEMBLY, *self.paths(), *self.ANSWERS
        )
        self.assertEqual(code, 1)
        self.assertTrue(self.deck().exists())
        self.assertIn("1 of 2 assemblies failed", err)
        self.assertIn("[99-999-99]", err)

    def test_assembly_file_skips_blank_lines_and_comments(self):
        listing = self.root / "list.txt"
        listing.write_text(f"# batch\n\n{SAMPLE_ASSEMBLY}\n  \n", encoding="utf-8")
        self.assertEqual(wig.read_assembly_file(listing), [SAMPLE_ASSEMBLY])
        code, _, err = self.run_cli("--assembly-file", str(listing), *self.paths(), *self.ANSWERS)
        self.assertEqual(code, 0, err)
        self.assertTrue(self.deck().exists())

    def test_duplicate_numbers_are_built_once(self):
        code, out, err = self.run_cli(
            "--assembly", SAMPLE_ASSEMBLY, SAMPLE_ASSEMBLY, *self.paths(), *self.ANSWERS
        )
        self.assertEqual(code, 0, err)
        self.assertEqual(out.count("created successfully"), 1)


class InitTests(CliTestCase):
    def test_init_creates_a_skeleton(self):
        target = self.root / "fresh"
        code, out, err = self.run_cli("--init", str(target))
        self.assertEqual(code, 0, err)
        for expected in ["Design Folder", "Part_Catalog_Images", "Templates/XX-XXX-XX REV XX.pptx",
                         "Templates/assembly-workbook-template.xlsx"]:
            self.assertTrue((target / expected).exists(), expected)
        self.assertEqual(len(list((target / "Design Folder").iterdir())), 0)
        self.assertEqual(len(Presentation(target / "Templates" / "XX-XXX-XX REV XX.pptx").slides), 10)

    def test_init_sample_is_immediately_buildable(self):
        target = self.root / "sampled"
        self.assertEqual(self.run_cli("--init", str(target), "--sample")[0], 0)
        code, _, err = self.run_cli(
            "--assembly", SAMPLE_ASSEMBLY,
            "--design-folder", str(target / "Design Folder"),
            "--part-images-folder", str(target / "Part_Catalog_Images"),
            "--template-path", str(target / "Templates" / "XX-XXX-XX REV XX.pptx"),
            *self.ANSWERS,
        )
        self.assertEqual(code, 0, err)

    def test_sample_without_init_is_an_error(self):
        code, _, err = self.run_cli("--sample")
        self.assertEqual(code, 1)
        self.assertIn("--init", err)


class ConfigTests(CliTestCase):
    def write(self, data) -> Path:
        path = self.root / "wig.json"
        path.write_text(json.dumps(data), encoding="utf-8")
        return path

    def test_slide_positions_and_relative_paths_apply(self):
        path = self.write({
            "slides": {"title": 0, "revision_log": 1, "welding": 2, "parts": 3,
                       "equipment": 4, "assembly": 5, "testing": 7, "packaging": 8},
            "paths": {"template_path": "t/deck.pptx"},
        })
        wig.load_config(path)
        self.assertEqual(wig.WELDING_SLIDE_INDEX, 2)
        self.assertEqual(wig.PARTS_TEMPLATE_SLIDE_INDEX, 3)
        self.assertEqual({s.sheet: s.template_index for s in wig.PROCESS_SECTIONS},
                         {"Assembly": 5, "Testing": 7, "Packaging": 8})
        self.assertEqual(wig.TEMPLATE_PATH, (self.root / "t" / "deck.pptx").resolve())

    def test_a_template_with_a_different_order_builds(self):
        """Drop the two optional pages so every later slide moves up by one."""
        template = self.made["template_path"]
        prs = Presentation(template)
        wig.delete_slide(prs, 7)
        wig.delete_slide(prs, 2)
        prs.save(template)
        config = self.write({"slides": {"welding": 2, "parts": 3, "equipment": 4,
                                        "assembly": 5, "testing": 6, "packaging": 7}})
        code, _, err = self.run_cli("--assembly", SAMPLE_ASSEMBLY, "--config", str(config),
                                    *self.paths(), *self.ANSWERS)
        self.assertEqual(code, 0, err)
        self.assertEqual(len(Presentation(self.deck()).slides), 5)

    def test_bad_configs_are_rejected_with_a_reason(self):
        cases = {
            "unknown top-level key": ({"slidez": {}}, "Unknown config key"),
            "unknown slide": ({"slides": {"cover": 0}}, "Unknown slide key"),
            "not a number": ({"slides": {"parts": "4"}}, "whole number"),
            "negative": ({"slides": {"parts": -1}}, "whole number"),
            "shared position": ({"slides": {"parts": 5}}, "own position"),
            "wrong order": ({"slides": {"parts": 10, "equipment": 4}}, "must appear in this order"),
            "title after welding": ({"slides": {"title": 4, "welding": 3, "parts": 5, "equipment": 6,
                                                "assembly": 7, "testing": 8, "packaging": 9}}, "before the welding"),
            "bad path key": ({"paths": {"nope": "x"}}, "may only contain"),
        }
        for label, (config, message) in cases.items():
            with self.subTest(label):
                with self.assertRaises(ValueError) as caught:
                    wig.load_config(self.write(config))
                self.assertIn(message, str(caught.exception))

    def test_invalid_config_leaves_defaults_untouched(self):
        before = wig.PARTS_TEMPLATE_SLIDE_INDEX
        with self.assertRaises(ValueError):
            wig.load_config(self.write({"slides": {"parts": 10, "equipment": 4}}))
        self.assertEqual(wig.PARTS_TEMPLATE_SLIDE_INDEX, before)

    def test_unreadable_config_is_an_error(self):
        code, _, err = self.run_cli("--config", str(self.root / "missing.json"), "--assembly", "1-1-1")
        self.assertEqual(code, 1)
        self.assertIn("Could not read config", err)


class ImportCliTests(CliTestCase):
    def page(self, steps=4) -> Path:
        rows = [(f"Step {i} text", png_bytes((40 * i, 90, 160))) for i in range(1, steps + 1)]
        return make_docx(self.root / "page.docx", rows)

    def test_import_then_build_end_to_end(self):
        code, out, err = self.run_cli("--import-onenote", str(self.page()), "--assembly", SAMPLE_ASSEMBLY,
                                      "--replace", *self.paths())
        self.assertEqual(code, 0, err)
        self.assertIn("Imported 4 Assembly step(s)", out)
        code, _, err = self.run_cli("--assembly", SAMPLE_ASSEMBLY, *self.paths(), *self.ANSWERS)
        self.assertEqual(code, 0, err)
        slide = next(s for s in Presentation(self.deck()).slides
                     if any(sh.has_text_frame and "Step 1 text" in sh.text_frame.text for sh in s.shapes))
        badges = [sh.text_frame.text for sh in slide.shapes if sh.name.startswith(wig.STEP_BADGE_NAME)]
        self.assertEqual(sorted(badges), ["1", "2", "3", "4"])

    def test_import_refuses_to_overwrite_the_sample_steps_without_replace(self):
        code, _, err = self.run_cli("--import-onenote", str(self.page()), "--assembly", SAMPLE_ASSEMBLY, *self.paths())
        self.assertEqual(code, 1)
        self.assertIn("--replace", err)

    def test_import_needs_exactly_one_assembly(self):
        code, _, err = self.run_cli("--import-onenote", str(self.page()), "--assembly", "1-1-1", "2-2-2", *self.paths())
        self.assertEqual(code, 1)
        self.assertIn("exactly one", err)

    def test_import_of_a_missing_file_is_a_clear_error(self):
        code, _, err = self.run_cli("--import-onenote", str(self.root / "nope.docx"), "--assembly", "9-9-9", *self.paths())
        self.assertEqual(code, 1)
        self.assertIn("not a readable Word", err)

    def test_testing_and_packaging_pages_import_and_build_into_the_deck(self):
        for section, count in (("testing", 2), ("packaging", 3)):
            code, out, err = self.run_cli("--import-onenote", str(self.page(count)), "--assembly", SAMPLE_ASSEMBLY,
                                          "--section", section, *self.paths())
            self.assertEqual(code, 0, err)
            self.assertIn(f"Imported {count} {section.title()} step(s)", out)
        code, _, err = self.run_cli("--assembly", SAMPLE_ASSEMBLY, *self.paths(), *self.ANSWERS)
        self.assertEqual(code, 0, err)
        prs = Presentation(self.deck())
        texts = [sh.text_frame.text for slide in prs.slides for sh in slide.shapes if sh.has_text_frame]
        self.assertIn("Testing Steps 1\u20132", texts)
        self.assertIn("Packaging Steps 1\u20133", texts)

    def test_import_into_a_new_assembly_creates_the_workbook(self):
        code, out, err = self.run_cli("--import-onenote", str(self.page(2)), "--assembly", "55-100-01", *self.paths())
        self.assertEqual(code, 0, err)
        self.assertIn("created", out)
        self.assertTrue((self.made["design_folder"] / "55" / "55-100-01" / "55-100-01.xlsx").exists())
        self.assertTrue((self.made["design_folder"] / "55" / "55-100-01" / "Assembly" / "2.jpg").exists())


class PdfTests(CliTestCase):
    def test_pdf_without_libreoffice_fails_before_building(self):
        with mock.patch.dict(os.environ, {"SOFFICE_BIN": str(self.root / "no-such-soffice")}):
            code, _, err = self.run_cli("--assembly", SAMPLE_ASSEMBLY, *self.paths(), *self.ANSWERS, "--pdf")
        self.assertEqual(code, 1)
        self.assertIn("LibreOffice", err)
        self.assertFalse(self.deck().exists())

    @unittest.skipUnless(wig.find_soffice(), "LibreOffice not installed")
    def test_pdf_export_writes_a_pdf_next_to_the_deck(self):
        code, out, err = self.run_cli("--assembly", SAMPLE_ASSEMBLY, *self.paths(), *self.ANSWERS, "--pdf")
        self.assertEqual(code, 0, err)
        pdf = self.deck().with_suffix(".pdf")
        self.assertTrue(pdf.exists())
        self.assertTrue(pdf.read_bytes().startswith(b"%PDF"))


if __name__ == "__main__":
    unittest.main()
