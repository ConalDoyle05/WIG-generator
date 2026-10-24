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

from openpyxl import load_workbook
from pptx import Presentation
from pptx.util import Inches

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(Path(__file__).resolve().parent))

import generate_work_instruction as wig  # noqa: E402
from docx_fixture import make_docx, png_bytes  # noqa: E402
from wig_scaffold import SAMPLE_ASSEMBLY, scaffold_project  # noqa: E402

GLOBALS = list(wig.CONFIGURABLE_GLOBALS) + ["FOOTER_ENABLED", "ALLOW_DRAFT"]


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

    def test_steps_per_slide_must_be_a_known_layout(self):
        code, _, err = self.run_cli("--assembly", SAMPLE_ASSEMBLY, *self.paths(), "--initials", "CD",
                                    "--welding", "N", "--steps-per-slide", "3", "--initial-release")
        self.assertEqual(code, 1)
        self.assertIn("--steps-per-slide must be one of: 2, 4, 8", err)
        with redirect_stderr(io.StringIO()), self.assertRaises(SystemExit):
            wig.main(["--steps-per-slide", "four"])


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


class PolishHelpers(CliTestCase):
    """Shared steps for import -> polish -> approve -> build. Holds no tests itself."""

    NOTES = ["Ins P-003 on fixture", "Pl P-004 on P-003 w P-005", "Tight w W1/2 P-001", "Insp for gaps NOPART"]

    def page(self, notes=None) -> Path:
        notes = notes or self.NOTES
        return make_docx(self.root / "page.docx", [(n, png_bytes((40 * i, 90, 160))) for i, n in enumerate(notes, 1)])

    def import_and_polish(self, notes=None):
        return self.run_cli("--import-onenote", str(self.page(notes)), "--assembly", SAMPLE_ASSEMBLY,
                            "--replace", "--polish", *self.paths())

    def workbook(self):
        return self.made["design_folder"] / "12" / SAMPLE_ASSEMBLY / f"{SAMPLE_ASSEMBLY}.xlsx"

    def statuses(self):
        ws = load_workbook(self.workbook())["Assembly"]
        return [row[3] for row in ws.iter_rows(min_row=2, values_only=True)]

    def approve_all(self):
        wb = load_workbook(self.workbook())
        ws = wb["Assembly"]
        for r in range(2, ws.max_row + 1):
            if ws.cell(row=r, column=1).value is not None:
                ws.cell(row=r, column=4).value = "Approved"
        wb.save(self.workbook())



class PolishCliTests(PolishHelpers):
    """Import -> polish -> approve -> build, as Conal would run it."""

    def test_import_and_polish_in_one_command(self):
        code, out, err = self.import_and_polish()
        self.assertEqual(code, 0, err)
        self.assertIn("Polished 12-345-67", out)
        self.assertEqual(self.statuses(), ["Ready", "Ready", "Ready", "Ready"])
        ws = load_workbook(self.workbook())["Assembly"]
        self.assertEqual(ws.cell(row=4, column=2).value, 'Tighten with a 1/2" wrench M5 x 16 socket head screw (P-001).')
        self.assertTrue((self.root / "project" / "glossary.json").exists())

    def test_unapproved_steps_block_a_real_build(self):
        self.import_and_polish()
        code, _, err = self.run_cli("--assembly", SAMPLE_ASSEMBLY, *self.paths(), *self.ANSWERS)
        self.assertEqual(code, 1)
        self.assertIn("4 step(s) are not approved yet: 1, 2, 3, 4", err)
        self.assertIn("--draft", err)
        self.assertFalse(self.deck().exists())

    def test_draft_build_is_watermarked_named_and_replaceable(self):
        self.import_and_polish(["HARD route the cable", "Tght w W1/2 P-001"])
        for _ in range(2):   # a second draft replaces the first
            code, out, err = self.run_cli("--assembly", SAMPLE_ASSEMBLY, *self.paths(), *self.ANSWERS, "--draft")
            self.assertEqual(code, 0, err)
        draft = self.deck().with_name(f"{SAMPLE_ASSEMBLY} REV A DRAFT.pptx")
        self.assertTrue(draft.exists())
        self.assertFalse(self.deck().exists())
        prs = Presentation(draft)
        for slide in prs.slides:
            self.assertTrue(any(sh.name == "WIG Draft Watermark" for sh in slide.shapes))
        texts = [sh.text_frame.text for slide in prs.slides for sh in slide.shapes if sh.has_text_frame]
        self.assertIn(wig.HARD_PLACEHOLDER, texts)
        total = len(prs.slides)
        for number, slide in enumerate(prs.slides, start=1):
            marks = [sh for sh in slide.shapes if sh.name == "WIG Draft Watermark"]
            self.assertEqual([m.text_frame.text for m in marks], [wig.DRAFT_WATERMARK])
            # The page-number routine must not have mistaken anything for the page number.
            self.assertIn(f"Page {number} of {total}", [sh.text_frame.text for sh in slide.shapes if sh.has_text_frame])

    def test_after_approval_the_real_build_works_and_ignores_drafts(self):
        self.import_and_polish()
        self.run_cli("--assembly", SAMPLE_ASSEMBLY, *self.paths(), *self.ANSWERS, "--draft")
        self.approve_all()
        code, out, err = self.run_cli("--assembly", SAMPLE_ASSEMBLY, *self.paths(), *self.ANSWERS)
        self.assertEqual(code, 0, err)
        self.assertTrue(self.deck().exists())
        prs = Presentation(self.deck())
        self.assertFalse(any(sh.name == "WIG Draft Watermark" for slide in prs.slides for sh in slide.shapes))
        # A revision archives the real REV A deck, never the draft.
        code, out, err = self.run_cli("--assembly", SAMPLE_ASSEMBLY, *self.paths(), "--initials", "CD", "--welding",
                                      "N", "--steps-per-slide", "4", "--revision", "B", "--change-description", "x")
        self.assertEqual(code, 0, err)
        cancelled = self.deck().parent / "CANCELLED"
        self.assertEqual(sorted(p.name for p in cancelled.iterdir()),
                         [f"{SAMPLE_ASSEMBLY} REV A.pptx", f"{SAMPLE_ASSEMBLY} REV A.wig.json"])
        self.assertTrue(self.deck().with_name(f"{SAMPLE_ASSEMBLY} REV A DRAFT.pptx").exists())

    def test_reimport_keeps_approvals_for_unchanged_notes(self):
        self.import_and_polish()
        self.approve_all()
        changed = list(self.NOTES)
        changed[1] = "Pl P-004 on P-003"
        code, out, err = self.run_cli("--import-onenote", str(self.page(changed)), "--assembly", SAMPLE_ASSEMBLY,
                                      "--replace", "--polish", *self.paths())
        self.assertEqual(code, 0, err)
        self.assertIn("Kept the description, status and approval of 3 step(s)", out)
        self.assertEqual(self.statuses(), ["Approved", "Ready", "Approved", "Approved"])

    def test_glossary_commands(self):
        glossary = str(self.root / "g.json")
        code, out, err = self.run_cli("--glossary", glossary, "--glossary-check")
        self.assertEqual(code, 1)
        self.assertIn("No glossary", err)
        self.assertEqual(self.run_cli("--glossary", glossary, "--glossary-add", "snug", "snug down")[0], 0)
        self.assertEqual(self.run_cli("--glossary", glossary, "--glossary-add-tool", "W12", "12 mm wrench")[0], 0)
        self.assertEqual(self.run_cli("--glossary", glossary, "--glossary-add-word", "gizmo")[0], 0)
        code, out, err = self.run_cli("--glossary", glossary, "--glossary-check")
        self.assertEqual(code, 0, err)
        self.assertIn("Glossary OK", out)
        code, _, err = self.run_cli("--glossary", glossary, "--glossary-add-tool", "W13")
        self.assertEqual(code, 1)
        self.assertIn("CODE NAME", err)

    def test_init_writes_a_starter_glossary(self):
        target = self.root / "fresh"
        self.assertEqual(self.run_cli("--init", str(target))[0], 0)
        self.assertTrue((target / "glossary.json").exists())

    def test_polish_without_a_workbook_is_a_clear_error(self):
        code, _, err = self.run_cli("--polish", "--assembly", "77-777-77", *self.paths())
        self.assertEqual(code, 1)
        self.assertIn("No workbook", err)


class ToolsPageTests(PolishHelpers):
    def test_tools_used_in_steps_are_listed_on_the_equipment_page(self):
        self.import_and_polish(["Tight w W1/2 P-001", "Seat P-004 w PH2 then W1/2"])
        page = make_docx(self.root / "testing.docx", [("Chk torque w TQS P-001", png_bytes())])
        self.run_cli("--import-onenote", str(page), "--assembly", SAMPLE_ASSEMBLY, "--section", "testing",
                     "--polish", *self.paths())
        wb = load_workbook(self.workbook())
        for sheet in ("Assembly", "Testing"):
            ws = wb[sheet]
            for r in range(2, ws.max_row + 1):
                if ws.cell(row=r, column=1).value is not None:
                    ws.cell(row=r, column=4).value = "Approved"
        wb.save(self.workbook())
        code, _, err = self.run_cli("--assembly", SAMPLE_ASSEMBLY, *self.paths(), *self.ANSWERS)
        self.assertEqual(code, 0, err)
        texts = [sh.text_frame.text for slide in Presentation(self.deck()).slides for sh in slide.shapes
                 if sh.has_text_frame and sh.text_frame.text.startswith("Tools")]
        self.assertEqual(texts, ['Tools\x0b- 1/2" wrench (W1/2)\x0b- #2 Phillips screwdriver (PH2)'
                                 '\x0b- Torque screwdriver (TQS)\x0b- Torque driver\x0b- Assembly fixture'])


class FooterTests(CliTestCase):
    def slides(self, *extra):
        code, _, err = self.run_cli("--assembly", SAMPLE_ASSEMBLY, *self.paths(), *self.ANSWERS, *extra)
        self.assertEqual(code, 0, err)
        name = f"{SAMPLE_ASSEMBLY} REV A{' DRAFT' if '--draft' in extra else ''}.pptx"
        return list(Presentation(self.deck().with_name(name)).slides)

    @staticmethod
    def named(slide, name):
        return [sh for sh in slide.shapes if sh.name == name]

    def test_every_slide_carries_the_document_id_and_page_n_of_n(self):
        slides = self.slides()
        for number, slide in enumerate(slides, start=1):
            footer = self.named(slide, wig.FOOTER_NAME)
            self.assertEqual(len(footer), 1)
            self.assertRegex(footer[0].text_frame.text, rf"^WI {SAMPLE_ASSEMBLY}  REV A  \|  \d\d/\d\d/\d{{4}}$")
            page = self.named(slide, wig.PAGE_NUMBER_NAME)
            self.assertEqual([p.text_frame.text for p in page], [f"Page {number} of {len(slides)}"])
            for shape in footer + page:
                self.assertGreaterEqual(shape.top, Inches(wig.FOOTER_TOP) - 1)
                self.assertLessEqual(shape.top + shape.height, Inches(7.5))

    def test_footer_sits_below_every_content_shape(self):
        for slide in self.slides():
            for shape in slide.shapes:
                if shape.name in (wig.FOOTER_NAME, wig.PAGE_NUMBER_NAME):
                    continue
                self.assertLessEqual(shape.top + shape.height, Inches(wig.FOOTER_TOP) + 1, shape.name)

    def test_draft_notice_sits_between_the_footer_texts(self):
        for slide in self.slides("--draft"):
            left, = self.named(slide, wig.FOOTER_NAME)
            mark, = self.named(slide, "WIG Draft Watermark")
            page, = self.named(slide, wig.PAGE_NUMBER_NAME)
            self.assertLessEqual(left.left + left.width, mark.left)
            self.assertLessEqual(mark.left + mark.width, page.left)

    def test_no_footer_keeps_the_plain_page_number(self):
        slides = self.slides("--no-footer")
        for number, slide in enumerate(slides, start=1):
            self.assertEqual(self.named(slide, wig.FOOTER_NAME), [])
            self.assertEqual([p.text_frame.text for p in self.named(slide, wig.PAGE_NUMBER_NAME)], [str(number)])


class CheckTests(CliTestCase):
    def check(self, *extra):
        return self.run_cli("--check", "--assembly", SAMPLE_ASSEMBLY, *self.paths(), *extra)

    def test_a_ready_assembly_passes_and_nothing_is_written(self):
        workbook = self.made["assembly_dir"] / f"{SAMPLE_ASSEMBLY}.xlsx"
        before = sorted(p.name for p in self.made["assembly_dir"].iterdir()), workbook.stat().st_mtime_ns
        code, out, err = self.check()
        self.assertEqual(code, 0, err + out)
        self.assertIn("Ready to build.", out)
        self.assertIn("5 part(s)", out)
        self.assertIn("Assembly: 4 step(s)", out)
        self.assertEqual((sorted(p.name for p in self.made["assembly_dir"].iterdir()), workbook.stat().st_mtime_ns), before)

    def test_problems_are_listed_one_by_one(self):
        (self.made["assembly_dir"] / "Assembly" / "3.png").unlink()
        (self.made["part_images_folder"] / "P-002.png").unlink()
        code, out, _ = self.check()
        self.assertEqual(code, 1)
        self.assertIn("NOT READY", out)
        self.assertIn("P-002", out)
        self.assertIn("Missing Assembly Images", out)

    def test_unapproved_steps_fail_the_check_but_pass_as_a_draft(self):
        page = make_docx(self.root / "page.docx", [("Tight w W1/2 P-001", png_bytes())])
        self.run_cli("--import-onenote", str(page), "--assembly", SAMPLE_ASSEMBLY, "--replace", "--polish", *self.paths())
        code, out, _ = self.check()
        self.assertEqual(code, 1)
        self.assertIn("not approved yet", out)
        self.assertIn("Assembly status: 1 Ready", out)
        code, out, _ = self.check("--draft")
        self.assertEqual(code, 0, out)
        self.assertIn("Ready to build a draft.", out)

    def test_hard_steps_get_their_own_message(self):
        page = make_docx(self.root / "page.docx", [("HARD route the cable", png_bytes())])
        self.run_cli("--import-onenote", str(page), "--assembly", SAMPLE_ASSEMBLY, "--replace", "--polish", *self.paths())
        _, out, _ = self.check()
        self.assertIn("step 1 is marked HARD", out)

    def test_a_short_template_and_a_missing_assembly_are_reported(self):
        prs = Presentation(self.made["template_path"])
        wig.delete_slide(prs, len(prs.slides) - 1)
        prs.save(self.made["template_path"])
        code, out, _ = self.run_cli("--check", "--assembly", SAMPLE_ASSEMBLY, "99-999-99", *self.paths())
        self.assertEqual(code, 1)
        self.assertIn("the slide layout expects at least 10", out)
        self.assertIn("Missing assembly folder for 99-999-99", out)
        self.assertIn("0 of 2 assemblies ready", out)

    def test_existing_deck_is_named(self):
        self.run_cli("--assembly", SAMPLE_ASSEMBLY, *self.paths(), *self.ANSWERS)
        _, out, _ = self.check()
        self.assertIn(f"Current deck: {SAMPLE_ASSEMBLY} REV A.pptx", out)


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
