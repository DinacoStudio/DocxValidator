import tempfile
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, RGBColor

import main
from docx_fixed import create_fixed_docx
from docx_notes import make_notes_docx
from metadata_utils import clear_python_docx_metadata


CORE = "docProps/core.xml"
APP = "docProps/app.xml"


def metadata_text(path, member, local_name):
    with ZipFile(path) as archive:
        root = ET.fromstring(archive.read(member))
    for element in root.iter():
        if element.tag.rsplit("}", 1)[-1] == local_name:
            return element.text or ""
    return None


class CliTests(unittest.TestCase):
    def test_parser_supports_profile_and_copy_switches(self):
        args = main.build_parser().parse_args([
            "input.docx", "--profile", "gost-19-201", "--make-notes",
            "--create-fixed",
        ])
        self.assertEqual("gost-19-201", args.profile)
        self.assertTrue(args.make_notes)
        self.assertTrue(args.create_fixed)

    def test_fixed_copy_is_rechecked_with_selected_profile(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "input.docx"
            source.touch()
            validator = MagicMock()
            validator.validate.side_effect = [[], []]
            with patch("main.profiles.get_rules", return_value=[object()]) as get_rules, \
                    patch("main.DocxValidator", return_value=validator), \
                    patch("main.create_fixed_docx", return_value={}):
                result = main.run([
                    str(source), "--profile", "gost-19-201", "--create-fixed"
                ])
        self.assertEqual(0, result)
        get_rules.assert_called_once_with("gost-19-201")
        self.assertEqual(2, validator.validate.call_count)
        self.assertTrue(
            validator.validate.call_args_list[1].args[0].endswith(
                "input_исправленный.docx"
            )
        )


class MetadataTests(unittest.TestCase):
    def test_cleaner_removes_only_known_python_docx_values(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "document.docx"
            doc = Document()
            doc.core_properties.title = "Пользовательский заголовок"
            doc.core_properties.author = "Пользователь"
            doc.save(path)

            clear_python_docx_metadata(path)

            self.assertEqual(
                "Пользовательский заголовок", metadata_text(path, CORE, "title")
            )
            self.assertEqual(
                "Пользователь", metadata_text(path, CORE, "creator")
            )
            self.assertEqual("", metadata_text(path, CORE, "description"))
            self.assertEqual("", metadata_text(path, APP, "Application"))
            self.assertEqual("", metadata_text(path, APP, "AppVersion"))
            Document(path)  # архив после атомарной перезаписи остаётся валидным

    def test_notes_copy_cleans_fingerprints_but_preserves_user_metadata(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "source.docx"
            target = Path(directory) / "notes.docx"
            doc = Document()
            doc.core_properties.title = "Название"
            doc.add_paragraph("Текст")
            doc.save(source)

            self.assertEqual((0, 0), make_notes_docx(str(source), [], str(target)))

            self.assertEqual("Название", metadata_text(target, CORE, "title"))
            self.assertEqual("", metadata_text(target, CORE, "creator"))
            self.assertEqual("", metadata_text(target, APP, "Application"))
            Document(target)

    def test_fixed_copy_preserves_color_and_sets_heading_grid_and_metadata(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "source.docx"
            target = Path(directory) / "fixed.docx"
            doc = Document()
            doc.core_properties.title = "Сохранить"
            heading = doc.add_paragraph("ВВЕДЕНИЕ")
            heading_run = heading.runs[0]
            heading_run.font.color.rgb = RGBColor(0x12, 0x34, 0x56)
            shading = OxmlElement("w:shd")
            shading.set(qn("w:fill"), "FFFFFF")
            heading_run._r.get_or_add_rPr().append(shading)
            body = doc.add_paragraph("Цветной основной текст")
            body.runs[0].font.color.rgb = RGBColor(0xAA, 0x22, 0x33)
            doc.save(source)

            create_fixed_docx(str(source), str(target))
            fixed = Document(target)

            self.assertAlmostEqual(
                1.25, fixed.paragraphs[0].paragraph_format.first_line_indent.cm,
                places=2,
            )
            self.assertEqual(
                RGBColor(0xAA, 0x22, 0x33), fixed.paragraphs[1].runs[0].font.color.rgb
            )
            self.assertFalse(fixed.element.xpath(".//w:shd[@w:fill='FFFFFF']"))
            section = fixed.sections[0]
            grid = section._sectPr.find(qn("w:docGrid"))
            pitch = int(grid.get(qn("w:linePitch")))
            usable = (
                section.page_height.twips
                - section.top_margin.twips
                - section.bottom_margin.twips
            )
            self.assertEqual(0, usable % pitch)
            self.assertEqual("Сохранить", metadata_text(target, CORE, "title"))
            self.assertNotIn(
                "python-docx",
                metadata_text(target, CORE, "creator")
                + metadata_text(target, CORE, "description"),
            )
            self.assertEqual("", metadata_text(target, APP, "Application"))


if __name__ == "__main__":
    unittest.main()
