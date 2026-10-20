"""Builds small Word files shaped like OneNote's page export, for the import tests."""

from __future__ import annotations

import io
import zipfile
from pathlib import Path
from typing import List, Optional, Sequence, Tuple, Union
from xml.sax.saxutils import escape

from PIL import Image

ROOT_NS = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"'
)


def png_bytes(color=(200, 30, 30), size=(60, 40), mode="RGB") -> bytes:
    buffer = io.BytesIO()
    Image.new(mode, size, color).save(buffer, "PNG")
    return buffer.getvalue()


def jpeg_bytes(size=(80, 40), orientation: Optional[int] = None, color=(10, 120, 200)) -> bytes:
    exif = Image.Exif()
    if orientation:
        exif[0x0112] = orientation
    buffer = io.BytesIO()
    Image.new("RGB", size, color).save(buffer, "JPEG", exif=exif)
    return buffer.getvalue()


Row = Tuple[Optional[str], Union[None, bytes, Sequence[bytes]]]


def make_docx(
    path: Path,
    rows: List[Row],
    *,
    padding_rows: int = 4,
    image_first: bool = False,
    title: str = "Test page",
    with_table: bool = True,
) -> Path:
    """rows are (text or None, photo bytes / list of photo bytes / None)."""
    media: List[Tuple[str, bytes]] = []

    def image_cell(photos) -> str:
        if photos is None:
            return "<w:tc><w:p/></w:tc>"
        if isinstance(photos, bytes):
            photos = [photos]
        runs = ""
        for data in photos:
            index = len(media) + 1
            media.append((f"media/image{index}.png", data))
            runs += (
                "<w:r><w:drawing><wp:inline><wp:extent cx=\"3086100\" cy=\"2000000\"/>"
                "<a:graphic><a:graphicData><pic:pic><pic:blipFill>"
                f'<a:blip r:embed="rId{index + 3}"/></pic:blipFill></pic:pic></a:graphicData></a:graphic>'
                "</wp:inline></w:drawing></w:r>"
            )
        return f"<w:tc><w:p>{runs}</w:p></w:tc>"

    def text_cell(text: Optional[str]) -> str:
        if text is None:
            return "<w:tc><w:p/></w:tc>"
        return f"<w:tc><w:p><w:r><w:t>{escape(text)}</w:t></w:r></w:p></w:tc>"

    def row_xml(text, photos) -> str:
        cells = [text_cell(text), image_cell(photos)]
        if image_first:
            cells.reverse()
        return "<w:tr>" + "".join(cells) + "</w:tr>"

    padding = '<w:tr><w:tc><w:p><w:r><w:t> </w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t> </w:t></w:r></w:p></w:tc></w:tr>'
    table = ""
    if with_table:
        table = (
            "<w:tbl><w:tblGrid><w:gridCol w:w=\"988\"/><w:gridCol w:w=\"5020\"/></w:tblGrid>"
            + "".join(row_xml(*row) for row in rows)
            + padding * padding_rows
            + "</w:tbl>"
        )
    document = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {ROOT_NS}><w:body>'
        f"<w:p><w:r><w:t>{escape(title)}</w:t></w:r></w:p>"
        "<w:p><w:r><w:t>Wednesday, July 15, 2026</w:t></w:r></w:p>"
        f"{table}<w:p/></w:body></w:document>"
    )
    rels = "".join(
        f'<Relationship Id="rId{i + 4}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="{name}"/>'
        for i, (name, _data) in enumerate(media)
    )
    path = Path(path)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>')
        archive.writestr("word/document.xml", document)
        archive.writestr(
            "word/_rels/document.xml.rels",
            '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            + rels + "</Relationships>",
        )
        for name, data in media:
            archive.writestr(f"word/{name}", data)
    return path
