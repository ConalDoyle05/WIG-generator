"""Photo handling shared by the generator and the OneNote import.

Phone photos are often stored sideways with an EXIF flag that says "rotate me",
and some viewers (LibreOffice, for one) ignore that flag, so the photo lands
sideways in the deck. Phone photos are also large, and iPhones save them as
HEIC, which PowerPoint can't embed. Everything here fixes that in memory:
your original files are never changed.

HEIC needs the pillow-heif package (it is in requirements.txt). Without it,
everything else still works and a HEIC photo gives a clear error.
"""

from __future__ import annotations

import io
from pathlib import Path
from typing import Tuple, Union

from PIL import Image, ImageOps, UnidentifiedImageError

try:
    from pillow_heif import register_heif_opener
    register_heif_opener()
    HEIC_SUPPORTED = True
except ImportError:  # optional: only HEIC photos need it
    HEIC_SUPPORTED = False

HEIC_EXTENSIONS = (".heic", ".heif")
# Embedded photos are shrunk to this many pixels on the long side. A step photo is
# at most about 4.5 inches wide on a slide, so this is still well over 300 dpi.
SLIDE_MAX_PX = 2000
# Formats python-pptx can embed as they are. Anything else is converted.
EMBEDDABLE_FORMATS = {"PNG", "JPEG"}
EXIF_ORIENTATION = 0x0112

Source = Union[str, io.BytesIO]


def _heic_error(name: str) -> ValueError:
    return ValueError(
        f"{name} is a HEIC photo, which needs the pillow-heif package.\n"
        "Install it with:  pip install pillow-heif   (it is listed in requirements.txt)"
    )


def open_image(source: Union[Path, bytes], name: str = "image") -> Image.Image:
    """Opens a photo from a path or bytes, with a clear error for unreadable files."""
    if isinstance(source, Path) and source.suffix.lower() in HEIC_EXTENSIONS and not HEIC_SUPPORTED:
        raise _heic_error(source.name)
    try:
        image = Image.open(source if isinstance(source, Path) else io.BytesIO(source))
        image.load()
    except (UnidentifiedImageError, OSError) as exc:
        if not HEIC_SUPPORTED and isinstance(source, bytes) and source[4:12] in (b"ftypheic", b"ftypheix", b"ftypmif1", b"ftypmsf1"):
            raise _heic_error(name)
        raise ValueError(f"{name[:1].upper()}{name[1:]} could not be read: {exc}")
    return image


def normalize(image: Image.Image, max_px: int) -> Image.Image:
    """Turns a photo upright from its EXIF flag, flattens transparency onto white, and shrinks it."""
    image = ImageOps.exif_transpose(image)
    if image.mode in ("RGBA", "LA") or (image.mode == "P" and "transparency" in image.info):
        rgba = image.convert("RGBA")
        flat = Image.new("RGB", rgba.size, (255, 255, 255))
        flat.paste(rgba, mask=rgba.getchannel("A"))
        image = flat
    else:
        image = image.convert("RGB")
    if max(image.size) > max_px:
        image.thumbnail((max_px, max_px), Image.LANCZOS)
    return image


def needs_conversion(image: Image.Image, max_px: int) -> bool:
    """True when a photo can't be embedded exactly as it is: rotated, oversized, or a format PowerPoint can't take."""
    orientation = image.getexif().get(EXIF_ORIENTATION, 1)
    return image.format not in EMBEDDABLE_FORMATS or orientation not in (1, None) or max(image.size) > max_px


def prepare_for_slide(path: Path, max_px: int = SLIDE_MAX_PX) -> Tuple[Source, Tuple[int, int]]:
    """Returns what to embed for a photo, and its upright size in pixels.

    A photo that is already fine is embedded untouched (its path is returned).
    Otherwise an upright, shrunk copy is made in memory: PNG stays PNG, so
    screenshots and diagrams stay crisp; everything else becomes JPEG.
    """
    path = Path(path)
    with open_image(path, path.name) as image:
        if not needs_conversion(image, max_px):
            return str(path), image.size
        keep_png = image.format == "PNG"
        fixed = normalize(image, max_px)
    buffer = io.BytesIO()
    if keep_png:
        fixed.save(buffer, "PNG", optimize=True)
    else:
        fixed.save(buffer, "JPEG", quality=90, optimize=True)
    buffer.seek(0)
    return buffer, fixed.size
