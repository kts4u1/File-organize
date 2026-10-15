"""Image loading/encoding helpers (EXIF rotation, HEIC from iPhones, JPEG output)."""

from __future__ import annotations

import io

from PIL import Image, ImageOps

try:  # iPhone photos arrive as .HEIC; support them when pillow-heif is installed.
    from pillow_heif import register_heif_opener

    register_heif_opener()
except ImportError:  # pragma: no cover - optional dependency
    pass

SUPPORTED_EXTENSIONS = ("jpg", "jpeg", "png", "heic", "heif", "webp", "bmp")


class ImageError(Exception):
    pass


def load_image(data: bytes) -> Image.Image:
    """Decode bytes into an upright RGB image."""
    try:
        img = Image.open(io.BytesIO(data))
        img = ImageOps.exif_transpose(img)
    except Exception as exc:  # PIL raises many types for unreadable files
        raise ImageError(f"이미지를 열 수 없습니다: {exc}") from exc
    if img.mode != "RGB":
        background = Image.new("RGB", img.size, (255, 255, 255))
        rgba = img.convert("RGBA")
        background.paste(rgba, mask=rgba.split()[3])
        img = background
    return img


def rotate_clockwise(img: Image.Image, degrees: int) -> Image.Image:
    """Lossless quarter-turn rotation."""
    op = {90: Image.Transpose.ROTATE_270, 180: Image.Transpose.ROTATE_180, 270: Image.Transpose.ROTATE_90}
    return img.transpose(op[degrees]) if degrees in op else img


def shrink(img: Image.Image, max_side: int) -> Image.Image:
    if max(img.size) <= max_side:
        return img
    out = img.copy()
    out.thumbnail((max_side, max_side), Image.LANCZOS)
    return out


def encode_jpeg(img: Image.Image, max_side: int = 1600, quality: int = 85) -> bytes:
    buf = io.BytesIO()
    shrink(img, max_side).save(buf, format="JPEG", quality=quality, optimize=True)
    return buf.getvalue()


def encode_png(img: Image.Image, max_side: int = 1600) -> bytes:
    buf = io.BytesIO()
    shrink(img, max_side).save(buf, format="PNG", optimize=True)
    return buf.getvalue()
