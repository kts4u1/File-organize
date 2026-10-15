"""Low-level, minimal-diff editing of an .xlsx package.

The company workbook carries SharePoint metadata (customXml), autofilters,
frozen panes and dozens of pictures. Round-tripping it through a high-level
library such as openpyxl silently drops some of that, so this module edits
the package at the XML level instead: only the parts that are actually
changed are re-serialised, every other part is written back byte-for-byte.
"""

from __future__ import annotations

import io
import os
import posixpath
import re
import tempfile
import zipfile
from typing import Iterable

from lxml import etree

NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}

REL_BASE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_WORKSHEET = f"{REL_BASE}/worksheet"
REL_DRAWING = f"{REL_BASE}/drawing"
REL_IMAGE = f"{REL_BASE}/image"
REL_STYLES = f"{REL_BASE}/styles"

CT_DRAWING = "application/vnd.openxmlformats-officedocument.drawing+xml"
CONTENT_TYPES = "[Content_Types].xml"


def q(prefix: str, tag: str) -> str:
    """Clark-notation tag name, e.g. q('main', 'row') -> '{ns}row'."""
    return f"{{{NS[prefix]}}}{tag}"


class XlsxPackage:
    """An .xlsx file held in memory as a dict of parts."""

    def __init__(self, data: bytes):
        self._order: list[str] = []
        self._parts: dict[str, bytes] = {}
        self._xml: dict[str, etree._Element] = {}
        self._dirty: set[str] = set()
        with zipfile.ZipFile(io.BytesIO(data)) as zf:
            for info in zf.infolist():
                if info.is_dir():
                    continue
                self._order.append(info.filename)
                self._parts[info.filename] = zf.read(info.filename)

    @classmethod
    def open(cls, path: str | os.PathLike) -> "XlsxPackage":
        with open(path, "rb") as fh:
            return cls(fh.read())

    # ------------------------------------------------------------------ parts
    def exists(self, name: str) -> bool:
        return name in self._parts

    def names(self) -> list[str]:
        return list(self._order)

    def read(self, name: str) -> bytes:
        if name in self._dirty:
            return self._serialize(name)
        return self._parts[name]

    def xml(self, name: str) -> etree._Element:
        """Parsed XML of a part. Call `touch` after modifying it."""
        if name not in self._xml:
            parser = etree.XMLParser(remove_blank_text=False, resolve_entities=False)
            self._xml[name] = etree.fromstring(self._parts[name], parser)
        return self._xml[name]

    def touch(self, name: str) -> None:
        self._dirty.add(name)

    def write(self, name: str, data: bytes) -> None:
        if name not in self._parts:
            self._order.append(name)
        self._parts[name] = data
        self._xml.pop(name, None)
        self._dirty.discard(name)

    def write_xml(self, name: str, root: etree._Element) -> None:
        if name not in self._parts:
            self._order.append(name)
            self._parts[name] = b""
        self._xml[name] = root
        self._dirty.add(name)

    def delete(self, name: str) -> None:
        self._parts.pop(name, None)
        self._xml.pop(name, None)
        self._dirty.discard(name)
        if name in self._order:
            self._order.remove(name)

    def _serialize(self, name: str) -> bytes:
        return etree.tostring(
            self._xml[name], xml_declaration=True, encoding="UTF-8", standalone=True
        )

    # ------------------------------------------------------------------ save
    def to_bytes(self) -> bytes:
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            order = list(self._order)
            if CONTENT_TYPES in order:  # Excel expects it first.
                order.remove(CONTENT_TYPES)
                order.insert(0, CONTENT_TYPES)
            for name in order:
                zf.writestr(name, self.read(name))
        return buf.getvalue()

    def save(self, path: str | os.PathLike) -> None:
        """Atomically write the package to `path` (temp file + rename)."""
        data = self.to_bytes()
        directory = os.path.dirname(os.path.abspath(path))
        fd, tmp = tempfile.mkstemp(prefix=".~toolcal_", suffix=".xlsx", dir=directory)
        try:
            with os.fdopen(fd, "wb") as fh:
                fh.write(data)
            os.replace(tmp, path)
        except BaseException:
            if os.path.exists(tmp):
                os.remove(tmp)
            raise

    # ----------------------------------------------------------- relationships
    @staticmethod
    def rels_name(part: str) -> str:
        directory, base = posixpath.split(part)
        return posixpath.join(directory, "_rels", base + ".rels")

    def rels(self, part: str, create: bool = False) -> etree._Element | None:
        name = self.rels_name(part)
        if not self.exists(name):
            if not create:
                return None
            root = etree.Element(q("rel", "Relationships"), nsmap={None: NS["rel"]})
            self.write_xml(name, root)
            return root
        return self.xml(name)

    @staticmethod
    def resolve(part: str, target: str) -> str:
        if target.startswith("/"):
            return target.lstrip("/")
        return posixpath.normpath(posixpath.join(posixpath.dirname(part), target))

    @staticmethod
    def relative(from_part: str, to_part: str) -> str:
        return posixpath.relpath(to_part, posixpath.dirname(from_part))

    def rel_targets(self, part: str) -> dict[str, tuple[str, str]]:
        """rId -> (type, absolute target part) for a part's relationships."""
        root = self.rels(part)
        out: dict[str, tuple[str, str]] = {}
        if root is None:
            return out
        for rel in root.findall(q("rel", "Relationship")):
            if rel.get("TargetMode") == "External":
                continue
            out[rel.get("Id")] = (rel.get("Type"), self.resolve(part, rel.get("Target")))
        return out

    def add_rel(self, part: str, rel_type: str, target_part: str) -> str:
        root = self.rels(part, create=True)
        used = {rel.get("Id") for rel in root.findall(q("rel", "Relationship"))}
        n = 1
        while f"rId{n}" in used:
            n += 1
        rid = f"rId{n}"
        rel = etree.SubElement(root, q("rel", "Relationship"))
        rel.set("Id", rid)
        rel.set("Type", rel_type)
        rel.set("Target", self.relative(part, target_part))
        self.touch(self.rels_name(part))
        return rid

    def remove_rel(self, part: str, rid: str) -> None:
        root = self.rels(part)
        if root is None:
            return
        for rel in root.findall(q("rel", "Relationship")):
            if rel.get("Id") == rid:
                root.remove(rel)
                self.touch(self.rels_name(part))

    def is_referenced(self, target_part: str) -> bool:
        """True if any relationship in the package still points at the part."""
        for name in self._order:
            if not name.endswith(".rels"):
                continue
            # "dir/_rels/file.ext.rels" describes "dir/file.ext".
            directory, base = posixpath.split(name)
            source = posixpath.join(posixpath.dirname(directory), base[: -len(".rels")])
            for _rtype, target in self.rel_targets(source).values():
                if target == target_part:
                    return True
        return False

    # ----------------------------------------------------------- content types
    def ensure_default_content_type(self, ext: str, content_type: str) -> None:
        root = self.xml(CONTENT_TYPES)
        for el in root.findall(q("ct", "Default")):
            if el.get("Extension", "").lower() == ext.lower():
                return
        el = etree.Element(q("ct", "Default"))
        el.set("Extension", ext)
        el.set("ContentType", content_type)
        # Defaults must come before Overrides.
        first_override = root.find(q("ct", "Override"))
        if first_override is not None:
            first_override.addprevious(el)
        else:
            root.append(el)
        self.touch(CONTENT_TYPES)

    def add_override(self, part: str, content_type: str) -> None:
        root = self.xml(CONTENT_TYPES)
        partname = "/" + part
        for el in root.findall(q("ct", "Override")):
            if el.get("PartName") == partname:
                return
        el = etree.SubElement(root, q("ct", "Override"))
        el.set("PartName", partname)
        el.set("ContentType", content_type)
        self.touch(CONTENT_TYPES)

    def unique_part_name(self, directory: str, stem: str, ext: str) -> str:
        pattern = re.compile(rf"^{re.escape(stem)}(\d+)\.", re.IGNORECASE)
        highest = 0
        for name in self._order:
            d, base = posixpath.split(name)
            if d == directory:
                m = pattern.match(base)
                if m:
                    highest = max(highest, int(m.group(1)))
        return posixpath.join(directory, f"{stem}{highest + 1}.{ext}")

    def iter_parts(self, prefix: str) -> Iterable[str]:
        return [n for n in self._order if n.startswith(prefix)]
