"""Find the white calibration label (교정필증) in a tool photo and crop it.

The label is a near-neutral white rectangle (low chroma, bright) on a
coloured tool body, which separates it well from yellow housings and from
the (light blue) sticky note carrying the engineer's name. Among the bright
neutral blobs the most label-like one is picked: solid, rectangular and
roughly 3:2. An optional hint box (from the vision model) breaks ties and
serves as a fallback when no blob qualifies.
"""

from __future__ import annotations

import math
from dataclasses import dataclass

import cv2
import numpy as np
from PIL import Image

WORK_SIDE = 1200  # detection runs on a copy of at most this size


@dataclass
class Box:
    left: float
    top: float
    right: float
    bottom: float
    angle: float = 0.0  # tilt of the long edge in degrees (image coords, y down)

    @property
    def width(self) -> float:
        return self.right - self.left

    @property
    def height(self) -> float:
        return self.bottom - self.top

    @property
    def center(self) -> tuple[float, float]:
        return (self.left + self.right) / 2, (self.top + self.bottom) / 2

    def iou(self, other: "Box") -> float:
        ix = max(0.0, min(self.right, other.right) - max(self.left, other.left))
        iy = max(0.0, min(self.bottom, other.bottom) - max(self.top, other.top))
        inter = ix * iy
        union = self.width * self.height + other.width * other.height - inter
        return inter / union if union > 0 else 0.0


@dataclass
class _Candidate:
    box: Box
    rect_size: tuple[float, float]  # (long, short) side of the rotated rectangle
    score: float


def box_from_fractions(img: Image.Image, left, top, right, bottom) -> Box | None:
    """Box from 0..1 fractions of the image size (as returned by the vision model)."""
    w, h = img.size
    vals = [min(max(float(v), 0.0), 1.0) for v in (left, top, right, bottom)]
    if vals[2] - vals[0] < 0.02 or vals[3] - vals[1] < 0.02:
        return None
    return Box(vals[0] * w, vals[1] * h, vals[2] * w, vals[3] * h)


def _edge_angle(points: np.ndarray) -> tuple[float, float, float]:
    """(angle of the long edge in degrees, long side, short side) of a rotated rect."""
    edges = [(points[i], points[(i + 1) % 4]) for i in range(4)]
    lengths = [float(np.hypot(*(b - a))) for a, b in edges]
    i = int(np.argmax(lengths))
    (x0, y0), (x1, y1) = edges[i]
    angle = math.degrees(math.atan2(y1 - y0, x1 - x0))
    while angle > 90:
        angle -= 180
    while angle <= -90:
        angle += 180
    return angle, max(lengths), min(lengths)


def _candidates(img: Image.Image) -> list[_Candidate]:
    scale = min(1.0, WORK_SIDE / max(img.size))
    small = img if scale == 1.0 else img.resize(
        (round(img.width * scale), round(img.height * scale)), Image.BILINEAR
    )
    rgb = np.asarray(small, dtype=np.int16)
    value = rgb.max(axis=2)
    chroma = value - rgb.min(axis=2)
    bright = np.percentile(value, 99.5)
    # Neutral (white/grey) and reasonably bright. Chroma tolerance grows with
    # brightness so the shaded side of a label still counts as white.
    mask = ((chroma < 0.09 * value + 4) & (value > 0.55 * bright)).astype(np.uint8) * 255
    h, w = mask.shape
    k_close = max(3, int(min(h, w) * 0.012)) | 1
    mask = cv2.morphologyEx(mask, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (3, 3)))
    mask = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, cv2.getStructuringElement(cv2.MORPH_RECT, (k_close, k_close)))
    contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    out: list[_Candidate] = []
    for contour in contours:
        area = cv2.contourArea(contour)
        frac = area / float(w * h)
        if frac < 0.004 or frac > 0.6:
            continue
        hull_area = cv2.contourArea(cv2.convexHull(contour))
        rect = cv2.minAreaRect(contour)
        rect_area = rect[1][0] * rect[1][1]
        if hull_area <= 0 or rect_area <= 0:
            continue
        solidity = area / hull_area
        rectangularity = hull_area / rect_area
        angle, long_side, short_side = _edge_angle(cv2.boxPoints(rect))
        aspect = long_side / max(short_side, 1.0)
        if solidity < 0.8 or rectangularity < 0.75 or not (1.05 <= aspect <= 2.8):
            continue
        aspect_pref = 1.0 if 1.2 <= aspect <= 2.0 else 0.7
        score = solidity * rectangularity * aspect_pref * (frac ** 0.25)
        x, y, bw, bh = cv2.boundingRect(contour)
        box = Box(x / scale, y / scale, (x + bw) / scale, (y + bh) / scale, angle)
        out.append(_Candidate(box, (long_side / scale, short_side / scale), score))
    return out


def find_label(img: Image.Image, hint: Box | None = None) -> Box | None:
    candidates = _candidates(img)
    if hint is not None:
        for cand in candidates:
            cand.score *= 0.3 + cand.box.iou(hint)
    if candidates:
        best = max(candidates, key=lambda c: c.score)
        if hint is None or best.box.iou(hint) > 0.05:
            return best.box
    return hint


def crop_label(img: Image.Image, box: Box, pad_ratio: float = 0.12) -> Image.Image:
    """Crop the label with some margin, straightening a small tilt."""
    work = img
    left, top, right, bottom = box.left, box.top, box.right, box.bottom
    if 0.7 <= abs(box.angle) <= 20:
        cx, cy = box.center
        work = img.rotate(box.angle, resample=Image.BICUBIC, center=(cx, cy), fillcolor=(255, 255, 255))
        # The axis-aligned bounding box of a tilted rectangle is slightly larger
        # than the rectangle itself; shrink it back after straightening.
        theta = math.radians(abs(box.angle))
        w, h = box.width, box.height
        denom = math.cos(theta) ** 2 - math.sin(theta) ** 2
        if denom > 0.5:
            true_w = (w * math.cos(theta) - h * math.sin(theta)) / denom
            true_h = (h * math.cos(theta) - w * math.sin(theta)) / denom
            if true_w > 0 and true_h > 0:
                left, right = cx - true_w / 2, cx + true_w / 2
                top, bottom = cy - true_h / 2, cy + true_h / 2
    pad_x = (right - left) * pad_ratio
    pad_y = (bottom - top) * pad_ratio
    crop = (
        max(0, int(left - pad_x)),
        max(0, int(top - pad_y)),
        min(work.width, int(math.ceil(right + pad_x))),
        min(work.height, int(math.ceil(bottom + pad_y))),
    )
    return work.crop(crop)
