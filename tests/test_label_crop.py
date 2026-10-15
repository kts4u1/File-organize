from conftest import make_tool_photo
from toolcal.label_crop import Box, crop_label, find_label

TRUTH = Box(250, 700, 650, 966)


def test_finds_label_not_sticky_note():
    box = find_label(make_tool_photo())
    assert box is not None
    assert box.iou(TRUTH) > 0.85


def test_tilted_photo_is_straightened():
    img = make_tool_photo(angle=7)
    box = find_label(img)
    assert box is not None
    assert 4 < abs(box.angle) < 10
    crop = crop_label(img, box, pad_ratio=0.0)
    w, h = crop.size
    assert abs(w / h - 400 / 266) < 0.15  # back to the label's own proportions


def test_hint_used_when_nothing_detected():
    from PIL import Image

    plain = Image.new("RGB", (800, 600), (200, 150, 40))
    hint = Box(100, 100, 300, 250)
    assert find_label(plain) is None
    assert find_label(plain, hint) == hint


def test_crop_has_margin_and_stays_inside():
    img = make_tool_photo()
    crop = crop_label(img, TRUTH, pad_ratio=0.1)
    assert crop.size == (480, 320)
    edge = crop_label(img, Box(0, 0, 200, 100), pad_ratio=0.5)
    assert edge.size == (300, 150)
