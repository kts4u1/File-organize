"""Read the engineer name, serial number and calibration dates from a photo.

Uses Claude's vision through the Anthropic Python SDK with structured JSON
output. Handwritten Korean names on sticky notes are beyond classic OCR
engines, which is why a vision model is used here.
"""

from __future__ import annotations

import base64
import json
import re
from dataclasses import dataclass, field
from typing import Protocol

from PIL import Image

from .imaging import encode_jpeg

DEFAULT_MODEL = "claude-opus-5-5"
DEFAULT_EFFORT = "medium"
API_IMAGE_SIDE = 1568  # long-edge size the API works best with

# Server-side refusal fallback ("default" routing) is available for these models.
_FALLBACK_MODELS = ("claude-opus-5", "claude-fable-5", "claude-sonnet-5-5")

IMAGE_TYPES = ("tool_photo", "smax_screenshot", "other")

SYSTEM_PROMPT = """\
You read photos for a medical-equipment service team's tool calibration records.

There are two kinds of images:

1. tool_photo - a photo of a measuring tool (multimeter, torque wrench, safety analyzer, ...).
   It usually shows
   - a handwritten sticky note with the owning engineer's Korean name (e.g. 강건욱) and the
     tool's serial number (digits, sometimes with letters such as 44350259SV);
   - a printed white calibration label titled "교정필증 / Calibration Label" with
     성적서번호 (Calibration Certificate No.), 교정일자 (Date of Calibration) and
     차기교정예정일자 (The due date of next Calibration), dates written like 2025.12.29.
   The serial number may instead be printed on the tool's own nameplate.
2. smax_screenshot - a screenshot of the SMAX/Salesforce "Tool Detail" page with fields such as
   Tool Owner, Model Number, Serial/Asset Number and Calibration Due Date (M/D/YYYY).

Transcribe exactly what is written. Do not guess values that are not visible: use an empty
string instead. Give dates as YYYY-MM-DD. For the serial number keep letters and leading zeros,
drop spaces and hyphens, and list other plausible readings of ambiguous handwritten characters
(0/6, 1/7, 5/6, 3/8, 4/9 ...) in serial_alternatives. For a SMAX screenshot, use the Tool Owner
for engineer_name (write it in Korean only if it appears in Korean) and the Serial/Asset Number
for serial_number. label_box is the calibration label's bounding box as fractions (0-1) of the
image width and height as given to you; all zeros when there is no label. rotation is how many
degrees the image must be turned clockwise for its text to read upright ("0" if it already does).
"""

OUTPUT_SCHEMA = {
    "type": "object",
    "properties": {
        "image_type": {"type": "string", "enum": list(IMAGE_TYPES)},
        "engineer_name": {"type": "string"},
        "serial_number": {"type": "string"},
        "serial_alternatives": {"type": "array", "items": {"type": "string"}},
        "certificate_no": {"type": "string"},
        "calibration_date": {"type": "string"},
        "next_due_date": {"type": "string"},
        "model": {"type": "string"},
        "label_box": {
            "type": "object",
            "properties": {
                "left": {"type": "number"},
                "top": {"type": "number"},
                "right": {"type": "number"},
                "bottom": {"type": "number"},
            },
            "required": ["left", "top", "right", "bottom"],
            "additionalProperties": False,
        },
        "rotation": {"type": "string", "enum": ["0", "90", "180", "270"]},
        "confidence": {"type": "string", "enum": ["high", "medium", "low"]},
        "notes": {"type": "string"},
    },
    "required": [
        "image_type",
        "engineer_name",
        "serial_number",
        "serial_alternatives",
        "certificate_no",
        "calibration_date",
        "next_due_date",
        "model",
        "label_box",
        "rotation",
        "confidence",
        "notes",
    ],
    "additionalProperties": False,
}


class RecognitionError(Exception):
    pass


@dataclass
class Recognition:
    image_type: str = "other"
    engineer_name: str = ""
    serial_number: str = ""
    serial_alternatives: list[str] = field(default_factory=list)
    certificate_no: str = ""
    calibration_date: str = ""
    next_due_date: str = ""
    model: str = ""
    label_box: tuple[float, float, float, float] | None = None
    rotation: int = 0  # clockwise degrees that make the image upright
    confidence: str = "low"
    notes: str = ""

    @classmethod
    def from_dict(cls, data: dict) -> "Recognition":
        box = data.get("label_box") or {}
        coords = tuple(float(box.get(k, 0) or 0) for k in ("left", "top", "right", "bottom"))
        image_type = data.get("image_type") if data.get("image_type") in IMAGE_TYPES else "other"
        try:
            rotation = int(data.get("rotation") or 0) % 360
        except (TypeError, ValueError):
            rotation = 0
        return cls(
            image_type=image_type,
            engineer_name=re.sub(r"\s+", "", str(data.get("engineer_name") or "")),
            serial_number=str(data.get("serial_number") or "").strip(),
            serial_alternatives=[str(s).strip() for s in data.get("serial_alternatives") or [] if str(s).strip()],
            certificate_no=str(data.get("certificate_no") or "").strip(),
            calibration_date=str(data.get("calibration_date") or "").strip(),
            next_due_date=str(data.get("next_due_date") or "").strip(),
            model=str(data.get("model") or "").strip(),
            label_box=coords if any(coords) else None,
            rotation=rotation if rotation in (90, 180, 270) else 0,
            confidence=str(data.get("confidence") or "low"),
            notes=str(data.get("notes") or "").strip(),
        )


class Recognizer(Protocol):
    def recognize(self, img: Image.Image, known_names: list[str] | None = None) -> Recognition: ...


class ClaudeRecognizer:
    """Vision recognizer backed by the Claude API."""

    def __init__(self, api_key: str | None = None, model: str = DEFAULT_MODEL,
                 effort: str = DEFAULT_EFFORT, client=None):
        if client is None:
            import anthropic

            # Without an explicit key the SDK resolves ANTHROPIC_API_KEY etc.
            client = anthropic.Anthropic(api_key=api_key) if api_key else anthropic.Anthropic()
        self.client = client
        self.model = model
        self.effort = effort

    def _request_kwargs(self, img: Image.Image, known_names: list[str] | None) -> dict:
        data = base64.standard_b64encode(encode_jpeg(img, API_IMAGE_SIDE, quality=90)).decode("ascii")
        prompt = "Read this image and fill in the fields."
        if known_names:
            prompt += (
                "\nEngineers on the tool list (use the exact spelling if the handwriting matches one of "
                "them; otherwise transcribe what is written): " + ", ".join(sorted(set(known_names)))
            )
        kwargs = dict(
            model=self.model,
            max_tokens=16000,
            system=SYSTEM_PROMPT,
            output_config={
                "effort": self.effort,
                "format": {"type": "json_schema", "schema": OUTPUT_SCHEMA},
            },
            messages=[
                {
                    "role": "user",
                    "content": [
                        {"type": "image", "source": {"type": "base64", "media_type": "image/jpeg", "data": data}},
                        {"type": "text", "text": prompt},
                    ],
                }
            ],
        )
        if self.model.startswith(_FALLBACK_MODELS):
            # If a safety classifier declines, retry server-side on Anthropic's
            # recommended fallback model instead of failing the photo.
            kwargs["betas"] = ["server-side-fallback-2026-07-01"]
            kwargs["fallbacks"] = "default"
        return kwargs

    def recognize(self, img: Image.Image, known_names: list[str] | None = None) -> Recognition:
        import anthropic

        kwargs = self._request_kwargs(img, known_names)
        try:
            if "betas" in kwargs:
                response = self.client.beta.messages.create(**kwargs)
            else:
                response = self.client.messages.create(**kwargs)
        except anthropic.AuthenticationError as exc:
            raise RecognitionError("Claude API 키가 올바르지 않습니다.") from exc
        except anthropic.PermissionDeniedError as exc:
            raise RecognitionError("Claude API 키에 이 모델을 사용할 권한이 없습니다.") from exc
        except anthropic.NotFoundError as exc:
            raise RecognitionError(f"모델을 찾을 수 없습니다: {self.model}") from exc
        except anthropic.RateLimitError as exc:
            raise RecognitionError("Claude API 사용량 한도를 초과했습니다. 잠시 후 다시 시도하세요.") from exc
        except anthropic.BadRequestError as exc:
            raise RecognitionError(f"Claude API 요청 오류: {exc.message}") from exc
        except anthropic.APIStatusError as exc:
            raise RecognitionError(f"Claude API 오류 ({exc.status_code}). 잠시 후 다시 시도하세요.") from exc
        except anthropic.APIConnectionError as exc:
            raise RecognitionError("Claude API에 연결할 수 없습니다. 인터넷/프록시 설정을 확인하세요.") from exc

        if response.stop_reason == "refusal":
            raise RecognitionError("Claude가 이 이미지 처리를 거절했습니다.")
        if response.stop_reason == "max_tokens":
            raise RecognitionError("응답이 잘렸습니다. 다시 시도하세요.")
        text = next((b.text for b in response.content if b.type == "text"), "")
        try:
            return Recognition.from_dict(json.loads(text))
        except (json.JSONDecodeError, TypeError, ValueError) as exc:
            raise RecognitionError(f"인식 결과를 해석할 수 없습니다: {text[:200]}") from exc


class NullRecognizer:
    """No automatic reading (no API key): the user fills in the target row by hand."""

    def recognize(self, img: Image.Image, known_names: list[str] | None = None) -> Recognition:
        return Recognition(image_type="tool_photo", notes="자동 인식 사용 안 함 (API 키 없음)")
