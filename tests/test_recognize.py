import json

import anthropic
import httpx2 as httpx
import pytest

from conftest import make_tool_photo
from toolcal.recognize import OUTPUT_SCHEMA, ClaudeRecognizer, RecognitionError

ANSWER = {
    "image_type": "tool_photo",
    "engineer_name": "박 민수",
    "serial_number": "31550075",
    "serial_alternatives": ["31550076"],
    "certificate_no": "DA25-1041-007",
    "calibration_date": "2025-12-29",
    "next_due_date": "2026-12-29",
    "model": "",
    "label_box": {"left": 0.27, "top": 0.58, "right": 0.72, "bottom": 0.8},
    "rotation": "90",
    "confidence": "high",
    "notes": "",
}


def _client(handler):
    return anthropic.Anthropic(api_key="test-key", http_client=httpx.Client(transport=httpx.MockTransport(handler)),
                               max_retries=0)


def _message(text, stop_reason="end_turn"):
    return {
        "id": "msg_test", "type": "message", "role": "assistant", "model": "claude-opus-5-5",
        "content": [{"type": "thinking", "thinking": "", "signature": "sig"}, {"type": "text", "text": text}],
        "stop_reason": stop_reason, "stop_sequence": None,
        "usage": {"input_tokens": 1500, "output_tokens": 200},
    }


def test_request_and_parse():
    seen = {}

    def handler(request: httpx.Request):
        seen["body"] = json.loads(request.content)
        seen["beta"] = request.headers.get("anthropic-beta", "")
        return httpx.Response(200, json=_message(json.dumps(ANSWER, ensure_ascii=False)))

    rec = ClaudeRecognizer(client=_client(handler)).recognize(make_tool_photo(), known_names=["박민수", "이영희"])
    body = seen["body"]
    assert body["model"] == "claude-opus-5-5"
    assert body["output_config"]["format"] == {"type": "json_schema", "schema": OUTPUT_SCHEMA}
    assert body["output_config"]["effort"] == "medium"
    assert body["fallbacks"] == "default"
    assert "server-side-fallback-2026-07-01" in seen["beta"]
    content = body["messages"][0]["content"]
    assert content[0]["type"] == "image" and content[0]["source"]["media_type"] == "image/jpeg"
    assert "박민수" in content[1]["text"]

    assert rec.engineer_name == "박민수"  # whitespace removed
    assert rec.serial_number == "31550075"
    assert rec.serial_alternatives == ["31550076"]
    assert rec.calibration_date == "2025-12-29"
    assert rec.label_box == (0.27, 0.58, 0.72, 0.8)
    assert rec.rotation == 90


def test_model_without_server_fallback_uses_plain_endpoint():
    seen = {}

    def handler(request):
        seen["body"] = json.loads(request.content)
        return httpx.Response(200, json=_message(json.dumps(ANSWER)))

    ClaudeRecognizer(client=_client(handler), model="claude-haiku-5-5").recognize(make_tool_photo())
    assert "fallbacks" not in seen["body"]


def test_refusal_is_reported():
    handler = lambda request: httpx.Response(200, json=_message("", stop_reason="refusal"))  # noqa: E731
    with pytest.raises(RecognitionError):
        ClaudeRecognizer(client=_client(handler)).recognize(make_tool_photo())


def test_auth_error_is_friendly():
    def handler(request):
        return httpx.Response(401, json={"type": "error", "error": {"type": "authentication_error",
                                                                    "message": "invalid x-api-key"}})

    with pytest.raises(RecognitionError, match="API 키"):
        ClaudeRecognizer(client=_client(handler)).recognize(make_tool_photo())
