"""POST /api/build — 슬라이드 JSON → PPTX (정밀 모드용 조립기)."""
from http.server import BaseHTTPRequestHandler
import json
import base64
import os
import sys

sys.path.insert(0, os.path.dirname(__file__))
import engine


class handler(BaseHTTPRequestHandler):
    def _send(self, code, obj):
        body = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_POST(self):
        try:
            length = int(self.headers.get("Content-Length", 0))
            raw = self.rfile.read(length) if length else b"{}"
            data = json.loads(raw)
            slides = data.get("slides")
            if not isinstance(slides, list) or not slides:
                return self._send(200, {"ok": False, "code": "bad_json",
                                        "error": "슬라이드 데이터가 비어 있어요."})
            if len(slides) > 80:
                return self._send(200, {"ok": False, "code": "too_many",
                                        "error": "슬라이드가 너무 많아요. 80장 이하로 나눠주세요."})
            pptx, report = engine.build_from_json(
                slides, data.get("font", "malgun"), data.get("aspect", "16:9"))
            self._send(200, {"ok": True, "report": report,
                             "pptx_base64": base64.b64encode(pptx).decode("ascii")})
        except engine.ConvertError as e:
            self._send(200, {"ok": False, "code": e.code, "error": e.message})
        except Exception:
            self._send(200, {"ok": False, "code": "failed",
                             "error": "PPT 생성 중 오류가 발생했어요. 다시 시도해주세요."})

    def do_GET(self):
        self._send(200, {"ok": True, "service": "ppt-fixer/build"})
