"""POST /api/convert — PDF → PPTX 고속 변환 (또는 정밀 모드 필요 응답)."""
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
            if length > 6 * 1024 * 1024:
                return self._send(200, {
                    "ok": False, "code": "too_big",
                    "error": "파일이 너무 커요. 4MB 이하의 PDF로 올려주세요."})
            raw = self.rfile.read(length) if length else b"{}"
            data = json.loads(raw)
            pdf_bytes = base64.b64decode(data.get("pdf_base64", ""))
            if len(pdf_bytes) > 5 * 1024 * 1024:
                return self._send(200, {
                    "ok": False, "code": "too_big",
                    "error": "파일이 너무 커요. 4MB 이하의 PDF로 올려주세요."})
            pptx, report = engine.convert_pdf_to_pptx(
                pdf_bytes, data.get("font", "malgun"), data.get("aspect", "16:9"))
            resp = {"ok": True, "report": report}
            if pptx:
                resp["pptx_base64"] = base64.b64encode(pptx).decode("ascii")
            self._send(200, resp)
        except engine.ConvertError as e:
            self._send(200, {"ok": False, "code": e.code, "error": e.message})
        except Exception:
            self._send(200, {"ok": False, "code": "failed",
                             "error": "변환 중 오류가 발생했어요. 잠시 후 다시 시도해주세요."})

    def do_GET(self):
        self._send(200, {"ok": True, "service": "ppt-fixer/convert"})
