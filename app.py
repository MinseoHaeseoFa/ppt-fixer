"""오타잡이 PPT — Vercel Python 런타임용 Flask 단일 앱.

엔트리포인트: pyproject.toml [tool.vercel] entrypoint = "app:app"
모든 요청(/, /api/convert, /api/build)을 이 앱이 처리한다.
"""
import base64
import os

from flask import Flask, Response, jsonify, request

import engine

app = Flask(__name__)
# Vercel 요청 본문 제한(4.5MB)보다 먼저 걸리도록 가드
app.config["MAX_CONTENT_LENGTH"] = 8 * 1024 * 1024

HERE = os.path.dirname(os.path.abspath(__file__))
INDEX_PATH = os.path.join(HERE, "index.html")


@app.get("/")
def index():
    try:
        with open(INDEX_PATH, "rb") as f:
            html = f.read()
    except OSError:
        return jsonify(ok=False, error="페이지를 찾을 수 없어요."), 500
    return Response(html, mimetype="text/html; charset=utf-8")


@app.get("/api/convert")
def convert_health():
    return jsonify(ok=True, service="ppt-fixer/convert")


@app.get("/api/build")
def build_health():
    return jsonify(ok=True, service="ppt-fixer/build")


@app.post("/api/convert")
def convert():
    """PDF → PPTX 고속 변환 (또는 정밀 모드 필요 응답)."""
    data = request.get_json(force=True, silent=True) or {}
    try:
        pdf_bytes = base64.b64decode(data.get("pdf_base64", ""))
    except Exception:
        return jsonify(ok=False, code="bad_input",
                       error="파일 데이터를 읽을 수 없어요. 다시 올려주세요.")
    if not pdf_bytes:
        return jsonify(ok=False, code="bad_input",
                       error="PDF 데이터가 비어 있어요.")
    if len(pdf_bytes) > 4 * 1024 * 1024:
        return jsonify(ok=False, code="too_big",
                       error="파일이 너무 커요. 3MB 이하의 PDF로 올려주세요.")
    try:
        pptx, report = engine.convert_pdf_to_pptx(
            pdf_bytes, data.get("font", "malgun"), data.get("aspect", "16:9"))
    except engine.ConvertError as e:
        return jsonify(ok=False, code=e.code, error=e.message)
    except Exception:
        return jsonify(ok=False, code="failed",
                       error="변환 중 오류가 발생했어요. 잠시 후 다시 시도해주세요.")
    resp = {"ok": True, "report": report}
    if pptx:
        resp["pptx_base64"] = base64.b64encode(pptx).decode("ascii")
    return jsonify(resp)


@app.post("/api/build")
def build():
    """slides JSON 또는 images → PPTX 조립."""
    data = request.get_json(force=True, silent=True) or {}
    aspect = data.get("aspect", "16:9")
    try:
        images = data.get("images")
        if isinstance(images, list) and images:
            if len(images) > 20:
                return jsonify(ok=False, code="too_many",
                               error="이미지는 최대 20장까지 올릴 수 있어요.")
            decoded = []
            for it in images:
                b64 = it.get("data", "") if isinstance(it, dict) else it
                name = it.get("name", "") if isinstance(it, dict) else ""
                try:
                    raw = base64.b64decode(b64)
                except Exception:
                    return jsonify(ok=False, code="bad_input",
                                   error="이미지 데이터를 읽을 수 없어요.")
                if len(raw) > 12 * 1024 * 1024:
                    return jsonify(ok=False, code="too_big",
                                   error="이미지가 너무 커요. 장당 10MB 이하로 올려주세요.")
                decoded.append((raw, name))
            pptx, report = engine.build_from_images(decoded, aspect)
        else:
            slides = data.get("slides")
            if not isinstance(slides, list) or not slides:
                return jsonify(ok=False, code="bad_json",
                               error="슬라이드 데이터가 비어 있어요.")
            if len(slides) > 80:
                return jsonify(ok=False, code="too_many",
                               error="슬라이드가 너무 많아요. 80장 이하로 나눠주세요.")
            font = data.get("font", "malgun")
            use_layout = any(isinstance(s, dict) and isinstance(s.get("elements"), list)
                             for s in slides)
            if use_layout:
                # 레이아웃 재현 모드: "image" 요소 크롭용 원본 페이지 이미지 준비
                sources = []
                pdf_b64 = data.get("pdf_base64", "")
                if pdf_b64:
                    try:
                        pdf_bytes = base64.b64decode(pdf_b64)
                    except Exception:
                        return jsonify(ok=False, code="bad_input",
                                       error="PDF 데이터를 읽을 수 없어요.")
                    need_pages = set()
                    for i, s in enumerate(slides):
                        if isinstance(s, dict):
                            for el in s.get("elements") or []:
                                if isinstance(el, dict) and el.get("type") == "image":
                                    need_pages.add(int(s.get("page", i + 1)))
                    for i in range(len(slides)):
                        pno = int(slides[i].get("page", i + 1)) \
                            if isinstance(slides[i], dict) else i + 1
                        if pno in need_pages:
                            try:
                                sources.append(engine.render_page_png(pdf_bytes, pno))
                            except engine.ConvertError:
                                sources.append(None)
                        else:
                            sources.append(None)
                elif isinstance(images, list) and images:
                    for it in images:
                        b64 = it.get("data", "") if isinstance(it, dict) else it
                        try:
                            sources.append(base64.b64decode(b64))
                        except Exception:
                            sources.append(None)
                elif isinstance(images, dict) and images:
                    # 새 프론트엔드: {"페이지번호": jpeg_b64}
                    for i, s in enumerate(slides):
                        pno = s.get("page", i + 1) if isinstance(s, dict) else i + 1
                        b64 = images.get(str(pno))
                        if b64 is None:
                            b64 = images.get(pno)
                        try:
                            sources.append(base64.b64decode(b64) if b64 else None)
                        except Exception:
                            sources.append(None)
                pptx, report = engine.build_from_layout(slides, sources, font, aspect)
            else:
                pptx, report = engine.build_from_json(slides, font, aspect)
        return jsonify(ok=True, report=report,
                       pptx_base64=base64.b64encode(pptx).decode("ascii"))
    except engine.ConvertError as e:
        return jsonify(ok=False, code=e.code, error=e.message)
    except Exception:
        return jsonify(ok=False, code="failed",
                       error="PPT 생성 중 오류가 발생했어요. 다시 시도해주세요.")


# 로컬 확인용 (Vercel에서는 사용하지 않음)
if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=False)
