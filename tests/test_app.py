"""Flask 앱 HTTP 계층 end-to-end 검증 (Vercel에서 실행될 app.py 그대로 테스트)."""
import base64
import io
import os
import sys
import zipfile

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import app as app_module

HERE = os.path.dirname(__file__)
SAMPLE_PDF = os.path.join(HERE, "sample.pdf")
BLANK_PDF = os.path.join(HERE, "blank.pdf")

client = app_module.app.test_client()


def b64(path):
    return base64.b64encode(open(path, "rb").read()).decode("ascii")


def test_index():
    r = client.get("/")
    assert r.status_code == 200, r.status_code
    assert "깨짐제로" in r.get_data(as_text=True)
    assert "text/html" in r.content_type
    print("[OK] GET / → index.html")


def test_health():
    for p in ("/api/convert", "/api/build"):
        r = client.get(p)
        assert r.status_code == 200 and r.get_json()["ok"] is True, p
    print("[OK] health checks")


def test_convert_fast():
    r = client.post("/api/convert", json={
        "pdf_base64": b64(SAMPLE_PDF), "font": "malgun", "aspect": "16:9"})
    assert r.status_code == 200, r.status_code
    d = r.get_json()
    assert d["ok"] is True, d
    assert d["report"]["mode"] == "fast", d["report"]
    assert d["report"]["pages"] == 3
    z = zipfile.ZipFile(io.BytesIO(base64.b64decode(d["pptx_base64"])))
    xml = z.read("ppt/slides/slide1.xml").decode("utf-8")
    assert "인공지능 활용 교육" in xml
    assert xml.count("<a:rPr") == xml.count("<a:ea ")
    print("[OK] POST /api/convert fast — 한글·ea서체 OK")


def test_convert_needs_ocr():
    r = client.post("/api/convert", json={"pdf_base64": b64(BLANK_PDF)})
    d = r.get_json()
    assert d["ok"] is True and d["report"]["mode"] == "needs_ocr", d
    assert "pptx_base64" not in d
    print("[OK] POST /api/convert image-only → needs_ocr")


def test_convert_errors():
    # 깨진 입력
    r = client.post("/api/convert", json={"pdf_base64": base64.b64encode(b"xxx").decode()})
    d = r.get_json()
    assert d["ok"] is False and "PDF" in d["error"], d
    # 빈 바디
    r = client.post("/api/convert", json={})
    d = r.get_json()
    assert d["ok"] is False, d
    # JSON 아님
    r = client.post("/api/convert", data="not json", content_type="application/json")
    d = r.get_json()
    assert d["ok"] is False, d
    print("[OK] POST /api/convert 에러 → 한글 메시지")


def test_build_slides():
    r = client.post("/api/build", json={
        "slides": [{"title": "테스트 제목", "bullets": ["가나다라", "마바사"]}],
        "font": "noto", "aspect": "16:9"})
    d = r.get_json()
    assert d["ok"] is True, d
    z = zipfile.ZipFile(io.BytesIO(base64.b64decode(d["pptx_base64"])))
    xml = z.read("ppt/slides/slide1.xml").decode("utf-8")
    assert "테스트 제목" in xml and 'typeface="Noto Sans KR"' in xml
    print("[OK] POST /api/build slides")


def test_build_images():
    from PIL import Image
    imgs = []
    for c in [(200, 60, 60), (60, 120, 200)]:
        buf = io.BytesIO()
        Image.new("RGB", (640, 360), c).save(buf, format="PNG")
        imgs.append({"data": base64.b64encode(buf.getvalue()).decode(),
                     "name": "t.png"})
    r = client.post("/api/build", json={"images": imgs, "aspect": "16:9"})
    d = r.get_json()
    assert d["ok"] is True and d["report"]["mode"] == "image", d
    assert d["report"]["pages"] == 2
    print("[OK] POST /api/build images")


def test_build_errors():
    r = client.post("/api/build", json={"slides": []})
    assert r.get_json()["ok"] is False
    r = client.post("/api/build", json={})
    assert r.get_json()["ok"] is False
    print("[OK] POST /api/build 에러 처리")


if __name__ == "__main__":
    for p in (SAMPLE_PDF, BLANK_PDF):
        assert os.path.exists(p), "먼저 tests/test_engine.py를 실행해 샘플 PDF를 생성하세요: " + p
    test_index()
    test_health()
    test_convert_fast()
    test_convert_needs_ocr()
    test_convert_errors()
    test_build_slides()
    test_build_images()
    test_build_errors()
    print("\nALL HTTP TESTS PASSED")
