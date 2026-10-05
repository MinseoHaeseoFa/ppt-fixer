"""엔진 검증: 한글 폰트 주입·레이아웃·엣지케이스."""
import io
import os
import sys
import zipfile

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import pymupdf
from PIL import Image, ImageDraw

import engine

HERE = os.path.dirname(__file__)
FONT = os.path.join(HERE, "NotoSansKR-sub.ttf")
EAFONT = "맑은 고딕"


def make_sample_pdf(path):
    doc = pymupdf.open()
    # 1p: 타이틀
    p = doc.new_page(width=960, height=540)
    p.insert_font(fontname="kr", fontfile=FONT)
    p.draw_rect(pymupdf.Rect(0, 0, 960, 540), color=None, fill=(0.08, 0.1, 0.25))
    p.insert_text((80, 220), "인공지능 활용 교육", fontname="kr", fontsize=54,
                  color=(1, 1, 1))
    p.insert_text((80, 300), "NotebookLM으로 만드는 강의 자료", fontname="kr",
                  fontsize=24, color=(1, 0.82, 0.4))
    # 2p: 불릿
    p = doc.new_page(width=960, height=540)
    p.insert_font(fontname="kr", fontfile=FONT)
    p.insert_text((70, 90), "오늘 배울 내용", fontname="kr", fontsize=36)
    bullets = ["프롬프트 작성 기본 원칙", "한글 자료 요약 기법", "슬라이드 자동 생성과 검수"]
    y = 170
    for b in bullets:
        p.insert_text((90, y), "•  " + b, fontname="kr", fontsize=22)
        y += 60
    # 3p: 이미지 + 캡션
    p = doc.new_page(width=960, height=540)
    p.insert_font(fontname="kr", fontfile=FONT)
    p.insert_text((70, 90), "시각 자료 예시", fontname="kr", fontsize=36)
    img = Image.new("RGB", (400, 220), (26, 35, 80))
    d = ImageDraw.Draw(img)
    d.ellipse([150, 40, 250, 140], fill=(255, 209, 102))
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    p.insert_image(pymupdf.Rect(90, 140, 490, 360), stream=buf.getvalue())
    p.insert_text((90, 410), "그림 1. 밤하늘 시각화 예시", fontname="kr", fontsize=18,
                  color=(0.4, 0.4, 0.4))
    doc.save(path)
    doc.close()
    return path


def make_blank_pdf(path):
    doc = pymupdf.open()
    p = doc.new_page(width=960, height=540)
    pix = p.get_pixmap(dpi=72)
    # 텍스트 없이 이미지만 있는 페이지로 저장
    doc2 = pymupdf.open()
    p2 = doc2.new_page(width=960, height=540)
    p2.insert_image(p2.rect, stream=pix.tobytes("png"))
    doc2.save(path)
    doc.close()
    doc2.close()
    return path


def check_pptx(pptx_bytes, expect_texts, expect_image=False, label="",
               expect_font=EAFONT):
    z = zipfile.ZipFile(io.BytesIO(pptx_bytes))
    names = z.namelist()
    assert any(n.startswith("ppt/slides/slide") for n in names), "slide missing"
    # 1) 모든 rPr에 ea 서체 명시
    bad = []
    for n in names:
        if not n.startswith("ppt/slides/slide"):
            continue
        xml = z.read(n).decode("utf-8")
        # rPr 개수 vs ea 개수
        n_rpr = xml.count("<a:rPr")
        n_ea = xml.count("<a:ea ")
        if n_rpr != n_ea:
            bad.append((n, n_rpr, n_ea))
        if 'typeface="%s"' % expect_font not in xml and n_rpr:
            bad.append((n, "no-" + expect_font))
    assert not bad, "ea font missing: %s" % bad
    # 2) 한글 텍스트 보존
    allxml = " ".join(z.read(n).decode("utf-8") for n in names
                      if n.startswith("ppt/slides/slide"))
    for t in expect_texts:
        assert t in allxml, "text lost: %s" % t
    # 3) 이미지
    if expect_image:
        assert any(n.startswith("ppt/media/") for n in names), "image missing"
    print("[OK] %s — rPr=ea 일치, 한글 %d개 보존%s" %
          (label, len(expect_texts), ", 이미지 포함" if expect_image else ""))


def main():
    os.chdir(HERE)
    pdf = make_sample_pdf("sample.pdf")
    data = open(pdf, "rb").read()

    # 진단
    diag = engine.diagnose(data)
    assert diag["extractable"] and diag["pages"] == 3, diag
    print("[OK] diagnose:", diag["pages"], "pages, avg_chars", diag["avg_chars"])

    # 고속 변환
    pptx, rep = engine.convert_pdf_to_pptx(data, "malgun", "16:9")
    assert rep["mode"] == "fast", rep
    open("sample_out.pptx", "wb").write(pptx)
    check_pptx(pptx, ["인공지능 활용 교육", "NotebookLM으로 만드는 강의 자료",
                      "프롬프트 작성 기본 원칙", "그림 1. 밤하늘 시각화 예시"],
               expect_image=True, label="fast convert")

    # 폰트 테마 변경
    pptx2, rep2 = engine.convert_pdf_to_pptx(data, "noto", "4:3")
    z = zipfile.ZipFile(io.BytesIO(pptx2))
    xml = z.read("ppt/slides/slide1.xml").decode("utf-8")
    assert 'typeface="Noto Sans KR"' in xml
    assert 'ppt/slides/slide1.xml' in z.namelist()
    print("[OK] font theme + 4:3")

    # 이미지 전용 PDF → needs_ocr
    blank = make_blank_pdf("blank.pdf")
    pptx3, rep3 = engine.convert_pdf_to_pptx(open(blank, "rb").read())
    assert pptx3 is None and rep3["mode"] == "needs_ocr", rep3
    print("[OK] image-only PDF → needs_ocr")

    # JSON 빌드
    slides = [
        {"title": "시작하며", "bullets": ["목표 설정", "자료 준비"], "notes": "오프닝 멘트"},
        {"title": "마무리", "bullets": ["질의응답"]},
    ]
    pptx4, rep4 = engine.build_from_json(slides, "pretendard", "16:9")
    check_pptx(pptx4, ["시작하며", "질의응답"], label="json build",
               expect_font="Pretendard")
    z4 = zipfile.ZipFile(io.BytesIO(pptx4))
    assert 'typeface="Pretendard"' in z4.read("ppt/slides/slide1.xml").decode("utf-8")
    print("[OK] build_from_json + Pretendard")

    # 엣지케이스
    for raw, code in [(b"hello", "not_pdf"), (b"%PDF-1.4 broken", "broken")]:
        try:
            engine.diagnose(raw)
            raise AssertionError("should raise " + code)
        except engine.ConvertError as e:
            assert e.code == code, (code, e.code)
    print("[OK] broken inputs → clean errors")

    # 암호 PDF
    doc = pymupdf.open()
    doc.new_page()
    doc.save("enc.pdf", encryption=pymupdf.PDF_ENCRYPT_AES_256, user_pw="1234")
    doc.close()
    try:
        engine.diagnose(open("enc.pdf", "rb").read())
        raise AssertionError("should raise encrypted")
    except engine.ConvertError as e:
        assert e.code == "encrypted", e.code
    print("[OK] encrypted PDF → clean error")

    # 이미지 담기
    imgs = []
    for i, color in enumerate([(200, 50, 50), (50, 120, 200)]):
        im = Image.new("RGB", (800, 450), color)
        d2 = ImageDraw.Draw(im)
        d2.text((300, 200), "slide %d" % (i + 1))
        b = io.BytesIO()
        im.save(b, format="PNG")
        imgs.append((b.getvalue(), "s%d.png" % (i + 1)))
    pptx5, rep5 = engine.build_from_images(imgs, "16:9")
    assert rep5["mode"] == "image" and rep5["pages"] == 2, rep5
    z5 = zipfile.ZipFile(io.BytesIO(pptx5))
    n_slides = len([n for n in z5.namelist()
                    if n.startswith("ppt/slides/slide") and n.endswith(".xml")])
    assert n_slides == 2, n_slides
    assert any(n.startswith("ppt/media/") for n in z5.namelist()), "media missing"
    print("[OK] build_from_images — 2 slides, images embedded")

    # 레이아웃 재현 모드 (정밀 모드 v2)
    src = Image.new("RGB", (1000, 1000), (240, 240, 240))
    ds = ImageDraw.Draw(src)
    ds.rectangle([600, 220, 880, 480], fill=(30, 60, 120))
    sbuf = io.BytesIO()
    src.save(sbuf, format="PNG")
    layout_slides = [
        {"page": 1, "elements": [
            {"type": "text", "x": 120, "y": 60, "w": 760, "h": 110,
             "text": "오타 수정 테스트", "font_size": 40, "bold": True,
             "color": "1A1A2E", "align": "center"},
            {"type": "text", "x": 100, "y": 220, "w": 440, "h": 120,
             "text": "NotebookLM 슬라이드", "font_size": 20},
            {"type": "shape", "shape": "rect", "x": 80, "y": 200,
             "w": 840, "h": 300, "fill": "F5F5F5"},
            {"type": "image", "x": 600, "y": 220, "w": 280, "h": 260},
            {"type": "bogus", "x": 0, "y": 0, "w": 10, "h": 10},
            {"type": "text", "x": 0, "y": 0, "w": 0, "h": 0, "text": "무시됨"},
        ]},
        {"page": 2, "notes": "발표 노트", "elements": [
            {"type": "text", "x": 100, "y": 100, "w": 800, "h": 100,
             "text": "둘째 장", "font_size": 32},
        ]},
    ]
    pptx6, rep6 = engine.build_from_layout(layout_slides, [sbuf.getvalue()], "malgun", "16:9")
    assert rep6["mode"] == "precise" and rep6["pages"] == 2 and rep6.get("layout"), rep6
    check_pptx(pptx6, ["오타 수정 테스트", "NotebookLM 슬라이드", "둘째 장"],
               expect_image=True, label="layout build")
    z6 = zipfile.ZipFile(io.BytesIO(pptx6))
    xml1 = z6.read("ppt/slides/slide1.xml").decode("utf-8")
    # 텍스트 박스 2개 + 도형 1개 + 그림 1개 (bogus/무효 요소 제외)
    assert xml1.count("<p:sp>") == 3, xml1.count("<p:sp>")  # 2 text + 1 shape
    assert "<p:pic>" in xml1, "cropped image not embedded"
    # 도형 테두리 없음 (영상의 검정 테두리 이슈 방지)
    assert "<a:noFill/>" in xml1, "shape border not removed"
    print("[OK] build_from_layout — 텍스트/도형/이미지 크롭, 한글 ea")

    # 페이지 렌더
    png = engine.render_page_png(data, 1)
    assert png[:8] == b"\x89PNG\r\n\x1a\n", "not png"
    im = Image.open(io.BytesIO(png))
    assert im.size[0] == 1500, im.size
    print("[OK] render_page_png — 1500px 렌더")

    # 빈 elements → 오류
    try:
        engine.build_from_layout([{"page": 1, "elements": []}], [])
        raise AssertionError("should raise bad_json")
    except engine.ConvertError as e:
        assert e.code == "bad_json", e.code
    print("[OK] layout empty → clean error")

    print("\nALL ENGINE TESTS PASSED")


if __name__ == "__main__":
    main()
