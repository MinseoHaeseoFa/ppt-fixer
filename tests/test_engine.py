"""엔진 검증: 한글 폰트 주입·레이아웃·엣지케이스."""
import io
import os
import sys
import zipfile

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "api"))

import pymupdf
from PIL import Image, ImageDraw

import engine

HERE = os.path.dirname(__file__)
FONT = os.path.join(HERE, "NotoSansKR.ttf")
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

    print("\nALL ENGINE TESTS PASSED")


if __name__ == "__main__":
    main()
