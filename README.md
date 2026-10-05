# 깨짐제로 PPT

NotebookLM 슬라이드 PDF를 한글 깨짐 없이 편집 가능한 PPTX로 변환하는 웹앱.

## 문제

- NotebookLM 슬라이드는 PDF로만 다운로드 → 편집 불가
- Canva·미리캔버스 등 기존 변환기는 한글이 'ㅁㅁㅁ'으로 깨지거나 레이아웃이 무너짐

## 해결 방식 (하이브리드 파이프라인)

1. **고속 모드**: PDF에서 텍스트·위치·이미지를 직접 추출 → 수 초 만에 PPTX 생성 (API 비용 0원)
2. **정밀 모드**: 텍스트 추출이 어려운 PDF는 Gemini 멀티모달이 슬라이드 구조를 JSON으로 파악 → PPTX 조립
   - Gemini API 키는 사용자 브라우저에만 저장 (서버로 전송 안 됨)

**한글 깨짐 방지 핵심**: 모든 텍스트 run에 OOXML 동아시아(`a:ea`) 서체를 명시적으로 지정 +
`lang="ko-KR"` 설정. 폰트 테마 3종 (맑은 고딕 / Noto Sans KR / Pretendard).

## 구조

```
index.html          # 프론트엔드 (단일 파일)
api/
  engine.py         # 변환 엔진 (진단·고속변환·JSON조립)
  convert.py        # POST /api/convert — PDF → PPTX
  build.py          # POST /api/build  — slides JSON → PPTX
requirements.txt
tests/              # 로컬 검증 스크립트
```

## 배포 (Vercel)

1. 이 저장소를 Vercel에 Import (Python 런타임 자동 감지)
2. 환경변수 불필요 (Gemini 키는 사용자 브라우저에서 입력)
3. `requirements.txt`의 의존성이 자동 설치됨

## 로컬 테스트

```bash
python3 -m venv .venv && .venv/bin/pip install -r requirements.txt
.venv/bin/python tests/test_engine.py
```

## API

**POST /api/convert** — `{pdf_base64, font, aspect}` →
`{ok, report:{mode:"fast"|"needs_ocr", ...}, pptx_base64?}`

**POST /api/build** — `{slides:[{title, bullets[], notes?}], font, aspect}` →
`{ok, report, pptx_base64}`
