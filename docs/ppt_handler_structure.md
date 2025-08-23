# ppt_handler.py 구조 분석

## 개요
PowerPoint 문서 처리 및 텍스트 프레임 업데이트를 위한 최적화된 모듈입니다. 번역 시 서식, 글머리 기호, 들여쓰기, 색상 등을 보존하면서 언어별 폰트를 적용합니다.

## 트리 구조

```
ppt_handler.py
├── 📦 Imports & Dependencies
│   ├── logging, re, typing
│   ├── dataclasses, pathlib
│   ├── pptx.dml.color.RGBColor
│   └── Local modules (config, dependencies, translation_engine, text_utils)
│
├── 📊 Data Classes
│   └── TranslationResult
│       ├── translated_count: int
│       ├── translated_notes_count: int
│       ├── total_shapes: int
│       └── errors: List[str]
│
├── 🔍 FormattingExtractor (서식 추출기)
│   ├── extract_paragraph_structure()
│   │   └── _extract_single_paragraph_info()
│   │       ├── _extract_xml_formatting()
│   │       ├── _extract_bullet_format()
│   │       │   └── _create_bullet_format()
│   │       └── _extract_run_info()
│   │           └── _extract_run_formatting()
│   │               └── _extract_font_color()
│   └── _get_namespace()
│
├── 🎨 FormattingApplier (서식 적용기)
│   ├── apply_paragraph_structure()
│   │   ├── _apply_paragraph_properties()
│   │   │   ├── _apply_xml_formatting()
│   │   │   └── _apply_bullet_format()
│   │   │       └── _apply_bullet_xml()
│   │   │           ├── _add_bullet_none()
│   │   │           ├── _add_bullet_char()
│   │   │           └── _add_bullet_autonum()
│   │   └── _apply_runs_with_formatting()
│   │       ├── _apply_multiple_runs()
│   │       └── _apply_run_formatting()
│   │           ├── _apply_rgb_color()
│   │           ├── _apply_theme_color()
│   │           └── _apply_scheme_color()
│
├── 📝 TextFrameUpdater (텍스트 프레임 업데이터)
│   ├── update_text_frame()
│   │   └── _choose_update_strategy()
│   │       ├── _update_matching_paragraphs()
│   │       └── _rebuild_with_structure()
│   ├── _update_with_hyperlinks_safe()
│   │   └── _apply_hyperlinks_to_paragraph()
│   ├── _has_hyperlinks()
│   └── _find_hyperlink_text()
│
├── 🧮 ComplexityAnalyzer (복잡도 분석기)
│   ├── slide_has_complex_formatting()
│   │   └── _text_frame_has_complex_formatting()
│   │       ├── _has_bullet_formatting()
│   │       └── _has_multiple_formatting_styles()
│
├── 🎯 TranslationStrategy (번역 전략)
│   ├── __init__(engine, text_updater)
│   ├── translate_slide()
│   │   ├── _translate_notes()
│   │   ├── _translate_individually() ← 복잡한 서식용
│   │   ├── _translate_with_context() ← 많은 텍스트용
│   │   └── _translate_with_batch() ← 일반적인 경우
│   ├── _apply_translations()
│   └── _apply_translation_to_item()
│
└── 🏛️ PowerPointTranslator (메인 클래스)
    ├── __init__(model_id, enable_polishing)
    ├── translate_presentation() ← 전체 번역
    ├── translate_specific_slides() ← 특정 슬라이드 번역
    ├── get_slide_count()
    └── get_slide_preview()
```

## 주요 클래스 상세 설명

### 1. TranslationResult (데이터 클래스)

번역 결과를 담는 데이터 클래스입니다.

```python
@dataclass
class TranslationResult:
    translated_count: int = 0          # 번역된 텍스트 수
    translated_notes_count: int = 0    # 번역된 노트 수
    total_shapes: int = 0              # 전체 도형 수
    errors: List[str] = None           # 오류 목록
```

### 2. FormattingExtractor (서식 추출기)

PowerPoint 요소에서 서식 정보를 추출하는 클래스입니다.

#### 주요 메서드:
- **`extract_paragraph_structure()`**: 단락 구조, 글머리 기호, 들여쓰기, 여백, 서식 추출
- **`_extract_single_paragraph_info()`**: 개별 단락 정보 추출
- **`_extract_xml_formatting()`**: XML 기반 서식 추출 (여백, 들여쓰기)
- **`_extract_bullet_format()`**: 글머리 기호 서식 추출 (문자, 자동 번호)
- **`_extract_run_info()`**: 개별 런(run) 정보 추출
- **`_extract_font_color()`**: 폰트 색상 정보 추출 (RGB, 테마, 스키마 색상 지원)

#### 색상 지원 타입:
- **RGB 색상**: 직접 RGB 값
- **테마 색상**: PowerPoint 테마 색상 + 밝기 조정
- **스키마 색상**: 색상 스키마 기반

### 3. FormattingApplier (서식 적용기)

PowerPoint 요소에 서식을 적용하는 클래스입니다.

#### 주요 메서드:
- **`apply_paragraph_structure()`**: 단락 구조 및 서식 적용 (언어별 폰트 포함)
- **`_apply_paragraph_properties()`**: 단락 수준 속성 적용
- **`_apply_xml_formatting()`**: XML 기반 서식 적용
- **`_apply_bullet_format()`**: 글머리 기호 서식 적용
- **`_apply_runs_with_formatting()`**: 런 서식과 함께 텍스트 적용
- **`_apply_font_color()`**: 폰트 색상 적용 (모든 색상 타입 지원)

#### 글머리 기호 타입:
- **buNone**: 글머리 기호 없음
- **buChar**: 문자 글머리 기호 (•, -, * 등)
- **buAutoNum**: 자동 번호 (1, 2, 3... 또는 a, b, c...)

### 4. TextFrameUpdater (텍스트 프레임 업데이터)

PowerPoint 텍스트 프레임을 번역으로 업데이트하는 클래스입니다.

#### 주요 메서드:
- **`update_text_frame()`**: 서식, 글머리 기호, 들여쓰기를 보존하며 텍스트 프레임 업데이트
- **`_choose_update_strategy()`**: 적절한 업데이트 전략 선택
- **`_update_matching_paragraphs()`**: 단락 수가 일치할 때 업데이트
- **`_rebuild_with_structure()`**: 구조를 보존하며 텍스트 프레임 재구성
- **`_update_with_hyperlinks_safe()`**: 하이퍼링크를 보존하며 안전하게 업데이트

#### 업데이트 전략:
1. **단일 단락**: 직접 서식 적용
2. **일치하는 단락 수**: 1:1 매핑으로 업데이트
3. **다른 구조**: 서식을 보존하며 재구성
4. **하이퍼링크 포함**: 안전한 하이퍼링크 보존 모드

### 5. ComplexityAnalyzer (복잡도 분석기)

슬라이드 복잡도를 분석하여 번역 전략을 결정하는 클래스입니다.

#### 주요 메서드:
- **`slide_has_complex_formatting()`**: 슬라이드의 복잡한 서식 여부 확인
- **`_text_frame_has_complex_formatting()`**: 텍스트 프레임의 복잡한 서식 여부 확인
- **`_has_bullet_formatting()`**: 글머리 기호 서식 여부 확인
- **`_has_multiple_formatting_styles()`**: 다중 서식 스타일 여부 확인

#### 복잡도 판단 기준:
- 들여쓰기 (level > 0)
- 글머리 기호 존재
- 다중 색상 사용
- 다중 서식 스타일 (굵기, 기울임 등)

### 6. TranslationStrategy (번역 전략)

다양한 번역 전략을 처리하는 클래스입니다.

#### 주요 메서드:
- **`translate_slide()`**: 적절한 전략을 사용하여 단일 슬라이드 번역
- **`_translate_individually()`**: 서식 보존을 위한 개별 번역
- **`_translate_with_context()`**: 컨텍스트 인식 번역
- **`_translate_with_batch()`**: 배치 번역
- **`_apply_translations()`**: 번역을 원본 도형에 적용

#### 번역 전략 선택:
1. **복잡한 서식** → Individual Translation (개별 번역)
2. **많은 텍스트** (> CONTEXT_THRESHOLD) → Context-aware Translation
3. **일반적인 경우** → Batch Translation

### 7. PowerPointTranslator (메인 클래스)

PowerPoint 번역의 메인 클래스입니다.

#### 주요 메서드:
- **`translate_presentation()`**: 전체 PowerPoint 프레젠테이션 번역
- **`translate_specific_slides()`**: 특정 슬라이드만 번역
- **`get_slide_count()`**: 슬라이드 총 개수 반환
- **`get_slide_preview()`**: 특정 슬라이드의 텍스트 미리보기

## 처리 흐름 (Processing Flow)

```
┌─────────────────────────────────────────────────────────────┐
│ PowerPointTranslator                                        │
│ ├── 1. 프레젠테이션 로드                                    │
│ ├── 2. 각 슬라이드별 처리                                   │
│ │   ├── SlideTextCollector로 텍스트 수집                   │
│ │   ├── ComplexityAnalyzer로 복잡도 분석                   │
│ │   └── TranslationStrategy로 번역 전략 선택               │
│ │       ├── 복잡한 서식 → Individual Translation          │
│ │       ├── 많은 텍스트 → Context-aware Translation       │
│ │       └── 일반적인 경우 → Batch Translation             │
│ ├── 3. 각 텍스트 요소별 처리                               │
│ │   ├── FormattingExtractor로 서식 추출                   │
│ │   ├── TranslationEngine으로 번역                        │
│ │   ├── FormattingApplier로 서식 적용                     │
│ │   └── TextFrameUpdater로 텍스트 업데이트                │
│ └── 4. 결과 저장 및 리포트                                 │
└─────────────────────────────────────────────────────────────┘
```

## 서식 보존 체계

### 단락 수준 (Paragraph Level)
- **글머리 기호 (Bullets)**: 문자, 자동 번호, 없음
- **들여쓰기 (Indentation)**: XML 기반 정확한 들여쓰기 값
- **여백 (Margins)**: 좌측 여백, 첫 줄 들여쓰기
- **정렬 (Alignment)**: 좌측, 중앙, 우측, 양쪽 정렬
- **간격 (Spacing)**: 단락 전후 간격, 줄 간격

### 런 수준 (Run Level)
- **폰트 (Font)**: 이름, 크기
- **스타일 (Style)**: 굵기, 기울임
- **색상 (Color)**: RGB, 테마, 스키마 색상

### 특수 요소
- **하이퍼링크 (Hyperlinks)**: URL 및 텍스트 보존
- **테이블 셀 (Table Cells)**: 셀 내 텍스트 서식
- **노트 (Notes)**: 슬라이드 노트 번역

## 언어별 폰트 매핑

| 언어 | 언어 코드 | 적용 폰트 |
|------|-----------|-----------|
| 한국어 | ko | 맑은 고딕 |
| 일본어 | ja | Yu Gothic UI |
| 중국어 | zh | Microsoft YaHei |
| 영어 | en | Amazon Ember |
| 기본값 | - | Arial |

## 주요 특징

### 🎨 서식 보존
- 글머리 기호, 들여쓰기, 여백 완벽 보존
- RGB, 테마, 스키마 색상 정확한 보존
- 폰트 크기, 굵기, 기울임 유지
- 하이퍼링크 URL 및 서식 보존

### 🌍 언어별 최적화
- 각 언어에 최적화된 폰트 자동 적용
- 언어별 텍스트 렌더링 품질 향상
- 다국어 환경에서의 가독성 보장

### 🧠 지능형 번역 전략
- **복잡한 서식**: 개별 번역으로 서식 완벽 보존
- **많은 텍스트**: 컨텍스트 인식으로 번역 품질 향상
- **일반적인 경우**: 배치 번역으로 효율성 극대화

### 🔧 견고한 오류 처리
- 각 단계별 세밀한 예외 처리
- 실패 시 자동 폴백 메커니즘
- 상세한 로깅 및 디버깅 정보 제공
- 부분 실패 시에도 최대한 번역 진행

## 성능 최적화

### 배치 처리
- `BATCH_SIZE` 설정으로 메모리 효율성 관리
- 대용량 프레젠테이션 처리 최적화

### 컨텍스트 인식
- `CONTEXT_THRESHOLD` 기준으로 전략 자동 선택
- 번역 품질과 성능의 균형 유지

### 메모리 관리
- 슬라이드별 순차 처리로 메모리 사용량 최소화
- 불필요한 객체 참조 해제

이 구조는 PowerPoint 번역 시 원본의 시각적 품질을 완벽하게 유지하면서도 효율적이고 안정적인 번역을 제공하도록 설계되었습니다.
