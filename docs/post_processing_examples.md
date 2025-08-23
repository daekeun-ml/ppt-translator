# PowerPoint Post-Processing Examples

## 후처리 기능 사용법

PowerPoint 번역 후 텍스트 박스의 자동 조정을 위한 후처리 기능을 제공합니다.

### 기본 설정 (.env 파일)

```env
# Post-processing Settings
ENABLE_TEXT_AUTOFIT=true
TEXT_LENGTH_THRESHOLD=10
```

- `ENABLE_TEXT_AUTOFIT`: 텍스트 자동 조정 활성화 여부 (기본값: true)
- `TEXT_LENGTH_THRESHOLD`: 자동 조정을 적용할 텍스트 길이 임계값 (기본값: 10글자)

### 사용 예시

#### 1. 번역과 함께 자동 후처리

```bash
# 전체 프레젠테이션 번역 (후처리 자동 적용)
uv run python server.py --translate --input-file presentation.pptx --target-language ko

# 특정 슬라이드 번역 (후처리 자동 적용)
uv run python server.py --translate-slides "1,3,5" --input-file presentation.pptx --target-language ko
```

#### 2. 후처리만 별도 실행

```bash
# 기존 PowerPoint 파일에 후처리만 적용
uv run python server.py --post-process-only translated_file.pptx

# 또는 독립 실행
uv run python post_processing.py --input-file translated_file.pptx
```

#### 3. 후처리 설정 커스터마이징

```bash
# 후처리 비활성화
uv run python server.py --translate --input-file presentation.pptx --target-language ko --disable-autofit

# 텍스트 길이 임계값 변경 (15글자로 설정)
uv run python server.py --translate --input-file presentation.pptx --target-language ko --text-threshold 15

# 후처리만 실행하면서 임계값 변경
uv run python post_processing.py --input-file presentation.pptx --text-threshold 20
```

#### 4. 디버그 모드

```bash
# 디버그 정보와 함께 후처리 실행
uv run python post_processing.py --input-file presentation.pptx --debug
```

### 후처리 기능 상세

#### 적용되는 설정

1. **Text Wrapping**: 텍스트 박스 내에서 텍스트 줄바꿈 활성화
2. **Shrink Text on Overflow**: 텍스트가 박스를 넘칠 때 자동으로 글자 크기 축소
3. **Margin Adjustment**: 텍스트 박스 여백 최적화

#### 처리 조건

- 텍스트 박스에 설정된 임계값(기본 10글자) 이상의 텍스트가 있는 경우
- `ENABLE_TEXT_AUTOFIT=true`로 설정된 경우

#### 출력 예시

```
Processing PowerPoint file: presentation.pptx
Text length threshold: 10 characters
Auto-fit enabled: true
Processing slide 1/5...
  → Processed 3 text boxes
Processing slide 2/5...
  → Processed 2 text boxes
...

Post-processing completed!
Total text boxes processed: 12
Output saved to: presentation_processed.pptx
```

### FastMCP 통합

FastMCP 서버를 통해 AI 어시스턴트와 함께 사용할 때도 후처리가 자동으로 적용됩니다:

```
사용자: "presentation.pptx를 한국어로 번역해주세요"
```

AI 어시스턴트가 번역을 완료한 후 자동으로 후처리를 적용하여 텍스트 박스를 최적화합니다.

### 문제 해결

#### 후처리가 적용되지 않는 경우

1. `.env` 파일에서 `ENABLE_TEXT_AUTOFIT=true` 확인
2. 텍스트 길이가 임계값 이상인지 확인
3. PowerPoint 파일이 올바른 형식(.pptx)인지 확인

#### 오류 발생 시

```bash
# 디버그 모드로 실행하여 상세 오류 확인
uv run python post_processing.py --input-file presentation.pptx --debug
```

일반적인 오류:
- 파일 권한 문제: 출력 파일 경로에 쓰기 권한이 있는지 확인
- 손상된 PowerPoint 파일: 원본 파일이 올바르게 열리는지 확인
- 메모리 부족: 큰 파일의 경우 충분한 메모리가 있는지 확인
