# onecell 자동 입력 도구

## 프로젝트 구조

```
onecell_project/
├── pyproject.toml
├── settings.ini
├── onecell_template.xlsx
├── build_windows.bat
├── onecell_tool.spec
└── src/onecell/
    ├── main.py      ← 진입점
    ├── app.py       ← GUI + 설정 + 판매가계산 + 템플릿쓰기
    ├── parser.py    ← Mechanism: SudoParser, ModuParser
    └── preset.py    ← Policy: 프리셋 정의 + REGISTRY
```

## 개발 환경 실행 (uv)

```bash
# 1. 최초 1회: 의존성 설치 + 패키지 editable 설치
uv sync

# 2. 실행
uv run onecell
```

## Windows .exe 빌드

```
build_windows.bat 더블클릭
→ dist/onecell_tool.exe 생성
```

## 새 업체 프리셋 추가 방법

1. `parser.py` — 형식이 다르면 `NewParser(BaseParser)` 추가
2. `preset.py` — `PresetConfig` 인스턴스 정의 후 `REGISTRY`에 등록
3. `app.py` 수정 불필요 (REGISTRY 동적 참조)

## 판매가 공식

```
판매가 = round((매입가 + 배송비) × 1.1 × (1 + 마진율/100) / 10) × 10
```