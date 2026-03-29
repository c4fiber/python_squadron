# onecell 자동 입력 도구 — 사용 설명서

## 폴더 구성

```
onecell_tool/
├── onecell_tool.exe       ← 실행 파일 (빌드 후 생성)
├── settings.ini           ← 설정 파일 (태그, 마진율)
├── main.py                ← 소스 코드
├── onecell_template.xlsx  ← 원셀 업로드 양식 (exe에 번들됨)
├── onecell_tool.spec      ← PyInstaller 빌드 설정
└── build_windows.bat      ← Windows 빌드 스크립트
```

## 사용 방법

1. `onecell_tool.exe` 실행
2. **설정** 영역에서 태그(`26신상` 등)와 마진율(%) 입력 후 **설정 저장** 클릭
3. **파일 선택...** 버튼으로 `product.xls` 파일 선택
4. **▶ 자동 입력** 버튼 클릭
5. 저장 팝업에서 결과 파일 위치 지정 후 저장

## 판매가 계산 공식

```
판매가 = round(매입가 × 1.1 × (1 + 마진율/100) / 10) × 10
```
- 매입가 × 1.1 : 부가세(VAT 10%) 포함
- × (1 + 마진율/100) : 마진 적용
- 10원 단위 반올림

## settings.ini 형식

```ini
[general]
tag = 26신상

[pricing]
margin_rate = 15
```

## 옵션 파싱 규칙

| 옵션명 원본 | 속성명1 | 속성명2 |
|---|---|---|
| 색상사이즈 | 색상 | 사이즈 |
| 품목사이즈 | 품목 | 사이즈 |
| 색상타입사이즈 (3개↑) | 색상 | 사이즈 (우선) |
| 사이즈 | 사이즈 | — |

옵션값은 개행(`\n`) 기준으로 분리합니다.  
색상값 없으면 `ONE COLOR`, 사이즈값 없으면 `ONE SIZE` 자동 입력.

## Windows 빌드 방법

```
build_windows.bat 더블클릭
→ dist/onecell_tool.exe 생성
```

## 업로드 자동화 (추후 구현)

`main.py`의 `OnecellUploader` 클래스에서 `login()`, `upload()` 메서드를 실제 로직으로 교체하면 됩니다.
