# Parcel Tracking Automation

[English](README.en.md)

엑셀 배송 목록에서 송장번호를 읽고, 배송조회 사이트에서 배송완료일을 조회한 뒤, 결과를 엑셀 복사본에 자동으로 반영하는 Windows용 파이썬 자동화 도구입니다.

## 주요 기능

- 파일 선택 창에서 원본 엑셀 파일을 선택합니다.
- 원본 파일과 같은 위치의 `updated` 폴더에 `updated_<원본파일명>` 복사본을 생성합니다.
- 첫 번째 시트에서 처리 대상 행을 찾고 송장번호를 수집합니다.
- Playwright로 Chrome을 열어 배송조회 사이트를 10건씩 조회합니다.
- 조회된 배송완료일을 엑셀의 지정 열에 입력하고 `log.txt` 처리 로그를 남깁니다.
- PyInstaller spec 파일(`OA_version1.3.spec`)로 실행 파일을 빌드할 수 있습니다.

## 프로젝트 구조

| 파일 | 역할 |
| --- | --- |
| `BatchRunner.py` | 전체 실행 흐름을 제어하는 진입점 |
| `ExcelAdapter.py` | 엑셀 파일 로드, 대상 행 추출, 배송완료일 저장 |
| `WebBrowserAdapter.py` | Playwright 기반 배송조회 사이트 자동화 |
| `DeliveryDateParser.py` | 조회 결과 문자열에서 배송완료일 파싱 |
| `ColorDetector.py` | 엑셀 행의 색상/미배송 사유 기준으로 처리 대상 판별 |
| `FileAdapter.py` | 원본 파일 복사 및 처리 로그 작성 |
| `Parcel.py`, `ParcelSheet.py` | 배송 데이터 모델 |
| `OA_version1.3.spec` | PyInstaller 빌드 설정 |

## 요구 사항

- Python 3.10 이상
- Google Chrome 또는 Playwright가 설치한 Chrome
- Windows 환경 권장
- `.xlsx` 형식의 엑셀 파일 권장

> 파일 선택 창은 `.xls`도 표시하지만, 내부에서 `openpyxl`을 사용하므로 구형 `.xls` 파일은 처리되지 않을 수 있습니다.

## 설치

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install openpyxl python-dotenv playwright pyinstaller
python -m playwright install chrome
```

이미 Google Chrome이 설치되어 있으면 마지막 명령은 환경에 따라 생략할 수 있습니다.

## 설정

실행 파일과 같은 위치 또는 소스 루트에 `.env` 파일이 필요합니다. 이 파일에는 운영 사이트 주소, 화면 selector, 엑셀 열 번호 등 배포 환경별 설정이 들어가므로 저장소나 공개 문서에 포함하지 않습니다.

설정 파일은 내부 배포 담당자에게 별도로 받아 사용하세요.

## 엑셀 처리 기준

1. 첫 번째 시트만 처리합니다.
2. 설정된 보고날짜 열에서 값이 `보고날짜`인 행을 헤더로 찾습니다.
3. 헤더 다음 행부터 보고날짜 열 값이 비어 있지 않은 동안 반복합니다.
4. 다음 조건 중 하나에 해당하는 행만 조회 대상으로 수집합니다.
   - 보고날짜 셀 배경이 비어 있거나 흰색으로 판단되는 행
   - 미배송 사유 상세 셀에 설정된 고객사명이 포함된 행
5. 재배송 송장번호 값이 12자리 숫자이면 이 값을 우선 사용하고, 그렇지 않으면 기본 송장번호를 사용합니다.
6. 조회 결과에서 날짜가 파싱되면 설정된 배송완료일 열에 날짜, 서식, 글꼴, 테두리를 적용합니다.

## 실행

```powershell
python BatchRunner.py
```

실행하면 파일 선택 창이 열립니다. 엑셀 파일을 선택하면 같은 폴더 아래에 다음 파일이 생성됩니다.

```text
updated/
  updated_<원본파일명>
  log.txt
```

## 빌드

```powershell
pyinstaller OA_version1.3.spec
```

빌드 결과는 일반적으로 `dist/OA_version1.3.exe`에 생성됩니다. spec 파일은 `.env`를 실행 파일에 포함하도록 설정되어 있으므로, 배포용 실행 파일을 만들 때는 운영 설정 파일 취급에 주의해야 합니다.

## 문제 해결

| 증상 | 확인할 내용 |
| --- | --- |
| Chrome을 찾을 수 없음 | Google Chrome 설치 여부를 확인하거나 `python -m playwright install chrome`을 실행합니다. |
| 환경 변수 오류 | `.env`가 실행 위치에 있는지, 내부 배포 설정이 누락되지 않았는지 확인합니다. |
| 배송완료일이 입력되지 않음 | 운영 설정이 실제 조회 화면 및 엑셀 양식과 맞는지 확인합니다. |
| 엑셀 파일을 열 수 없음 | 구형 `.xls` 대신 `.xlsx`로 저장한 파일을 사용합니다. |
