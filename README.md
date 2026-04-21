# EFC/LSS Collection Dashboard

`DynamicList.CSV` 기준으로 EFC/LSS 타리프 대비 실제 징수율을 확인하는 대시보드입니다. 화면은 `index.html`, 데이터는 `data.json`으로 분리되어 있으며 Google 로그인 후 데이터를 로드합니다.

## 실행

GitHub Pages 배포 URL에서 `index.html` 또는 `dashboard.html`을 열면 됩니다. Google OAuth 로그인이 필요하므로 로컬 파일 직접 열기보다 배포 URL에서 확인하는 것을 권장합니다.

상세 사용법은 대시보드 우측 상단의 `Guide` 버튼 또는 `guide.html`에서 확인할 수 있습니다. 가이드는 한국어/영어 전환을 지원합니다.

대시보드와 가이드 모두 `EN`/`KO` 버튼으로 한국어/영어 전환을 지원합니다.

대시보드 본문은 Google 로그인 후 표시됩니다. OAuth 앱에는 아래 값이 등록되어 있어야 합니다.

- 승인된 JavaScript 원본: `https://jkpark-create.github.io`
- 승인된 리디렉션 URI: `https://jkpark-create.github.io/EFC_LSS_monitoring/`
- 허용 도메인: `ekmtc.com`
- OAuth Client ID를 교체해야 하면 `build_dashboard.py`의 `GOOGLE_CLIENT_ID` 값을 수정한 뒤 대시보드를 다시 생성합니다.
- 앱은 Google OAuth에 항상 `https://jkpark-create.github.io/EFC_LSS_monitoring/`를 `redirect_uri`로 전달합니다.

CSV가 갱신되면 아래 명령으로 대시보드를 다시 생성합니다.

```bash
python3 build_dashboard.py
```

이 명령은 `index.html`, `dashboard.html`, `data.json`을 함께 생성합니다.

## ICC 매일 갱신

ICC `On-Demand Data`에서 엑셀을 내려받아 `DynamicList.CSV`를 교체하고 대시보드를 다시 생성하는 자동화는 `icc_daily_update.py`로 실행합니다.

최초 1회만 Playwright를 설치합니다.

```powershell
py -m pip install -r requirements.txt
py -m playwright install chromium
```

기본 조건은 아래와 같습니다.

- Document Name: `[영업팀] LSS & EFC 징수금액조회`
- 주차: 실행일 기준 ICC 금주를 종료년주로 사용하고 최근 4개 주를 조회합니다. 예를 들어 2026-04-21에 실행하면 금주가 16주이므로 `시작년주 202613`, `종료년주 202616`입니다.
- 조직: `O`
- 구분: `D`
- 다운로드 버튼: `Excel Down`

ICC 화면을 열어 다운로드까지 실행하려면 사내 ICC URL을 넘겨 실행합니다.

```powershell
.\run_icc_daily_update.ps1 -Url "https://ICC_ON_DEMAND_DATA_URL"
```

이미 내려받은 엑셀/CSV 파일만 반영할 때는 아래처럼 실행합니다.

```powershell
.\run_icc_daily_update.ps1 -DownloadFile ".\downloads\DynamicList.xlsx"
```

조건 계산만 확인할 때는 다음 명령을 사용합니다.

```powershell
py .\icc_daily_update.py --date 2026-04-21 --dry-run
```

Windows 작업 스케줄러에 매일 걸 때는 예를 들어 아래처럼 등록합니다.

```powershell
schtasks /Create /TN "EFC LSS ICC Update" /SC DAILY /ST 08:30 /TR "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"C:\Users\JKPARK\OneDrive\Documents\Claude\EFC_LSS_monitoring\run_icc_daily_update.ps1`" -Url `"https://ICC_ON_DEMAND_DATA_URL`" -Headless"
```

ICC 화면의 입력 컨트롤이 일반 HTML 입력이 아닌 커스텀 위젯이면 selector 환경변수로 직접 지정할 수 있습니다. 시작/종료년주가 한 칸이면 `ICC_SELECTOR_START`, `ICC_SELECTOR_END`에 각각 `202613`, `202616` 형태로 입력하고, 두 칸이면 `ICC_SELECTOR_START_YEAR`, `ICC_SELECTOR_START_WEEK`, `ICC_SELECTOR_END_YEAR`, `ICC_SELECTOR_END_WEEK`를 사용합니다. 그 외 예: `ICC_SELECTOR_ORG`, `ICC_SELECTOR_DIVISION`, `ICC_SELECTOR_DOCUMENT`, `ICC_SELECTOR_EXCEL_DOWN`.

## 계산 기준

- `CN -> JP`는 LSS 인상 타리프 `20 DRY USD 150`, `40 DRY USD 300`을 적용합니다.
- 중국 외 기점 EFC는 원산지 `CN/KR/JP/US`를 제외하고 목적지별 EFC 타리프를 적용합니다.
- 실제 징수액은 `20 lss + 40 lss` 또는 `20 efc + 40 efc` 합계입니다.
- 기대액은 `20갯수 * 20DRY tariff + 40갯수 * 40DRY tariff`입니다.
- 징수율은 `실제 징수액 / 타리프 기대액`입니다.
- Gap은 `실제 징수액 - 타리프 기대액`입니다. 음수는 shortfall(부족), 양수는 초과징수입니다.
- `20 o/f`와 `40 o/f`가 모두 빈칸이거나 0인 행은 모수에서 제외합니다.
- 원본 파일의 마지막 합계 행은 자동 제외합니다.
- 원본에는 `실적년/실적월/실적년주차`만 있고 ETD POL 일자가 없어, 효력일 기준의 행 단위 제외는 적용하지 않았습니다.
- 원본에 RF/RH 여부 컬럼이 없어 RF/RH 50% 할증은 계산에 반영하지 않았습니다.

## Drill-down

대시보드는 `선적지 -> 선적지-도착지 -> 고객` 레이어로 집계됩니다. 고객 기준은 `booking shipper`와 `handling consignee`를 전환해서 볼 수 있습니다.
