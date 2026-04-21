# EFC/LSS Collection Dashboard

`DynamicList.CSV` 기준으로 EFC/LSS 타리프 대비 실제 징수율을 확인하는 단일 HTML 대시보드입니다.

## 실행

브라우저에서 `index.html` 또는 `dashboard.html`을 열면 됩니다. 별도 서버나 패키지 설치가 필요하지 않습니다.

CSV가 갱신되면 아래 명령으로 대시보드를 다시 생성합니다.

```bash
python3 build_dashboard.py
```

## 계산 기준

- `CN -> JP`는 LSS 인상 타리프 `20 DRY USD 150`, `40 DRY USD 300`을 적용합니다.
- 중국 외 기점 EFC는 원산지 `CN/KR/JP/US`를 제외하고 목적지별 EFC 타리프를 적용합니다.
- 실제 징수액은 `20 lss + 40 lss` 또는 `20 efc + 40 efc` 합계입니다.
- 기대액은 `20갯수 * 20DRY tariff + 40갯수 * 40DRY tariff`입니다.
- 징수율은 `실제 징수액 / 타리프 기대액`입니다.
- `20 o/f`와 `40 o/f`가 모두 빈칸이거나 0인 행은 모수에서 제외합니다.
- 원본 파일의 마지막 합계 행은 자동 제외합니다.
- 원본에는 `실적년/실적월/실적년주차`만 있고 ETD POL 일자가 없어, 효력일 기준의 행 단위 제외는 적용하지 않았습니다.
- 원본에 RF/RH 여부 컬럼이 없어 RF/RH 50% 할증은 계산에 반영하지 않았습니다.

## Drill-down

대시보드는 `선적지 -> 선적지-도착지 -> 고객` 레이어로 집계됩니다. 고객 기준은 `booking shipper`와 `handling consignee`를 전환해서 볼 수 있습니다.
