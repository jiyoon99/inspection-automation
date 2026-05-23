# Notebook Inspection Automation

Windows 노트북 검수 과정을 자동 진단, 수동 체크, 최종 판정, 리포트 저장 흐름으로 정리한 PowerShell GUI 도구입니다.

## Problem

- 노트북 검수 시 모델명, 시리얼, CPU, RAM/SSD, 배터리, 장치 인식, 포트, 키보드 상태를 반복 확인해야 합니다.
- 사람이 항목을 직접 기록하면 누락이나 판정 기준 차이가 생길 수 있습니다.
- 검수 결과를 이력과 리포트로 남길 필요가 있습니다.

## Solution

- 조회 가능한 하드웨어 정보는 자동 수집합니다.
- 작업자가 직접 확인해야 하는 항목은 GUI에서 수동 체크합니다.
- 자동 진단 결과와 수동 체크 결과를 합쳐 `합격 / 재검수 / 불합격`으로 판정합니다.
- CSV 이력과 PDF 리포트를 저장합니다.

## Tech Stack

| Area | Stack |
| --- | --- |
| Runtime | PowerShell |
| GUI | Windows Forms |
| Windows APIs | WMI/CIM, powercfg, device queries |
| Report | CSV, PDF print output |
| Platform | Windows |

## Skills

- PowerShell Windows Forms GUI 구성
- WMI/CIM 기반 하드웨어 정보 수집
- `powercfg` 기반 배터리 리포트 분석
- 자동 진단과 수동 체크를 결합한 판정 로직
- CSV 이력 저장
- PDF 출력 흐름 구성
- 공개 저장소에서 실제 장비 정보 제외

## Key Features

- 검수 기본 정보 입력
- 모델명, 시리얼, CPU, RAM/SSD 자동 수집
- 배터리 효율과 사이클 수 확인
- Wi-Fi, 카메라, 스피커 장치 인식 확인
- 키보드 입력 테스트
- USB/HDMI/LAN/오디오잭/충전포트 체크
- 자동 진단 결과 요약
- 최종 판정 자동화
- CSV 이력 저장
- PDF 리포트 출력

## Judgment Rules

| Condition | Result |
| --- | --- |
| 하나라도 `불량`이 있음 | `불합격` |
| `불량`은 없고 `확인필요`가 있음 | `재검수` |
| 모든 항목이 `정상` | `합격` |
| 배터리 효율 80% 미만 | `재검수` 기준 반영 |

## Preview

![Notebook Inspection Automation GUI preview](docs/assets/inspection-automation-preview.svg)

공개 저장소용 샘플 화면입니다. 실제 시리얼, 검수 이력, 리포트 파일은 포함하지 않았습니다.

## Run

배치 파일 실행:

```bat
run_inspection_app.bat
```

PowerShell 직접 실행:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\driver_gui.ps1
```

## Output

```text
data/
reports/
```

- `data/inspections.csv`: 검수 이력
- `reports/`: PDF 출력 결과

## Project Structure

```text
inspection-automation/
├── driver_gui.ps1
├── run_inspection_app.bat
├── README.md
└── .gitignore
```

## Safety

- 실제 검수 결과는 로컬 `data/`와 `reports/`에 저장됩니다.
- 공개 저장소에는 실제 장비 시리얼, 검수자명, 검수 이력, PDF 리포트를 포함하지 않습니다.
- Windows 권한 또는 장치 드라이버 상태에 따라 일부 자동 진단 항목은 조회가 실패할 수 있습니다.
