# 네트워크 루틴

Windows에서 내부망/외부망 전환을 간단한 GUI로 처리하는 프로그램입니다.

## 사용자용 요약

- 내부망 IP/DNS를 GUI에서 입력하고 저장합니다.
- 외부망은 DHCP 자동 설정으로 전환합니다.
- 출근/퇴근 시간을 요일별로 지정할 수 있습니다.
- 회사 Wi-Fi 이름이 맞으면 시간표보다 먼저 내부망으로 판단할 수 있습니다.
- 자동 루틴을 켜면 출근/퇴근 경계 시각에만 자동 전환을 시도합니다.
- 로그인, 잠금 해제, 절전 복귀 계열 시점에도 현재 시간대 기준으로 한 번 더 보정합니다.
- 회사 Wi-Fi 이름 기준을 켠 경우 Wi-Fi 연결/해제 시점에도 다시 판단합니다.
- 같은 시간대에서 자동 전환은 한 번만 처리합니다.
- 수동 전환하면 현재 시간대에서는 자동이 다시 덮어쓰지 않습니다.
- 한 번 설정해두면 평소에는 프로그램을 열어둘 필요가 없습니다.
- 배터리 사용 중에도 자동 루틴은 계속 동작합니다.
- 수동 전환 버튼으로 언제든 즉시 변경할 수 있습니다.

## 배포 파일

- 최종 실행 파일: `release_final\NetworkRoutine.exe`
- 일반 사용자는 이 `exe`만 있으면 됩니다.
- `network_routine_settings.json`은 첫 실행 후 exe와 같은 폴더에 자동 생성됩니다.
- 로그 파일은 기본적으로 생성하지 않습니다.

## 사용 방법

1. `NetworkRoutine.exe`를 원하는 폴더에 둡니다.
2. 실행 시 관리자 권한 허용을 누릅니다.
3. 사용할 어댑터를 선택합니다.
4. 내부망 `IP 주소 / 서브넷 마스크 / 게이트웨이 / DNS`를 입력합니다.
5. 필요하면 `회사 Wi-Fi 이름 우선`을 켜고 회사 Wi-Fi 이름을 입력합니다.
6. 근무 시간과 요일을 설정합니다.
7. `자동 루틴 사용`을 켜고 `저장 및 반영`을 누릅니다.
8. 작업이 등록된 뒤에는 창을 닫아도 자동 루틴은 계속 동작합니다.

## 내부망 입력 규칙

- 입력한 값만 반영합니다.
- IP를 바꾸려면 `IP 주소`와 `서브넷 마스크`를 함께 입력해야 합니다.
- `게이트웨이`, `기본 DNS`, `보조 DNS`는 비워둘 수 있습니다.
- DNS만 입력하면 DNS만 바꾸고, IP 관련 값이 비어 있으면 IP는 건드리지 않습니다.

## 회사 Wi-Fi 이름 기준

- `회사 Wi-Fi 이름 우선`을 켜면 현재 연결된 Wi-Fi 이름이 입력한 값과 같을 때 내부망을 우선 적용합니다.
- 여러 이름을 쓰려면 쉼표로 구분해 입력하면 됩니다.
- 현재 Wi-Fi 이름이 맞지 않거나 확인되지 않으면 기존 시간표 기준으로 판단합니다.
- 이 옵션을 켜면 Wi-Fi 연결/해제 이벤트도 자동 루틴 재검사 트리거로 같이 등록됩니다.

## 등록 확인

자동 루틴을 켠 뒤 아래 작업들이 등록되어 있어야 합니다.

- `NetworkRoutine_Reconcile_Schedule`
- `NetworkRoutine_Reconcile_Logon`
- `NetworkRoutine_Reconcile_Unlock`
- `NetworkRoutine_Reconcile_ConsoleConnect`
- `NetworkRoutine_Reconcile_WifiChange` (`회사 Wi-Fi 이름 우선` 사용 시)

확인 명령:

```powershell
schtasks /Query /TN "NetworkRoutine_Reconcile_Schedule" /V /FO LIST
schtasks /Query /TN "NetworkRoutine_Reconcile_Logon" /V /FO LIST
schtasks /Query /TN "NetworkRoutine_Reconcile_Unlock" /V /FO LIST
schtasks /Query /TN "NetworkRoutine_Reconcile_ConsoleConnect" /V /FO LIST
schtasks /Query /TN "NetworkRoutine_Reconcile_WifiChange" /V /FO LIST
```

정상 기준:

- 정시 작업: `마지막 결과: 0`
- 로그온 작업: 첫 로그인 전에는 `267011 (0x41303)`일 수 있음
- `실행할 작업`은 `NetworkRoutine.exe --reconcile` 이어야 함
- `전원 관리:`는 비어 있거나 배터리 제한이 없어야 함

## 개발

개발 실행:

```powershell
py -3.13 -m pip install -r .\requirements-dev.txt
python .\network_routine.py
```

빌드:

```powershell
.\build_exe.ps1
```

기본 빌드 기준은 `Python 3.13`입니다.

## 개발자 로그

- 기본값은 비활성화입니다.
- 필요할 때만 아래처럼 켜서 실행하면 `network_routine.log`를 남깁니다.

```powershell
$env:NETWORK_ROUTINE_DEV_LOG = "1"
python .\network_routine.py
```
