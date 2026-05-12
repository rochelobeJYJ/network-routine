# 네트워크 루틴

Windows에서 내부망/외부망 전환을 간단한 GUI로 처리하는 프로그램입니다.

## 핵심 기능

- 내부망 IP/DNS를 GUI에서 입력하고 저장합니다.
- 외부망은 DHCP 자동 설정으로 전환합니다.
- 자동 판단 기준을 `시간표만`, `시간표 + 네트워크 이름`, `네트워크 이름만` 중에서 고를 수 있습니다.
- 내부망으로 볼 Wi-Fi 이름과 외부망으로 볼 Wi-Fi 이름을 각각 여러 개 등록할 수 있습니다.
- 자동 루틴을 켜면 작업 스케줄러에 등록되어, 창을 닫아도 계속 동작합니다.
- 재부팅 후에도 `부팅`, `로그온`, `잠금 해제`, `콘솔 복귀`, `Wi-Fi 연결/해제`, `시간표 경계 시각` 기준으로 다시 검사합니다.
- 수동 전환을 하면 현재 판단 구간에서는 자동이 다시 덮어쓰지 않습니다.

## 배포 구조

- 사용자가 직접 실행하는 파일은 `NetworkRoutine.exe` 하나입니다.
- 다른 PC에서도 Python 설치 없이 `NetworkRoutine.exe`만 실행하면 됩니다.
- 자동 루틴을 켜고 `저장 및 반영`을 누르면 프로그램이 아래 고정 위치를 사용합니다.
- 실행 파일: `C:\Program Files\NetworkRoutine\NetworkRoutine.exe`
- 설정 파일: `C:\ProgramData\NetworkRoutine\network_routine_settings.json`

즉, 바탕화면에 둔 exe를 옮기거나 버전 파일명을 바꿔도, 자동 루틴은 고정 설치 경로 기준으로 계속 유지됩니다.

## 사용 방법

1. `NetworkRoutine.exe`를 실행합니다.
2. 관리자 권한 허용을 누릅니다.
3. 사용할 네트워크 어댑터를 선택합니다.
4. 내부망 `IP 주소 / 서브넷 마스크 / 게이트웨이 / DNS`를 입력합니다.
5. 자동 기준을 선택합니다.
6. 필요하면 내부망 이름과 외부망 이름을 입력합니다.
7. 시간표를 쓸 경우 출근/퇴근 시간을 설정합니다.
8. `자동 루틴 사용`을 켜고 `저장 및 반영`을 누릅니다.

처음 `저장 및 반영`을 누를 때는 관리자 권한으로 고정 경로 설치와 작업 등록을 함께 진행합니다.

## 자동 기준 설명

### 시간표만

- Wi-Fi 이름은 무시합니다.
- 시간표 기준으로 내부망/외부망을 자동 전환합니다.

### 시간표 + 네트워크 이름

- 내부망 이름이 맞으면 시간과 무관하게 내부망으로 판단합니다.
- 외부망 이름이 맞으면 시간과 무관하게 외부망으로 판단합니다.
- 현재 연결이 없어도 주변에서 감지된 Wi-Fi 목록에 이름이 보이면 그 기준으로 먼저 판단합니다.
- 둘 다 아니면 시간표 기준으로 판단합니다.

### 네트워크 이름만

- 내부망 이름이 맞으면 내부망으로 판단합니다.
- 외부망 이름이 맞으면 외부망으로 판단합니다.
- 현재 연결이 없어도 주변에서 감지된 Wi-Fi 목록에 이름이 보이면 그 기준으로 판단합니다.
- 둘 다 아니면 현재 상태를 그대로 유지합니다.

네트워크 이름 기준을 쓰려면 Windows 위치 서비스가 켜져 있어야 주변 Wi-Fi 검색이 가능합니다. 권한이 꺼져 있으면 앱이 안내 창을 띄우고 위치 설정 화면을 열어줍니다.

## 네트워크 이름 입력 규칙

- 여러 개를 쓰려면 쉼표로 구분해 입력하면 됩니다.
- 예시
- 내부망 이름: `회사1, 회사2`
- 외부망 이름: `집WiFi, 핫스팟`
- 같은 이름을 내부망과 외부망에 동시에 넣을 수는 없습니다.

## 내부망 입력 규칙

- 입력한 값만 반영합니다.
- IP를 바꾸려면 `IP 주소`와 `서브넷 마스크`를 함께 입력해야 합니다.
- `게이트웨이`, `기본 DNS`, `보조 DNS`는 비워둘 수 있습니다.
- DNS만 입력하면 DNS만 바꾸고, IP 관련 값이 비어 있으면 IP는 건드리지 않습니다.

## 등록 확인

자동 루틴을 켠 뒤 아래 작업들이 등록되어 있어야 합니다.

- `NetworkRoutine_Reconcile_Schedule`
- `NetworkRoutine_Reconcile_Startup`
- `NetworkRoutine_Reconcile_Logon`
- `NetworkRoutine_Reconcile_Unlock`
- `NetworkRoutine_Reconcile_ConsoleConnect`
- `NetworkRoutine_Reconcile_WifiChange`

참고:
- `Schedule`은 `네트워크 이름만` 모드에서는 사용하지 않을 수 있습니다.
- `WifiChange`는 `시간표만` 모드에서는 사용하지 않을 수 있습니다.

확인 명령:

```powershell
schtasks /Query /TN "NetworkRoutine_Reconcile_Schedule" /V /FO LIST
schtasks /Query /TN "NetworkRoutine_Reconcile_Startup" /V /FO LIST
schtasks /Query /TN "NetworkRoutine_Reconcile_Logon" /V /FO LIST
schtasks /Query /TN "NetworkRoutine_Reconcile_Unlock" /V /FO LIST
schtasks /Query /TN "NetworkRoutine_Reconcile_ConsoleConnect" /V /FO LIST
schtasks /Query /TN "NetworkRoutine_Reconcile_WifiChange" /V /FO LIST
```

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

## 개발자 로그

- 기본값은 비활성화입니다.
- 필요할 때만 아래처럼 켜서 실행하면 `network_routine.log`가 생성됩니다.

```powershell
$env:NETWORK_ROUTINE_DEV_LOG = "1"
python .\network_routine.py
```
