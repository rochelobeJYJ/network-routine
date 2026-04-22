# 네트워크 루틴

Windows에서 내부망/외부망 전환을 GUI로 처리하는 최소 프로그램입니다.

## 현재 구조

- `network_routine.py`
  - 관리자 권한 확보
  - 내부망 값 입력/저장
  - 수동 내부망/외부망 전환
  - 요일별 출근/퇴근 시간 설정
  - 작업 스케줄러 등록/해제
- `requirements.txt`
  - 소스 실행용 최소 의존성
- `requirements-dev.txt`
  - 빌드 포함 개발 의존성
- `network_routine_settings.json`
  - 첫 실행 후 자동 생성
- `build_exe.ps1`
  - PyInstaller로 `exe` 빌드

## 동작 방식

- `자동 루틴 사용`을 켜고 저장하면 작업 스케줄러에 2개 작업을 등록합니다.
- 1분마다 현재 시간표를 확인해 내부망/외부망 중 맞는 쪽으로 전환합니다.
- 로그인 직후에도 한 번 더 현재 시간 기준으로 상태를 맞춥니다.
- 자동 루틴을 끄고 저장하면 등록한 작업을 제거합니다.
- 배터리 사용 중이어도 자동 루틴은 계속 동작하도록 등록합니다.

## 내부망 입력 규칙

- 입력한 값만 반영합니다.
- IP를 바꾸려면 `IP 주소`와 `서브넷 마스크`를 함께 입력해야 합니다.
- `게이트웨이`, `기본 DNS`, `보조 DNS`는 비워둘 수 있습니다.
- DNS만 입력하면 DNS만 바꾸고, IP 관련 값이 비어 있으면 IP는 건드리지 않습니다.

## 실행

개발 중:

```powershell
py -3.13 -m pip install -r .\requirements-dev.txt
python .\network_routine.py
```

빌드:

```powershell
.\build_exe.ps1
```

기본값은 `Python 3.13`으로 빌드합니다.

빌드 후 결과물:

- `release_final\NetworkRoutine.exe`

## 루틴 등록 확인

자동 루틴을 켜고 저장한 뒤 아래 2개 작업이 보여야 정상입니다.

- `NetworkRoutine_Reconcile_Minutely`
- `NetworkRoutine_Reconcile_Logon`

확인 명령:

```powershell
schtasks /Query /TN "NetworkRoutine_Reconcile_Minutely" /V /FO LIST
schtasks /Query /TN "NetworkRoutine_Reconcile_Logon" /V /FO LIST
```

## 메모

- 기본 저장 파일은 실행 파일과 같은 폴더에 생성됩니다.
- 로그 파일은 기본적으로 생성하지 않습니다.
- 개발자가 로그가 필요하면 실행 전에 `NETWORK_ROUTINE_DEV_LOG=1` 환경 변수를 주면 `network_routine.log`를 저장합니다.
- 자동 루틴이 켜져 있으면 수동으로 다른 쪽으로 바꿔도 다음 검사 시 시간표 기준으로 다시 맞춰집니다.
