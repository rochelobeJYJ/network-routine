from __future__ import annotations

import argparse
import base64
import ctypes
import datetime as dt
import html
import ipaddress
import json
import logging
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import tkinter as tk
from tkinter import messagebox, ttk
import xml.etree.ElementTree as ET

import win32com.client


APP_TITLE = "네트워크 루틴"
APP_FOLDER = "network_routine"
SETTINGS_NAME = "network_routine_settings.json"
LOG_NAME = "network_routine.log"
DEV_LOG_ENV = "NETWORK_ROUTINE_DEV_LOG"
TASK_SCHEDULE = "NetworkRoutine_Reconcile_Schedule"
TASK_LOGON = "NetworkRoutine_Reconcile_Logon"
TASK_UNLOCK = "NetworkRoutine_Reconcile_Unlock"
TASK_CONSOLE = "NetworkRoutine_Reconcile_ConsoleConnect"
LEGACY_TASK_MINUTE = "NetworkRoutine_Reconcile_Minutely"
CREATE_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)
TASK_CREATE_OR_UPDATE = 6
TASK_LOGON_INTERACTIVE_TOKEN = 3
TASK_RUNLEVEL_HIGHEST = 1
TASK_INSTANCES_IGNORE_NEW = 2
TASK_ACTION_EXEC = 0
TASK_TRIGGER_WEEKLY = 3
TASK_TRIGGER_LOGON = 9
TASK_TRIGGER_SESSION_STATE_CHANGE = 11
TASK_SESSION_STATE_CONSOLE_CONNECT = 1
TASK_SESSION_STATE_UNLOCK = 8
UTF8_BOM = b"\xef\xbb\xbf"
ROUTINE_TASKS = [TASK_SCHEDULE, TASK_LOGON, TASK_UNLOCK, TASK_CONSOLE]
TASKS_TO_DELETE = [LEGACY_TASK_MINUTE, *ROUTINE_TASKS]
DAY_XML_NAMES = {
    "mon": "Monday",
    "tue": "Tuesday",
    "wed": "Wednesday",
    "thu": "Thursday",
    "fri": "Friday",
    "sat": "Saturday",
    "sun": "Sunday",
}
DAY_BITMASKS = {
    "sun": 0x01,
    "mon": 0x02,
    "tue": 0x04,
    "wed": 0x08,
    "thu": 0x10,
    "fri": 0x20,
    "sat": 0x40,
}

DAY_KEYS = ["mon", "tue", "wed", "thu", "fri", "sat", "sun"]
DAY_LABELS = {
    "mon": "월",
    "tue": "화",
    "wed": "수",
    "thu": "목",
    "fri": "금",
    "sat": "토",
    "sun": "일",
}

DEFAULT_SETTINGS = {
    "adapter_name": "Wi-Fi",
    "automation_enabled": False,
    "schedule": {
        "base_start": "08:20",
        "base_end": "16:20",
        "days": {
            "mon": {"enabled": True, "start": "08:20", "end": "16:20"},
            "tue": {"enabled": True, "start": "08:20", "end": "16:20"},
            "wed": {"enabled": True, "start": "08:20", "end": "16:20"},
            "thu": {"enabled": True, "start": "08:20", "end": "16:20"},
            "fri": {"enabled": True, "start": "08:20", "end": "16:20"},
            "sat": {"enabled": False, "start": "08:20", "end": "16:20"},
            "sun": {"enabled": False, "start": "08:20", "end": "16:20"},
        },
    },
    "internal": {
        "ip": "",
        "mask": "",
        "gateway": "",
        "dns1": "",
        "dns2": "",
    },
    "last_handled_segment_id": "",
    "last_handled_mode": "",
    "last_handled_at": "",
    "manual_override_segment_id": "",
    "manual_override_mode": "",
    "manual_override_at": "",
    "last_applied_mode": "",
    "last_applied_at": "",
    "last_message": "아직 적용 기록이 없습니다.",
}


def app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def settings_path() -> Path:
    return app_dir() / SETTINGS_NAME


def log_path() -> Path:
    return app_dir() / LOG_NAME


def migrate_text_file_to_utf8_sig(path: Path) -> None:
    if not path.exists():
        return

    data = path.read_bytes()
    if not data or data.startswith(UTF8_BOM):
        return

    for encoding in ("utf-8", "cp949"):
        try:
            text = data.decode(encoding)
            path.write_text(text, encoding="utf-8-sig")
            return
        except UnicodeDecodeError:
            continue


def is_dev_logging_enabled() -> bool:
    value = os.environ.get(DEV_LOG_ENV, "").strip().lower()
    return value in {"1", "true", "yes", "on"}


def configure_logging() -> None:
    root_logger = logging.getLogger()
    root_logger.handlers.clear()
    root_logger.setLevel(logging.INFO)

    if not is_dev_logging_enabled():
        root_logger.addHandler(logging.NullHandler())
        return

    migrate_text_file_to_utf8_sig(log_path())
    file_handler = logging.FileHandler(str(log_path()), encoding="utf-8-sig")
    file_handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
    root_logger.addHandler(file_handler)
    logging.info(
        "개발자 로그 저장 활성화: %s=1 일 때만 로그 파일이 생성됩니다.",
        DEV_LOG_ENV,
    )


def now_text() -> str:
    return dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def deep_merge(default: object, loaded: object) -> object:
    if isinstance(default, dict) and isinstance(loaded, dict):
        merged = {}
        for key, default_value in default.items():
            if key in loaded:
                merged[key] = deep_merge(default_value, loaded[key])
            else:
                merged[key] = default_value
        for key, loaded_value in loaded.items():
            if key not in merged:
                merged[key] = loaded_value
        return merged
    return loaded if loaded is not None else default


def load_settings() -> dict:
    path = settings_path()
    if not path.exists():
        return json.loads(json.dumps(DEFAULT_SETTINGS))

    try:
        with path.open("r", encoding="utf-8") as file:
            loaded = json.load(file)
        return deep_merge(DEFAULT_SETTINGS, loaded)
    except Exception as exc:
        logging.exception("설정 파일을 읽지 못했습니다.")
        raise RuntimeError(f"설정 파일을 읽지 못했습니다: {exc}") from exc


def save_settings(data: dict) -> None:
    path = settings_path()
    temp_path = path.with_suffix(".tmp")
    with temp_path.open("w", encoding="utf-8") as file:
        json.dump(data, file, ensure_ascii=False, indent=2)
    temp_path.replace(path)


def is_admin() -> bool:
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def relaunch_as_admin(extra_args: list[str] | None = None) -> None:
    args = extra_args if extra_args is not None else sys.argv[1:]

    if getattr(sys, "frozen", False):
        executable = sys.executable
        params = subprocess.list2cmdline(args)
    else:
        executable = sys.executable
        params = subprocess.list2cmdline([str(Path(__file__).resolve()), *args])

    result = ctypes.windll.shell32.ShellExecuteW(
        None,
        "runas",
        executable,
        params,
        str(app_dir()),
        1,
    )
    if result <= 32:
        raise RuntimeError("관리자 권한으로 다시 실행하지 못했습니다.")
    raise SystemExit(0)


def ensure_admin() -> None:
    if not is_admin():
        relaunch_as_admin()


def run_command(command: list[str], check: bool = True) -> subprocess.CompletedProcess[str]:
    logging.info("실행: %s", command)
    result = subprocess.run(
        command,
        capture_output=True,
        text=True,
        creationflags=CREATE_NO_WINDOW,
        check=False,
    )
    if check and result.returncode != 0:
        message = (result.stderr or result.stdout or "").strip() or "알 수 없는 오류"
        raise RuntimeError(message)
    return result


def run_powershell_script(script: str, check: bool = True) -> subprocess.CompletedProcess[str]:
    encoded = base64.b64encode(script.encode("utf-16le")).decode("ascii")
    return run_command(
        ["powershell", "-NoProfile", "-EncodedCommand", encoded],
        check=check,
    )


def run_powershell_json(script: str) -> dict:
    result = run_powershell_script(script, check=True)
    output = result.stdout.strip()
    if not output:
        return {}
    return json.loads(output)


def ensure_list(value: object) -> list[str]:
    if value is None:
        return []
    if isinstance(value, list):
        return [str(item) for item in value if str(item).strip()]
    if isinstance(value, str):
        return [value] if value.strip() else []
    return [str(value)]


def list_adapters() -> list[str]:
    powershell = [
        "powershell",
        "-NoProfile",
        "-Command",
        (
            "Get-NetAdapter | "
            "Where-Object { $_.Status -ne 'Disabled' -and $_.Name } | "
            "Sort-Object Name | "
            "Select-Object -ExpandProperty Name"
        ),
    ]
    try:
        result = run_command(powershell, check=True)
        names = [line.strip() for line in result.stdout.splitlines() if line.strip()]
        return names or ["Wi-Fi"]
    except Exception:
        logging.exception("네트워크 어댑터 목록을 읽지 못했습니다.")
        return ["Wi-Fi"]


def validate_ipv4(value: str, label: str) -> None:
    if not value:
        return
    try:
        ipaddress.IPv4Address(value)
    except Exception as exc:
        raise ValueError(f"{label} 형식이 올바르지 않습니다.") from exc


def parse_time(value: str) -> dt.time:
    try:
        return dt.datetime.strptime(value.strip(), "%H:%M").time()
    except ValueError as exc:
        raise ValueError("시간은 HH:MM 형식으로 입력해주십시오.") from exc


def mode_label(mode: str) -> str:
    if mode == "internal":
        return "내부망"
    if mode == "external":
        return "외부망"
    return "-"


def internal_interval_for_date(settings: dict, date_value: dt.date) -> tuple[dt.datetime, dt.datetime] | None:
    day_key = DAY_KEYS[date_value.weekday()]
    day_config = settings["schedule"]["days"][day_key]
    if not day_config["enabled"]:
        return None

    start = parse_time(day_config["start"])
    end = parse_time(day_config["end"])
    if start == end:
        return None

    start_dt = dt.datetime.combine(date_value, start)
    end_date = date_value if start < end else date_value + dt.timedelta(days=1)
    end_dt = dt.datetime.combine(end_date, end)
    return start_dt, end_dt


def merged_internal_intervals(
    settings: dict,
    center: dt.datetime | None = None,
    days_before: int = 8,
    days_after: int = 8,
) -> list[tuple[dt.datetime, dt.datetime]]:
    center = center or dt.datetime.now()
    base_date = center.date()
    intervals: list[tuple[dt.datetime, dt.datetime]] = []

    for offset in range(-days_before, days_after + 1):
        interval = internal_interval_for_date(settings, base_date + dt.timedelta(days=offset))
        if interval is not None:
            intervals.append(interval)

    intervals.sort(key=lambda item: item[0])
    merged: list[list[dt.datetime]] = []
    for start_dt, end_dt in intervals:
        if not merged or start_dt > merged[-1][1]:
            merged.append([start_dt, end_dt])
            continue
        if end_dt > merged[-1][1]:
            merged[-1][1] = end_dt

    return [(start_dt, end_dt) for start_dt, end_dt in merged]


def current_segment_info(settings: dict, when: dt.datetime | None = None) -> dict:
    when = when or dt.datetime.now()
    intervals = merged_internal_intervals(settings, when)

    for start_dt, end_dt in intervals:
        if start_dt <= when < end_dt:
            return {
                "mode": "internal",
                "segment_id": f"internal|{start_dt.isoformat()}|{end_dt.isoformat()}",
                "start": start_dt,
                "end": end_dt,
            }

    previous_end: dt.datetime | None = None
    next_start: dt.datetime | None = None
    for start_dt, end_dt in intervals:
        if end_dt <= when:
            previous_end = end_dt
            continue
        if start_dt > when:
            next_start = start_dt
            break

    if previous_end is None and next_start is None:
        segment_id = "external|always"
    else:
        previous_key = previous_end.isoformat() if previous_end else "past"
        next_key = next_start.isoformat() if next_start else "future"
        segment_id = f"external|{previous_key}|{next_key}"

    return {
        "mode": "external",
        "segment_id": segment_id,
        "start": previous_end,
        "end": next_start,
    }


def desired_mode_for(settings: dict, when: dt.datetime | None = None) -> str:
    return current_segment_info(settings, when)["mode"]


def schedule_summary(settings: dict) -> str:
    items = []
    for day in DAY_KEYS:
        day_config = settings["schedule"]["days"][day]
        if day_config["enabled"]:
            items.append(f"{DAY_LABELS[day]} {day_config['start']}~{day_config['end']}")
    return ", ".join(items) if items else "선택된 요일이 없습니다."


def normalize_internal_commands(adapter: str, internal: dict) -> list[list[str]]:
    adapter_arg = f'name="{adapter}"'
    commands: list[list[str]] = []

    ip = internal["ip"].strip()
    mask = internal["mask"].strip()
    gateway = internal["gateway"].strip()
    dns_values = [internal["dns1"].strip(), internal["dns2"].strip()]
    dns_values = [value for value in dns_values if value]

    if ip or mask or gateway:
        if not ip or not mask:
            raise ValueError("IP 주소와 서브넷 마스크는 함께 입력해야 합니다.")
        address_command = [
            "netsh",
            "interface",
            "ipv4",
            "set",
            "address",
            adapter_arg,
            "source=static",
            f"address={ip}",
            f"mask={mask}",
        ]
        if gateway:
            address_command.append(f"gateway={gateway}")
        commands.append(address_command)

    if dns_values:
        commands.append(
            [
                "netsh",
                "dnsclient",
                "set",
                "dnsservers",
                adapter_arg,
                "source=static",
                f"address={dns_values[0]}",
            ]
        )
        for index, value in enumerate(dns_values[1:], start=2):
            commands.append(
                [
                    "netsh",
                    "dnsclient",
                    "add",
                    "dnsservers",
                    adapter_arg,
                    f"index={index}",
                    f"address={value}",
                ]
            )

    return commands


def external_commands(adapter: str) -> list[list[str]]:
    adapter_arg = f'name="{adapter}"'
    return [
        ["netsh", "interface", "ipv4", "set", "address", adapter_arg, "source=dhcp"],
        ["netsh", "dnsclient", "set", "dnsservers", adapter_arg, "source=dhcp"],
    ]


def validate_settings(settings: dict) -> None:
    adapter = settings["adapter_name"].strip()
    if not adapter:
        raise ValueError("적용할 네트워크 어댑터를 선택해주십시오.")

    internal = settings["internal"]
    validate_ipv4(internal["ip"].strip(), "IP 주소")
    validate_ipv4(internal["mask"].strip(), "서브넷 마스크")
    validate_ipv4(internal["gateway"].strip(), "게이트웨이")
    validate_ipv4(internal["dns1"].strip(), "기본 DNS")
    validate_ipv4(internal["dns2"].strip(), "보조 DNS")

    for day in DAY_KEYS:
        day_config = settings["schedule"]["days"][day]
        parse_time(day_config["start"])
        parse_time(day_config["end"])


def clear_stale_manual_override(settings: dict, current_segment_id: str) -> bool:
    if settings.get("manual_override_segment_id") == current_segment_id:
        return False
    if not settings.get("manual_override_segment_id"):
        return False

    settings["manual_override_segment_id"] = ""
    settings["manual_override_mode"] = ""
    settings["manual_override_at"] = ""
    return True


def record_result(
    settings: dict,
    mode: str,
    message: str,
    *,
    segment_info: dict | None = None,
    mark_handled: bool = False,
) -> None:
    current_segment_id = segment_info["segment_id"] if segment_info else ""
    if current_segment_id:
        clear_stale_manual_override(settings, current_segment_id)

    settings["last_applied_mode"] = mode
    settings["last_applied_at"] = now_text()
    settings["last_message"] = message

    if mark_handled and segment_info is not None:
        settings["last_handled_segment_id"] = segment_info["segment_id"]
        settings["last_handled_mode"] = segment_info["mode"]
        settings["last_handled_at"] = settings["last_applied_at"]

    save_settings(settings)


def mark_manual_override(settings: dict, mode: str, when: dt.datetime | None = None) -> None:
    segment_info = current_segment_info(settings, when)
    settings["manual_override_segment_id"] = segment_info["segment_id"]
    settings["manual_override_mode"] = mode
    settings["manual_override_at"] = now_text()
    save_settings(settings)


def read_network_state(adapter: str) -> dict:
    adapter_escaped = adapter.replace("'", "''")
    script = f"""
$alias = '{adapter_escaped}'
$iface = Get-NetIPInterface -AddressFamily IPv4 -InterfaceAlias $alias -ErrorAction Stop | Select-Object -First 1
$ipInfo = Get-NetIPAddress -InterfaceAlias $alias -AddressFamily IPv4 -ErrorAction SilentlyContinue |
    Where-Object {{ $_.IPAddress -and $_.IPAddress -notlike '169.254.*' }} |
    Select-Object -ExpandProperty IPAddress
$dnsInfo = Get-DnsClientServerAddress -InterfaceAlias $alias -AddressFamily IPv4 -ErrorAction SilentlyContinue
$dns = @()
if ($dnsInfo) {{ $dns = @($dnsInfo.ServerAddresses) }}
[ordered]@{{
    dhcp = "$($iface.Dhcp)"
    ips = @($ipInfo)
    dns = @($dns)
}} | ConvertTo-Json -Compress
"""
    state = run_powershell_json(script)
    state.setdefault("ips", [])
    state.setdefault("dns", [])
    return state


def mode_matches_current_state(settings: dict, desired_mode: str) -> bool:
    try:
        state = read_network_state(settings["adapter_name"])
    except Exception:
        logging.exception("현재 네트워크 상태 확인 실패")
        return settings.get("last_applied_mode") == desired_mode

    dhcp_enabled = str(state.get("dhcp", "")).lower() == "enabled"
    actual_ips = ensure_list(state.get("ips"))
    actual_dns = ensure_list(state.get("dns"))

    if desired_mode == "external":
        return dhcp_enabled

    internal = settings["internal"]
    has_expected_ip = bool(internal["ip"].strip())
    expected_dns = [value for value in [internal["dns1"].strip(), internal["dns2"].strip()] if value]

    if has_expected_ip and internal["ip"].strip() not in actual_ips:
        return False

    for index, value in enumerate(expected_dns):
        if index >= len(actual_dns) or actual_dns[index] != value:
            return False

    if has_expected_ip or expected_dns:
        return True
    return settings.get("last_applied_mode") == desired_mode


def apply_mode(
    mode: str,
    settings: dict,
    reason: str = "",
    *,
    segment_info: dict | None = None,
    mark_handled: bool = False,
) -> str:
    adapter = settings["adapter_name"].strip()
    if mode == "internal":
        commands = normalize_internal_commands(adapter, settings["internal"])
        if not commands:
            raise ValueError("내부망으로 바꿀 설정값이 없습니다. IP 또는 DNS를 입력해주십시오.")
        label = mode_label("internal")
    elif mode == "external":
        commands = external_commands(adapter)
        label = mode_label("external")
    else:
        raise ValueError("알 수 없는 전환 모드입니다.")

    for command in commands:
        run_command(command, check=True)

    suffix = f" ({reason})" if reason else ""
    message = f"{label} 적용 완료{suffix}"
    record_result(settings, mode, message, segment_info=segment_info, mark_handled=mark_handled)
    logging.info(message)
    return message


def reconcile_now(settings: dict) -> str:
    if not settings.get("automation_enabled"):
        return "자동 루틴이 꺼져 있어 검사만 생략했습니다."

    segment_info = current_segment_info(settings)
    desired = segment_info["mode"]
    current_segment_id = segment_info["segment_id"]
    clear_stale_manual_override(settings, current_segment_id)

    if settings.get("manual_override_segment_id") == current_segment_id:
        manual_mode = settings.get("manual_override_mode") or desired
        message = f"수동 전환 유지: {mode_label(manual_mode)}"
        settings["last_message"] = message
        save_settings(settings)
        return message

    if settings.get("last_handled_segment_id") == current_segment_id:
        return f"이미 처리한 시간대 유지: {mode_label(desired)}"

    if mode_matches_current_state(settings, desired):
        record_result(
            settings,
            desired,
            f"현재 상태 확인: {mode_label(desired)}",
            segment_info=segment_info,
            mark_handled=True,
        )
        return f"현재 상태 확인: {mode_label(desired)}"
    return apply_mode(
        desired,
        settings,
        reason="자동 루틴",
        segment_info=segment_info,
        mark_handled=True,
    )


def build_task_runner_action(extra_args: list[str]) -> tuple[str, str]:
    if getattr(sys, "frozen", False):
        return sys.executable, subprocess.list2cmdline(extra_args)
    else:
        pythonw = Path(sys.executable).with_name("pythonw.exe")
        launcher = str(pythonw if pythonw.exists() else sys.executable)
        arguments = subprocess.list2cmdline([str(Path(__file__).resolve()), *extra_args])
        return launcher, arguments


def ps_literal(value: str) -> str:
    return "'" + value.replace("'", "''") + "'"


def current_task_user() -> str:
    domain = os.environ.get("USERDOMAIN", "").strip()
    user = os.environ.get("USERNAME", "").strip()
    if domain and user:
        return f"{domain}\\{user}"
    if user:
        return user
    return run_command(["whoami"], check=True).stdout.strip()


def task_service():
    service = win32com.client.Dispatch("Schedule.Service")
    service.Connect()
    return service


def current_identity() -> tuple[str, str]:
    domain = os.environ.get("USERDOMAIN", "").strip()
    user = os.environ.get("USERNAME", "").strip()
    user_id = current_task_user()
    if not domain and "\\" in user_id:
        domain = user_id.split("\\", 1)[0]
    if not user and "\\" in user_id:
        user = user_id.split("\\", 1)[1]
    return user_id, user or user_id


def delete_task(task_name: str) -> None:
    try:
        root = task_service().GetFolder("\\")
        root.DeleteTask(task_name, 0)
        logging.info("작업 삭제: %s", task_name)
    except Exception:
        result = run_command(["schtasks", "/Delete", "/TN", task_name, "/F"], check=False)
        if result.returncode == 0:
            logging.info("작업 삭제: %s", task_name)


def create_logon_trigger(task_definition, user_id: str):
    trigger = task_definition.Triggers.Create(TASK_TRIGGER_LOGON)
    trigger.Enabled = True
    trigger.UserId = user_id
    return trigger


def create_session_trigger(task_definition, user_id: str, state_change: int):
    trigger = task_definition.Triggers.Create(TASK_TRIGGER_SESSION_STATE_CHANGE)
    trigger.Enabled = True
    trigger.UserId = user_id
    trigger.StateChange = state_change
    return trigger


def next_weekly_boundary(day_key: str, time_value: dt.time, now: dt.datetime | None = None) -> dt.datetime:
    now = now or dt.datetime.now()
    target_weekday = DAY_KEYS.index(day_key)
    days_ahead = (target_weekday - now.weekday()) % 7
    candidate_date = now.date() + dt.timedelta(days=days_ahead)
    candidate = dt.datetime.combine(candidate_date, time_value)
    if candidate <= now:
        candidate += dt.timedelta(days=7)
    return candidate


def build_schedule_trigger_specs(settings: dict) -> list[dict]:
    specs: list[dict] = []
    seen: set[tuple[str, str]] = set()

    for day_index, day_key in enumerate(DAY_KEYS):
        day_config = settings["schedule"]["days"][day_key]
        if not day_config["enabled"]:
            continue

        start = parse_time(day_config["start"])
        end = parse_time(day_config["end"])
        if start == end:
            continue

        start_spec = (day_key, start.strftime("%H:%M"))
        if start_spec not in seen:
            seen.add(start_spec)
            specs.append({"day_key": day_key, "time": start})

        end_day_key = DAY_KEYS[(day_index + (0 if start < end else 1)) % 7]
        end_spec = (end_day_key, end.strftime("%H:%M"))
        if end_spec not in seen:
            seen.add(end_spec)
            specs.append({"day_key": end_day_key, "time": end})

    return specs


def create_weekly_trigger(task_definition, day_key: str, time_value: dt.time):
    trigger = task_definition.Triggers.Create(TASK_TRIGGER_WEEKLY)
    trigger.StartBoundary = next_weekly_boundary(day_key, time_value).strftime("%Y-%m-%dT%H:%M:%S")
    trigger.Enabled = True
    trigger.WeeksInterval = 1
    trigger.DaysOfWeek = DAY_BITMASKS[day_key]
    return trigger


def session_trigger_xml(trigger_kind: str, user_id_xml: str) -> str:
    if trigger_kind == "unlock":
        state_change = "SessionUnlock"
    elif trigger_kind == "console":
        state_change = "ConsoleConnect"
    else:
        raise ValueError("알 수 없는 세션 상태 트리거입니다.")

    return f"""    <SessionStateChangeTrigger>
      <Enabled>true</Enabled>
      <UserId>{user_id_xml}</UserId>
      <StateChange>{state_change}</StateChange>
    </SessionStateChangeTrigger>"""


def weekly_trigger_xml(day_key: str, time_value: dt.time) -> str:
    start_boundary = next_weekly_boundary(day_key, time_value).strftime("%Y-%m-%dT%H:%M:%S")
    day_xml_name = DAY_XML_NAMES[day_key]
    return f"""    <CalendarTrigger>
      <StartBoundary>{start_boundary}</StartBoundary>
      <Enabled>true</Enabled>
      <ScheduleByWeek>
        <WeeksInterval>1</WeeksInterval>
        <DaysOfWeek>
          <{day_xml_name}/>
        </DaysOfWeek>
      </ScheduleByWeek>
    </CalendarTrigger>"""


def configure_task_definition(task_definition, user_id: str) -> None:
    settings = task_definition.Settings
    settings.Enabled = True
    settings.StartWhenAvailable = True
    settings.DisallowStartIfOnBatteries = False
    settings.StopIfGoingOnBatteries = False
    settings.AllowDemandStart = True
    settings.Hidden = False
    settings.RunOnlyIfIdle = False
    settings.ExecutionTimeLimit = "PT72H"
    settings.MultipleInstances = TASK_INSTANCES_IGNORE_NEW

    principal = task_definition.Principal
    principal.UserId = user_id
    principal.LogonType = TASK_LOGON_INTERACTIVE_TOKEN
    principal.RunLevel = TASK_RUNLEVEL_HIGHEST


def configure_task_action(task_definition, command: str, arguments: str) -> None:
    action = task_definition.Actions.Create(TASK_ACTION_EXEC)
    action.Path = command
    action.Arguments = arguments
    action.WorkingDirectory = str(app_dir())


def create_task_via_com(task_name: str, trigger_kind: str, command: str, arguments: str) -> None:
    service = task_service()
    root = service.GetFolder("\\")
    task_definition = service.NewTask(0)

    user_id, _ = current_identity()
    configure_task_definition(task_definition, user_id)

    if trigger_kind == "logon":
        create_logon_trigger(task_definition, user_id)
    elif trigger_kind == "unlock":
        create_session_trigger(task_definition, user_id, TASK_SESSION_STATE_UNLOCK)
    elif trigger_kind == "console":
        create_session_trigger(task_definition, user_id, TASK_SESSION_STATE_CONSOLE_CONNECT)
    else:
        raise ValueError("알 수 없는 작업 트리거입니다.")

    configure_task_action(task_definition, command, arguments)

    root.RegisterTaskDefinition(
        task_name,
        task_definition,
        TASK_CREATE_OR_UPDATE,
        "",
        "",
        TASK_LOGON_INTERACTIVE_TOKEN,
    )
    logging.info("작업 등록: %s", task_name)


def create_schedule_task_via_com(task_name: str, settings: dict, command: str, arguments: str) -> None:
    service = task_service()
    root = service.GetFolder("\\")
    task_definition = service.NewTask(0)

    user_id, _ = current_identity()
    configure_task_definition(task_definition, user_id)

    trigger_specs = build_schedule_trigger_specs(settings)
    if not trigger_specs:
        raise ValueError("등록할 자동 경계 시간이 없습니다.")

    for spec in trigger_specs:
        create_weekly_trigger(task_definition, spec["day_key"], spec["time"])

    configure_task_action(task_definition, command, arguments)

    root.RegisterTaskDefinition(
        task_name,
        task_definition,
        TASK_CREATE_OR_UPDATE,
        "",
        "",
        TASK_LOGON_INTERACTIVE_TOKEN,
    )
    logging.info("작업 등록(일정): %s", task_name)


def build_task_xml_document(task_name: str, trigger_blocks: list[str], command: str, arguments: str) -> str:
    user_id, _ = current_identity()
    working_directory = str(app_dir())
    author = html.escape(user_id)
    command_xml = html.escape(command)
    arguments_xml = html.escape(arguments)
    working_dir_xml = html.escape(working_directory)
    user_id_xml = html.escape(user_id)

    return f"""<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.4" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Author>{author}</Author>
    <URI>\\{html.escape(task_name)}</URI>
  </RegistrationInfo>
  <Triggers>
{chr(10).join(trigger_blocks)}
  </Triggers>
  <Principals>
    <Principal id="Author">
      <UserId>{user_id_xml}</UserId>
      <LogonType>InteractiveToken</LogonType>
      <RunLevel>HighestAvailable</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>true</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>false</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>PT72H</ExecutionTimeLimit>
    <Priority>7</Priority>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>{command_xml}</Command>
      <Arguments>{arguments_xml}</Arguments>
      <WorkingDirectory>{working_dir_xml}</WorkingDirectory>
    </Exec>
  </Actions>
</Task>
"""


def build_task_xml_content(task_name: str, trigger_kind: str, command: str, arguments: str) -> str:
    user_id, _ = current_identity()
    user_id_xml = html.escape(user_id)

    if trigger_kind == "logon":
        trigger_block = f"""    <LogonTrigger>
      <Enabled>true</Enabled>
      <UserId>{user_id_xml}</UserId>
    </LogonTrigger>"""
    elif trigger_kind in {"unlock", "console"}:
        trigger_block = session_trigger_xml(trigger_kind, user_id_xml)
    else:
        raise ValueError("알 수 없는 작업 트리거입니다.")

    return build_task_xml_document(task_name, [trigger_block], command, arguments)


def build_schedule_task_xml_content(task_name: str, settings: dict, command: str, arguments: str) -> str:
    trigger_blocks = [weekly_trigger_xml(spec["day_key"], spec["time"]) for spec in build_schedule_trigger_specs(settings)]
    if not trigger_blocks:
        raise ValueError("등록할 자동 경계 시간이 없습니다.")
    return build_task_xml_document(task_name, trigger_blocks, command, arguments)


def create_task_via_xml(task_name: str, trigger_kind: str, command: str, arguments: str) -> None:
    xml_content = build_task_xml_content(task_name, trigger_kind, command, arguments)
    temp_path: str | None = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="w",
            encoding="utf-16",
            suffix=".xml",
            delete=False,
        ) as file:
            file.write(xml_content)
            temp_path = file.name
        run_command(["schtasks", "/Create", "/TN", task_name, "/XML", temp_path, "/F"], check=True)
        logging.info("작업 등록(XML): %s", task_name)
    finally:
        if temp_path:
            try:
                Path(temp_path).unlink(missing_ok=True)
            except Exception:
                logging.exception("임시 XML 삭제 실패: %s", temp_path)


def create_schedule_task_via_xml(task_name: str, settings: dict, command: str, arguments: str) -> None:
    xml_content = build_schedule_task_xml_content(task_name, settings, command, arguments)
    temp_path: str | None = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="w",
            encoding="utf-16",
            suffix=".xml",
            delete=False,
        ) as file:
            file.write(xml_content)
            temp_path = file.name
        run_command(["schtasks", "/Create", "/TN", task_name, "/XML", temp_path, "/F"], check=True)
        logging.info("작업 등록(XML 일정): %s", task_name)
    finally:
        if temp_path:
            try:
                Path(temp_path).unlink(missing_ok=True)
            except Exception:
                logging.exception("임시 XML 삭제 실패: %s", temp_path)


def create_task_via_schtasks(task_name: str, trigger_kind: str, command: str, arguments: str) -> None:
    task_run = f"{subprocess.list2cmdline([command])} {arguments}".strip()
    if trigger_kind == "logon":
        schedule_args = ["/SC", "ONLOGON"]
    else:
        raise RuntimeError("세션 상태 트리거는 schtasks 직접 등록이 지원되지 않습니다.")

    run_command(
        ["schtasks", "/Create", "/TN", task_name, *schedule_args, "/TR", task_run, "/RL", "HIGHEST", "/F"],
        check=True,
    )
    logging.info("작업 등록(SCHTASKS): %s", task_name)


def create_task(task_name: str, trigger_kind: str, command: str, arguments: str) -> None:
    try:
        create_task_via_com(task_name, trigger_kind, command, arguments)
        return
    except Exception:
        logging.exception("COM 작업 등록 실패, XML 등록 시도: %s", task_name)

    try:
        create_task_via_xml(task_name, trigger_kind, command, arguments)
        return
    except Exception:
        logging.exception("XML 작업 등록 실패, SCHTASKS 등록 시도: %s", task_name)

    create_task_via_schtasks(task_name, trigger_kind, command, arguments)


def create_schedule_task(task_name: str, settings: dict, command: str, arguments: str) -> None:
    if not build_schedule_trigger_specs(settings):
        logging.info("등록할 자동 경계 시간이 없어 정시 작업은 만들지 않습니다: %s", task_name)
        return

    try:
        create_schedule_task_via_com(task_name, settings, command, arguments)
        return
    except Exception:
        logging.exception("COM 일정 작업 등록 실패, XML 등록 시도: %s", task_name)

    try:
        create_schedule_task_via_xml(task_name, settings, command, arguments)
        return
    except Exception as exc:
        logging.exception("XML 일정 작업 등록 실패: %s", task_name)
        raise RuntimeError(f"자동 경계 시간 작업 등록에 실패했습니다: {exc}") from exc


def sync_tasks(settings: dict) -> None:
    for task_name in TASKS_TO_DELETE:
        delete_task(task_name)

    if not settings.get("automation_enabled"):
        return

    command, arguments = build_task_runner_action(["--reconcile"])
    create_schedule_task(TASK_SCHEDULE, settings, command, arguments)
    create_task(TASK_LOGON, "logon", command, arguments)
    create_task(TASK_UNLOCK, "unlock", command, arguments)
    create_task(TASK_CONSOLE, "console", command, arguments)


def parse_list_fields(output: str) -> dict[str, str]:
    fields: dict[str, str] = {}
    for line in output.splitlines():
        if ":" not in line:
            continue
        key, value = line.split(":", 1)
        fields[key.strip()] = value.strip()
    return fields


def get_field(fields: dict[str, str], *names: str) -> str:
    for name in names:
        if name in fields:
            return fields[name]
    return ""


def parse_task_xml(task_name: str) -> ET.Element | None:
    result = run_command(["schtasks", "/Query", "/TN", task_name, "/XML"], check=False)
    if result.returncode != 0 or not result.stdout.strip():
        return None
    try:
        return ET.fromstring(result.stdout)
    except ET.ParseError:
        logging.exception("작업 XML 파싱 실패: %s", task_name)
        return None


def inspect_task(task_name: str) -> dict:
    result = run_command(["schtasks", "/Query", "/TN", task_name, "/V", "/FO", "LIST"], check=False)
    if result.returncode != 0:
        return {"exists": False, "task_name": task_name}

    fields = parse_list_fields(result.stdout)
    xml_root = parse_task_xml(task_name)
    ns = {"ns": "http://schemas.microsoft.com/windows/2004/02/mit/task"}

    command = ""
    arguments = ""
    disallow_battery = None
    stop_on_battery = None
    if xml_root is not None:
        command = xml_root.findtext(".//ns:Exec/ns:Command", "", ns).strip()
        arguments = xml_root.findtext(".//ns:Exec/ns:Arguments", "", ns).strip()
        disallow_text = xml_root.findtext(".//ns:DisallowStartIfOnBatteries", "", ns).strip().lower()
        stop_text = xml_root.findtext(".//ns:StopIfGoingOnBatteries", "", ns).strip().lower()
        disallow_battery = disallow_text == "true" if disallow_text else None
        stop_on_battery = stop_text == "true" if stop_text else None

    action_text = get_field(fields, "실행할 작업", "Task To Run") or " ".join(
        part for part in [command, arguments] if part
    ).strip()

    command_name = Path(command).name if command else ""
    if command_name.lower() == "networkroutine.exe":
        target_type = "배포 exe"
    elif command_name.lower() == "pythonw.exe":
        target_type = "개발 python"
    elif command_name:
        target_type = command_name
    else:
        target_type = "-"

    if disallow_battery is False and stop_on_battery is False:
        battery_status = "허용"
    elif disallow_battery is None and stop_on_battery is None:
        battery_status = "미확인"
    else:
        battery_status = "제한"

    return {
        "exists": True,
        "task_name": task_name,
        "state": get_field(fields, "상태", "Status"),
        "next_run": get_field(fields, "다음 실행 시간", "Next Run Time"),
        "last_run": get_field(fields, "마지막 실행 시간", "Last Run Time"),
        "last_result": get_field(fields, "마지막 결과", "Last Result"),
        "action_text": action_text,
        "target_type": target_type,
        "battery_status": battery_status,
    }


class NetworkRoutineApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("900x540")
        self.root.minsize(840, 500)
        self.refresh_job: str | None = None
        self.startup_reconcile_job: str | None = None

        self.settings = load_settings()
        self.adapter_choices = list_adapters()

        self.adapter_var = tk.StringVar(value=self.settings["adapter_name"])
        self.automation_var = tk.BooleanVar(value=self.settings["automation_enabled"])
        self.base_start_var = tk.StringVar(value=self.settings["schedule"]["base_start"])
        self.base_end_var = tk.StringVar(value=self.settings["schedule"]["base_end"])

        internal = self.settings["internal"]
        self.ip_var = tk.StringVar(value=internal["ip"])
        self.mask_var = tk.StringVar(value=internal["mask"])
        self.gateway_var = tk.StringVar(value=internal["gateway"])
        self.dns1_var = tk.StringVar(value=internal["dns1"])
        self.dns2_var = tk.StringVar(value=internal["dns2"])

        self.day_enabled_vars: dict[str, tk.BooleanVar] = {}
        self.day_start_vars: dict[str, tk.StringVar] = {}
        self.day_end_vars: dict[str, tk.StringVar] = {}
        for day in DAY_KEYS:
            day_config = self.settings["schedule"]["days"][day]
            self.day_enabled_vars[day] = tk.BooleanVar(value=day_config["enabled"])
            self.day_start_vars[day] = tk.StringVar(value=day_config["start"])
            self.day_end_vars[day] = tk.StringVar(value=day_config["end"])

        self.status_var = tk.StringVar()
        self.runtime_var = tk.StringVar()
        self.task_var = tk.StringVar()

        self.build_ui()
        self.refresh_runtime_labels()
        self.schedule_startup_reconcile()
        self.schedule_refresh()

    def build_ui(self) -> None:
        outer = ttk.Frame(self.root, padding=12)
        outer.pack(fill="both", expand=True)

        top_row = ttk.Frame(outer)
        top_row.pack(fill="x", pady=(0, 8))
        top_row.columnconfigure(0, weight=3)
        top_row.columnconfigure(1, weight=4)

        top = ttk.LabelFrame(top_row, text="실행")
        top.grid(row=0, column=0, sticky="nsew", padx=(0, 8))

        net_frame = ttk.LabelFrame(top_row, text="내부망")
        net_frame.grid(row=0, column=1, sticky="nsew")

        ttk.Label(top, text="어댑터").grid(row=0, column=0, padx=6, pady=6, sticky="w")
        self.adapter_combo = ttk.Combobox(
            top,
            textvariable=self.adapter_var,
            state="readonly",
            values=self.all_adapter_values(),
            width=22,
        )
        self.adapter_combo.grid(row=0, column=1, padx=6, pady=6, sticky="ew")
        ttk.Button(top, text="새로고침", command=self.refresh_adapters).grid(
            row=0, column=2, padx=6, pady=6, sticky="ew"
        )
        top.columnconfigure(1, weight=1)

        ttk.Checkbutton(top, text="자동 루틴 사용", variable=self.automation_var).grid(
            row=1, column=0, columnspan=2, padx=6, pady=6, sticky="w"
        )

        button_row = ttk.Frame(top)
        button_row.grid(row=2, column=0, columnspan=3, sticky="w", padx=6, pady=(0, 6))
        ttk.Button(button_row, text="지금 내부망", command=lambda: self.manual_apply("internal")).pack(
            side="left", padx=(0, 6)
        )
        ttk.Button(button_row, text="지금 외부망", command=lambda: self.manual_apply("external")).pack(
            side="left"
        )

        ttk.Label(
            top,
            text="자동 루틴은 출근/퇴근 경계 시각과 로그인·복귀 시점에만 자동 처리합니다.",
        ).grid(row=3, column=0, columnspan=3, padx=6, pady=(0, 6), sticky="w")

        compact_fields = [
            ("IP", self.ip_var, 0, 0),
            ("마스크", self.mask_var, 0, 2),
            ("게이트웨이", self.gateway_var, 1, 0),
            ("DNS1", self.dns1_var, 1, 2),
            ("DNS2", self.dns2_var, 2, 0),
        ]
        for label, variable, row, column in compact_fields:
            ttk.Label(net_frame, text=label).grid(row=row, column=column, padx=6, pady=6, sticky="w")
            ttk.Entry(net_frame, textvariable=variable, width=18).grid(
                row=row, column=column + 1, padx=(0, 8), pady=6, sticky="ew"
            )
        ttk.Label(net_frame, text="비워두면 해당 값은 변경하지 않습니다.").grid(
            row=2, column=2, columnspan=2, padx=6, pady=6, sticky="w"
        )
        net_frame.columnconfigure(1, weight=1)
        net_frame.columnconfigure(3, weight=1)

        schedule = ttk.LabelFrame(outer, text="근무 시간")
        schedule.pack(fill="x", pady=(0, 8))

        ttk.Label(schedule, text="기본 출근").grid(row=0, column=0, padx=6, pady=6, sticky="w")
        ttk.Entry(schedule, textvariable=self.base_start_var, width=10).grid(
            row=0, column=1, padx=6, pady=6, sticky="w"
        )
        ttk.Label(schedule, text="기본 퇴근").grid(row=0, column=2, padx=6, pady=6, sticky="w")
        ttk.Entry(schedule, textvariable=self.base_end_var, width=10).grid(
            row=0, column=3, padx=6, pady=6, sticky="w"
        )
        ttk.Button(schedule, text="선택 요일에 시간 적용", command=self.apply_base_schedule).grid(
            row=0, column=4, padx=6, pady=6, sticky="e"
        )

        ttk.Label(schedule, text="요일별 개별 수정").grid(
            row=1, column=0, columnspan=5, padx=6, pady=(0, 2), sticky="w"
        )

        days_grid = ttk.Frame(schedule)
        days_grid.grid(row=2, column=0, columnspan=5, padx=6, pady=(0, 6), sticky="ew")
        days_grid.columnconfigure(0, weight=1)
        days_grid.columnconfigure(1, weight=1)

        for index, day in enumerate(DAY_KEYS):
            cell = ttk.Frame(days_grid)
            cell.grid(
                row=index // 2,
                column=index % 2,
                padx=(0, 10) if index % 2 == 0 else (0, 0),
                pady=3,
                sticky="ew",
            )
            ttk.Checkbutton(cell, text=DAY_LABELS[day], variable=self.day_enabled_vars[day]).grid(
                row=0, column=0, padx=(0, 4), sticky="w"
            )
            ttk.Entry(cell, textvariable=self.day_start_vars[day], width=7).grid(
                row=0, column=1, padx=(0, 4), sticky="w"
            )
            ttk.Label(cell, text="~").grid(row=0, column=2, padx=(0, 4), sticky="w")
            ttk.Entry(cell, textvariable=self.day_end_vars[day], width=7).grid(
                row=0, column=3, sticky="w"
            )

        bottom = ttk.Frame(outer)
        bottom.pack(fill="x", pady=(0, 8))
        ttk.Label(bottom, text="수동 전환하면 현재 시간대에는 자동이 다시 덮어쓰지 않습니다.").pack(
            side="left"
        )
        ttk.Button(bottom, text="창 닫기", command=self.root.destroy).pack(side="right")
        ttk.Button(bottom, text="저장 및 반영", command=self.save_and_apply).pack(side="right", padx=(0, 6))

        status = ttk.LabelFrame(outer, text="상태")
        status.pack(fill="x")

        ttk.Label(status, textvariable=self.runtime_var, wraplength=840, justify="left").pack(
            fill="x", padx=8, pady=(8, 4)
        )
        ttk.Label(status, textvariable=self.task_var, wraplength=840, justify="left").pack(
            fill="x", padx=8, pady=(0, 4)
        )
        ttk.Label(status, textvariable=self.status_var, wraplength=840, justify="left").pack(
            fill="x", padx=8, pady=(0, 8)
        )

    def schedule_startup_reconcile(self) -> None:
        if self.startup_reconcile_job:
            self.root.after_cancel(self.startup_reconcile_job)
        self.startup_reconcile_job = self.root.after(700, self._run_startup_reconcile)

    def _run_startup_reconcile(self) -> None:
        self.startup_reconcile_job = None
        self.run_background_reconcile_if_enabled("시작 시 확인", update_status_on_noop=True)

    def run_background_reconcile_if_enabled(self, source: str, update_status_on_noop: bool = False) -> None:
        try:
            saved = load_settings()
            validate_settings(saved)
            self.settings = saved

            if not saved.get("automation_enabled"):
                self.refresh_runtime_labels()
                return

            message = reconcile_now(saved)
            self.settings = load_settings()
            self.refresh_runtime_labels()

            is_noop = message.startswith("현재 시간 기준 유지:")
            if update_status_on_noop or not is_noop:
                self.set_status(f"{source}: {message}")
        except Exception as exc:
            logging.exception("%s 실패", source)
            self.set_status(f"{source} 실패: {exc}")

    def all_adapter_values(self) -> list[str]:
        values = list(self.adapter_choices)
        current = self.adapter_var.get().strip()
        if current and current not in values:
            values.append(current)
        return values

    def refresh_adapters(self) -> None:
        self.adapter_choices = list_adapters()
        self.adapter_combo.configure(values=self.all_adapter_values())
        self.set_status("어댑터 목록을 새로 읽었습니다.")

    def apply_base_schedule(self) -> None:
        try:
            start = self.base_start_var.get().strip()
            end = self.base_end_var.get().strip()
            parse_time(start)
            parse_time(end)

            for day in DAY_KEYS:
                if self.day_enabled_vars[day].get():
                    self.day_start_vars[day].set(start)
                    self.day_end_vars[day].set(end)

            self.refresh_runtime_labels()
            self.set_status("선택된 요일에 기본 시간을 반영했습니다.")
        except Exception as exc:
            self.show_error(exc)

    def collect_form_settings(self) -> dict:
        settings = load_settings()
        settings["adapter_name"] = self.adapter_var.get().strip()
        settings["automation_enabled"] = self.automation_var.get()
        settings["schedule"]["base_start"] = self.base_start_var.get().strip()
        settings["schedule"]["base_end"] = self.base_end_var.get().strip()
        settings["internal"] = {
            "ip": self.ip_var.get().strip(),
            "mask": self.mask_var.get().strip(),
            "gateway": self.gateway_var.get().strip(),
            "dns1": self.dns1_var.get().strip(),
            "dns2": self.dns2_var.get().strip(),
        }
        for day in DAY_KEYS:
            settings["schedule"]["days"][day] = {
                "enabled": self.day_enabled_vars[day].get(),
                "start": self.day_start_vars[day].get().strip(),
                "end": self.day_end_vars[day].get().strip(),
            }
        validate_settings(settings)
        return settings

    def persist_settings(self, sync_schedule_enabled: bool = True) -> dict:
        previous = load_settings()
        updated = self.collect_form_settings()
        state_related_changed = (
            previous["adapter_name"] != updated["adapter_name"]
            or previous["internal"] != updated["internal"]
            or previous["schedule"] != updated["schedule"]
        )
        if state_related_changed:
            updated["last_applied_mode"] = ""
            updated["last_handled_segment_id"] = ""
            updated["last_handled_mode"] = ""
            updated["last_handled_at"] = ""
            updated["manual_override_segment_id"] = ""
            updated["manual_override_mode"] = ""
            updated["manual_override_at"] = ""

        save_settings(updated)
        self.settings = updated
        if sync_schedule_enabled:
            sync_tasks(updated)
        return updated

    def manual_apply(self, mode: str) -> None:
        try:
            settings = self.persist_settings(sync_schedule_enabled=True)
            message = apply_mode(mode, settings, reason="수동 전환")
            mark_manual_override(settings, mode)
            self.settings = load_settings()
            self.refresh_runtime_labels()
            self.set_status(message)
        except Exception as exc:
            self.show_error(exc)

    def save_and_apply(self) -> None:
        try:
            settings = self.persist_settings(sync_schedule_enabled=True)
            if settings["automation_enabled"]:
                message = reconcile_now(settings)
            else:
                message = "설정을 저장했고 자동 루틴 작업은 제거했습니다."
            self.settings = load_settings()
            self.refresh_runtime_labels()
            self.set_status(message)
            messagebox.showinfo(APP_TITLE, message)
        except Exception as exc:
            self.show_error(exc)

    def current_runtime_text(self) -> str:
        try:
            preview = self.collect_form_settings()
            segment_info = current_segment_info(preview)
            current_target = segment_info["mode"]
            current_label = mode_label(current_target)
            summary = schedule_summary(preview)
            if preview.get("manual_override_segment_id") == segment_info["segment_id"]:
                control_text = f"수동 우선 {mode_label(preview.get('manual_override_mode') or current_target)}"
            elif preview.get("last_handled_segment_id") == segment_info["segment_id"]:
                control_text = f"현재 시간대 자동 처리 완료 {current_label}"
            else:
                control_text = f"현재 시간대 자동 처리 대기 {current_label}"
            return (
                f"예상 모드: {current_label} | "
                f"자동 루틴: {'켜짐' if preview['automation_enabled'] else '꺼짐'} | "
                f"어댑터: {preview['adapter_name'] or '-'} | "
                f"시간표: {summary} | "
                f"제어: {control_text}"
            )
        except Exception as exc:
            return f"입력값 확인 필요: {exc}"

    def refresh_runtime_labels(self) -> None:
        self.runtime_var.set(self.current_runtime_text())
        self.task_var.set(self.current_task_text())
        last_mode = self.settings.get("last_applied_mode", "")
        last_mode_label = (
            "내부망" if last_mode == "internal" else "외부망" if last_mode == "external" else "-"
        )
        self.status_var.set(
            f"마지막 적용: {last_mode_label} | "
            f"시각: {self.settings.get('last_applied_at') or '-'} | "
            f"메시지: {self.settings.get('last_message') or '-'}"
        )

    def schedule_refresh(self) -> None:
        if self.refresh_job:
            self.root.after_cancel(self.refresh_job)
        self.refresh_job = self.root.after(30000, self._scheduled_refresh)

    def _scheduled_refresh(self) -> None:
        self.refresh_job = None
        self.refresh_runtime_labels()
        self.schedule_refresh()

    def set_status(self, text: str) -> None:
        self.settings = load_settings()
        self.task_var.set(self.current_task_text())
        self.status_var.set(
            f"마지막 적용: "
            f"{'내부망' if self.settings.get('last_applied_mode') == 'internal' else '외부망' if self.settings.get('last_applied_mode') == 'external' else '-'} | "
            f"시각: {self.settings.get('last_applied_at') or '-'} | "
            f"메시지: {text}"
        )

    def current_task_text(self) -> str:
        schedule = inspect_task(TASK_SCHEDULE)
        logon = inspect_task(TASK_LOGON)
        unlock = inspect_task(TASK_UNLOCK)
        console = inspect_task(TASK_CONSOLE)

        def format_task(prefix: str, info: dict) -> str:
            if not info["exists"]:
                return f"{prefix}: 미등록"
            state = info.get("state") or "-"
            result = info.get("last_result") or "-"
            battery = info.get("battery_status") or "-"
            target = info.get("target_type") or "-"
            next_run = info.get("next_run") or "-"
            return (
                f"{prefix}: 등록됨/{state}/결과 {result}/대상 {target}/배터리 {battery}/다음 {next_run}"
            )

        def format_resume_task() -> str:
            infos = [("잠금해제", unlock), ("콘솔", console)]
            existing = [(label, info) for label, info in infos if info["exists"]]
            if not existing:
                return "복귀: 미등록"

            states = ",".join(f"{label} {info.get('state') or '-'}" for label, info in existing)
            results = ",".join(f"{label} {info.get('last_result') or '-'}" for label, info in existing)
            targets = {info.get("target_type") or "-" for _, info in existing}
            batteries = {info.get("battery_status") or "-" for _, info in existing}
            target_text = next(iter(targets)) if len(targets) == 1 else "혼합"
            battery_text = next(iter(batteries)) if len(batteries) == 1 else "혼합"
            return (
                f"복귀: {len(existing)}/{len(infos)} 등록/상태 {states}/결과 {results}/"
                f"대상 {target_text}/배터리 {battery_text}"
            )

        return (
            f"작업 상태 | {format_task('정시', schedule)} | "
            f"{format_task('로그온', logon)} | {format_resume_task()}"
        )

    def show_error(self, exc: Exception) -> None:
        logging.exception("오류 발생")
        messagebox.showerror(APP_TITLE, str(exc))
        self.set_status(f"오류: {exc}")


def run_reconcile_cli() -> int:
    try:
        ensure_admin()
        settings = load_settings()
        validate_settings(settings)
        logging.info(reconcile_now(settings))
        return 0
    except Exception:
        logging.exception("자동 루틴 실행 실패")
        return 1


def main() -> int:
    configure_logging()

    parser = argparse.ArgumentParser()
    parser.add_argument("--reconcile", action="store_true", help="현재 시간 기준으로 자동 전환")
    args = parser.parse_args()

    if args.reconcile:
        return run_reconcile_cli()

    ensure_admin()

    root = tk.Tk()
    style = ttk.Style(root)
    if "vista" in style.theme_names():
        style.theme_use("vista")
    app = NetworkRoutineApp(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
