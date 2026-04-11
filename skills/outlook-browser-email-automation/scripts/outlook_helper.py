import argparse
import base64
import json
import os
import re
import shutil
import subprocess
import sys
import time
import urllib.parse
from contextlib import suppress
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any
from zoneinfo import ZoneInfo

import requests

PLAYWRIGHT_COMMANDS = {"recent-mail", "ensure-calendar", "web-login", "ensure-session"}
PLAYWRIGHT_REEXEC_ENV = "MAIL_AUTOMATION_PLAYWRIGHT_REEXEC"
PLAYWRIGHT_FALLBACK_PYTHONS = ("3.12", "3.13", "3.11")

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8")


WINDOWS_ENV_VAR_RE = re.compile(r"%([^%]+)%")


def expand_path(value: str | os.PathLike[str]) -> Path:
    raw = os.path.expanduser(os.path.expandvars(str(value)))
    raw = WINDOWS_ENV_VAR_RE.sub(lambda match: os.environ.get(match.group(1), match.group(0)), raw)
    return Path(raw)


def default_state_root() -> Path:
    return expand_path(os.getenv("LOCALAPPDATA", str(Path.home() / ".codex")))


def default_edge_user_data_dir() -> Path:
    if os.name == "nt":
        return expand_path(os.getenv("LOCALAPPDATA", "")) / "Microsoft" / "Edge" / "User Data"
    return Path.home() / ".config" / "microsoft-edge"


def command_name(argv: list[str]) -> str:
    for arg in argv[1:]:
        if not arg.startswith("-"):
            return arg
    return ""


def can_use_python_launcher(version: str) -> bool:
    py = shutil.which("py")
    if not py:
        return False
    probe = [
        py,
        f"-{version}",
        "-c",
        "import playwright; from zoneinfo import ZoneInfo; ZoneInfo('Europe/London')",
    ]
    result = subprocess.run(
        probe,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        check=False,
    )
    return result.returncode == 0


def maybe_reexec_for_playwright():
    if os.name != "nt":
        return
    if os.environ.get(PLAYWRIGHT_REEXEC_ENV):
        return
    if sys.version_info < (3, 14):
        return
    if command_name(sys.argv) not in PLAYWRIGHT_COMMANDS:
        return
    py = shutil.which("py")
    if not py:
        return
    version = next((item for item in PLAYWRIGHT_FALLBACK_PYTHONS if can_use_python_launcher(item)), "")
    if not version:
        return
    env = os.environ.copy()
    env[PLAYWRIGHT_REEXEC_ENV] = "1"
    result = subprocess.run(
        [py, f"-{version}", str(Path(__file__).resolve()), *sys.argv[1:]],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        env=env,
        check=False,
    )
    if result.stdout:
        sys.stdout.write(result.stdout)
    if result.stderr:
        sys.stderr.write(result.stderr)
    raise SystemExit(result.returncode)


maybe_reexec_for_playwright()

from playwright.sync_api import Error as PlaywrightError
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright


OUTLOOK_WEB_URL = "https://outlook.office.com/mail/"
OUTLOOK_REST_BASE = "https://outlook.office.com/api/v2.0"
DEFAULT_LOCAL_TIME_ZONE = "Europe/London"
DEFAULT_OUTLOOK_TIME_ZONE = "GMT Standard Time"
DEFAULT_TOKEN_REFRESH_WINDOW_SECONDS = 3600
CONFIG_ENV = "MAIL_AUTOMATION_CONFIG"
CONFIG_PATH = expand_path(os.getenv(CONFIG_ENV, str(Path(__file__).with_name("mail_automation_web_config.json"))))
STATE_DIR = default_state_root() / "codex-mail-automation-web"
PROFILE_DIR = STATE_DIR / "edge-profile"
STORAGE_STATE_PATH = STATE_DIR / "storage_state.json"
MAIL_PREVIEW_CHARS = 320
COMPACT_JSON_KWARGS = {"ensure_ascii": False, "separators": (",", ":")}
MAIL_LIST_SELECTORS = (
    "[role='main'] div[data-convid]",
    "div[data-convid]",
    "[data-testid='virtuoso-scroller'] div[data-convid]",
    "[role='option'][data-convid]",
    "[role='option']",
)
MAILBOX_READY_SELECTORS = (
    "[role='main'] [role='listbox']",
    "[role='main'] [role='grid']",
    "div[data-app-section='MailModule']",
    "div[data-convid]",
    "[data-testid='virtuoso-scroller']",
)
SIGNIN_HINT_SELECTORS = (
    "input[type='email']",
    "input[name='loginfmt']",
    "input[type='password']",
    "form[action*='login.microsoftonline.com']",
)
DEBUG_DIR = Path(__file__).with_name("output") / "playwright"
HTTP_SESSION = requests.Session()
HTTP_SESSION.trust_env = False
TOKEN_REFRESH_ERROR_MARKERS = (
    "Missing Outlook Web storage state",
    "No valid Outlook Web API token found",
    "Stored Outlook Web API token is missing or expired",
)
LAST_AUTO_REFRESH_FAILURE = ""


if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8")


def load_config() -> dict[str, Any]:
    defaults = {
        "outlook_url": OUTLOOK_WEB_URL,
        "edge_channel": "msedge",
        "local_time_zone": DEFAULT_LOCAL_TIME_ZONE,
        "outlook_time_zone": DEFAULT_OUTLOOK_TIME_ZONE,
        "token_refresh_window_seconds": DEFAULT_TOKEN_REFRESH_WINDOW_SECONDS,
        "headless": True,
        "browser_mode": "system-edge-default",
        "edge_user_data_dir": str(default_edge_user_data_dir()),
        "edge_profile_directory": "Default",
    }
    if not CONFIG_PATH.exists():
        return defaults
    raw = json.loads(CONFIG_PATH.read_text(encoding="utf-8-sig"))
    config = {**defaults, **raw}
    for key in ("edge_user_data_dir",):
        if key in config:
            config[key] = str(expand_path(str(config[key])))
    try:
        ZoneInfo(str(config["local_time_zone"]))
    except Exception as exc:
        raise RuntimeError(f"Unsupported local_time_zone in {CONFIG_PATH}: {config['local_time_zone']}") from exc
    try:
        config["token_refresh_window_seconds"] = max(0, int(config.get("token_refresh_window_seconds", DEFAULT_TOKEN_REFRESH_WINDOW_SECONDS)))
    except Exception as exc:
        raise RuntimeError(
            f"Unsupported token_refresh_window_seconds in {CONFIG_PATH}: {config.get('token_refresh_window_seconds')}"
        ) from exc
    return config


def ensure_state_dir():
    STATE_DIR.mkdir(parents=True, exist_ok=True)
    DEBUG_DIR.mkdir(parents=True, exist_ok=True)


def local_zone(config: dict[str, Any]) -> ZoneInfo:
    return ZoneInfo(str(config["local_time_zone"]))


def storage_state_data() -> dict[str, Any]:
    if not STORAGE_STATE_PATH.exists():
        raise RuntimeError(
            f"Missing Outlook Web storage state: {STORAGE_STATE_PATH}. "
            "Run `python outlook_helper.py web-login` once to refresh login state."
        )
    return json.loads(STORAGE_STATE_PATH.read_text(encoding="utf-8-sig"))


def decode_jwt_payload(token: str) -> dict[str, Any]:
    parts = token.split(".")
    if len(parts) != 3:
        raise RuntimeError("Stored Outlook token is not a valid JWT")
    payload = parts[1]
    payload += "=" * (-len(payload) % 4)
    return json.loads(base64.urlsafe_b64decode(payload))


def find_access_token(storage_state: dict[str, Any], audience: str) -> str:
    now_ts = int(time.time())
    best_token = ""
    best_exp = 0
    audience_prefix = audience.rstrip("/") + "/"
    audience_value = audience.rstrip("/")
    for origin in storage_state.get("origins") or []:
        for item in origin.get("localStorage") or []:
            name = str(item.get("name") or "")
            if "|accesstoken|" not in name or audience_prefix not in name:
                continue
            value = item.get("value")
            if not value:
                continue
            with suppress(Exception):
                token = json.loads(value)["secret"]
                payload = decode_jwt_payload(token)
                exp = int(payload.get("exp") or 0)
                if payload.get("aud") != audience_value:
                    continue
                if exp > now_ts + 120 and exp > best_exp:
                    best_token = token
                    best_exp = exp
    if best_token:
        return best_token
    raise RuntimeError(
        "No valid Outlook Web API token found in storage_state.json. "
        "Run `python outlook_helper.py web-login` to refresh login state."
    )


def outlook_api_headers(storage_state: dict[str, Any]) -> dict[str, str]:
    token = find_access_token(storage_state, "https://outlook.office.com/")
    return {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
        "Content-Type": "application/json",
    }


def outlook_token_status(storage_state: dict[str, Any] | None = None) -> dict[str, Any]:
    payload: dict[str, Any] = {"present": False}
    try:
        current_state = storage_state or storage_state_data()
        token = find_access_token(current_state, "https://outlook.office.com/")
        claims = decode_jwt_payload(token)
        exp = int(claims.get("exp") or 0)
        expires_at = datetime.fromtimestamp(exp, tz=timezone.utc)
        payload.update(
            {
                "present": True,
                "audience": str(claims.get("aud") or ""),
                "expires_at_utc": expires_at.isoformat(),
                "expires_at_local": expires_at.astimezone(local_zone(load_config())).isoformat(),
                "seconds_remaining": max(0, exp - int(time.time())),
            }
        )
    except RuntimeError as exc:
        payload["error"] = str(exc)
    return payload


def browser_context_from_storage_state(playwright, config: dict[str, Any], *, headed: bool):
    browser = playwright.chromium.launch(
        channel=str(config["edge_channel"]),
        headless=not headed,
    )
    storage_state = str(STORAGE_STATE_PATH) if STORAGE_STATE_PATH.exists() else None
    context = browser.new_context(
        storage_state=storage_state,
        viewport={"width": 1600, "height": 1200},
    )
    return browser, context


def should_attempt_automatic_refresh(exc: RuntimeError) -> bool:
    return any(marker in str(exc) for marker in TOKEN_REFRESH_ERROR_MARKERS)


def set_auto_refresh_failure(reason: str):
    global LAST_AUTO_REFRESH_FAILURE
    LAST_AUTO_REFRESH_FAILURE = reason


def apply_storage_state_to_context(context, storage_state: dict[str, Any], config: dict[str, Any]):
    cookies = storage_state.get("cookies") or []
    if cookies:
        context.add_cookies(cookies)
    page = first_page(context)
    seen_origins: set[str] = set()
    for origin in storage_state.get("origins") or []:
        origin_url = str(origin.get("origin") or "").strip()
        local_storage = origin.get("localStorage") or []
        if not origin_url or not local_storage or origin_url in seen_origins:
            continue
        seen_origins.add(origin_url)
        page.goto(origin_url, wait_until="domcontentloaded", timeout=120000)
        page.evaluate(
            """items => {
                for (const item of items) {
                    localStorage.setItem(String(item.name ?? ""), String(item.value ?? ""));
                }
            }""",
            local_storage,
        )
    goto_outlook_mail(page, config)


def seed_dedicated_profile_from_storage_state(playwright, config: dict[str, Any], storage_state: dict[str, Any]) -> bool:
    seed_config = {**config, "browser_mode": "dedicated-profile", "headless": True}
    context = None
    try:
        context = persistent_context(playwright, seed_config, headed=False)
        apply_storage_state_to_context(context, storage_state, config)
        save_storage_state(context)
        return True
    except (RuntimeError, PlaywrightError, PermissionError):
        return False
    finally:
        with suppress(Exception):
            if context is not None:
                context.close()


def preferred_login_config(config: dict[str, Any]) -> dict[str, Any]:
    if str(config.get("browser_mode")) == "system-edge-default" and (PROFILE_DIR.exists() or STORAGE_STATE_PATH.exists()):
        return {**config, "browser_mode": "dedicated-profile", "headless": False}
    return config


def ensure_fresh_storage_state(
    *,
    allow_auto_refresh: bool = True,
    min_valid_seconds: int | None = None,
    refresh_timeout_seconds: int = 90,
) -> dict[str, Any]:
    config = load_config()
    refresh_window_seconds = (
        config["token_refresh_window_seconds"] if min_valid_seconds is None else max(0, int(min_valid_seconds))
    )
    try:
        current_state = storage_state_data()
    except RuntimeError as exc:
        if allow_auto_refresh and should_attempt_automatic_refresh(exc):
            if try_refresh_storage_state_automatically(timeout_seconds=refresh_timeout_seconds):
                return storage_state_data()
            if LAST_AUTO_REFRESH_FAILURE.startswith("interactive_auth_required:"):
                stage = LAST_AUTO_REFRESH_FAILURE.split(":", 1)[1]
                raise RuntimeError(
                    "Stored Outlook Web API token is missing or expired, and automatic refresh requires "
                    f"interactive sign-in ({stage}). Run `python outlook_helper.py web-login` to refresh login state."
                ) from exc
        raise
    if not allow_auto_refresh or refresh_window_seconds <= 0:
        return current_state
    token_status = outlook_token_status(current_state)
    seconds_remaining = int(token_status.get("seconds_remaining") or 0) if token_status.get("present") else 0
    if seconds_remaining > refresh_window_seconds:
        return current_state
    if try_refresh_storage_state_automatically(timeout_seconds=refresh_timeout_seconds):
        return storage_state_data()
    if token_status.get("present") and seconds_remaining > 0:
        return current_state
    if LAST_AUTO_REFRESH_FAILURE.startswith("interactive_auth_required:"):
        stage = LAST_AUTO_REFRESH_FAILURE.split(":", 1)[1]
        raise RuntimeError(
            "Stored Outlook Web API token is missing or expired, and automatic refresh requires "
            f"interactive sign-in ({stage}). Run `python outlook_helper.py web-login` to refresh login state."
        )
    return current_state


def try_refresh_storage_state_automatically(*, timeout_seconds: int = 90) -> bool:
    set_auto_refresh_failure("")
    config = load_config()
    ensure_state_dir()
    with sync_playwright() as playwright:
        attempts: list[tuple[str, dict[str, Any]]] = []
        if PROFILE_DIR.exists():
            attempts.append(("dedicated-profile", {**config, "browser_mode": "dedicated-profile", "headless": True}))
        if STORAGE_STATE_PATH.exists():
            attempts.append(("storage-state", config))
        if not attempts:
            set_auto_refresh_failure("no_saved_session")
            return False
        first_failure = ""
        interactive_failure = ""
        for strategy, strategy_config in attempts:
            browser = None
            context = None
            try:
                if strategy == "dedicated-profile":
                    context = persistent_context(playwright, strategy_config, headed=False)
                else:
                    browser, context = browser_context_from_storage_state(playwright, strategy_config, headed=False)
                page = first_page(context)
                goto_outlook_mail(page, config)
                if not wait_for_mailbox_for_automatic_refresh(page, timeout_seconds=timeout_seconds):
                    stage = detect_login_stage(page) or "unknown"
                    interactive_failure = f"interactive_auth_required:{stage}"
                    continue
                if sign_in_required(page):
                    stage = detect_login_stage(page) or "sign-in"
                    interactive_failure = f"interactive_auth_required:{stage}"
                    continue
                save_storage_state(context)
                current_state = context.storage_state()
                if strategy != "dedicated-profile":
                    seed_dedicated_profile_from_storage_state(playwright, config, current_state)
                return True
            except (RuntimeError, PlaywrightError, PermissionError) as exc:
                first_failure = first_failure or f"{strategy}:{type(exc).__name__}"
            finally:
                with suppress(Exception):
                    if context is not None:
                        context.close()
                with suppress(Exception):
                    if browser is not None:
                        browser.close()
        set_auto_refresh_failure(interactive_failure or first_failure or "refresh_failed")
        return False


def outlook_api_headers_with_auto_refresh(*, allow_auto_refresh: bool = True) -> dict[str, str]:
    try:
        return outlook_api_headers(ensure_fresh_storage_state(allow_auto_refresh=allow_auto_refresh))
    except RuntimeError as exc:
        if allow_auto_refresh and should_attempt_automatic_refresh(exc):
            if try_refresh_storage_state_automatically():
                return outlook_api_headers(storage_state_data())
            if LAST_AUTO_REFRESH_FAILURE.startswith("interactive_auth_required:"):
                stage = LAST_AUTO_REFRESH_FAILURE.split(":", 1)[1]
                raise RuntimeError(
                    "Stored Outlook Web API token is missing or expired, and automatic refresh requires "
                    f"interactive sign-in ({stage}). Run `python outlook_helper.py web-login` to refresh login state."
                ) from exc
        raise


def request_json(method: str, url: str, *, headers: dict[str, str], params: dict[str, Any] | None = None, payload: dict[str, Any] | None = None, allow_auto_refresh: bool = False) -> Any:
    response = HTTP_SESSION.request(method, url, headers=headers, params=params, json=payload, timeout=45)
    if response.status_code == 401 and allow_auto_refresh:
        if try_refresh_storage_state_automatically():
            refreshed_headers = outlook_api_headers_with_auto_refresh(allow_auto_refresh=False)
            response = HTTP_SESSION.request(
                method,
                url,
                headers=refreshed_headers,
                params=params,
                json=payload,
                timeout=45,
            )
        elif LAST_AUTO_REFRESH_FAILURE.startswith("interactive_auth_required:"):
            stage = LAST_AUTO_REFRESH_FAILURE.split(":", 1)[1]
            raise RuntimeError(
                "Stored Outlook Web API token is missing or expired, and automatic refresh requires "
                f"interactive sign-in ({stage}). Run `python outlook_helper.py web-login` to refresh login state."
            )
    if response.status_code == 401:
        raise RuntimeError(
            "Stored Outlook Web API token is missing or expired. "
            "Run `python outlook_helper.py web-login` to refresh login state."
        )
    response.raise_for_status()
    if response.status_code == 204 or not response.text:
        return None
    return response.json()


def api_datetime_to_local_text(value: str, zone: ZoneInfo) -> str:
    if not value:
        return ""
    dt = datetime.fromisoformat(value.replace("Z", "+00:00"))
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    return dt.astimezone(zone).isoformat(sep=" ", timespec="seconds")


def local_text_to_utc_iso(local_text: str, zone: ZoneInfo) -> str:
    dt = datetime.strptime(local_text, "%Y-%m-%d %H:%M").replace(tzinfo=zone)
    return dt.astimezone(timezone.utc).isoformat().replace("+00:00", "Z")


def outlook_show_as(value: Any) -> str:
    mapping = {
        0: "Free",
        1: "Tentative",
        2: "Busy",
        3: "Oof",
        4: "WorkingElsewhere",
    }
    if isinstance(value, int):
        return mapping.get(value, "Busy")
    return "Busy"


def build_rest_message(item: dict[str, Any], zone: ZoneInfo) -> dict[str, Any]:
    sender = ((item.get("From") or {}).get("EmailAddress") or {})
    preview = " ".join(str(item.get("BodyPreview") or "").split())
    meeting_type = str(item.get("MeetingMessageType") or "")
    return {
        "received": api_datetime_to_local_text(str(item.get("ReceivedDateTime") or ""), zone),
        "sender": str(sender.get("Name") or ""),
        "sender_email": str(sender.get("Address") or ""),
        "subject": str(item.get("Subject") or ""),
        "topic": normalize_topic(str(item.get("Subject") or "")),
        "message_class": "outlook-rest.message",
        "importance": item.get("Importance"),
        "flag_status": ((item.get("Flag") or {}).get("FlagStatus") if isinstance(item.get("Flag"), dict) else None),
        "is_meeting": bool(meeting_type and meeting_type.lower() != "none"),
        "preview": preview[:MAIL_PREVIEW_CHARS],
        "entry_id": str(item.get("Id") or ""),
        "conversation_id": str(item.get("ConversationId") or ""),
        "web_link": str(item.get("WebLink") or ""),
    }


def fetch_recent_messages_http(hours: int = 24, max_items: int = 250) -> list[dict[str, Any]]:
    max_items = max(0, int(max_items))
    hours = max(0, int(hours))
    if max_items == 0 or hours == 0:
        return []
    config = load_config()
    zone = local_zone(config)
    headers = outlook_api_headers_with_auto_refresh()
    top = max(1, min(max_items, 250))
    params: dict[str, Any] | None = {
        "$top": top,
        "$orderby": "ReceivedDateTime DESC",
    }
    url = f"{OUTLOOK_REST_BASE}/me/mailfolders/inbox/messages"
    cutoff = datetime.now(zone) - timedelta(hours=hours)
    messages = []
    while url and len(messages) < max_items:
        data = request_json("GET", url, headers=headers, params=params, allow_auto_refresh=True)
        reached_cutoff = False
        for item in data.get("value") or []:
            message = build_rest_message(item, zone)
            if not message["subject"] or not message["received"]:
                continue
            received = datetime.fromisoformat(message["received"])
            if received < cutoff:
                reached_cutoff = True
                break
            messages.append(message)
            if len(messages) >= max_items:
                break
        if reached_cutoff:
            break
        url = str(data.get("@odata.nextLink") or "")
        params = None
    return messages


def event_time_range_label_local(start_local: str, end_local: str) -> str:
    start_dt = datetime.strptime(start_local, "%Y-%m-%d %H:%M")
    end_dt = datetime.strptime(end_local, "%Y-%m-%d %H:%M")
    return f"{start_dt.strftime('%H:%M')} 到 {end_dt.strftime('%H:%M')}"


def calendar_view_events(headers: dict[str, str], zone: ZoneInfo, entry: dict[str, Any]) -> list[dict[str, Any]]:
    start_local = str(entry["start_local"])
    local_start = datetime.strptime(start_local, "%Y-%m-%d %H:%M").replace(tzinfo=zone)
    day_start = local_start.replace(hour=0, minute=0, second=0, microsecond=0)
    day_end = day_start + timedelta(days=1)
    params = {
        "startDateTime": day_start.astimezone(timezone.utc).isoformat().replace("+00:00", "Z"),
        "endDateTime": day_end.astimezone(timezone.utc).isoformat().replace("+00:00", "Z"),
        "$orderby": "Start/DateTime",
        "$select": "Id,Subject,Start,End,Type,ShowAs,WebLink,Location,Categories,ReminderMinutesBeforeStart,IsReminderOn",
    }
    data = request_json("GET", f"{OUTLOOK_REST_BASE}/me/calendarview", headers=headers, params=params)
    return data.get("value") or []


def rest_event_to_local_signature(item: dict[str, Any], zone: ZoneInfo) -> dict[str, str]:
    start = item.get("Start") or {}
    end = item.get("End") or {}
    start_local = api_datetime_to_local_text(str(start.get("DateTime") or ""), zone)
    end_local = api_datetime_to_local_text(str(end.get("DateTime") or ""), zone)
    return {
        "id": str(item.get("Id") or ""),
        "subject": str(item.get("Subject") or ""),
        "start_local": start_local[:16],
        "end_local": end_local[:16],
        "time_range": event_time_range_label_local(start_local[:16], end_local[:16]),
        "web_link": str(item.get("WebLink") or ""),
    }


def ensure_calendar_http(spec_path: Path):
    spec = json.loads(spec_path.read_text(encoding="utf-8-sig"))
    if not isinstance(spec, list):
        raise RuntimeError(f"Calendar spec must be a JSON array: {spec_path}")
    config = load_config()
    zone = local_zone(config)
    outlook_time_zone = str(config.get("outlook_time_zone") or DEFAULT_OUTLOOK_TIME_ZONE)
    headers = outlook_api_headers_with_auto_refresh()
    changes = []
    for index, entry in enumerate(spec):
        if not isinstance(entry, dict):
            raise RuntimeError(f"Calendar entry #{index + 1} must be an object")
        subject = str(entry.get("subject") or "").strip()
        if not subject:
            raise RuntimeError(f"Calendar entry #{index + 1} is missing subject")
        for key in ("start_local", "end_local"):
            if not entry.get(key):
                raise RuntimeError(f"Calendar entry #{index + 1} is missing {key}")
        start_dt = datetime.strptime(str(entry["start_local"]), "%Y-%m-%d %H:%M")
        end_dt = datetime.strptime(str(entry["end_local"]), "%Y-%m-%d %H:%M")
        if end_dt <= start_dt:
            raise RuntimeError(f"Calendar entry #{index + 1} end_local must be after start_local")
        existing = [rest_event_to_local_signature(item, zone) for item in calendar_view_events(headers, zone, entry)]
        exact_duplicate = next(
            (
                item
                for item in existing
                if same_subject(item["subject"], subject)
                and item["start_local"] == str(entry["start_local"])
                and item["end_local"] == str(entry["end_local"])
            ),
            None,
        )
        if exact_duplicate:
            changes.append(
                {
                    "subject": subject,
                    "status": "skipped_duplicate",
                    "display_start_local": entry["start_local"],
                    "display_end_local": entry["end_local"],
                    "web_link": exact_duplicate.get("web_link", ""),
                }
            )
            continue
        conflicts = [
            item
            for item in existing
            if same_subject(item["subject"], subject)
            and (item["start_local"] != str(entry["start_local"]) or item["end_local"] != str(entry["end_local"]))
        ]
        payload = {
            "Subject": subject,
            "Start": {"DateTime": str(entry["start_local"]).replace(" ", "T"), "TimeZone": outlook_time_zone},
            "End": {"DateTime": str(entry["end_local"]).replace(" ", "T"), "TimeZone": outlook_time_zone},
            "ShowAs": outlook_show_as(entry.get("busy_status")),
            "IsReminderOn": bool(entry.get("reminder_set", False)),
            "ReminderMinutesBeforeStart": int(entry.get("reminder_minutes") or 0),
        }
        if entry.get("location"):
            payload["Location"] = {"DisplayName": str(entry["location"])}
        if entry.get("body"):
            payload["Body"] = {"ContentType": "Text", "Content": str(entry["body"])}
        categories = entry.get("categories")
        if categories:
            payload["Categories"] = categories if isinstance(categories, list) else [str(categories)]
        created = request_json("POST", f"{OUTLOOK_REST_BASE}/me/events", headers=headers, payload=payload)
        change = {
            "subject": subject,
            "status": "created",
            "display_start_local": entry["start_local"],
            "display_end_local": entry["end_local"],
            "web_link": str(created.get("WebLink") or ""),
        }
        if conflicts:
            change["note"] = "existing_same_subject_different_time_detected"
        changes.append(change)
    return changes


def resolved_user_data_dir(config: dict[str, Any]) -> Path:
    mode = str(config.get("browser_mode") or "dedicated-profile")
    if mode == "system-edge-default":
        return Path(str(config["edge_user_data_dir"]))
    return PROFILE_DIR


def persistent_context(playwright, config: dict[str, Any], *, headed: bool):
    ensure_state_dir()
    user_data_dir = resolved_user_data_dir(config)
    args = []
    if str(config.get("browser_mode")) == "system-edge-default":
        args.append(f"--profile-directory={config['edge_profile_directory']}")
    try:
        return playwright.chromium.launch_persistent_context(
            user_data_dir=str(user_data_dir),
            channel=str(config["edge_channel"]),
            headless=not headed if headed else bool(config.get("headless", True)),
            viewport={"width": 1600, "height": 1200},
            args=args,
        )
    except PlaywrightError as exc:
        if str(config.get("browser_mode")) == "system-edge-default":
            raise RuntimeError(
                "Unable to use your current Edge profile while Edge is open. "
                "Close all Microsoft Edge windows, then retry the command."
            ) from exc
        raise


def first_page(context):
    if context.pages:
        return context.pages[0]
    return context.new_page()


def selector_exists(page, selectors: tuple[str, ...], *, timeout_ms: int = 1000) -> bool:
    for selector in selectors:
        with suppress(PlaywrightTimeoutError):
            page.locator(selector).first.wait_for(state="visible", timeout=timeout_ms)
            return True
    return False


def wait_for_mailbox_ready(page, *, timeout_seconds: int = 60) -> bool:
    deadline = time.time() + timeout_seconds
    while time.time() < deadline:
        if selector_exists(page, MAILBOX_READY_SELECTORS, timeout_ms=700):
            return True
        page.wait_for_timeout(800)
    return selector_exists(page, MAILBOX_READY_SELECTORS, timeout_ms=500)


def sign_in_required(page) -> bool:
    return selector_exists(page, SIGNIN_HINT_SELECTORS, timeout_ms=400) or "login.microsoftonline.com" in page.url


def goto_outlook_mail(page, config: dict[str, Any]):
    page.goto(str(config["outlook_url"]), wait_until="domcontentloaded", timeout=120000)
    with suppress(PlaywrightTimeoutError):
        page.wait_for_load_state("networkidle", timeout=5000)


def save_storage_state(context):
    ensure_state_dir()
    context.storage_state(path=str(STORAGE_STATE_PATH))


def web_login(timeout_seconds: int = 900):
    config = load_config()
    login_config = preferred_login_config(config)
    with sync_playwright() as playwright:
        context = persistent_context(playwright, login_config, headed=True)
        page = first_page(context)
        goto_outlook_mail(page, config)
        print(
            "Browser opened for Outlook Web login. Complete sign-in in the browser window.",
            file=sys.stderr,
        )
        if not wait_for_mailbox_with_login_assist(page, timeout_seconds=timeout_seconds):
            context.close()
            raise RuntimeError(
                f"Outlook Web login did not reach inbox within {timeout_seconds}s. "
                "If sign-in is still pending, rerun `python outlook_helper.py web-login`."
            )
        save_storage_state(context)
        if str(login_config.get("browser_mode")) != "dedicated-profile":
            seed_dedicated_profile_from_storage_state(playwright, config, context.storage_state())
        context.close()


def web_status() -> dict[str, Any]:
    config = load_config()
    token_status = outlook_token_status()
    seconds_remaining = int(token_status.get("seconds_remaining") or 0) if token_status.get("present") else 0
    return {
        "config_path": str(CONFIG_PATH),
        "config_present": CONFIG_PATH.exists(),
        "browser_mode": str(config["browser_mode"]),
        "preferred_login_mode": str(preferred_login_config(config)["browser_mode"]),
        "token_refresh_window_seconds": int(config["token_refresh_window_seconds"]),
        "refresh_recommended": bool(token_status.get("present")) and seconds_remaining <= int(config["token_refresh_window_seconds"]),
        "edge_user_data_dir": str(resolved_user_data_dir(config)),
        "profile_dir": str(PROFILE_DIR),
        "profile_present": PROFILE_DIR.exists(),
        "storage_state_path": str(STORAGE_STATE_PATH),
        "storage_state_present": STORAGE_STATE_PATH.exists(),
        "outlook_url": str(config["outlook_url"]),
        "token_status": token_status,
    }


def reset_web_session():
    with suppress(FileNotFoundError):
        STORAGE_STATE_PATH.unlink()
    with suppress(FileNotFoundError):
        if PROFILE_DIR.exists():
            shutil.rmtree(PROFILE_DIR)


def require_logged_in_mail_page(page):
    if wait_for_mailbox_ready(page, timeout_seconds=45):
        return
    stage = detect_login_stage(page)
    if sign_in_required(page):
        raise RuntimeError(
            "Outlook Web sign-in is required. Run `python outlook_helper.py web-login`, "
            f"finish login in the browser, then retry. Detected stage: {stage or 'sign-in'}."
        )
    raise RuntimeError(
        f"Outlook Web inbox did not load successfully. Current URL: {page.url}"
    )


def dismiss_time_zone_dialog(page):
    page.evaluate(
        """
        () => {
          const dlg = document.querySelector('[role="dialog"]');
          if (!dlg) return false;
          const text = dlg.innerText || '';
          if (!text.includes('你的时区发生了变化')) return false;
          const buttons = dlg.querySelectorAll('button,[role="button"]');
          if (buttons[2]) {
            buttons[2].click();
            return true;
          }
          return false;
        }
        """
    )


def open_calendar_compose(page, entry: dict[str, Any]):
    params = {
        "path": "/calendar/action/compose",
        "rru": "addevent",
        "subject": str(entry.get("subject") or ""),
        "startdt": datetime.strptime(entry["start_local"], "%Y-%m-%d %H:%M").strftime("%Y-%m-%dT%H:%M:%S"),
        "enddt": datetime.strptime(entry["end_local"], "%Y-%m-%d %H:%M").strftime("%Y-%m-%dT%H:%M:%S"),
    }
    if entry.get("location"):
        params["location"] = str(entry["location"])
    if entry.get("body"):
        params["body"] = str(entry["body"])
    url = "https://outlook.office.com/calendar/deeplink/compose?" + urllib.parse.urlencode(params)
    page.goto(url, wait_until="domcontentloaded", timeout=120000)
    page.wait_for_timeout(5000)
    dismiss_time_zone_dialog(page)
    page.wait_for_timeout(500)


def event_time_range_label(entry: dict[str, Any]) -> str:
    start_dt = datetime.strptime(entry["start_local"], "%Y-%m-%d %H:%M")
    end_dt = datetime.strptime(entry["end_local"], "%Y-%m-%d %H:%M")
    return f"{start_dt.strftime('%H:%M')} 到 {end_dt.strftime('%H:%M')}"


def collect_existing_events_in_compose(page) -> list[dict[str, str]]:
    js = """
    () => Array.from(document.querySelectorAll('div[title]'))
      .map(el => {
        const title = (el.getAttribute('title') || '').trim();
        if (!title) return null;
        const parts = title.split('\\n').map(s => s.trim()).filter(Boolean);
        if (parts.length < 2) return null;
        return {
          subject: parts[0],
          time_range: parts[parts.length - 1],
          location: parts.length > 2 ? parts.slice(1, -1).join(' | ') : '',
          title,
        };
      })
      .filter(Boolean)
    """
    return page.evaluate(js)


def same_subject(left: str, right: str) -> bool:
    return " ".join((left or "").split()).strip().lower() == " ".join((right or "").split()).strip().lower()


def has_exact_duplicate(existing_events: list[dict[str, str]], entry: dict[str, Any]) -> bool:
    wanted_subject = str(entry.get("subject") or "")
    wanted_range = event_time_range_label(entry)
    for item in existing_events:
        if same_subject(item.get("subject", ""), wanted_subject) and item.get("time_range", "") == wanted_range:
            return True
    return False


def find_same_subject_conflicts(existing_events: list[dict[str, str]], entry: dict[str, Any]) -> list[dict[str, str]]:
    wanted_subject = str(entry.get("subject") or "")
    wanted_range = event_time_range_label(entry)
    conflicts = []
    for item in existing_events:
        if same_subject(item.get("subject", ""), wanted_subject) and item.get("time_range", "") != wanted_range:
            conflicts.append(item)
    return conflicts


def click_save_in_compose(page):
    clicked = page.evaluate(
        """
        () => {
          const buttons = Array.from(document.querySelectorAll('button,[role="button"]'));
          const saveButton = buttons.find(el => (el.innerText || '').includes('保存'));
          if (!saveButton) return false;
          saveButton.click();
          return true;
        }
        """
    )
    if not clicked:
        raise RuntimeError("Could not find Save button in Outlook Web calendar compose page")
    page.wait_for_timeout(8000)


def detect_login_stage(page) -> str:
    url = page.url.lower()
    text = ""
    with suppress(Exception):
        text = page.locator("body").inner_text(timeout=1500)
    lowered = text.lower()
    if "fido/get" in url or "人脸、指纹、pin" in text or "security key" in lowered:
        return "windows-hello"
    if "选择帐户" in text or "pick an account" in lowered:
        return "account-picker"
    if "保持登录状态" in text or "stay signed in" in lowered:
        return "stay-signed-in"
    if "输入密码" in text or "enter password" in lowered or "input[type='password']" in lowered:
        return "password"
    if "login.microsoftonline.com" in url:
        return "microsoft-login"
    return ""


def try_advance_login(page) -> bool:
    stage = detect_login_stage(page)
    if stage == "account-picker":
        tiles = page.locator("div.table")
        count = tiles.count()
        for idx in range(count):
            tile = tiles.nth(idx)
            tile_id = tile.get_attribute("id") or ""
            if tile_id == "otherTile":
                continue
            tile.click(timeout=3000)
            return True
    if stage == "stay-signed-in":
        for selector in ("#idSIButton9", "input[type='submit']"):
            locator = page.locator(selector).first
            if locator.count():
                with suppress(Exception):
                    locator.click(timeout=3000)
                    return True
    return False


def wait_for_mailbox_with_login_assist(page, *, timeout_seconds: int = 60) -> bool:
    deadline = time.time() + timeout_seconds
    last_stage = ""
    while time.time() < deadline:
        if wait_for_mailbox_ready(page, timeout_seconds=2):
            return True
        stage = detect_login_stage(page)
        if stage and stage != last_stage:
            last_stage = stage
            if stage == "windows-hello":
                print(
                    "Outlook Web is waiting for Windows Hello / security confirmation. "
                    "Complete the system prompt, then the browser should continue.",
                    file=sys.stderr,
                )
        if try_advance_login(page):
            page.wait_for_timeout(1500)
            continue
        page.wait_for_timeout(1000)
    return wait_for_mailbox_ready(page, timeout_seconds=2)


def wait_for_mailbox_for_automatic_refresh(page, *, timeout_seconds: int = 30) -> bool:
    deadline = time.time() + timeout_seconds
    while time.time() < deadline:
        if wait_for_mailbox_ready(page, timeout_seconds=2):
            return True
        stage = detect_login_stage(page)
        if stage in {"windows-hello", "password", "microsoft-login"}:
            return False
        if sign_in_required(page):
            if not try_advance_login(page):
                return False
            page.wait_for_timeout(1500)
            continue
        page.wait_for_timeout(1000)
    return wait_for_mailbox_ready(page, timeout_seconds=2)


def collect_mail_rows(page, max_items: int) -> list[dict[str, Any]]:
    js = """
    ({ selectors, maxItems }) => {
      const result = [];
      const seen = new Set();
      const visible = (el) => {
        if (!el) return false;
        const style = window.getComputedStyle(el);
        if (style.display === "none" || style.visibility === "hidden") return false;
        const rect = el.getBoundingClientRect();
        return rect.width > 0 && rect.height > 0;
      };
      const pushCandidate = (el, sourceSelector) => {
        if (!visible(el)) return;
        const key = el.getAttribute("data-convid")
          || el.getAttribute("data-itemid")
          || el.getAttribute("id")
          || `${sourceSelector}:${result.length}`;
        if (seen.has(key)) return;
        seen.add(key);
        const text = (el.innerText || "").split(/\\n+/).map(s => s.trim()).filter(Boolean);
        if (!text.length) return;
        const descendantData = Array.from(el.querySelectorAll("[title],[aria-label],[datetime],[data-testid]"))
          .slice(0, 30)
          .map(node => ({
            title: node.getAttribute("title") || "",
            aria_label: node.getAttribute("aria-label") || "",
            datetime: node.getAttribute("datetime") || "",
            data_testid: node.getAttribute("data-testid") || "",
            text: (node.textContent || "").trim(),
          }));
        result.push({
          key,
          source_selector: sourceSelector,
          text_lines: text.slice(0, 12),
          text_blob: text.join("\\n"),
          aria_label: el.getAttribute("aria-label") || "",
          title: el.getAttribute("title") || "",
          data_convid: el.getAttribute("data-convid") || "",
          data_itemid: el.getAttribute("data-itemid") || "",
          descendant_data: descendantData,
        });
      };
      for (const selector of selectors) {
        const els = Array.from(document.querySelectorAll(selector));
        for (const el of els) {
          pushCandidate(el, selector);
          if (result.length >= maxItems * 3) {
            return result;
          }
        }
      }
      return result;
    }
    """
    rows = page.evaluate(js, {"selectors": list(MAIL_LIST_SELECTORS), "maxItems": max_items})
    return [row for row in rows if row.get("text_lines")]


def parse_datetime_text(text: str, zone: ZoneInfo) -> str:
    now = datetime.now(zone)
    patterns = (
        "%d/%m/%Y %H:%M",
        "%d/%m/%Y %I:%M %p",
        "%Y-%m-%d %H:%M",
        "%a %d %b %Y %H:%M",
        "%A, %d %B %Y %H:%M",
        "%d %b %Y %H:%M",
        "%d %B %Y %H:%M",
    )
    cleaned = " ".join((text or "").replace(",", " ").split())

    weekday_map = {
        "周一": 0,
        "周二": 1,
        "周三": 2,
        "周四": 3,
        "周五": 4,
        "周六": 5,
        "周日": 6,
        "周天": 6,
        "星期一": 0,
        "星期二": 1,
        "星期三": 2,
        "星期四": 3,
        "星期五": 4,
        "星期六": 5,
        "星期日": 6,
        "星期天": 6,
        "Mon": 0,
        "Tue": 1,
        "Wed": 2,
        "Thu": 3,
        "Fri": 4,
        "Sat": 5,
        "Sun": 6,
    }
    for label, weekday in weekday_map.items():
        if cleaned.startswith(label):
            rest = cleaned[len(label) :].strip()
            md_match = re.match(r"(?P<month>\d{1,2})/(?P<day>\d{1,2})(?:\s+(?P<hour>\d{1,2}):(?P<minute>\d{2}))?$", rest)
            if md_match:
                month = int(md_match.group("month"))
                day = int(md_match.group("day"))
                hour = int(md_match.group("hour") or 0)
                minute = int(md_match.group("minute") or 0)
                year = now.year
                with suppress(ValueError):
                    dt = datetime(year, month, day, hour, minute, tzinfo=zone)
                    if dt > now + timedelta(days=2):
                        dt = datetime(year - 1, month, day, hour, minute, tzinfo=zone)
                    return dt.isoformat(sep=" ", timespec="seconds")
            hm_match = re.match(r"(?P<hour>\d{1,2}):(?P<minute>\d{2})$", rest)
            if hm_match:
                target = weekday
                days_back = (now.weekday() - target) % 7
                candidate = now - timedelta(days=days_back)
                candidate = candidate.replace(
                    hour=int(hm_match.group("hour")),
                    minute=int(hm_match.group("minute")),
                    second=0,
                    microsecond=0,
                )
                if candidate > now:
                    candidate -= timedelta(days=7)
                return candidate.isoformat(sep=" ", timespec="seconds")

    for pattern in patterns:
        with suppress(ValueError):
            return datetime.strptime(cleaned, pattern).replace(tzinfo=zone).isoformat(
                sep=" ",
                timespec="seconds",
            )
    return ""


def first_nonempty(*values: str) -> str:
    for value in values:
        if value and value.strip():
            return value.strip()
    return ""


def is_icon_or_avatar_line(line: str) -> bool:
    value = (line or "").strip()
    if not value:
        return True
    if len(value) <= 3 and re.fullmatch(r"[A-Z]{1,3}", value):
        return True
    if re.fullmatch(r"[\W_]+", value, flags=re.UNICODE):
        return True
    if all(ord(ch) > 0xE000 or (0xE000 <= ord(ch) <= 0xF8FF) for ch in value):
        return True
    return False


def cleaned_content_lines(row: dict[str, Any]) -> list[str]:
    lines = []
    for raw in row.get("text_lines") or []:
        line = str(raw or "").strip()
        if not line or is_icon_or_avatar_line(line):
            continue
        lines.append(line)
    return lines


def parse_received(row: dict[str, Any], zone: ZoneInfo) -> str:
    candidates: list[str] = []
    for key in ("title", "aria_label"):
        value = str(row.get(key) or "").strip()
        if value:
            candidates.append(value)
    for item in row.get("descendant_data") or []:
        for key in ("datetime", "title", "aria_label", "text"):
            value = str(item.get(key) or "").strip()
            if value:
                candidates.append(value)
    for value in candidates:
        exact = parse_datetime_text(value, zone)
        if exact:
            return exact
        iso_match = re.search(r"\\d{4}-\\d{2}-\\d{2}T\\d{2}:\\d{2}:\\d{2}", value)
        if iso_match:
            with suppress(ValueError):
                dt = datetime.fromisoformat(iso_match.group(0)).replace(tzinfo=zone)
                return dt.isoformat(sep=" ", timespec="seconds")
    for line in reversed(row.get("text_lines") or []):
        exact = parse_datetime_text(line, zone)
        if exact:
            return exact
    return ""


def normalize_topic(subject: str) -> str:
    prefixes = (
        "RE:",
        "FW:",
        "FWD:",
        "SV:",
        "答复:",
        "回复:",
        "转发:",
        "RE：",
        "FW：",
        "FWD：",
        "SV：",
        "答复：",
        "回复：",
        "转发：",
    )
    value = (subject or "").strip()
    changed = True
    while changed and value:
        changed = False
        upper = value.upper()
        for prefix in prefixes:
            if upper.startswith(prefix.upper()):
                value = value[len(prefix) :].strip()
                changed = True
                break
    return " ".join(value.split())


def row_to_message(row: dict[str, Any], zone: ZoneInfo) -> dict[str, Any] | None:
    lines = cleaned_content_lines(row)
    if len(lines) < 2:
        return None
    date_index = -1
    for idx, line in enumerate(lines):
        if parse_datetime_text(line, zone):
            date_index = idx
            break

    if date_index >= 2:
        sender = lines[0]
        subject = lines[1]
        preview_parts = lines[date_index + 1 :]
    else:
        sender = lines[0]
        subject = lines[1] if len(lines) > 1 else ""
        preview_parts = lines[2:]

    preview = " ".join(preview_parts).strip()
    if not subject:
        return None
    return {
        "received": parse_received(row, zone),
        "sender": sender,
        "sender_email": "",
        "subject": subject,
        "topic": normalize_topic(subject),
        "message_class": "outlook-web.message",
        "importance": None,
        "flag_status": None,
        "is_meeting": False,
        "preview": preview[:MAIL_PREVIEW_CHARS],
        "entry_id": first_nonempty(str(row.get("data_itemid") or ""), str(row.get("data_convid") or ""), str(row.get("key") or "")),
        "conversation_id": str(row.get("data_convid") or ""),
        "web_link": "",
    }


def fetch_recent_messages_web(hours: int = 24, max_items: int = 250) -> list[dict[str, Any]]:
    return fetch_recent_messages_http(hours, max_items)


def ensure_calendar_web(spec_path: Path):
    return ensure_calendar_http(spec_path)


def ensure_outlook_session(min_valid_seconds: int | None = None, timeout_seconds: int = 90) -> dict[str, Any]:
    ensure_fresh_storage_state(
        allow_auto_refresh=True,
        min_valid_seconds=min_valid_seconds,
        refresh_timeout_seconds=timeout_seconds,
    )
    status = web_status()
    status["session_ready"] = bool(status.get("token_status", {}).get("present"))
    return status


def main():
    parser = argparse.ArgumentParser()
    sub = parser.add_subparsers(dest="cmd", required=True)

    recent = sub.add_parser("recent-mail")
    recent.add_argument("--hours", type=int, default=24)
    recent.add_argument("--max-items", type=int, default=250)
    recent.add_argument("--mode", choices=("web", "auto"), default="web")

    ensure = sub.add_parser("ensure-calendar")
    ensure.add_argument("--spec", required=True)
    ensure.add_argument("--mode", choices=("web", "auto"), default="web")

    login = sub.add_parser("web-login")
    login.add_argument("--timeout-seconds", type=int, default=900)

    ensure_session = sub.add_parser("ensure-session")
    ensure_session.add_argument("--min-valid-seconds", type=int, default=None)
    ensure_session.add_argument("--timeout-seconds", type=int, default=90)

    sub.add_parser("web-status")
    sub.add_parser("web-logout")

    args = parser.parse_args()

    if args.cmd == "recent-mail":
        print(json.dumps(fetch_recent_messages_web(args.hours, args.max_items), **COMPACT_JSON_KWARGS))
        return

    if args.cmd == "ensure-calendar":
        print(json.dumps(ensure_calendar_web(Path(args.spec)), **COMPACT_JSON_KWARGS))
        return

    if args.cmd == "web-login":
        web_login(timeout_seconds=args.timeout_seconds)
        print(json.dumps({"logged_in": True, "profile_dir": str(PROFILE_DIR)}, **COMPACT_JSON_KWARGS))
        return

    if args.cmd == "ensure-session":
        print(
            json.dumps(
                ensure_outlook_session(
                    min_valid_seconds=args.min_valid_seconds,
                    timeout_seconds=args.timeout_seconds,
                ),
                **COMPACT_JSON_KWARGS,
            )
        )
        return

    if args.cmd == "web-status":
        print(json.dumps(web_status(), **COMPACT_JSON_KWARGS))
        return

    if args.cmd == "web-logout":
        reset_web_session()
        print(json.dumps({"logged_out": True, "profile_dir": str(PROFILE_DIR)}, **COMPACT_JSON_KWARGS))
        return


if __name__ == "__main__":
    try:
        sys.exit(main())
    except (RuntimeError, PlaywrightError) as exc:
        print(str(exc), file=sys.stderr)
        sys.exit(1)
