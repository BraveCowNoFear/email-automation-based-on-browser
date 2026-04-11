#!/usr/bin/env python3
"""Produce a compact Outlook mail handoff for Codex automation runs."""

from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
from pathlib import Path
from typing import Any


COMPACT_JSON_KWARGS = {"ensure_ascii": False, "separators": (",", ":")}
KEEP_FIELDS = (
    "received",
    "sender",
    "sender_email",
    "subject",
    "topic",
    "importance",
    "flag_status",
    "is_meeting",
    "preview",
    "entry_id",
    "conversation_id",
    "web_link",
)


def helper_path() -> Path:
    if os.getenv("MAIL_AUTOMATION_HELPER"):
        return Path(os.path.expandvars(os.path.expanduser(os.environ["MAIL_AUTOMATION_HELPER"]))).resolve()
    return Path(__file__).resolve().with_name("outlook_helper.py")


def run_helper(args: list[str]) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        [sys.executable, str(helper_path()), *args],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=False,
    )


def compact_message(message: dict[str, Any]) -> dict[str, Any]:
    return {key: message.get(key) for key in KEEP_FIELDS if message.get(key) not in (None, "")}


def dedupe(messages: list[dict[str, Any]]) -> list[dict[str, Any]]:
    seen: set[str] = set()
    result: list[dict[str, Any]] = []
    for item in messages:
        key = str(item.get("entry_id") or "").strip()
        if not key:
            parts = [
                str(item.get(part) or "").strip().lower()
                for part in ("sender_email", "sender", "topic", "subject", "received", "preview")
            ]
            key = "|".join(part for part in parts if part)
        if not key:
            key = json.dumps(item, ensure_ascii=False, sort_keys=True)
        if key in seen:
            continue
        seen.add(key)
        result.append(compact_message(item))
    return result


def parse_json(stdout: str) -> Any:
    try:
        return json.loads(stdout)
    except json.JSONDecodeError as exc:
        raise RuntimeError(f"helper returned non-JSON output: {stdout[:500]}") from exc


def fetch_recent_mail(hours: int, max_items: int, retries: int) -> tuple[list[dict[str, Any]], list[str]]:
    errors: list[str] = []
    for _ in range(retries + 1):
        proc = run_helper(["recent-mail", "--hours", str(hours), "--max-items", str(max_items)])
        if proc.returncode == 0:
            data = parse_json(proc.stdout)
            if not isinstance(data, list):
                raise RuntimeError("recent-mail returned JSON that is not a list")
            return data, errors
        errors.append((proc.stderr or proc.stdout or f"exit {proc.returncode}").strip())
    return [], errors


def ensure_session(min_valid_seconds: int) -> dict[str, Any]:
    proc = run_helper(["ensure-session", "--min-valid-seconds", str(min_valid_seconds)])
    payload: dict[str, Any] = {
        "ok": proc.returncode == 0,
        "error": (proc.stderr or proc.stdout or "").strip() if proc.returncode else "",
    }
    if proc.returncode == 0 and proc.stdout.strip():
        parsed = parse_json(proc.stdout)
        if isinstance(parsed, dict):
            payload.update(parsed)
    return payload


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--hours", type=int, default=24)
    parser.add_argument("--max-items", type=int, default=250)
    parser.add_argument("--retries", type=int, default=1)
    parser.add_argument("--ensure-session", action="store_true")
    parser.add_argument("--min-valid-seconds", type=int, default=10800)
    args = parser.parse_args()
    hours = max(0, args.hours)
    max_items = max(0, args.max_items)
    retries = max(0, args.retries)

    output: dict[str, Any] = {
        "fetch_status": "failed",
        "helper": str(helper_path()),
        "hours": hours,
        "max_items": max_items,
    }
    if args.ensure_session:
        output["session"] = ensure_session(args.min_valid_seconds)
        if not output["session"].get("ok"):
            output["failure_note"] = output["session"].get("error") or "session preflight failed"
            print(json.dumps(output, **COMPACT_JSON_KWARGS))
            return 1

    try:
        messages, errors = fetch_recent_mail(hours, max_items, retries)
        compact = dedupe(messages)
        output.update(
            {
                "fetch_status": "success",
                "raw_count": len(messages),
                "deduped_count": len(compact),
                "messages": compact,
                "errors_before_success": errors,
            }
        )
    except Exception as exc:
        output["failure_note"] = str(exc)
    else:
        if errors:
            output["failure_note"] = errors[-1]

    print(json.dumps(output, **COMPACT_JSON_KWARGS))
    return 0 if output["fetch_status"] == "success" else 1


if __name__ == "__main__":
    raise SystemExit(main())
