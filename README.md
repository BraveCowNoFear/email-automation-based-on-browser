# Email automation based on browser

Browser-backed Outlook Web automation for recurring email briefings and calendar updates.

This project avoids desktop Outlook COM/MAPI and avoids creating a Microsoft Graph app registration. It reuses a logged-in Outlook Web session, extracts the Outlook Web access token from Playwright storage state, and calls Outlook REST endpoints used by the web client.

## What it includes

- `scripts/outlook_helper.py`: CLI for login refresh, session checks, recent mail extraction, and calendar event creation.
- `scripts/outlook_briefing_data.py`: compact handoff generator for Codex or other automation agents.
- `skills/outlook-browser-email-automation`: installable Codex skill that wraps the scripts and briefing policy.
- `automation-examples/outlook-browser-email.toml`: short automation prompt that calls the skill instead of embedding the full workflow.

## Setup

Install dependencies:

```powershell
python -m pip install -r requirements.txt
```

Copy and edit the example config only if the defaults are not enough:

```powershell
Copy-Item .\scripts\mail_automation_web_config.example.json .\scripts\mail_automation_web_config.json
```

The default state directory is under `%LOCALAPPDATA%\codex-mail-automation-web`. Do not commit generated config, browser state, or debug output.

Useful config fields:

- `local_time_zone`: IANA timezone used to interpret local mail and calendar times, for example `Europe/London`.
- `outlook_time_zone`: Outlook/Exchange timezone name used when creating events, for example `GMT Standard Time`.
- `edge_user_data_dir` and `edge_profile_directory`: browser profile used for the initial web login.

## First login

```powershell
python .\scripts\outlook_helper.py web-login
```

Finish the Microsoft login flow in the browser. The helper exits after Outlook Mail loads and saves a local storage state.

## Commands

Check session:

```powershell
python .\scripts\outlook_helper.py web-status
```

Ensure the session remains valid for at least three hours:

```powershell
python .\scripts\outlook_helper.py ensure-session --min-valid-seconds 10800
```

Fetch recent mail:

```powershell
python .\scripts\outlook_helper.py recent-mail --hours 24 --max-items 250
```

Generate a compact automation handoff:

```powershell
python .\scripts\outlook_briefing_data.py --ensure-session --hours 24 --max-items 250
```

## Install the Codex skill

Copy `skills/outlook-browser-email-automation` into your Codex skills directory, usually:

```powershell
$codexHome = if ($env:CODEX_HOME) { $env:CODEX_HOME } else { Join-Path $env:USERPROFILE ".codex" }
New-Item -ItemType Directory -Force -Path (Join-Path $codexHome "skills") | Out-Null
Copy-Item .\skills\outlook-browser-email-automation (Join-Path $codexHome "skills") -Recurse -Force
```

Then use:

```text
Use $outlook-browser-email-automation to review recent Outlook mail and update calendar items.
```

## Notes

- This is browser-session automation, not an official Microsoft Graph integration.
- If Outlook requires password, Windows Hello, or MFA, run `web-login` manually.
- Keep saved browser state private; it can grant access to mail.
- Calendar creation is conservative and deduplicates exact same-title same-time events.
