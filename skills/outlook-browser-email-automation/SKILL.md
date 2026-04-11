---
name: outlook-browser-email-automation
description: Browser-based Outlook Web automation for recurring mail briefings, session refresh, recent-mail extraction, actionable-email triage, and calendar event creation without Outlook COM or Microsoft Graph app registration. Use when Codex needs to run or maintain the user's Outlook email automation, merge Outlook briefing/session cron prompts, refresh a browser-backed Outlook session, summarize recent Outlook mail, or create/update Outlook calendar events from email.
---

# Outlook Browser Email Automation

## Workflow

Use the bundled scripts from this skill directory unless the user explicitly points to another checkout.

1. Read the relevant automation memory first if the run is from a Codex automation:
   - `<CODEX_HOME>\automations\outlook-browser-email\memory.md`
   - fall back to the legacy memories under `outlook\memory.md` and `outlook-session\memory.md` if the merged memory is absent.
2. Run the compact data handoff:

```powershell
$codexHome = if ($env:CODEX_HOME) { $env:CODEX_HOME } else { Join-Path $env:USERPROFILE ".codex" }
$skillDir = Join-Path $codexHome "skills\outlook-browser-email-automation"
python (Join-Path $skillDir "scripts\outlook_briefing_data.py") --ensure-session --hours 24 --max-items 250
```

If `CODEX_HOME` is unset, use the local user Codex directory, usually `%USERPROFILE%\.codex` on Windows.

3. If the handoff reports `fetch_status=failed`, run `scripts/outlook_helper.py web-status` once to classify login/token state, then stop with a concise Chinese failure note.
4. For successful fetches, load `references/briefing-policy.md` and produce the final Chinese briefing from the compact handoff.
5. Only create calendar entries for clear commitments. Write a temporary JSON spec and run:

```powershell
python "<skill-dir>\scripts\outlook_helper.py" ensure-calendar --spec "<spec.json>"
```

6. Report exactly what calendar work was created, skipped as duplicate, updated, or rejected.

## Commands

- Manual login refresh: `python <skill-dir>\scripts\outlook_helper.py web-login`
- Session status: `python <skill-dir>\scripts\outlook_helper.py web-status`
- Session preflight only: `python <skill-dir>\scripts\outlook_helper.py ensure-session --min-valid-seconds 10800`
- Recent mail only: `python <skill-dir>\scripts\outlook_helper.py recent-mail --hours 24 --max-items 250`
- Reset saved web session: `python <skill-dir>\scripts\outlook_helper.py web-logout`

## Boundaries

- Use Outlook Web login state and Outlook REST endpoints exposed to the web client.
- Do not use raw Outlook COM, desktop Outlook automation, MAPI, or a custom Microsoft Entra app unless the user explicitly asks to replace this approach.
- Use Playwright only for web login/session refresh paths; normal mail and calendar operations should use the helper's REST-backed commands.
- Keep output concise and Chinese by default.
