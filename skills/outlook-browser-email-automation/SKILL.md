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
2. Read stable user facts before deciding whether an email applies to the user:
   - `<CODEX_HOME>\memories\PERSONAL_PROFILE.md`
   - If mail branches by course, tutorial group, college, year, identity, or similar user-specific context, search the global memory directory for the relevant fact before creating or rejecting a calendar entry.
   - Current known course fact: the user is tutorial group 1 for 1P4H / Wouter Mostert tutorial grouping.
3. Run the compact data handoff:

```powershell
$codexHome = if ($env:CODEX_HOME) { $env:CODEX_HOME } else { Join-Path $env:USERPROFILE ".codex" }
$skillDir = Join-Path $codexHome "skills\outlook-browser-email-automation"
python (Join-Path $skillDir "scripts\outlook_briefing_data.py") --ensure-session --hours 72 --max-items 250
```

If `CODEX_HOME` is unset, use the local user Codex directory, usually `%USERPROFILE%\.codex` on Windows.

4. If the handoff reports `fetch_status=failed`, run `scripts/outlook_helper.py web-status` once to classify login/token state, then stop with a concise Chinese failure note.
5. For successful fetches, load `references/briefing-policy.md` and produce the final Chinese briefing from the compact handoff.
6. Only create calendar entries for clear commitments. Calendar specs must include the relevant person/professor lastname in the event `subject` when known, include the physical/online `location` when known, and mention source sender, full person name, lastname, location, and any inference in `body`. Write a temporary JSON spec and run:

```powershell
python "<skill-dir>\scripts\outlook_helper.py" ensure-calendar --spec "<spec.json>"
```

7. Report exactly what calendar work was created, skipped as duplicate, updated, or rejected.

## Commands

- Manual login refresh: `python <skill-dir>\scripts\outlook_helper.py web-login`
- Session status: `python <skill-dir>\scripts\outlook_helper.py web-status`
- Session preflight only: `python <skill-dir>\scripts\outlook_helper.py ensure-session --min-valid-seconds 10800`
- Recent mail only: `python <skill-dir>\scripts\outlook_helper.py recent-mail --hours 72 --max-items 250`
- Reset saved web session: `python <skill-dir>\scripts\outlook_helper.py web-logout`

## Boundaries

- This skill is only for the user's school Outlook Web mailbox / school-account automation path. Do not use it for personal Outlook / Hotmail mailboxes or consumer-email verification codes unless the user explicitly asks to repurpose it.
- Use Outlook Web login state and Outlook REST endpoints exposed to the web client.
- Do not use raw Outlook COM, desktop Outlook automation, MAPI, or a custom Microsoft Entra app unless the user explicitly asks to replace this approach.
- Use Playwright only for web login/session refresh paths; normal mail and calendar operations should use the helper's REST-backed commands.
- Keep output concise and Chinese by default.
