# Briefing Policy

## Mail Selection

Keep only user-actionable mail:

- deadlines, meetings, approvals, bills, security/account issues, follow-ups, reply-needed mail
- class, tutorial, revision, travel, delivery, lodging, or operational notices affecting the user
- marketing/newsletters only when they contain a concrete deadline or action

Ignore routine promotions, FYI newsletters, and duplicate notification noise unless they affect a required action.

## Deduplication And Grouping

- Prefer `entry_id`.
- If no `entry_id`, dedupe by sender + normalized `topic` + received time.
- Group final output by sender + topic, not read/unread state.
- Use exact absolute dates. If a date or time is inferred, say it is inferred.

## Calendar Rules

Create or update events only for clear commitments:

- scheduled meeting, class, tutorial, revision session, interview, trip, appointment
- explicit deadline or due date
- date-only commitment where a placeholder is useful

Skip vague, tentative, promotional, or info-only mentions. For date-only commitments, use a local 09:00-09:30 placeholder unless automation memory clearly suggests another default.

Before reporting calendar work, verify exact date, local time, timezone assumption, placeholder usage, and duplicate handling. Reject unsupported calendar proposals even if the mail looks important.

For tutorial, class, meeting, and appointment events, include the relevant person or professor lastname in the calendar subject when known. Include the physical or online location in the calendar location field when known. Put source sender, full person name, lastname, location, and any date/time inference in the body so later audits can understand why the event was created.

## Final Output

Write in concise Chinese, in this order:

1. Short heading
2. Grouped items: key point, why it matters, needed action
3. `Priority summary`
4. `Recommended next actions`
5. `Calendar changes`
6. One-line failure note only if fetch failed after retry

If there is no relevant mail, say so plainly. Do not restate commands or raw email text unless needed for the user to act.
