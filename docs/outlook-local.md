# Local Outlook MCP

This implementation uses Classic Outlook's COM object model through the installed
Windows PowerShell, using the logged-in Windows user's default Outlook profile.
No pywin32, Graph registration, passwords, packages, or models are installed.
New Outlook is not supported by this adapter.

## Use in MobinAid

Open Connections (اتصال‌ها) and select اتصال Outlook همین رایانه. This performs a
real stdio MCP health check and calendar sync before reporting success. Then ask
in chat about your emails or calendar. Each actual tool execution is displayed in
the activity/evidence panel. Disconnect stops further polling and removes pending
and due Outlook reminders. Reconnect is required after backend restart.

Five tools: outlook_status, outlook_search_emails, outlook_read_email,
outlook_read_attachment, outlook_calendar. The Python entry point is
`python -m app.outlook` from `enterprise_agent/backend`; it is an actual
stdio FastMCP server. MobinAid launches this via the existing MCP SDK.

Mail search uses literal case-insensitive substring matching in subject, sender
name and plain body of Inbox and Sent Items, scanning at most 1000 items and
returning at most 50. It does not claim whole-mailbox/archive search. Tool responses
explicitly identify scope and possible truncation. No address-book enumeration,
mail sending, editing, deletion, arbitrary commands, or attachment execution APIs.

Date-only mail listing/existence questions for today/yesterday/trailing seven days
are compiled into timezone-aware ranges (Asia/Tehran day boundaries) and answered
from actual Inbox metadata, without model-generated dates or invented summaries.
The answer identifies the exact interval, folder scope, and incomplete results.
Numeric day ranges (1–31 days, Persian/Arabic/Latin digits) and read/unread filters
are also compiled directly. The unread flag uses Outlook's actual UnRead property;
listing does not mark messages read. Identical repeated model searches stop with
an explicit no-progress error instead of continuing until the run deadline.
General search supports folder selection and offset pagination; mailbox changes
between requests can shift pages. Email body reads use offset/body_limit (1600
characters default, 2400 maximum) and expose total length and next_offset.
The agent uses session-local short references instead of native Outlook EntryIDs,
small tool pages and explicit previews. Full bounded results remain in UI evidence;
old message bodies are not repeatedly accumulated in model context. A preview is
not represented as a complete email read. No automatic model download or context
size increase is used. Follow-up IDs do not persist across separate MCP sessions.

Simple latest-email requests use a deterministic search → read workflow. Displaying
the actual subject/body does not wait for model generation. Explicit summary requests
still use the configured local model; if it fails, the response identifies that
failure and displays source text instead of inventing a summary. In the real probe,
CPU model generation timed out; only source retrieval/listing and bounded read
workflows are verified end to end. Do not claim general summarization reliability
from the read-only probe.

Attachment bytes (file attachments <=5MB) are transferred through MCP, sanitized
into the existing owner-scoped attachment storage, and only extracted supported
text is passed to the local model. Files are never executed. PDF extraction and
image understanding retain the existing unsupported-capability errors. Unsupported
formats are rejected. Download bytes are not placed in UI debug output or prompts.

The default calendar is polled every 60 seconds for seven days ahead. Recurrences
are expanded; cancelled meetings and all-day entries are excluded. A reminder is
scheduled 15 minutes before each upcoming meeting and displayed in chat while the
backend and chat are running. Rescheduled/cancelled pending occurrences are reconciled;
acknowledged occurrences are deduplicated. More than 50 calendar entries fails
explicitly rather than committing a partial sync. Outlook failure is exposed in
Connections; previously cached reminders may be stale until a successful sync.
This does not promise delivery while Windows, Outlook, or MobinAid is stopped.

## Boundaries

This is a single-workstation, single-Windows-user adapter, not a multiuser server.
Only one browser owner can bind to Outlook per backend process. Bind endpoints
retain existing session/origin/API-token checks. Run on loopback only. Do not deploy
it on a company shared server: per-user remote hosting needs proper organizational
authentication and delegated mailbox authorization. Outlook security prompts or
organizational policies are not suppressed. Account settings and execution policy
are not changed. The language model receives untrusted source text with explicit
instructions; summaries still require review and are not guaranteed error-free.

## Verification

From backend, with existing Python and PYTHONPATH pointing at backend:
`python ../scripts/probe_outlook.py` performs real discovery, status, search, email
read, calendar read and prints counts only. No mail contents are logged by the probe.
`python -m unittest discover -s tests -q` includes reminder reconciliation, ownership,
deduplication and denied operations/ranges.
