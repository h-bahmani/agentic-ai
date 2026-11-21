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

Six tools: outlook_status, outlook_search_emails, outlook_read_email,
outlook_read_attachment, outlook_calendar, outlook_resolve_person. The Python entry point is
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
the actual subject/body does not wait for model generation. The summary operation
currently returns source excerpts, not a verified semantic summary. Model fallback
is used to interpret requests outside the explicit filter grammar, and can fail.

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


---

# Outlook: general planning and verified execution

The production entry point `outlook_run` now compiles ordinary requests into a
single typed `MailPlan`. It does not dispatch natural-language requests to the
old example-specific helpers. The direct card/reference path uses `read_selected_mail`
and validates the conversation-scoped message ID.

## Pipeline

1. Explicit filter slots are compiled first. Only requests outside this grammar use
   a schema-constrained Ollama request to propose operation, folder, person and role,
   topic, relative time, read state and latest flag.
2. Pydantic rejects unknown operations, folders, types and ranges. Person/topic must
   be literal source spans. Unsupported extra read-state filters, obvious operation
   changes and outgoing/incoming inconsistencies are rejected. One correction attempt
   is allowed, then an explicit error; no fallback to a fabricated answer/search.
3. Identities are resolved from actual mailbox sender display names/SMTP addresses
   and recipient display names. Exchange sender addresses use PrimarySmtpAddress
   where available. No employee name/address registry or hardcoded company people.
   Exact address/local-part matching precedes cross-script display-name matching.
   Zero/multiple candidates require clarification. Approximate matching is not an
   authoritative corporate directory, and partial mailbox coverage limits discovery.
4. A deterministic executor performs paging, filters, counts and body reads through
   actual MCP. Responses are built from these results, not generated business facts.
5. Identity discovery results are hidden from final email cards; only final selected
   messages or actually read messages are displayed. Technical evidence remains.

## Limits, not guarantees

Qwen3:4b planning has failed on unseen phrasing in real tests. A typed schema alone
does not prove semantic correctness. Added validation catches several failure modes,
not every possible interpretation error. Unsupported conditions should clarify.
Persian-calendar absolute dates, multiple people in one role, attachment filtering
and other advanced filters are not generally implemented. Semantic summary is not complete:
the summary operation explicitly uses source excerpts. Tasks remain source-based
candidate extraction, not confirmed assignment or completion status. Calendar/reminder
support exists separately, but general calendar planning is currently clarified
rather than silently routed to mailbox search. No write operation is enabled.

The planner can take tens of seconds on CPU. No new model or package is installed.
No claim of zero errors, full language understanding or general enterprise readiness.

## Analytical request coverage (2026-11-21)

Rolling-month count/list/body requests compile a real 30-day range and state this
interpretation explicitly (not a calendar month). Unsupported multi-month ranges
clarify rather than silently becoming all-time searches. Count plus «و چیه» retains
both count and listing; explicit summary retains count plus a per-message source
excerpt table with date and correspondent.

«خلاصه متن همش باشه» (or standalone «خلاصه») reuses the previous conversation's
verified plan with its relative time frozen to exact UTC bounds. It re-fetches
matching mail and reads bodies, rather than regenerating unrelated filters.
Summary/read can inspect up to 50 messages; per-body reading remains bounded to
24k characters, with incompleteness disclosed. Priority/tasks remain capped at 10.
This is source-excerpt review, not evaluated semantic summarization.

Live chat acceptance found seven Sent Items in the trailing 30 days and read all
seven for the summary follow-up; the prior all-time count of twelve was not a
month-filtered answer. Tests also verify twelve-message body coverage and exact
date/filter preservation on follow-up. Regression suite: 90 tests.

| Request | Current behavior | Remaining limitation |
| --- | --- | --- |
| Prioritize today's mail needing action | Searches today's Inbox and reads up to 10 message bodies; reports candidates based on Outlook High or explicit request/deadline phrases | Not confirmed assignment, urgency or pending status; read/scan limits are disclosed |
| Identify information-only mail requiring no reply | Explicit unsupported-capability response; no person lookup | Needs evaluated body, recipient and thread analysis before making a no-response claim |
| Latest correspondence for every project | Explicit unsupported-capability response; no person lookup | Needs a verified project catalogue and message-to-project mapping |

These are capability boundaries, not successful implementations of all three
analyses. Passing regression tests does not establish general semantic accuracy.
The mail route does not require Ollama health for deterministic retrieval; model
fallback still requires an available inference provider. Standalone correspondence
questions route to Outlook even in a fresh conversation.

Live acceptance used the actual stdio MCP and Classic Outlook with model generation
disabled: both unsupported requests returned explicit limits without tool calls;
today's priority request executed search and body read, without resolving a person.
At that check one message was found today and no explicit action/High candidate was
detected. This is a snapshot, not a permanent expected mailbox count. Unit suite: 79.


## Verification commands

From the enterprise_agent directory, using the existing Python:

```powershell
& 'D:\CompanyMicroApiTest\.venv\Scripts\python.exe' scripts/verify.py outlook
& 'D:\CompanyMicroApiTest\.venv\Scripts\python.exe' scripts/verify.py plan --text 'Your mailbox question'
```

Outlook mode checks real MCP status, without reading message bodies. Plan mode uses
the actual configured model and prints its plan without executing mailbox operations.
Health mode checks running services. Chat mode explicitly executes the supplied query
against the running application and requires its own connection/session authorization.
It does not transfer an existing browser's Outlook binding automatically.

Unit coverage is consolidated in backend/tests/test_outlook.py (including planner and executor tests). Legacy parser-only tests were removed together
with the inactive parsers, not counted as proof of general model accuracy.

Measured model results before cleanup: 3/4 planning paraphrases passed; combined
count-and-review remained rejected. Actual sent counting and m.arvinfar resolution
passed end to end. CPU planning took approximately 10–32s. This cleanup does not
claim to fix the model's remaining semantic limitations.

## Composable search filters (current)

Explicit From/To/CC, topic/subject, folder, read state and dates are extracted as
independent source slots before model fallback. This is a field grammar, not
whole-question example matching or model training. Named unknown/unsupported
conditions still require clarification; no claim of unrestricted language support.
Sender SMTP addresses, To and CC recipient roles are read from actual Outlook
metadata. Recipient matching distinguishes MAPI Type 1 (To) and Type 2 (CC).
Filters combine with AND. Topic can target subject or general text. Sent-folder
dates use SentOn; received-folder dates use ReceivedTime.

Relative dates and Gregorian YYYY-MM-DD ranges (inclusive calendar days, Tehran,
maximum 31 days) are supported. Persian-calendar absolute dates are explicitly
rejected pending a verified converter; no guessed conversion. Multiple people
in a role, OR/NOT clauses, and attachment-condition queries are not generally
supported and must clarify.

Identity-discovery search results always have show_cards=false, including early
ambiguity/error returns. Sender references such as a displayed email address plus
“از این” do not invoke selected-message follow-up. Summary reports show a compact
source-excerpt table and do not duplicate the discovery/result cards. They remain
extractive, not semantic summaries.

Real API/MCP/Outlook verification after this change passed both sender-summary
phrasings, Deleted Items listing, and a topic+Gregorian-date+CC combination from
an observed message. Unit suite: 69 tests, including a compositional matrix of
people/topics and CC/date constraints. No package/model installed or message modified.

## Explicit addresses and Address Book resolution

Relative sender clauses («ایمیل‌هایی که account به من فرستاده») retain the
named sender for combined count/body requests. Before execution, an independent
check rejects explicit account tokens dropped by the compiled/model plan. Each
filtered search page must also report the expected SMTP sender for every row;
a missing or mismatched sender address stops counting/body retrieval rather than
presenting unrelated Inbox messages. This check does not prove understanding of
all natural-language clauses.

Live full-chat acceptance of the user's relative-sender/count/body request on
2026-11-21 found seven matching emails and read all seven bodies in about eight
seconds, without model generation. All searched rows had the requested sender.
Regression suite: 88 tests, including dropped-account and wrong-sender rejection.

Count-and-list requests return the count and selected mail cards together. A
standalone list follow-up such as «و چیاست» reuses the last successful list/count
snapshot in the same conversation, explicitly labelled as previous results. It
does not regenerate a plan or repeat a mailbox search. New Outlook requests clear
this snapshot before executing, and fresh conversations cannot access it. Cards
are limited to 10 with the displayed count stated separately; incomplete scans
are labelled as minimum counts. Synthetic card selection and cached displays are
excluded from the UI's real-tool-call count.

Live acceptance on 2026-11-21 verified combined count/list and count → follow-up
against real Outlook: seven matching messages, seven cards, and no new model/tool
call for the standalone follow-up. Regression suite: 85 tests.

Possessive identity requests such as «ایمیل‌های Some Full Name را پیدا کن» or
«ایمیل‌های some.account را پیدا کن» preserve the entire Latin name/account as
one identity slot. Without an explicit direction they search received mail from
that identity in Inbox; an explicit sent folder selects the recipient instead.
Explicit From/To/CC clauses take precedence. The result states the actual folder
and resolved address filter. This path needs no model generation and does not
contain a company employee lookup table. Unknown/ambiguous directory entries
still require clarification. Live verification of the user's full-name and
local-account inputs returned the same actual SMTP sender and seven Inbox
messages for each input on 2026-11-21. Regression suite: 82 tests.

A full SMTP address is passed directly into the sender/To/CC search filter;
no prior message or local identity cache is required. Zero matches means zero
matches in the searched scope, not an unresolved person. Display names/accounts
use outlook_resolve_person (read-only MAPI CreateRecipient/Resolve) against the
user's Outlook Address Book, without enumerating/exporting the GAL. Unresolved
or ambiguous names can fall back to observed-message identities or clarification.

SMTP extraction supports PR_SENDER_SMTP_ADDRESS, Exchange user and Exchange
distribution list PrimarySmtpAddress, then address-entry SMTP properties. The
fix is applicable to organizational/shared/group identities as well as people.
The tool never creates/sends a mail item or changes the Address Book.

Real chat verification: both full-address and display-name queries for the
observed HC-RE mailbox returned the same one message (~2s each). Test suite: 75.
“Important” is not inferred from a sender's identity; actual Outlook Importance
metadata is reported separately from unverified semantic importance.
