# Architecture

## Scope and deployment boundary

Stages 1–2 introduced the application package and reusable Google clients. Stage 3
adds run results, opt-in operational events and reproducible verification. It does
not add a queue, daemon, reminders, UI, new delivery policy or automatic deployment.
The stacked refactor remains under review; rollout is paused. Production, its
editable installation (where present), private configuration, credentials, delivery
history and scheduler are not changed by work in a separate development worktree.
Do not point a production editable install at development sources. Do not merge,
install into production, copy launchers or resume schedules as part of verification.

## Dependency direction

```text
installed cartoes.main / cartoes command / python -m resumos_cartoes
    -> cli -> service.run -> service.run_result
                 |             |-- operations (RunResult, JSON Lines allowlist)
                 |             |-- config / scheduling
                 |             |-- workbook -> models -> Decimal totals
                 |             |-- rendering / previews (local files)
                 |             |-- delivery / alerts -> gmail adapter
                 |             |-- drive adapter
                 |             `-- ledger / locking
                 |                         |
                 `-------------------------v
                         automation_core
                    google_auth / gmail / drive
                             -> Google SDK (lazy imports)
```

`automation_core` never imports cards, Excel, Hermes, CLI parsing or application
configuration. Each automation installs the library into its **own** virtualenv;
there is no shared virtualenv and no source-path injection. Package metadata keeps
runtime dependencies; the verification snapshot constrains their resolved versions.
The library is local, not published on PyPI: install both local projects for cards.

| Application modules | Ownership |
| --- | --- |
| `models`, `errors` | Report values, exact Decimal calculations, controlled application diagnostics |
| `config`, `scheduling` | Private configuration/default paths, card/period/date validation |
| `workbook` | Excel structure and cell validation; no cached-formula fallback |
| `rendering`, `previews` | Original HTML/text content and local preview cleanup |
| `gmail`, `drive` | Application-specific sender, path/MIME and diagnostic policy |
| `delivery`, `alerts` | Required To/Cc, deduplication, no Bcc, batch sends and owner alerts |
| `ledger`, `locking` | Existing SQLite schema, claims, reconciliation and exclusion |
| `service`, `cli` | One orchestration and the unchanged integer/CLI contract |
| `operations` | Small immutable result and validated, best-effort event serialization |

There is no duplicate logging orchestrator. `run()` is a thin integer-returning
wrapper over `run_result()`. The installed `cartoes.main` remains the same CLI
function. No module configures global logging or attaches handlers at import time.

## Delivery boundaries

The workbook is read once into a byte snapshot for parsing and its audit hash.
The original financial calculations, source selection, rendering, CLI defaults,
schedule checks and recipient protections are unchanged. Preview/history/resolve
never send email. A Drive-backed preview can still read Drive; offline verification
therefore supplies synthetic **local** input, not just `--dry-run`.

Gmail send calls are not retried. A provider ID is evidence of API acceptance,
not delivery or reading. `SendResult.message_id` is not the RFC `Message-ID` used
by the existing ledger. Provider IDs are not added to the ledger or operational logs.
The optional provider `threadId` becomes `None` if absent, empty or not a string;
a valid message acknowledgement remains successful.

Existing conservative ledger policy is deliberately retained: 4xx rejection is
`failed`; uncertain/network outcomes and post-acceptance persistence failures block
resends as `unknown`. The library's local `InvalidMessage` means no provider send
was attempted, but cards may already have claimed the item. Cards still records
that exception as `unknown`, requiring explicit reconciliation. This is conservative,
not a new automatic retry opportunity. The adapter regression test proves this
without changing delivery code or the SQLite schema.

## Extending with another automation (fake service only)

Create another installable package with its own virtualenv, configuration and tests.
Depend on `automation-core`, not `resumos-cartoes`; leave recipient policy, content
and scheduling in the new application. This example is entirely local:

```python
from email.message import EmailMessage
from unittest.mock import MagicMock
from automation_core.gmail import GmailClient

fake = MagicMock()
fake.users.return_value.getProfile.return_value.execute.return_value = {
    "emailAddress": "robot@example.com"
}
fake.users.return_value.messages.return_value.send.return_value.execute.return_value = {
    "id": "synthetic-ack"
}
message = EmailMessage()
message["From"] = "robot@example.com"
message["To"] = "team@example.com"
message["Subject"] = "Synthetic release notice"
message.set_content("A local fake-service example; no email is sent.")
client = GmailClient(fake)
try:
    result = client.send_message(message)
    assert result.message_id == "synthetic-ack"
    fake.users.return_value.messages.return_value.send.return_value.execute.assert_called_once_with(num_retries=0)
finally:
    client.close()
```

Real service construction, when separately authorized, uses the shared
`google_auth.build_service(api, version, token_file)` with an explicit path supplied
by the caller. No credentials or production defaults belong in reusable examples.
Drive extension tests likewise inject fake services/media downloads; exact-parent,
exact-name pagination and byte transport are shared, workbook layout is not.

See [operations](operations.md) for result semantics, logging and the installation
verification command, and [the shared API](../libs/automation_core/README.md) for
transport error contracts.
