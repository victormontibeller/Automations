# automation-core

Small, installable Google integrations for independent Python automations.
Requires Python 3.11+. No application configuration discovery, business rules,
Excel parsing, scheduler integration, delivery ledger or logging framework.

## Install and test

This package is local to this repository and **is not published on PyPI**. Install
it into each consuming application's own virtualenv. From the repository root:

```bash
# Cards environment, regular installs of both local packages:
cartoes/.venv/bin/python -m pip install ./libs/automation_core ./cartoes
cartoes/.venv/bin/python -m pip check
cartoes/.venv/bin/python -m unittest discover -s libs/automation_core/tests -v
cartoes/.venv/bin/python -m unittest discover -s cartoes/tests -v
```

An independent consumer can install `./libs/automation_core` alone; the library
has no dependency on `resumos-cartoes`, `openpyxl` or Hermes. Do not share application
virtualenvs or inject sibling source directories into `sys.path`. Reinstall regular
installs after editing source. Development installation is not production deployment.
For the shared pinned snapshot and isolated wheel/editable gates, use the
[operations guide](../../docs/operations.md); rollout remains paused.

## API and ownership

- `google_auth.build_service(api, version, token_file)` takes an explicit token
  path and builds a Google service with discovery caching disabled. It only reads
  the supplied authorized-user credentials; it neither discovers a default token
  nor saves refreshed credentials. The Google transport may refresh credentials
  in memory during use. Failures raise `GoogleAuthError` with a controlled message.
  Token provisioning, scopes and path defaults are the caller's responsibility.
- `gmail.GmailClient(service)` owns the supplied Gmail service and obtains the
  authenticated profile once. `email_address` exposes that profile address;
  applications may impose their own stricter sender policy. Initialization failure
  closes the transport and raises `GmailProfileError`.
- `GmailClient.send_message(message)` takes a caller-composed `EmailMessage`.
  It checks serializability and address headers, but does not choose, deduplicate,
  add or remove recipients. Display names, multiple recipients, and To/Cc/Bcc-only
  recipient lists are supported. Caller-provided Bcc is deliberately preserved
  in the raw MIME handed to Gmail. Subject, content and attachments belong to the
  caller. Invalid input raises `InvalidMessage` before a send is attempted.
- Successful sends return immutable `gmail.SendResult(message_id, thread_id=None)`.
  These are **provider identifiers**, not the message's RFC `Message-ID`, delivery
  receipts or evidence of reading. An absent/invalid provider message ID is an
  uncertain outcome, never success. An absent, empty or non-string optional
  `threadId` becomes `None` without invalidating an acknowledged message.
- HTTP 4xx send failures raise `GmailSendRejected`; transport errors, 5xx failures
  and missing acknowledgements raise `GmailSendUncertain`. API execution uses
  `num_retries=0`: the library never loops, reconnects or resends. Consumers must
  reconcile uncertain sends before deciding whether to retry. An exception is not
  evidence that no message was accepted.
- `drive.DriveClient(service)` owns a supplied Drive service. `root_id()` resolves
  My Drive. `find_folder(parent_id, name)` restricts to folder MIME type;
  `find_file(parent_id, name)` finds any uniquely named item and leaves MIME policy
  to the caller. Both escape query literals, match the exact parent/name, exclude
  trash, follow all pages, return `None` for absence, and raise `AmbiguousItem` for
  multiple matches. They never search alternate paths or choose the first duplicate.
- `DriveClient.download_bytes(file_id)` returns raw media bytes without file-type
  assumptions, conversion or disk writes. Root/list/download errors raise
  `DriveError`. API calls and download chunks use `num_retries=0`.
- Call `close()` in `finally` when finished with either client. Cleanup is idempotent
  and best effort: transport-close errors are suppressed so they cannot replace
  the original failure or turn successful acceptance into a retry signal. Do not
  share the owned service between clients or reuse a client after closing it.

Errors retain their original exception as a cause for debugging. Their public
messages are controlled; consumers should not print chained tracebacks/provider
payloads into operational notifications.

## Reuse without cards or live APIs

`tests/test_gmail.py` sends a synthetic **release notice** through an injected fake
service, verifies the complete serialized MIME, preserves caller To/Cc/Bcc and
checks returned provider IDs. `tests/test_drive.py` finds **design assets** and
returns a synthetic SVG in multiple media chunks. Neither knows about billing
months, participant lists or spreadsheet files. `test_dependency_direction.py`
checks source imports and imports the installed library outside the checkout with
application/Google SDK imports blocked; SDK imports remain lazy.

The cards application owns its Hermes token fallback, required personal Cc,
no-Bcc policy, workbook paths, Excel MIME validation, error translations and ledger.
Adding another consumer should reuse these transports, not copy cards' policies
into this library. No speculative multi-provider backend interface is required.
