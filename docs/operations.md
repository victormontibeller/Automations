# Operations and safe verification

The refactor rollout is **paused pending review and explicit deployment approval**.
These commands are for a separate development worktree, never the production
checkout. Development approval does not authorize Gmail/Drive access, changes to
private configuration/history, scheduler changes, merges or production installation.
Historical deployment instructions in the cards guide are not a rollout instruction.

## Reproducible verification (recommended)

Prerequisites: Linux/macOS, Git, Python 3.11+ with `venv` and pip 22.3+ available in
the specified interpreter, and an existing scratch directory outside the checkout.
Use your development environment's interpreter if the system Python has no pip.
From the **development repository root**:

```bash
# Set these to the development interpreter and an existing scratch directory.
PYTHON="$PWD/cartoes/.venv/bin/python"
SCRATCH="${TMPDIR:?Set TMPDIR to a disposable directory outside this checkout}"
"$PYTHON" tools/verify.py --python "$PYTHON" --scratch "$SCRATCH" --mode both
```

The stdlib-only driver:

1. Copies an explicit public file allowlist to scratch: immediate Python sources
   and tests in named code directories, plus named build/docs/example/fixture assets.
   It does not recurse into unknown directories (including nested secrets or venvs),
   and never copies runtime `config.json`, `.env`, tokens, input workbooks, outputs
   or delivery history. Linked files/trees are excluded, including linked ancestors;
   a symlinked source root or ancestor is rejected. New public packages/assets must
   be added to the allowlist. Keep the checkout unchanged while it is snapshotted.
2. Creates independent clean virtualenvs; it does not install into the supplied
   interpreter, touch production or change the development venv's installation mode.
3. Installs the pinned runtime and build/bootstrap closure from
   `requirements/verification.txt`. It uses the same snapshot as CI; dependencies
   remain declared in both `pyproject.toml` files. Only this bootstrap/install phase
   may use the network. Review dependency updates as a new snapshot, not a silent
   upgrade. Pins provide a version snapshot, not a hash-locked/offline artifact store.
4. Builds **both local wheels**, then installs them with `--no-index --no-deps` into
   the clean wheel environment. Separately installs **both local packages editable**
   into another clean environment. Builds occur only in the scratch source copy;
   build isolation is disabled after installing the pinned build tools.
5. Runs `pip check`, verifies the exact installed dependency closure, import origins
   and editable metadata, and invokes module/console CLIs outside the checkout.
6. Installs a disposable site-local network-denying audit hook, including for child
   interpreters, then runs both unittest suites outside the checkout. The suites
   include architecture/dependency direction, original regressions, complete golden
   HTML/text comparisons, synthetic workbook previews and operational contracts.
   Private-workbook tests are skipped because runtime inputs were never copied.
7. Runs `git diff --check` on the development worktree, then removes its temporary
   environments, source copies, preview files and synthetic ledgers.

Use `--mode wheel` or `--mode editable` to run one gate. CI runs both modes on
Linux Python 3.11/3.12 and macOS Python 3.12. A local pass on a different interpreter
is **not** evidence that the remote matrix has passed. Review the actual CI results
before approving deployment. Network denial is a regression guard, not an OS sandbox.

## Direct development installation and tests

For ordinary development in this worktree's independent environment, the equivalent
pinned local installation is:

```bash
# Run only in the development checkout; both packages are repository-local.
cartoes/.venv/bin/python -m pip install -r requirements/verification.txt
cartoes/.venv/bin/python -m pip install --no-build-isolation \
  -c requirements/verification.txt ./libs/automation_core ./cartoes
cartoes/.venv/bin/python -m pip check
cartoes/.venv/bin/python -m unittest discover -s libs/automation_core/tests -v
cartoes/.venv/bin/python -m unittest discover -s cartoes/tests -v
```

After source edits, reinstall **both** regular packages before testing. For editable
work only, replace the second command's two local paths with
`-e ./libs/automation_core -e ./cartoes`. A production editable install still points
to its own checkout; it must never be repointed to this development worktree.
Direct tests retain the optional private-workbook behavior of the original suite;
the isolated gate above deliberately excludes it and enforces no network.
No command copies or rewrites runtime configuration or credentials.

## Structured per-run API

`resumos_cartoes.service.run_result(config_path, **options)` accepts the same run
options as `run()`, plus `log_json=False`, and returns immutable `RunResult`.
It still prints the original human diagnostics/output. `run()` and all CLI entry
points return the result's integer `exit_code` (0 success, 1 handled failure).
CLI argument parsing errors retain argparse's exit code 2; `--help` and parsing
errors occur before orchestration and do not emit run events.

| Field | Meaning |
| --- | --- |
| `run_id` | New UUID4 for each invocation; correlation only, not a delivery/idempotency key |
| `status` | `succeeded` or `failed`; not a delivery/read receipt |
| `mode` | `preview`, `send`, `test`, `history` or `resolve`; test means `--send --test-to` |
| `card`, `month` | Allowlisted card and valid calendar `YYYY-MM`, or null when unavailable/invalid |
| `sent` | Messages accepted in a successfully completed batch, including test messages; not confirmed final delivery |
| `skipped` | Previously accepted ledger entries skipped in a completed batch |
| `no_movement` | Selected report blocks without purchases/payments, once reports were successfully read/filtered |
| `previews` | Report previews written by a completed preview write; not the index file count |
| `elapsed_seconds` | Nonnegative monotonic duration through processing and any owner-alert attempt |
| `error_category` | null, `report_error` or `operational_error`; never exception text/type |
| `alert_status` | `not_requested`, `accepted` or `failed`; alert failure never masks the original error |

**Null counts are intentional.** Preview and maintenance runs have null sent/skipped
counts (not applicable). Maintenance has null preview/no-movement counts. A failed
batch reports null sent/skipped: earlier messages may already be accepted, including
when ledger persistence fails after an acknowledgement. Do not interpret null as
zero, subtract it to calculate a resend list, or retry from these logs. Consult the
existing ledger and Gmail through the existing authorized reconciliation workflow.
A completed empty batch reports zero sent/skipped. Preview counts stay null if the
write fails, since some files may already exist. No-movement remains available after
successful parsing/filtering even if a later send fails. Counts exclude owner alerts.

Ordinary exceptions follow the existing controlled-error and single owner-alert
policy. Interruption/process termination is not caught or disguised as a completed
result; it may leave only a start event. This result API is not a persistent audit
store and does not replace SQLite claims, locks or reconciliation.

## Opt-in JSON Lines

Add `--log-json` to an existing command. For a **synthetic local** configuration:

```bash
cartoes/.venv/bin/python -m resumos_cartoes \
  --config "$SCRATCH/synthetic/config.json" --card black --month 2026-09 \
  --dry-run --log-json
```

The command above assumes you created a matching synthetic local workbook/config;
the verification gate's installed-preview test creates them automatically. `--dry-run`
is not a network-disable flag: never use a production/Drive config for offline tests.

Successful logging writes one `run_started` and one `run_finished` JSON object to
**stderr**, correlated by `run_id`. Human stdout is unchanged. Human error/alert
lines also remain on stderr for compatibility; the entire stderr stream is therefore
not exclusively JSON. Collectors must select and validate JSON event lines rather
than treating human diagnostics as privacy-safe telemetry. Default output is unchanged.

Start fields: `event`, `run_id`, `mode`, `card`, `month`. Finish adds the remaining
result fields from the table. A scheduled start has null month until the business
period is validated. A valid requested month is available even if configuration
loading fails. Malformed fields or extra keys reject the entire structured record.
Values are checked as well as keys, including UUID, calendar period, modes/statuses,
fixed error categories, finite durations and nonnegative integer counts.

Structured records never contain recipients, names, subjects, financial values,
configuration/token paths, tokens, raw exceptions, arbitrary exception class names,
provider IDs or stack traces. Logging is best effort: write, flush, serialization
or validation failure cannot change accepted-send results or trigger another send.
Native file-descriptor sinks bypass Python's JSON write buffering, so a closed
stderr pipe cannot leave pending log bytes that change the interpreter's exit code
at shutdown. Custom/in-memory sinks retain their write/flush interface. This does
not redirect stdout/stderr, change human diagnostics or suppress application errors.
There is no global logging setup and no accumulating handlers across repeated calls.
A logging failure can lose an event; the absence of an event never proves no send.

See [architecture](architecture.md) for module boundaries, a runnable fake-client
extension example, and the deliberately unchanged conservative ledger behavior for
local `InvalidMessage` failures.
