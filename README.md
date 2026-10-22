# Identity Reconciliation

A Google Sheets + Apps Script tool that pulls account rosters from several
identity systems, builds one row per person, and reports where those systems
disagree with each other.

It answers a question most organizations cannot answer quickly: **when someone
leaves, do they actually lose access everywhere?**

**This tool is read-only.** It observes and reports. It never provisions,
deprovisions, or writes back to any connected system — the directory stays the
source of truth, and a human decides what to do about each finding.

## The problem

Offboarding is rarely one action. It is a disable in the primary directory, a
license removal in a second tenant, a seat revoked in a SaaS product, a
contract marked ended in an HR or student system. Each step lives with a
different person, and the ones that get skipped are invisible — because
nothing in any individual system looks wrong. The account that still works is
only anomalous *relative to the others*.

This joins them and names the discrepancies.

## What it reports

| Severity | Finding |
|---|---|
| High | Disabled in the primary directory but still enabled in Microsoft |
| High | Disabled in the primary directory but still assigned to services |
| High | Enabled account in a secondary system with no active person record |
| High | Roster says inactive but accounts remain active |
| Medium | Service assignment with no active person record |
| Medium | Org unit implies a license that is not assigned |
| Low | Enabled account dormant past a configurable threshold |

Rules are specific and named rather than a generic field-by-field diff. A
diff-everything pass over four systems produces thousands of meaningless
discrepancies and gets ignored within a week; each rule here describes a
situation somebody should actually act on.

## Providers

| Provider | Source | Auth |
|---|---|---|
| `GoogleWorkspaceProvider` | Admin SDK Directory API | The executing account (read-only scope) |
| `MicrosoftGraphProvider` | Microsoft Graph `/users` | App registration, client credentials |
| `RosterSheetProvider` | An exported Google Sheet | Viewer access to that sheet |
| `CsvTabProvider` | Pasted CSV tabs | None |

Two of these deserve a note.

**The roster provider reads an exported Sheet rather than calling a roster API
directly.** Roster vendors tend to have rate-limited, awkwardly paginated, or
paid APIs — and most of them will already export to a Sheet on a schedule.
Pointing at that export costs one Script Property, consumes no API quota, and
means this tool needs no credentials for the roster system at all.

**The CSV tab provider exists because plenty of SaaS products cannot be read
from.** Many expose a SCIM endpoint that is write-only by design, built for an
identity provider to push provisioning *into* the product rather than for a
third party to read users back out. Others offer a console CSV export and
nothing else. Rather than leave those systems out of the reconciliation
entirely — which is how orphaned licenses go unnoticed for years — paste the
export into a tab named `Import: <ServiceName>` and it joins the report. Only
an email column is required, located by header name rather than position, so
the vendor's own export format works unmodified.

## Setup

1. Create a Google Sheet and open **Extensions → Apps Script**.
2. Copy the files from `src/` into the script project (or push with
   [clasp](https://github.com/google/clasp)).
3. Enable the **Admin SDK API** advanced service. The executing account needs
   a role with User Read privileges.
4. Add secrets under **Project Settings → Script Properties**:

   | Key | Needed for |
   |---|---|
   | `MS_TENANT_ID` | Microsoft |
   | `MS_CLIENT_ID` | Microsoft |
   | `MS_CLIENT_SECRET` | Microsoft |
   | `ROSTER_SHEET_ID` | Roster |

5. Reload the sheet, then run **Reconciliation → First-time setup**.
6. Run **Reconciliation → Run full refresh**.
7. Optionally **Install 6-hour refresh**.

Unset providers are skipped with a log line. A deployment with only Google
Workspace configured still produces a working Master and Findings pass.

## Configuration

Secrets live in Script Properties. Everything else lives in a **Settings** tab,
so whoever tunes the rules does not need access to credentials:

| Setting | Meaning |
|---|---|
| `anchorSource` | Which provider is the authoritative list of people |
| `disabledOuPattern` | Regex for org unit paths that mean "disabled" |
| `rosterInactivePattern` | Regex for roster statuses that mean "gone" |
| `dormantAfterDays` | Days without login before reporting; `0` disables |

Entitlement rules live in a **LicenseRules** tab (`orgPathPrefix`,
`expectedLicense`) and are empty by default. Which org unit implies which
license is a policy decision specific to each organization, and a guessed
default would produce a page of confident, wrong findings on first run.

## Design notes

**Email is the join key, normalized in exactly one place.** Directory systems
disagree about case and whitespace, and one stray capital letter silently
turns one person into two rows.

**Master is anchored, not unioned.** Every Master row comes from the
authoritative source. Records that exist only in a secondary system are not
rows — they are orphans, and they surface in Findings where they belong.

**Active status is determined by exclusion.** Roster systems label the departed
precisely ("Withdrawn", "Terminated") but describe the present tense
inconsistently ("Enrolled", "Current", "FT", blank). Matching the inactive
pattern and defaulting everything else to active is more robust than
enumerating every way a system might say "fine".

**Paging is explicit everywhere.** Directory caps at 500 users per page; Graph
pages with an opaque `@odata.nextLink` that cannot be reconstructed from the
original query. A naive single call against either returns a clean-looking
report that is quietly missing most of the organization.

**Reads and writes are batched.** Apps Script charges a round trip per
`getValue()`/`setValue()`, so a per-cell loop over a few thousand accounts
will exceed the six-minute execution limit.

**Disabled is two different things.** Some directories set a suspended flag;
others park the account in an org unit. Organizations frequently do both, so
the model treats either as disabled.

## Limitations

- Google Workspace license data requires the Licensing API, which is not wired
  up — the column exists and is left blank rather than guessed.
- Matching is by email only. Organizations where a person's address differs
  across systems need a secondary key (an employee id carried in
  `externalId`), which the schema accommodates but no rule currently uses.
- Apps Script's six-minute execution limit applies. Very large directories may
  need per-provider runs on separate triggers rather than one full refresh.

## License

MIT
