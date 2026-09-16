# Agent team — role definitions and playbooks

This folder defines a small team of specialist agents that run the daily
operation under human control. The definitions are plain markdown so they
can be loaded into any agent runtime — a local, self-hosted model included.
Nothing here depends on a particular vendor.

## Operating model

The **orchestrator** is the session the analyst types into. It routes each
ask to one specialist, collects the result, and is the only party that asks
the human questions or takes an approval-gated action. Specialists never
talk to the human directly.

| The human says | Route to | Then the orchestrator |
|---|---|---|
| update the roster / run the refresh / what failed | `morning-review` | runs apply after human OK |
| change X on the roster / who changed / add these people | `wave-file-steward` | runs the apply command after human OK |
| why is this person…, what does a status mean, which source is right | `training-registration-expert` | relays the finding |
| today's numbers / how many / by leader / pull a list | `data-analyst` | relays the tables |
| build / rebuild / fix a tracker, census, dashboard, workbook | `report-builder` | confirms before any NEW file is published |
| add a column / new view / script errors / change the pipeline | `pipeline-engineer` | reviews the diff; nothing runs against live data |

## The stop-and-return contract

Specialists cannot ask the human anything. When a specialist reaches an
action on its "Never" list, or would overwrite a hand-entered value, it stops
and returns:

```
DECISION NEEDED
- what: <one line>
- evidence: <path / query / rows>
- recommended: <one line>
```

The orchestrator puts that in front of the human and acts on the answer.
Apply steps, roster writes, overwrites of manual edits and first-time
publishes are always orchestrator actions taken after an explicit OK.

## Files

- `roles/` — one charter per specialist: when to use, allowed and denied
  tools, what it owns, what it never does, its output shape, and the
  stop-and-return contract.
- `playbooks/` — human-triggered procedures the orchestrator runs:
  `start-day` (morning brief) and `end-day` (wrap-up).

## Adapting to your runtime

1. Load each role file as that agent's system prompt.
2. Map **Allowed tools** / **Denied tools** onto your framework's tool
   permissions. Read-only roles must not be able to write files.
3. Give every role the same "operations handbook" as preloaded context:
   data roots, pipeline step order, the list of scripts allowed to write the
   roster, and output style. That document is org-specific and is not in
   this repository.
4. Put a **pre-write guard** in front of file writes and shell commands. Ours
   denies any assistant-tooling file or AI wording bound for the shared
   workspace, and asks before a roster write that is not one of the
   allowlisted writer scripts. The implementation is runtime-specific and not
   included; the rules are described in `docs/ENGINEERING_GUIDE.md`.
5. Keep approvals in the orchestrator. If your runtime lets specialists ask
   the human directly, disable it; the contract above is simpler to audit.
