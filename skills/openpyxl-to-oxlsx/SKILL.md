---
name: openpyxl-to-oxlsx
description: "Trigger: migrate openpyxl to oxlsx, replace openpyxl. Audit compatibility, translate Python workbook code, and verify behavioral parity."
license: MIT OR Apache-2.0
metadata:
  author: alvaroroco
  version: "1.0"
---

# Migrate openpyxl to oxlsx

## Activation Contract

Use when auditing or migrating existing Python openpyxl workflows to oxlsx. Keep audits read-only unless implementation is requested. Do not turn a Python library migration into a Rust rewrite.

## Hard Rules

- Identify both installed versions; the bundled mapping targets oxlsx 0.2.0. Never assume drop-in compatibility or perform a blind import replacement.
- Preserve required workbook behavior, not merely cell values. Never remove unsupported options or formatting to make the migrated code run.
- Retain openpyxl where required features lack a verified equivalent. Ask only when a real tradeoff requires the user's decision.
- Keep originals and dependencies until parity is established; write trial outputs to separate paths. Do not run user scripts against production destinations for comparison.

## Decision Gates

| Evidence | Action |
|---|---|
| Supported API with matching semantics | Translate and test the exercised behavior |
| Different setter, return type, or memory behavior | Adapt explicitly using the mapping |
| Required unsupported or unverified feature | Report the gap; leave that path on openpyxl |

## Execution Steps

1. Inventory imports, wrappers, load flags, cell/style writes, sheet creation, iteration, pandas calls, and workbook features. Inspect representative files; do not infer requirements solely from imports.
2. Read [API mapping](references/api-mapping.md). Classify each exercised behavior as compatible, adaptable, blocked, or unverified before editing.
3. Capture baseline outputs in isolated temporary storage. Follow [parity validation](references/validation.md); use its before/after example only for supported paths.
4. Migrate one coherent path at a time, updating imports and call sites together. Preserve explicit failure behavior and resource limits.
5. Compare original and migrated outputs independently. Remove openpyxl from application dependencies only after all runtime consumers are migrated; it may remain a test oracle.

## Output Contract

Report migrated paths, retained openpyxl paths and reasons, parity checks actually run, and unresolved differences. For blocked migrations, give the exact feature gap and smallest safe next decision; do not claim completion or performance gains without evidence.

## References

- [API mapping and compatibility boundaries](references/api-mapping.md)
- [Before/after example and parity validation](references/validation.md)
