# Archived automation scripts

These files were moved out of active automation on **2026-04-22** as part of
the v2.1 Change Order (see `/change_order.md`). They are kept here so the
future Fee Analysis automation (Phase 3+) can revive them without having to
rebuild from scratch.

- **fee_memo.py** — generates a Fee Analysis Memo .docx from a COMPS dataset.
  Was called by `render_proposal_package.render_fee_memo()` (also removed).
- **refresh_comps_cache.py** — pulled comp rows from Jobs 26KS DB +
  OneDrive folder walk + Master-Business-Plan.xlsm into a cached JSON
  blob. Ran on the `excel-comps-sync` scheduled task (now disabled).
- **comps_cache.json** — last known good comps snapshot.

To revive: move these three files back up one level, re-enable
`excel-comps-sync`, and re-wire the `render_proposal_package` pipeline
to call `render_fee_memo` after the contract render.

None of this is in scope for Phase 1 / Phase 2 (n8n). Fee automation is
a separate future workstream per the change order.
