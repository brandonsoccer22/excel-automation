---
name: excel-automation
description: Automate Microsoft Excel desktop on Windows with Python, xlwings, and pywin32. Use for native workbook editing, editable Power Query M, worksheet and Data Model loads, refresh verification, PivotTables, slicers, charts, dashboards, and PDF exports.
---

# Excel automation

Run in a signed-in Windows desktop session with desktop Excel installed. Use Windows Python in the same session as Excel, including when Codex is orchestrating an authorized remote Windows host. Install missing dependencies into the selected environment using `python -m pip install -r requirements.txt` from this skill directory. Check Excel availability and feature support on the actual host. Report the missing runtime if Windows COM is unavailable.

Use xlwings for workbook, sheet, range, table, and chart operations. Use its native `.api` objects and pywin32 COM for Excel-specific operations. Use Microsoft Learn's [Excel object model](https://learn.microsoft.com/en-us/office/vba/api/overview/excel/object-model) and [Power Query M reference](https://learn.microsoft.com/en-us/powerquery-m/) as authorities. Keep workbook changes in Excel's object model. Use only programmatic automation through that API. Do not use screenshots, mouse clicks, keyboard simulation, or computer-use tools.

## Workbook ownership

- Resolve each input and output to a full path. Compare Excel's `Workbook.FullName` with the requested path. Inspect matching open workbooks in accessible Excel instances, including their `Saved` and `ReadOnly` state.
- Prefer an isolated, script-owned Excel instance for unattended work. If the target is already open, use the exact open workbook when its live state is intended, or stop and resolve any conflict. Never close a user-owned workbook or discard its changes. Save user-owned workbooks only when requested and account for pre-existing edits. A native `SaveCopyAs` can capture the open state without changing the active workbook's filename.
- Preserve original files unless explicitly asked to save in place. Refuse existing output paths by default. For unattended new outputs, save/export to a private staging directory under the output parent, release the staged workbook, then publish via Windows `os.rename` so a concurrent destination collision fails. Match the format and extension, including macro-enabled formats. Verify the final path by reopening it.
- Keep Trust Center, Protected View, macro, credential, and privacy protections intact. Review query sources and refresh scope before connecting to external data. Open with link updates disabled unless intended. Keep Excel alerts enabled and treat security prompts, authentication prompts, locks, and unsupported providers as blockers.
- Save and verify before closing script-owned workbooks. On failure discard changes only to disposable workbooks or copies that the script owns. Never terminate a shared Excel process.

## Choose the workflow

- For editable queries, query lifecycle, worksheet loads, or Data Model loads, read [Power Query](references/power-query.md). Reuse [power_query.py](scripts/power_query.py) for query and load primitives.
- For formulas, names, connections, PivotTables, charts, slicers, dashboards, and PDF see [Workbook automation](references/workbooks.md).
- For a native Windows acceptance run, execute `python scripts/smoke_test.py "C:\\absolute\\path\\to\\test-parent"` after dependency setup. The parent must exist. The script creates a unique folder, tests creation and repeated refresh, and reopens the new filename to verify persisted M, a native query table, and expected data. Record the Excel version and observed results. A syntax check or mocked run cannot establish Excel compatibility.
