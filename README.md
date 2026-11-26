# Excel Automation for Codex

Use Codex to automate desktop Microsoft Excel on Windows through Python, xlwings, and Excel's native COM API. The skill covers workbook edits and copies, editable Power Query M, worksheet and Data Model loads, PivotTables, charts, slicers, and PDF exports.

## Requirements

- A signed-in Windows desktop session with Microsoft Excel installed.
- Codex, Git, and a Windows Python environment available in that session.
- Permission to open the target workbooks and access any query data sources.

The included scripts require Windows Excel COM, even if you have Excel with Power Query on another platform.

## Install for your Windows user

Run this in PowerShell. If the skill is already installed from this repository, use the update command below.

```powershell
$skills = Join-Path $HOME '.agents\skills'
$skill = Join-Path $skills 'excel-automation'
New-Item -ItemType Directory -Path $skills -Force | Out-Null
git clone https://github.com/brandonsoccer22/excel-automation.git $skill
```

Use the same Windows Python interpreter for dependency installation, preflight, and workbook commands. If you use a virtual environment, activate it first or set `$python` to its full `python.exe` path.

```powershell
$python = (Get-Command python.exe -CommandType Application).Source
& $python -m pip install -r (Join-Path $skill 'requirements.txt')
& $python (Join-Path $skill 'scripts\preflight.py')
```

The preflight prints the Python executable and install prefix, checks that xlwings and pywin32 load, and reports the version of Excel started through COM. If it fails, use the reported stage, path, and error to diagnose the issue. [Setup and troubleshooting](SKILL.md#setup-and-com-preflight)

## Use in Codex

Start a new Codex chat and invoke `$excel-automation`. Give Codex full workbook paths and say how you want it to save the result. For example:

> Use $excel-automation to open C:\Reports\Sales.xlsx, refresh the Sales Summary query, and save the result as C:\Reports\Sales-reviewed.xlsx. Reopen the new workbook and verify the saved M formula, worksheet load destination, and refreshed totals.

## Run the smoke test

From the same PowerShell session, run:

```powershell
& $python (Join-Path $skill 'scripts\smoke_test.py') $env:TEMP
```

The script creates `original.xlsx` and `revised.xlsx` in a new folder under the existing temp directory. It creates a native Power Query, loads and refreshes it, modifies its M, and reopens the output to check the saved query, native connection, and calculated results. A successful run prints `PASS` and the revised workbook path.

## Update

```powershell
git -C (Join-Path $HOME '.agents\skills\excel-automation') pull --ff-only
```

## Skill references

- [Skill instructions](SKILL.md)
- [Power Query authoring, loading, and verification](references/power-query.md)
- [Workbook automation](references/workbooks.md)
