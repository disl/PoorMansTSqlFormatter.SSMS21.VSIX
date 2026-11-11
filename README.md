# Poor Man's T-SQL Formatter for SSMS 22

A Visual Studio extension (VSIX) that integrates the [Poor Man's T-SQL Formatter](https://github.com/TaoK/PoorMansTSqlFormatter) into **SQL Server Management Studio 22** (VS 2026 shell).

It reformats T-SQL code cleanly and **preserves comments** (`--` and `/* */`), and exports query results to Excel **with correct data types** (numbers stay numbers, dates stay dates – independent of the regional settings).

## Features

- **Docked toolbar** "Poor Man's T-SQL Formatter" below the menu bar (between "Standard" and "SQL-Editor")
  - `Format Selection`
  - `Format Whole Document`
  - `Formatting Settings…`
  - `Export to Excel` *(new in 1.5)*
- **Type-safe Excel export** of the grid results (see [Export to Excel](#export-to-excel))
- Custom toolbar icons
- Keyboard shortcuts
- Comment-preserving, configurable formatting (indentation, keyword casing, etc.)
- **Localized UI** — German texts on German systems (`de-DE`, `de-AT`, `de-CH`, …), English everywhere else (toolbar/menu captions, message boxes, settings dialog)
- Per-user install — no admin rights required

> **SSMS 22.12 users:** version 1.4 crashes SSMS a few seconds after formatting (stack overflow caused by a self-referencing menu in SSMS 22.12). **Please update to 1.5.**

## System requirements

- SQL Server Management Studio 22 (SSMS 22)
- Windows 10 / 11

## Installation

Download the latest `.vsix` from the [Releases](https://github.com/dislhaas/dislPoorMansTSqlFormatter-SSMS22/releases) page.

1. Close SSMS.
2. Double-click the `.vsix` file, then click **Install**.
3. Start SSMS.

The toolbar "Poor Man's T-SQL Formatter" appears docked below the menu bar. It can be toggled via **View > Toolbars**.

> **Important:** On the very first start, click one menu item under **Tools > Poor Man's T-SQL Formatter** (e.g. "Formatting Settings…") once. This activates the keyboard shortcuts permanently.

## Usage

| Command | Description |
|---|---|
| Format Selection | formats only the selected text |
| Format Whole Document | formats the whole active script |
| Formatting Settings… | opens the options (indentation, casing, …) |
| Export to Excel | exports all grid results of the active query window to an `.xlsx` file and opens it |

| Shortcut | Action |
|---|---|
| `Ctrl+Alt+F` | Format Selection |
| `Ctrl+Alt+D` | Format Whole Document |
| `Ctrl+Alt+K` | Formatting Settings |
| `Ctrl+Alt+E` | Export to Excel ¹ |

¹ Only assigned if `Ctrl+Alt+E` is not already used by another SSMS command (checked once on first load). Otherwise assign it yourself via **Tools > Options > Environment > Keyboard** (search the command list for `Excel`).

## Export to Excel

Copying grid results into Excel or saving them as CSV causes well-known problems: decimal point vs. comma (`1.5` becomes a date or text in a German Excel), dates arrive as text, leading zeros of article numbers are lost, long numbers turn into `1,23457E+18`.

**Export to Excel** avoids all of this: it does not export the *displayed text*, but the *typed values* SSMS keeps internally, and writes them as real number/date/text cells into a native `.xlsx` file. Numbers and dates are stored culture-independently, so Excel displays them in the user's regional format (e.g. `1.234,56` and `31.12.2026` on German systems).

- All result sets of the active query window are exported, one worksheet each (`Result 1`, `Result 2`, …). More than 1,048,575 rows are continued on additional sheets.
- Header row bold, frozen, with AutoFilter; column widths fitted.
- The query is **not** executed again – exactly what is in the grid is exported.
- The file is written to `%TEMP%\SSMS-Excel-Export\SSMS_Export_<timestamp>.xlsx` and opened with the program registered for `.xlsx` (usually Excel). Use *Save As* in Excel to keep it.
- Large exports run in the background with a progress dialog and can be cancelled.
- No dependencies: no Excel interop, no Open XML SDK – Excel is only needed to *open* the file.

| SQL Server type | Excel cell |
|---|---|
| `tinyint`, `smallint`, `int`, `bigint` (≤ 15 digits) | number, format `0` |
| `decimal` / `numeric(p,s)` | number with *s* decimal places |
| `money`, `smallmoney` | number, 2–4 decimal places |
| `float`, `real` | number (exact round-trip, `real` without float artifacts such as `0.100000001`) |
| `bit` | number `0` / `1` (as shown in the grid) |
| `date` | date (short date format of the current culture) |
| `datetime`, `datetime2`, `smalldatetime` | date + time (milliseconds when present) |
| `time` | time |
| `datetimeoffset` | text incl. offset (Excel has no time zones) |
| `uniqueidentifier` | text (upper case, as in SSMS) |
| `binary`, `varbinary`, `image`, `rowversion` | text `0x…` |
| `char`, `varchar`, `nchar`, `nvarchar`, `xml`, … | text – never interpreted as number, date or formula (`00123`, `=…` stay as they are) |
| `NULL` | empty cell |

**No silent data loss** – the following cases are exported as text and reported after the export:

- numbers with **more than 15 significant digits** (`bigint` IDs, `decimal(38,…)`): Excel only keeps 15 digits and would otherwise change the value,
- dates **before 1900-03-01** (not representable in Excel),
- texts longer than **32,767 characters** are truncated (Excel cell limit).

Requires **Results to Grid** (`Ctrl+D`); with *Results to Text/File* there is nothing to export.

## Build from source

Required: **Visual Studio 2026** (v18.0 — the same shell SSMS 22 is built on) with the **Visual Studio extension development** workload. The project uses `Microsoft.VisualStudio.SDK` and `Microsoft.VSSDK.BuildTools` via NuGet.

Command line:

```
msbuild PoorMansTSqlFormatter.SSMS21.VSIX\PoorMansTSqlFormatter.SSMS21.VSIX.csproj /t:Restore /p:Configuration=Release /p:VisualStudioVersion=18.0
msbuild PoorMansTSqlFormatter.SSMS21.VSIX\PoorMansTSqlFormatter.SSMS21.VSIX.csproj /t:Build   /p:Configuration=Release /p:VisualStudioVersion=18.0
```

(`VisualStudioVersion=18.0` is required so the VSSDK BuildTools resolve the VS 18 shell.)

The resulting `PoorMansTSqlFormatter.SSMS21.VSIX.vsix` is in `PoorMansTSqlFormatter.SSMS21.VSIX\bin\Release\`.

> **Build gotchas:**
> - **Restore and Build separately.** The combined `/t:Restore;Build` target list puts MSBuild into restore mode, which skips the NuGet `VSToolsPath` property — the build then falls back to the VS-installed VSSDK targets and `CreatePkgDef` fails with a `Microsoft.VisualStudio.Shell.Framework 18.0` resolution error. Run the two commands above in sequence.
> - **Stale manifest.** When bumping the version number, delete `obj\Release\extension.vsixmanifest`, `bin\Release\extension.vsixmanifest` and the old `.vsix` before rebuilding — otherwise the old version number lands in the VSIX and the installer refuses the update.

## Uninstall

- **Option A:** Open **Extensions > Manage Extensions**, select the extension, click **Uninstall**, restart SSMS.
- **Option B:** Close SSMS and run the VSIX uninstaller:
  ```
  "%ProgramFiles%\Microsoft SQL Server Management Studio 22\Release\Common7\IDE\VSIXInstaller.exe" /q /u:PoorMansTSqlFormatter.SSMS21.VSIX.93aae78f-16c6-49f1-92ff-d53a474d56b9
  ```

## Changelog

### 1.5
- **New:** `Export to Excel` button – type-safe export of all grid results to `.xlsx` (see [Export to Excel](#export-to-excel)).
- **Fix:** SSMS 22.12 crashed a few seconds after formatting (stack overflow `0xC00000FD`). The German caption localization walked all SSMS menus recursively; SSMS 22.12 contains a self-referencing menu (*Build > Build*), which caused endless recursion. Captions are now set directly on the commands.

### 1.4
- Docked toolbar, custom icons, keyboard shortcuts, localized UI (German/English).

## License

**AGPL-3.0** — see [LICENSE.txt](LICENSE.txt).

This project is a modified, SSMS-22-specific build of the [Poor Man's T-SQL Formatter](https://github.com/TaoK/PoorMansTSqlFormatter) by [Tao Klerks](https://github.com/TaoK) (originally authored by Joseph Shakely). The formatter engine (`PoorMansTSqlFormatterLib`, `PoorMansTSqlFormatterLibShared`) is unmodified upstream code; the SSMS 22 VSIX integration is the contribution of this repository.

*LinqBridge.dll* is included under its [BSD-style license](PoorMansTSqlFormatterLib/References/LinqBridge/COPYING.txt).
