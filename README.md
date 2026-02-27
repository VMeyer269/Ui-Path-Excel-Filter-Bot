# UiPath Excel Filter Bot (Invoke Code, C#)

EN: Compares two monthly Excel datasets in UiPath (`dtAlt`, `dtNeu`) and produces a change table (`dtChanges`) using a single `Invoke Code` activity in C#.
DE: Vergleicht zwei monatliche Excel-Datensaetze in UiPath (`dtAlt`, `dtNeu`) und erzeugt eine Aenderungstabelle (`dtChanges`) mit einer einzigen `Invoke Code`-Activity in C#.

German version: `README.de.md`

## Repository Content

- `uipath/InvokeCode_ExcelAbgleich.cs` - production-ready Invoke Code snippet
- `uipath/README_UiPath_ExcelAbgleich.md` - detailed German documentation

## What It Detects

- `NEU`: balancing group exists only in the new month
- `WEG`: balancing group exists only in the old month
- `RLM_NEU` / `RLM_WEG`: RLM appears/disappears within existing balancing groups
- `SLP_NEU` / `SLP_WEG`: SLP appears/disappears within existing balancing groups
- `WECHSEL_TAGES_STUNDEN`: switch between daily and hourly regime

## Technical Highlights

- Safe row-copying across varying schemas (no `ItemArray` mismatch issues)
- Case-insensitive type detection (`RLM`/`SLP`)
- Optional deduplication and deterministic sorting
- Works directly in UiPath `Invoke Code` with `DataTable` arguments

## Quick Start (UiPath)

1. Add an `Invoke Code` activity (`CSharp`).
2. Add In arguments:
   - `dtAlt` (`System.Data.DataTable`)
   - `dtNeu` (`System.Data.DataTable`)
3. Add Out argument:
   - `dtChanges` (`System.Data.DataTable`)
4. Paste code from `uipath/InvokeCode_ExcelAbgleich.cs`.
5. Ensure namespaces:
   - `System`
   - `System.Data`
   - `System.Linq`
   - `System.Collections.Generic`

## Required Input Columns

- `BILANZKREIS`
- `FALLGRUPPE`

## Notes

- Columns starting with `Unnamed` are removed automatically.
- Regime detection is pattern-based and depends on values in `FALLGRUPPE`.
- Remove or anonymize any sensitive company/customer data before sharing.
