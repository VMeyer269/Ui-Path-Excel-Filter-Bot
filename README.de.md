# UiPath Excel Filter Bot (Invoke Code, C#)

Vergleicht zwei monatliche Excel-Datensaetze in UiPath (`dtAlt`, `dtNeu`) und erzeugt eine Aenderungstabelle (`dtChanges`) mit einer einzigen `Invoke Code`-Activity in C#.

English version: `README.md`

## Repository-Inhalt

- `uipath/InvokeCode_ExcelAbgleich.cs` - produktionsreifes Invoke-Code-Snippet
- `uipath/README_UiPath_ExcelAbgleich.md` - detaillierte technische Doku (Deutsch)

## Was erkannt wird

- `NEU`: Bilanzkreis existiert nur im neuen Monat
- `WEG`: Bilanzkreis existiert nur im alten Monat
- `RLM_NEU` / `RLM_WEG`: RLM kommt innerhalb bestehender Bilanzkreise hinzu/faellt weg
- `SLP_NEU` / `SLP_WEG`: SLP kommt innerhalb bestehender Bilanzkreise hinzu/faellt weg
- `WECHSEL_TAGES_STUNDEN`: Wechsel zwischen Tages- und Stundenregime

## Technische Highlights

- Robustes Zeilenkopieren bei unterschiedlichen Schemas (kein `ItemArray`-Mismatch)
- Case-insensitive Typ-Erkennung (`RLM`/`SLP`)
- Optionale Deduplizierung und deterministische Sortierung
- Direkt in UiPath `Invoke Code` mit `DataTable`-Argumenten nutzbar

## Quick Start (UiPath)

1. `Invoke Code`-Activity einfuegen (`CSharp`).
2. In-Argumente:
   - `dtAlt` (`System.Data.DataTable`)
   - `dtNeu` (`System.Data.DataTable`)
3. Out-Argument:
   - `dtChanges` (`System.Data.DataTable`)
4. Code aus `uipath/InvokeCode_ExcelAbgleich.cs` einfuegen.
5. Namespaces sicherstellen:
   - `System`
   - `System.Data`
   - `System.Linq`
   - `System.Collections.Generic`

## Benoetigte Input-Spalten

- `BILANZKREIS`
- `FALLGRUPPE`

## Hinweise

- Spalten mit Prefix `Unnamed` werden automatisch entfernt.
- Regime-Erkennung basiert auf Textmustern in `FALLGRUPPE`.
- Vor Veroeffentlichung sensible Unternehmens- oder Kundendaten anonymisieren.
