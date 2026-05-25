# Office Scripts

TypeScript scripts for Excel Online automation in `presenze_sern_2026_v2.xlsx`.

## Scripts

### aggiornaPresenze.ts
**Called by:** Power Automate weekly flow (every Monday)  
**Parameters:** `jsonTurni` (shifts JSON), `jsonAssenze` (time-off JSON)  
**What it does:**
- Reads shifts and absences from Microsoft Shifts via Power Automate
- Writes to `TabellaPresenze`: OreOrdinarie, Malattia, Permessi, Ferie, CongParentale
- Calculates meal vouchers (BuoniPasto) — €8/day for shifts ≥4h, excludes Sofia Ricci
- Covers current month + previous month simultaneously
- Subtracts 30-min lunch break for shifts >4h

**Key logic:**
- Shifts encode full-day absences starting at 22:00 UTC prior day → `dataEffettiva()` adds 1 day when UTC hours ≥ 20
- Laura Avanzi Monday = 6h (part-time), all others = 8h
- userId → surname mapping hardcoded (see top of script)

### compilaGriglia.ts
**Called by:** Power Automate weekly flow (after aggiornaPresenze)  
**Parameters:** none (reads TabellaPresenze directly)  
**What it does:**
- Reads TabellaPresenze and TabellaRimborsi
- Writes data to monthly grid sheets (e.g. "Maggio 2026")
- Populates: ore ordinarie, assenze, allowance missioni, rimborsi km, altri rimborsi, buoni pasto
- Allowance missione: €77.47/giorno programma, €38.74 (0.5) o €77.47 (1) per giorni viaggio

## How to update scripts in Excel Online

1. Open `presenze_sern_2026_v2.xlsx` in Excel Online
2. Tab **Automatizza** → **Tutti gli script**
3. Click the script name → editor opens
4. Select all (`Ctrl+A`), paste updated code from this repo
5. Save (`Ctrl+S`)

## userId → employee mapping

| userId | Surname |
|--------|---------|
| 04fe1e3f-... | Catellani |
| b1ab83a9-... | Mereu |
| d28776ec-... | Avanzi |
| dfef8343-... | Roli |
| 4b01936b-... | Gravina |
| 89f1124d-... | Ricci |
