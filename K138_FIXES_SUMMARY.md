# K138 Fixes Summary (2026-02-05)

## Changes Implemented

### 1. Date fields – no duplication, correct semantics
- **notice_date** (header "Avis de saisie"): Set to the **date the letter was generated** (today when the script runs).
- **seizure_date_line** (body): Date of the seizure from SAISIE (DATE / HEURE INTERCEPTION).
- Each date appears once in its designated place.

### 2. lieu_interception (ciblage)
- **Extracted and saved** for analysis in the hidden folder.
- **Not used on the K138 form** – the location field uses a fixed value instead.

### 3. seizure_location – fixed value
- **Always printed** as: `MONTREAL POSTAL FACILITY, ETC LÉO-BLANCHETTE / POSTAL CUSTOMS`
- This is the value shown at the "at" / à location on the form.

### 4. description_item (heroine etc.) → ITEM SEIZED
- `DESCRIPTION DE L'ITEM À SAISIR` is extracted and written to **ITEM SEIZED / MARCHANDISE SAISIE** in the description block.

### 5. SIED value → SEIZURE NUMBER
- SIED # (or BOND ROOM LEDGER #) is used in **SEIZURE NUMBER / NUMÉRO DE SAISIE**.
- Later, the full ICES number (e.g. `3952-25-2325`) can be added when the case is submitted.

### 6. Hidden folder for analysis/aggregation
- All extracted values are saved in `.extracted_data/` at the project root.
- Files: `k138_values_{stem}_{timestamp}.csv` and `saisie_extract_{stem}_{timestamp}.csv`
- Used for analysis and future aggregation with other forms.

### 7. Inventory number – no spaces
- Inventory numbers are normalized with spaces removed (e.g. `AB 123 456 789 CA` → `AB123456789CA`).

---

## CSV fields saved (including hidden folder)

| Field                      | Used on K138 form? | Notes                                      |
|----------------------------|--------------------|--------------------------------------------|
| notice_to                  | Yes                | Address block                              |
| notice_date                | Yes                | Letter generation date (header)             |
| seizure_date_line          | Yes                | Seizure date (body)                        |
| seizure_year_left/right    | Yes                | Year display                               |
| seizure_location           | Yes                | Always fixed value                         |
| lieu_interception          | No                 | Extracted; saved for analysis only         |
| description_inventory      | Yes                | No spaces                                  |
| description_item           | Yes                | ITEM SEIZED content                        |
| description_seizure_number | Yes                | SIED or BOND ROOM LEDGER #                  |
| legal_notice               | Yes                | Cannabis/knife/etc. notice text            |
| seizing_officer            | Yes                | 5-digit officer number                     |
| form_type                  | Yes                | Form type for notice selection             |
