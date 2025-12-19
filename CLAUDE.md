# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

```bash
# Install dependencies
pip install -r requirements.txt

# Run the application
streamlit run app.py
```

The app runs at http://localhost:8501

## Architecture

This is a single-file Streamlit application ([app.py](app.py)) that generates Division of Interest (DOI) Excel reports for oil/gas mineral ownership data.

### Report Types

**Tract-Based Ownership**: Organizes data by tract (land parcel). Each tract shows all owners with their interests, with a TOTALS row per tract. Generates sheets: Tract List, LORI, NPRI, ORI, WI, Unit Recap.

**Unit-Based DOI**: Organizes data by owner. Applies tract allocation factors to calculate UNIT NRI (Net Revenue Interest). Each owner shows all their tracts with a TOTAL row per owner. Requires both a Combined Data file and a Schedule file with tract allocations. The output includes the original Tract List sheet from the Schedule file, followed by LORI, NPRI, ORI, WI, and Unit Recap sheets.

### Interest Types
- **MI** (Mineral Interest) → LORI sheet (Landowner Royalty Interests)
- **NPRI** → Non-Participating Royalty Interests
- **ORI** → Overriding Royalty Interests
- **WI** → Working Interests

### Key Functions

- `load_combined_data()`: Parses the Combined Excel file, finds the data sheet, validates required columns (OWNER, TYPE, TRACT)
- `load_tract_allocations()`: Extracts tract allocation factors from Schedule file's "Tract List" sheet
- `create_tract_based_workbook()`: Generates the Tract-Based report organized by tract
- `create_unit_based_workbook()`: Generates the Unit-Based report organized by owner, applying allocation factors
- `normalize_tract()`: Standardizes tract identifiers (e.g., "1.0" → "1")
- `tract_sort_key()`: Natural sorting for tract numbers (handles both numeric "1, 2, 10" and text "Oram 1, Oram 2")

### Excel Output Formatting

Reports use consistent styling defined in the workbook creation functions:
- Times New Roman 10pt font
- Light blue headers (`DDEBF7`)
- Gray tract/owner info boxes (`F2F2F2`)
- 8 decimal places for NRI values
- 6 decimal places for acre values
- Landscape orientation, narrow margins, fit-to-width printing

### Data Flow

1. User uploads Combined Data file (required) + Schedule file (Unit-Based only)
2. `load_combined_data()` parses and validates the ownership data
3. For Unit-Based: `load_tract_allocations()` extracts allocation factors
4. Workbook creation function iterates through interest types, creating one sheet per type
5. Each sheet groups data by tract (Tract-Based) or owner (Unit-Based)
6. Unit Recap sheet summarizes NRI totals across all interest types
7. For Unit-Based: validates that total UNIT NRI equals 1.00000000
