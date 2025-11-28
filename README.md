# DOI Generator

A Streamlit application for generating Tract-Based Ownership and Unit-Based Division of Interest (DOI) reports.

## Installation

```bash
pip install -r requirements.txt
```

## Running the App

```bash
streamlit run app.py
```

The app will open in your browser at http://localhost:8501

## Report Types

### Tract-Based Ownership
- Organizes data **by tract** (matching the format of Barb_Tract_Based_DOI.xlsx)
- Each tract has:
  - Tract info box (Tract No., Gross Acres, Legal Description)
  - All owners for that tract
  - TOTALS row
- Generates sheets: Tract List, LORI, NPRI, ORI, WI, Unit Recap

### Unit-Based DOI
- Organizes data **by owner**
- Applies tract allocation factors to calculate UNIT NRI
- Each owner has:
  - Owner name/address box
  - All tracts for that owner
  - TOTAL row with UNIT NRI total
- Generates sheets: LORI, NPRI, ORI, WI, Unit Recap

## Required Files

### Combined Data File (both report types)
Excel file with a "Combined" sheet containing:
- OWNER - Owner name
- TYPE - Interest type (MI, NPRI, ORI, WI)
- TRACT - Tract number
- TRACT NRI - Tract Net Revenue Interest
- DECIMAL INTEREST - Owner's decimal interest
- LEASE NO. - Lease number
- REQ - Requirement number
- LEASE ROYALTY - Lease royalty rate (for MI)
- NPRI BURDENS - NPRI burden amount (for MI)
- NET ACRES - Net acres
- And other interest-specific fields

### Schedule File (Unit-Based DOI only)
Excel file with a "Tract List" sheet containing:
- Tract numbers
- Legal descriptions
- Acres
- Tract Allocation factors

## Output Format

The generated Excel files match the format of existing DOI schedules:
- Times New Roman 10pt font
- Light blue headers (DDEBF7)
- Gray tract/owner info boxes (F2F2F2)
- 8 decimal places for NRI values
- 6 decimal places for acre values
