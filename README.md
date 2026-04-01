# VHA Proposed Rates Generator

## What it does
| Step | Manual today | Automated |
|---|---|---|
| 1 | Business team cleans raw L&V Excel | ✅ Auto-cleans headers, formats, leading zeros |
| 2 | Build pivot table by benefit/plan | ✅ Auto-groups by division cluster + benefit |
| 3 | Manually enter lives & volumes in Proposed Rates template | ✅ Auto-populates |
| 4 | Apply rate adjustments | ✅ Configurable % per benefit in the UI |
| 5 | Calculate premiums | ✅ Auto-calculated |

---

## Setup

```bash
pip install -r requirements.txt
python app.py
```

Then open **http://localhost:5050** in your browser.

---

## Usage

1. **Upload** the raw `VHA_LivesAndVolumes.xlsx` (first sheet only – "LivesAndVolumes original")
2. **Review** the cleaned data and benefit summary
3. **Adjust** proposed rate changes (%) per benefit if needed
4. **Download** the completed `VHA_Proposed_Rates.xlsx`

---

## Division Groups

| Group Label | Divisions | Plans |
|---|---|---|
| 1,8 | Toronto Office (1) + OPSEU (8) | A, B, E, H, I |
| 2,4,9 | Toronto Field (2) + Durham Field (4) + Toronto Central OT (9) | C, D, J |
| 1 | Toronto Office only (for STD) | A, B, I |
| 8 | OPSEU only (for STD) | E, H |
| All | All divisions (for AD&D) | All |

---

## Benefit Grouping Logic

| Benefit | Basis | Rate Source |
|---|---|---|
| Basic Life (LIFE) | Per $1,000 salary | Config / previous period |
| Dependent Life (DEPL) | Per Member | Config / previous period |
| AD&D | Per $1,000 salary | Config / previous period |
| LTD | Per $100 salary | Config / previous period |
| STD | Per $10 salary | Config / previous period |
| EHC | Per Member (Single/Family) | **From raw data** |
| Dental | Per Member (Single/Family) | **From raw data** |

---

## Files

```
manulife-processor/
├── app.py          → Flask API + routes
├── processor.py    → Parse, clean, group logic  
├── generator.py    → Excel output formatting
├── requirements.txt
└── templates/
    └── index.html  → React frontend (Manulife UI)
```
"# RateSheet" 
