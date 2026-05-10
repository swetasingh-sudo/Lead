# IS Lead Dashboard — Quick Start

## 1. Install dependencies (one time)
```bash
pip install -r requirements.txt
```

## 2. Run
```bash
streamlit run is_lead_dashboard.py
```

Your browser opens automatically at http://localhost:8501

## 3. Upload files
- **Sidebar → File 1**: Drag & drop your **new leads** file (created yesterday)
- **Sidebar → File 2**: Drag & drop your **activity leads** file (worked yesterday)
- Accepts `.xlsx`, `.xls`, `.csv`
- Dashboard updates instantly — no page refresh needed

## What the dashboard checks
| Tab | What's covered |
|-----|---------------|
| Overview | KPI summary, call response split, stage distribution, live alerts |
| Response Time | Avg & median RT (business hours only), by-agent breakdown, distribution histogram |
| Call Status | Responded/Not Responded splits, preferred language completion by agent |
| Field Completion | Basic fields (all leads) + Extended fields (MQL + responded), by-agent drill-down |
| Task Audit | Flags leads with no task where call wasn't responded or "Need Another Call: Yes" |
| Lead Qualification | MQL/MUL split, MUL disqualification reasons, raw stage leads |
| By Agent | Scorecard per agent + comparison table |
| All Leads | Fully filterable master table with CSV download |

## Business hours filter
Adjust the "Business Hours" sliders in the sidebar (default 9 AM – 9 PM).
This controls which leads are included in response time and raw-stage calculations.

## Missing data handling
- Missing columns → dashboard shows "—" for that field, no crash
- Missing rows / nulls → excluded from rate calculations, shown in health score
- Data Health panel at the top shows a score and lists any missing/sparse columns
