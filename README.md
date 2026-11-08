# ISO 27001 Risk Assessment and Annex A Gap Analysis

An Excel-based **information security risk assessment** and **Annex A gap analysis** for a **fictional** company ("Example Fintech Ltd"). It shows the governance, risk and compliance (GRC) side of cyber security: identifying assets, scoring risks, choosing treatments, and measuring control gaps against ISO/IEC 27001:2022.

> **Learning exercise.** The company, assets, scores and control statuses are invented. This is not a real audit and not certification advice.

## What's inside

| File | Purpose |
|---|---|
| `iso27001_risk_assessment.xlsx` | The workbook (5 sheets, 123 formulas, 2 charts) |
| `gap_analysis_summary.md` | Readable summary of the results, for viewing on GitHub |
| `build_workbook.py` | Script that generates both files from editable data lists |

**Workbook sheets**

1. **Read Me**: scope, method, legend, limitations
2. **Assets**: 10 assets with Confidentiality/Integrity/Availability ratings and a criticality formula
3. **Risk Register**: 13 risks, Likelihood x Impact (1-5), inherent and residual scores, treatment (mitigate / accept / transfer / avoid), linked Annex A controls, owners, status
4. **Annex A Controls**: 37 controls sampled from ISO/IEC 27001:2022 Annex A, with status, evidence and planned action
5. **Gap Summary**: formula-driven counts, percentages by theme, and two charts

## Headline results (from the fictional data)

- **Controls:** 10 implemented, 16 partially implemented, 10 not implemented, 1 not applicable; **27.8% fully implemented** (of 36 applicable)
- **Risks:** 7 Critical and 2 High before treatment, falling to 11 Medium and 2 Low after planned controls (residual is a target, not a measurement)
- **Biggest gaps:** detection and monitoring (A.8.15, A.8.16), incident management (A.5.24, A.5.26), vulnerability and change management (A.8.8, A.8.9, A.8.32), continuity (A.5.30), suppliers (A.5.19)

See [`gap_analysis_summary.md`](gap_analysis_summary.md) for the full tables and a suggested remediation order.

## How the scoring works

- Score = Likelihood x Impact. Rating: **15-25 Critical, 10-14 High, 5-9 Medium, 1-4 Low**
- Rating cells are colour-coded; scores accept whole numbers 1-5 only; statuses and treatments use dropdown lists
- `% implemented = Implemented / (all controls - Not applicable)`. Partial controls do not count as implemented

## Use it for your own scenario

1. Open the workbook and change the blue input cells, or edit the data lists in `build_workbook.py` and run `python build_workbook.py` (needs `openpyxl`).
2. Keep control references consistent between the Risk Register and the Annex A sheet.
3. Optional: import the workbook into **Power BI** to build a dashboard. A Power BI file is not included here.

## Limitations

- Only 37 of the 93 Annex A controls are included. A real engagement also covers ISMS clauses 4-10, scope definition, the Statement of Applicability and evidence testing.
- Control names are short references to ISO/IEC 27001:2022 Annex A. Read the standard itself for the authoritative wording.
- Likelihood and impact scores are judgement-based, which is normal in qualitative risk assessment, but real ones need stakeholder input and evidence.

## Author

Arik Abdullah - [github.com/Arik-Abdullah](https://github.com/Arik-Abdullah)
