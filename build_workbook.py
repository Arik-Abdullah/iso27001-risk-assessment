#!/usr/bin/env python3
"""Builds the ISO 27001 risk assessment + Annex A gap analysis workbook for a FICTIONAL company.

  python build_workbook.py     -> iso27001_risk_assessment.xlsx + gap_analysis_summary.md

All scores, statuses and evidence below are invented for a learning exercise.
Edit the data lists to model a different organisation, then re-run.
"""
from collections import Counter

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

COMPANY = "Example Fintech Ltd (FICTIONAL)"
FONT = "Arial"
STATUSES = ["Implemented", "Partially implemented", "Not implemented", "Not applicable"]
THEMES = ["Organizational", "People", "Physical", "Technological"]
RATINGS = ["Critical", "High", "Medium", "Low"]

# ------------------------------------------------------------------ data ---
ASSETS = [  # id, asset, type, owner, C, I, A (1-3)
    ("A-01", "Customer database (payment metadata)", "Data", "Head of Engineering", 3, 3, 2),
    ("A-02", "Core payment API servers", "System", "CTO", 3, 3, 3),
    ("A-03", "Employee laptops", "Hardware", "IT Manager", 2, 2, 2),
    ("A-04", "Corporate email and collaboration (SaaS)", "Service", "IT Manager", 2, 2, 2),
    ("A-05", "Source code repositories", "Data", "Head of Engineering", 3, 3, 2),
    ("A-06", "Cloud object storage", "Service", "DevOps Lead", 3, 2, 2),
    ("A-07", "HR records", "Data", "HR Manager", 3, 2, 1),
    ("A-08", "Office network and Wi-Fi", "Infrastructure", "IT Manager", 2, 2, 2),
    ("A-09", "Backup system", "System", "DevOps Lead", 2, 3, 3),
    ("A-10", "Customer support platform (third-party SaaS)", "Service", "Support Lead", 2, 2, 2),
]

RISKS = [  # id, asset, threat, vulnerability, L, I, treatment, controls, resL, resI, owner, status
    ("R-01", "A-04", "Phishing leads to compromised staff credentials",
     "Awareness training is annual only; MFA not enforced for all users", 4, 4, "Mitigate",
     "A.6.3, A.8.5", 2, 4, "IT Manager", "In progress"),
    ("R-02", "A-02", "Ransomware encrypts production servers",
     "Slow patching; flat network between servers", 3, 5, "Mitigate",
     "A.8.8, A.8.13, A.8.22", 2, 4, "CTO", "Open"),
    ("R-03", "A-02", "Known vulnerability in an internet-facing service is exploited",
     "No formal patch SLA or vulnerability scanning cadence", 4, 4, "Mitigate",
     "A.8.8, A.8.32", 2, 4, "CTO", "Open"),
    ("R-04", "A-01", "Insider copies customer data",
     "Broad database access; no periodic access reviews", 3, 5, "Mitigate",
     "A.5.15, A.8.2, A.8.15", 2, 4, "Head of Engineering", "Open"),
    ("R-05", "A-03", "Lost or stolen laptop exposes company data",
     "Disk encryption not enforced on all devices", 3, 3, "Mitigate",
     "A.8.1, A.8.24, A.7.9", 1, 3, "IT Manager", "In progress"),
    ("R-06", "A-06", "Misconfigured cloud storage exposes data publicly",
     "No baseline configuration or regular configuration review", 3, 5, "Mitigate",
     "A.5.23, A.8.9", 2, 4, "DevOps Lead", "Open"),
    ("R-07", "A-10", "Breach at third-party support vendor exposes customer data",
     "No supplier security assessment or contractual security terms", 3, 4, "Mitigate",
     "A.5.19, A.5.20", 2, 3, "Support Lead", "Open"),
    ("R-08", "A-09", "Backups fail or cannot be restored after an incident",
     "Restores are never tested; no documented recovery plan", 3, 5, "Mitigate",
     "A.8.13, A.5.30", 2, 3, "DevOps Lead", "Open"),
    ("R-09", "A-01", "Security incident goes undetected for days",
     "Limited log collection; no alert monitoring or incident process", 4, 4, "Mitigate",
     "A.8.15, A.8.16, A.5.24, A.5.26", 3, 3, "CTO", "Open"),
    ("R-10", "A-05", "Credentials or API keys committed to code repositories",
     "No secret scanning; inconsistent code review", 4, 3, "Mitigate",
     "A.8.4, A.8.28", 2, 3, "Head of Engineering", "In progress"),
    ("R-11", "A-08", "Unauthorised device joins office network",
     "Shared Wi-Fi password; office holds no sensitive workloads", 2, 3, "Accept",
     "A.8.20, A.8.22", 2, 3, "IT Manager", "Closed"),
    ("R-12", "A-07", "Unauthorised disclosure of HR records",
     "Shared drive permissions not reviewed", 2, 4, "Mitigate",
     "A.5.15, A.5.34", 1, 4, "HR Manager", "Open"),
    ("R-13", "A-02", "DDoS attack makes the payment API unavailable",
     "No upstream DDoS protection or capacity plan", 3, 3, "Transfer",
     "A.5.19, A.8.6", 2, 3, "CTO", "Open"),
]

CONTROLS = [  # ref, theme, name, status, owner, evidence, planned action
    ("A.5.1", "Organizational", "Policies for information security", "Implemented", "CEO",
     "Information security policy approved (fictional v1.0)", ""),
    ("A.5.2", "Organizational", "Information security roles and responsibilities", "Implemented", "CEO",
     "Security lead role defined in org chart", ""),
    ("A.5.9", "Organizational", "Inventory of information and other associated assets", "Partially implemented", "IT Manager",
     "Hardware list exists; data and SaaS assets not inventoried", "Extend the asset register to data and SaaS"),
    ("A.5.10", "Organizational", "Acceptable use of information and other associated assets", "Implemented", "HR Manager",
     "Acceptable use policy in onboarding pack", ""),
    ("A.5.15", "Organizational", "Access control", "Partially implemented", "Head of Engineering",
     "Role-based access in main apps; no periodic reviews", "Quarterly access review for databases and shared drives"),
    ("A.5.16", "Organizational", "Identity management", "Implemented", "IT Manager",
     "Central identity provider for staff accounts", ""),
    ("A.5.19", "Organizational", "Information security in supplier relationships", "Not implemented", "Support Lead",
     "No supplier risk assessment process", "Create a supplier security questionnaire and review the top 5 vendors"),
    ("A.5.20", "Organizational", "Addressing information security within supplier agreements", "Partially implemented", "Legal",
     "NDAs in place; no security clauses", "Add security and breach-notification clauses to contracts"),
    ("A.5.23", "Organizational", "Information security for use of cloud services", "Partially implemented", "DevOps Lead",
     "Cloud accounts use MFA; no formal cloud policy", "Write a cloud security standard and baseline"),
    ("A.5.24", "Organizational", "Information security incident management planning and preparation", "Not implemented", "CTO",
     "No documented incident process", "Write incident response procedure and escalation matrix"),
    ("A.5.26", "Organizational", "Response to information security incidents", "Not implemented", "CTO",
     "No playbooks", "Create playbooks for phishing, ransomware and data leak"),
    ("A.5.30", "Organizational", "ICT readiness for business continuity", "Not implemented", "CTO",
     "No continuity or recovery plan", "Define recovery objectives and test a restore"),
    ("A.5.34", "Organizational", "Privacy and protection of PII", "Partially implemented", "Legal",
     "Privacy notice published; no data-flow mapping", "Map personal-data flows"),
    ("A.5.36", "Organizational", "Compliance with policies, rules and standards for information security", "Not implemented", "CEO",
     "No internal compliance checks", "Schedule an annual internal review"),
    ("A.6.1", "People", "Screening", "Implemented", "HR Manager",
     "Background checks for all new hires", ""),
    ("A.6.3", "People", "Information security awareness, education and training", "Partially implemented", "HR Manager",
     "Annual training only; no phishing simulations", "Add onboarding training and quarterly phishing tests"),
    ("A.6.5", "People", "Responsibilities after termination or change of employment", "Implemented", "HR Manager",
     "Offboarding checklist removes access within 1 day", ""),
    ("A.6.7", "People", "Remote working", "Partially implemented", "IT Manager",
     "Remote work policy exists; VPN not mandatory", "Require VPN or zero-trust access for remote staff"),
    ("A.7.2", "Physical", "Physical entry", "Implemented", "Office Manager",
     "Badge access to the office", ""),
    ("A.7.4", "Physical", "Physical security monitoring", "Not applicable", "Office Manager",
     "Production systems are cloud-hosted; office holds no sensitive workloads", ""),
    ("A.7.9", "Physical", "Security of assets off-premises", "Partially implemented", "IT Manager",
     "Laptop policy exists; encryption not verified", "Verify and enforce full-disk encryption"),
    ("A.8.1", "Technological", "User endpoint devices", "Partially implemented", "IT Manager",
     "Antivirus on all laptops; no central device management", "Deploy MDM and baseline settings"),
    ("A.8.2", "Technological", "Privileged access rights", "Partially implemented", "CTO",
     "Admin accounts limited; shared admin credentials found", "Remove shared admin accounts; use a vault"),
    ("A.8.4", "Technological", "Access to source code", "Implemented", "Head of Engineering",
     "Repository permissions by team; branch protection on main", ""),
    ("A.8.5", "Technological", "Secure authentication", "Partially implemented", "IT Manager",
     "MFA for admins only", "Enforce MFA for all users"),
    ("A.8.6", "Technological", "Capacity management", "Partially implemented", "DevOps Lead",
     "Basic autoscaling; no capacity planning", "Document capacity thresholds and alerts"),
    ("A.8.7", "Technological", "Protection against malware", "Implemented", "IT Manager",
     "Endpoint protection and email filtering active", ""),
    ("A.8.8", "Technological", "Management of technical vulnerabilities", "Not implemented", "CTO",
     "No scanning cadence or patch SLA", "Monthly scans and a patch SLA by severity"),
    ("A.8.9", "Technological", "Configuration management", "Not implemented", "DevOps Lead",
     "No baselines for cloud or servers", "Adopt secure configuration baselines"),
    ("A.8.13", "Technological", "Information backup", "Partially implemented", "DevOps Lead",
     "Nightly backups run; restores never tested", "Quarterly restore test"),
    ("A.8.15", "Technological", "Logging", "Partially implemented", "DevOps Lead",
     "Application logs kept 30 days; no central collection", "Centralise logs and extend retention"),
    ("A.8.16", "Technological", "Monitoring activities", "Not implemented", "CTO",
     "No alerting on security events", "Introduce a SIEM or managed monitoring"),
    ("A.8.20", "Technological", "Networks security", "Implemented", "IT Manager",
     "Firewall rules reviewed at setup", ""),
    ("A.8.22", "Technological", "Segregation of networks", "Not implemented", "DevOps Lead",
     "Flat network for servers", "Segment production, staging and office networks"),
    ("A.8.24", "Technological", "Use of cryptography", "Partially implemented", "CTO",
     "TLS everywhere; no key management standard", "Document key management"),
    ("A.8.28", "Technological", "Secure coding", "Partially implemented", "Head of Engineering",
     "Code review on main; no secure coding guidelines or secret scanning", "Add secret scanning and guidelines"),
    ("A.8.32", "Technological", "Change management", "Not implemented", "CTO",
     "No formal change approval process", "Introduce change tickets and approvals"),
]


def rating(score):
    return "Critical" if score >= 15 else "High" if score >= 10 else "Medium" if score >= 5 else "Low"


# --------------------------------------------------------------- styling ---
def f(**kw):
    return Font(name=FONT, size=kw.pop("size", 10), **kw)


HEADER_FILL = PatternFill("solid", fgColor="1F3864")
INPUT_BLUE = "0000FF"
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="top", wrap_text=True)
RATING_FILLS = {"Critical": "F4B6B6", "High": "F8CBAD", "Medium": "FFE699", "Low": "C6E0B4"}


def title(ws, text, sub=None):
    ws["A1"] = text
    ws["A1"].font = f(size=14, bold=True)
    if sub:
        ws["A2"] = sub
        ws["A2"].font = f(italic=True, color="595959")


def header(ws, row, labels):
    for c, label in enumerate(labels, 1):
        cell = ws.cell(row=row, column=c, value=label)
        cell.font = f(bold=True, color="FFFFFF")
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        cell.border = BORDER
    ws.row_dimensions[row].height = 32


def widths(ws, ws_widths):
    for i, w in enumerate(ws_widths, 1):
        ws.column_dimensions[chr(64 + i)].width = w


def style_cell(cell, input_cell=False, center=False):
    cell.font = f(color=INPUT_BLUE) if input_cell else f()
    cell.alignment = CENTER if center else WRAP
    cell.border = BORDER


def rating_formatting(ws, rng):
    for name, color in RATING_FILLS.items():
        ws.conditional_formatting.add(
            rng, CellIsRule(operator="equal", formula=[f'"{name}"'],
                            fill=PatternFill("solid", bgColor=color, fgColor=color)))


# ----------------------------------------------------------------- build ---
def build(path="iso27001_risk_assessment.xlsx"):
    wb = Workbook()

    # ---- Read Me
    ws = wb.active
    ws.title = "Read Me"
    ws.column_dimensions["A"].width = 120
    lines = [
        (f"ISO/IEC 27001 Risk Assessment and Annex A Gap Analysis - {COMPANY}", "title"),
        ("LEARNING EXERCISE. The company, assets, scores and control statuses are invented. This is not a real audit "
         "and does not certify or advise on any real organisation.", "warn"),
        ("", None),
        ("Sheets", "h"),
        ("Assets: asset register with Confidentiality / Integrity / Availability ratings (1-3) and a criticality formula.", None),
        ("Risk Register: 13 risks scored by Likelihood x Impact (1-5 each), inherent vs residual, with treatment and linked Annex A controls.", None),
        ("Annex A Controls: a 37-control sample of ISO/IEC 27001:2022 Annex A with implementation status, evidence and planned action.", None),
        ("Gap Summary: formulas that count statuses and ratings, plus two charts. Everything here updates when you edit the other sheets.", None),
        ("", None),
        ("Method", "h"),
        ("Score = Likelihood x Impact. Rating: 15-25 Critical, 10-14 High, 5-9 Medium, 1-4 Low.", None),
        ("Likelihood: 1 rare, 2 unlikely, 3 possible, 4 likely, 5 almost certain. Impact: 1 negligible, 2 minor, 3 moderate, 4 major, 5 severe.", None),
        ("Residual risk is the score expected after the planned controls are in place (a target, not a measured result).", None),
        ("Treatment options: Mitigate, Accept, Transfer, Avoid.", None),
        ("% implemented = Implemented / (all controls - Not applicable). Partially implemented controls do not count as implemented.", None),
        ("", None),
        ("Legend", "h"),
        ("Blue text = input you can change (scores, statuses, owners). Black text = formula or fixed label. Do not overwrite formulas.", None),
        ("Dropdown lists: Treatment, Status, Control status and Asset ID. Scores accept whole numbers 1-5 only.", None),
        ("Example row to copy: R-01 (Risk Register, row 5) shows the expected format for a new risk.", None),
        ("", None),
        ("Source and limitations", "h"),
        ("Control references and short names follow ISO/IEC 27001:2022 Annex A. For authoritative wording, read the standard itself (purchased from ISO or a national body).", None),
        ("Only a sample of the 93 Annex A controls is included. A real gap analysis also covers the ISMS clauses 4-10, scope, Statement of Applicability and evidence review.", None),
        ("Author: Arik Abdullah - github.com/Arik-Abdullah", None),
    ]
    for r, (text, kind) in enumerate(lines, 1):
        c = ws.cell(row=r, column=1, value=text)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        c.font = {"title": f(size=14, bold=True), "h": f(bold=True, size=11, color="1F3864"),
                  "warn": f(bold=True, color="C00000")}.get(kind, f())

    # ---- Assets
    wa = wb.create_sheet("Assets")
    title(wa, "Asset Register", f"{COMPANY}. C/I/A rated 1 (low) to 3 (high).")
    header(wa, 4, ["Asset ID", "Asset", "Type", "Owner", "Confidentiality (1-3)", "Integrity (1-3)",
                   "Availability (1-3)", "Criticality (max of C/I/A)"])
    a_first, a_last = 5, 4 + len(ASSETS)
    for i, row in enumerate(ASSETS):
        r = a_first + i
        for c, v in enumerate(row, 1):
            cell = wa.cell(row=r, column=c, value=v)
            style_cell(cell, input_cell=True, center=c in (1, 5, 6, 7))
        wa.cell(row=r, column=8, value=f"=MAX(E{r}:G{r})")
        style_cell(wa.cell(row=r, column=8), center=True)
    widths(wa, [10, 44, 16, 22, 16, 14, 16, 18])
    dv = DataValidation(type="whole", operator="between", formula1="1", formula2="3", allow_blank=False)
    dv.error, dv.errorTitle = "Enter a whole number from 1 to 3", "Invalid rating"
    wa.add_data_validation(dv)
    dv.add(f"E{a_first}:G{a_last}")
    wa.freeze_panes = "C5"

    # ---- Risk Register
    wr = wb.create_sheet("Risk Register")
    title(wr, "Risk Register", "Score = Likelihood x Impact. Rating: 15+ Critical, 10-14 High, 5-9 Medium, 1-4 Low.")
    header(wr, 4, ["Risk ID", "Asset ID", "Asset", "Threat", "Vulnerability", "Likelihood (1-5)", "Impact (1-5)",
                   "Inherent score", "Inherent rating", "Treatment", "Annex A controls", "Residual likelihood",
                   "Residual impact", "Residual score", "Residual rating", "Risk owner", "Status"])
    r_first, r_last = 5, 4 + len(RISKS)
    for i, (rid, aid, threat, vuln, L, I, treat, ctrls, rl, ri, owner, status) in enumerate(RISKS):
        r = r_first + i
        values = {1: rid, 2: aid, 4: threat, 5: vuln, 6: L, 7: I, 10: treat, 11: ctrls, 12: rl, 13: ri, 16: owner, 17: status}
        for c, v in values.items():
            style_cell(wr.cell(row=r, column=c, value=v), input_cell=True, center=c in (1, 2, 6, 7, 10, 12, 13, 17))
        formulas = {
            3: f"=INDEX(Assets!$B${a_first}:$B${a_last},MATCH(B{r},Assets!$A${a_first}:$A${a_last},0))",
            8: f"=F{r}*G{r}",
            9: f'=IF(H{r}>=15,"Critical",IF(H{r}>=10,"High",IF(H{r}>=5,"Medium","Low")))',
            14: f"=L{r}*M{r}",
            15: f'=IF(N{r}>=15,"Critical",IF(N{r}>=10,"High",IF(N{r}>=5,"Medium","Low")))',
        }
        for c, v in formulas.items():
            style_cell(wr.cell(row=r, column=c, value=v), center=c != 3)
    widths(wr, [8, 9, 26, 38, 40, 11, 9, 10, 11, 11, 24, 11, 10, 10, 11, 20, 12])
    rating_formatting(wr, f"I{r_first}:I{r_last}")
    rating_formatting(wr, f"O{r_first}:O{r_last}")
    for rng, formula in ((f"F{r_first}:G{r_last}", None), (f"L{r_first}:M{r_last}", None)):
        d = DataValidation(type="whole", operator="between", formula1="1", formula2="5", allow_blank=False)
        d.error, d.errorTitle = "Enter a whole number from 1 to 5", "Invalid score"
        wr.add_data_validation(d)
        d.add(rng)
    for rng, items in ((f"J{r_first}:J{r_last}", '"Mitigate,Accept,Transfer,Avoid"'),
                       (f"Q{r_first}:Q{r_last}", '"Open,In progress,Closed"'),
                       (f"B{r_first}:B{r_last}", f"=Assets!$A${a_first}:$A${a_last}")):
        d = DataValidation(type="list", formula1=items, allow_blank=False)
        wr.add_data_validation(d)
        d.add(rng)
    wr.freeze_panes = "D5"

    # ---- Controls
    wc = wb.create_sheet("Annex A Controls")
    title(wc, "Annex A Control Assessment (sample)", f"{COMPANY}. Statuses are invented for the exercise.")
    header(wc, 4, ["Ref", "Theme", "Control", "Status", "Owner", "Evidence / notes", "Planned action"])
    c_first, c_last = 5, 4 + len(CONTROLS)
    for i, row in enumerate(CONTROLS):
        r = c_first + i
        for c, v in enumerate(row, 1):
            style_cell(wc.cell(row=r, column=c, value=v), input_cell=c in (4, 5, 6, 7), center=c == 1)
    widths(wc, [8, 16, 46, 22, 20, 46, 44])
    d = DataValidation(type="list", formula1='"' + ",".join(STATUSES) + '"', allow_blank=False)
    wc.add_data_validation(d)
    d.add(f"D{c_first}:D{c_last}")
    for name, color in (("Implemented", "C6E0B4"), ("Partially implemented", "FFE699"),
                        ("Not implemented", "F4B6B6"), ("Not applicable", "D9D9D9")):
        wc.conditional_formatting.add(
            f"D{c_first}:D{c_last}",
            CellIsRule(operator="equal", formula=[f'"{name}"'], fill=PatternFill("solid", bgColor=color, fgColor=color)))
    wc.freeze_panes = "D5"

    # ---- Gap Summary
    wg = wb.create_sheet("Gap Summary")
    title(wg, "Gap Summary", "All values are formulas driven by the other sheets.")
    CS = f"'Annex A Controls'!$D${c_first}:$D${c_last}"
    CT = f"'Annex A Controls'!$B${c_first}:$B${c_last}"
    header(wg, 4, ["Control status", "Controls"])
    for i, s in enumerate(STATUSES):
        r = 5 + i
        wg.cell(row=r, column=1, value=s)
        wg.cell(row=r, column=2, value=f"=COUNTIF({CS},A{r})")
    wg["A9"], wg["B9"] = "Total assessed", "=SUM(B5:B8)"
    wg["A10"], wg["B10"] = "Applicable controls (excl. N/A)", "=B9-B8"
    wg["A11"], wg["B11"] = "% fully implemented", "=IF(B10=0,0,B5/B10)"
    wg["A12"], wg["B12"] = "% at least partially implemented", "=IF(B10=0,0,(B5+B6)/B10)"
    for r in range(5, 13):
        style_cell(wg.cell(row=r, column=1))
        style_cell(wg.cell(row=r, column=2), center=True)
    wg["B11"].number_format = wg["B12"].number_format = "0.0%"
    for c in ("A9", "B9", "A10", "B10"):
        wg[c].font = f(bold=True)

    header(wg, 15, ["Theme"] + STATUSES + ["Total", "% implemented (of applicable)"])
    for i, th in enumerate(THEMES):
        r = 16 + i
        wg.cell(row=r, column=1, value=th)
        for c in range(2, 6):
            col = chr(64 + c)
            wg.cell(row=r, column=c, value=f"=COUNTIFS({CT},$A{r},{CS},{col}$15)")
        wg.cell(row=r, column=6, value=f"=SUM(B{r}:E{r})")
        wg.cell(row=r, column=7, value=f"=IF((F{r}-E{r})=0,0,B{r}/(F{r}-E{r}))")
        wg.cell(row=r, column=7).number_format = "0.0%"
    wg["A20"] = "All themes"
    for c in range(2, 7):
        col = chr(64 + c)
        wg.cell(row=20, column=c, value=f"=SUM({col}16:{col}19)")
    wg["G20"] = "=IF((F20-E20)=0,0,B20/(F20-E20))"
    wg["G20"].number_format = "0.0%"
    for r in range(16, 21):
        for c in range(1, 8):
            style_cell(wg.cell(row=r, column=c), center=c > 1)
    for c in range(1, 8):
        wg.cell(row=20, column=c).font = f(bold=True)

    RI = f"'Risk Register'!$I${r_first}:$I${r_last}"
    RR = f"'Risk Register'!$O${r_first}:$O${r_last}"
    header(wg, 23, ["Risk rating", "Inherent", "Residual"])
    for i, name in enumerate(RATINGS):
        r = 24 + i
        wg.cell(row=r, column=1, value=name)
        wg.cell(row=r, column=2, value=f"=COUNTIF({RI},A{r})")
        wg.cell(row=r, column=3, value=f"=COUNTIF({RR},A{r})")
    wg["A28"], wg["B28"], wg["C28"] = "Total risks", "=SUM(B24:B27)", "=SUM(C24:C27)"
    for r in range(24, 29):
        for c in range(1, 4):
            style_cell(wg.cell(row=r, column=c), center=c > 1)
    for c in ("A28", "B28", "C28"):
        wg[c].font = f(bold=True)
    widths(wg, [34, 14, 22, 16, 16, 10, 24])

    ch1 = BarChart()
    ch1.type, ch1.title, ch1.legend = "col", "Controls by status", None
    ch1.add_data(Reference(wg, min_col=2, min_row=4, max_row=8), titles_from_data=True)
    ch1.set_categories(Reference(wg, min_col=1, min_row=5, max_row=8))
    ch1.height, ch1.width = 7.5, 14
    wg.add_chart(ch1, "I3")
    ch2 = BarChart()
    ch2.type, ch2.title = "col", "Risks by rating: inherent vs residual"
    ch2.add_data(Reference(wg, min_col=2, max_col=3, min_row=23, max_row=27), titles_from_data=True)
    ch2.set_categories(Reference(wg, min_col=1, min_row=24, max_row=27))
    ch2.height, ch2.width = 7.5, 14
    wg.add_chart(ch2, "I20")

    wb.save(path)
    return path


def write_summary(path="gap_analysis_summary.md"):
    """Plain-text summary computed in Python (also used to cross-check the workbook formulas)."""
    status = Counter(c[3] for c in CONTROLS)
    applicable = len(CONTROLS) - status["Not applicable"]
    inh = Counter(rating(r[4] * r[5]) for r in RISKS)
    res = Counter(rating(r[8] * r[9]) for r in RISKS)
    out = [f"# Gap Analysis Summary - {COMPANY}", "",
           "> Learning exercise with invented data. Numbers match the workbook's `Gap Summary` sheet.", "",
           "## Control status (37-control Annex A sample)", "", "| Status | Controls |", "|---|---|"]
    out += [f"| {s} | {status[s]} |" for s in STATUSES]
    out += ["", f"- Fully implemented: **{status['Implemented'] / applicable:.1%}** of {applicable} applicable controls",
            f"- At least partially implemented: **{(status['Implemented'] + status['Partially implemented']) / applicable:.1%}**", "",
            "## By theme", "", "| Theme | Implemented | Partial | Not implemented | N/A | % implemented |", "|---|---|---|---|---|---|"]
    for th in THEMES:
        cs = [c for c in CONTROLS if c[1] == th]
        n = Counter(c[3] for c in cs)
        app = len(cs) - n["Not applicable"]
        out.append(f"| {th} | {n['Implemented']} | {n['Partially implemented']} | {n['Not implemented']} | "
                   f"{n['Not applicable']} | {n['Implemented'] / app:.0%} |")
    out += ["", "## Risks by rating", "", "| Rating | Inherent | Residual (target) |", "|---|---|---|"]
    out += [f"| {r} | {inh[r]} | {res[r]} |" for r in RATINGS]
    out += ["", "## Highest inherent risks", "", "| ID | Threat | Score | Rating | Treatment |", "|---|---|---|---|---|"]
    for r in sorted(RISKS, key=lambda x: -x[4] * x[5])[:6]:
        out.append(f"| {r[0]} | {r[2]} | {r[4] * r[5]} | {rating(r[4] * r[5])} | {r[6]} |")
    gaps = [c for c in CONTROLS if c[3] == "Not implemented"]
    out += ["", "## Priority gaps (controls not implemented)", ""]
    out += [f"- **{c[0]} {c[2]}**: {c[6]}" for c in gaps]
    out += ["", "## Suggested sequence", "",
            "1. Detection and response: A.8.15, A.8.16, A.5.24, A.5.26 (risk R-09 and support for all incident scenarios)",
            "2. Vulnerability and change management: A.8.8, A.8.9, A.8.32 (risks R-02, R-03, R-06)",
            "3. Resilience: A.5.30 and a tested restore for A.8.13 (risk R-08)",
            "4. Suppliers: A.5.19 and A.5.20 (risk R-07)",
            "5. Network segmentation: A.8.22 (risk R-02)", ""]
    with open(path, "w") as fh:
        fh.write("\n".join(out))
    return path


if __name__ == "__main__":
    print("Wrote", build())
    print("Wrote", write_summary())
