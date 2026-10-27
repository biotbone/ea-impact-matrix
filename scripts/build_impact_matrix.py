#!/usr/bin/env python3
"""
build_impact_matrix.py

Generates an Enterprise Architecture Project Impact Matrix workbook from a JSON
config. Run it with no arguments to emit a blank fill-in template.

    python build_impact_matrix.py portfolio.json -o Impact_Matrix_v1.xlsx
    python build_impact_matrix.py --template -o EA_Impact_Matrix_TEMPLATE.xlsx
    python build_impact_matrix.py portfolio.json --check      # validate against the method, write nothing
    python build_impact_matrix.py --intake                    # print the list of documents the matrix needs

Requires: openpyxl. After generating, recalculate the file so cached formula
values exist (LibreOffice, or just open and save in Excel).
"""

import argparse
import json
import sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation

FONT = "Arial"
NAVY, BLUE_HDR, LIGHT, BAND, YELLOW, GREY = (
    "1F3864", "2E5496", "D9E2F3", "F2F5FA", "FFF2CC", "808080")
BLUE_TXT, GREEN_TXT, BLACK = "0000FF", "008000", "000000"
GATE_FILL = {"Met": "C6E0B4", "Partial": "FFE699", "Not started": "F2F2F2",
             "Not met": "F8CBAD", "Not applicable": "E7E6E6"}
LVL = {"Low": 1, "Medium": 2, "High": 3, "Critical": 4}
RISK_THRESHOLD = 9   # likelihood x impact at or above this is red; set portfolio.risk_threshold to change
RISK_RESPONSES = ["Mitigate", "Accept", "Transfer", "Avoid"]
DECISIONS = ["Approved", "Approved with conditions", "Deferred", "Held", "Rejected"]
GATE_STATUSES = ["Met", "Partial", "Not started", "Not met", "Not applicable"]

_thin = Side(style="thin", color="BFBFBF")
BOX = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)

DIMS = ["Business value",
        "Cost optimization and savings",
        "Time criticality",
        "Risk reduction and resilience",
        "Security and compliance",
        "Delivery confidence",
        "Strategic and architecture alignment"]

DIM_KEYS = ["business_value", "cost_optimization", "time_criticality",
            "risk_reduction", "security_compliance", "delivery_confidence",
            "strategic_alignment"]

LENS_PRESETS = {
    "balanced":        [0.20, 0.20, 0.20, 0.10, 0.10, 0.10, 0.10],
    "run_cost":        [0.10, 0.35, 0.15, 0.10, 0.05, 0.15, 0.10],
    "deadline_risk":   [0.10, 0.10, 0.35, 0.20, 0.05, 0.10, 0.10],
    "delivery_speed":  [0.15, 0.10, 0.15, 0.10, 0.05, 0.35, 0.10],
    "security":        [0.10, 0.10, 0.10, 0.20, 0.35, 0.05, 0.10],
}

ANCHORS = [
 ("Business value", "No measurable business outcome; internal convenience only.",
  "Efficiency gain for one team, no quantified baseline.",
  "Quantified efficiency or quality gain for a defined process.",
  "Removes a live operational blind spot or unlocks decisions across teams.",
  "Direct revenue, contractual, or client-facing outcome."),
 ("Cost optimization and savings", "No savings; adds run cost.",
  "Soft labour savings only, not booked.",
  "Modest run-cost reduction, single digit percent of the affected spend.",
  "Material and measurable reduction, or a direct enabler of one.",
  "Primary purpose is cost reduction with a named savings target."),
 ("Time criticality", "No deadline. Can slip a year with no consequence.",
  "Preferable this year, no consequence if deferred.",
  "A live problem is costing time or risk every month it waits.",
  "Dependency window or vendor change makes deferral expensive.",
  "Hard external deadline or an already unsupported component."),
 ("Risk reduction and resilience", "No change to risk posture.",
  "Marginal improvement.",
  "Reduces the frequency or blast radius of a known failure mode.",
  "Removes a recurring incident cause or restores a lost safety net.",
  "Removes a single point of failure with client-visible consequences."),
 ("Security and compliance", "Increases attack surface with no offsetting control.",
  "Net neutral, or new surface offset by new controls.",
  "Minor control improvement.",
  "Removes standing credentials, adds auditability or least privilege.",
  "Closes an audit or regulatory finding."),
 ("Delivery confidence", "Novel, undefined scope, unproven pattern, external dependency.",
  "Large scope with material unknowns.",
  "Well understood work with normal coordination overhead.",
  "Proven pattern already running elsewhere in the estate.",
  "Small, bounded, single team, reversible in hours."),
 ("Strategic and architecture alignment", "Conflicts with target-state architecture.",
  "Tactical, creates no reusable asset.",
  "Consistent with standards, no reuse beyond itself.",
  "Advances a target-state standard reusable by other workloads.",
  "Establishes or completes an enterprise standard."),
]

GATES = ["Alignment with a named business capability",
         "Architecture standards compliance",
         "Security review completed",
         "Operational ownership defined (named team, not a role)",
         "Success KPIs agreed with a measured baseline",
         "Rollback strategy documented and tested",
         "Cost model and payback agreed with FinOps",
         "Data, retention and privacy position cleared",
         "Change and CAB approval path confirmed",
         "Architecture decision recorded and traceable (ADR or decision log)"]

# Optional traceability fields, section 9 of every project sheet: (label, config key, hint, expected)
TRACE_FIELDS = [
    ("Business capability or value stream", "capability", "Name the capability this changes. Gate 1 depends on it.", True),
    ("Architecture principles addressed", "principles", "Optional. The principles this initiative follows or tests.", False),
    ("Gap closed (baseline to target)", "gap_closed", "One line: where it is today, where it ends up.", False),
    ("Benefits owner (named)", "benefit_owner", "The person who answers for the benefit after delivery.", True),
    ("Benefits review date", "benefits_review", "When the KPIs above are checked against the business case.", False),
    ("Standards references", "standards_refs", "Optional. TOGAF phase or artefact, COBIT objective IDs.", False),
]

# Documents the matrix is built from. Single source of truth for the Read Me register,
# the --check rules and the --intake list (references/required-documents.md is generated from it).
DOC_STATUSES = ["Received", "Requested", "Not available", "Not applicable"]
# (code, level, scope, name, feeds, look_for, if_missing)
DOC_CATALOG = [
    ("PROPOSAL", "Required", "Each initiative",
     "Proposal, business case or statement of work",
     "Project profile (name, sponsor, scope, deadline, status, basis); Business value and Cost assessments; the questions the proposal asks the client go to 05 Open Questions",
     "Scope in and out, stated benefits, stated risks, cost, duration, assumptions, questions put to the client",
     "The initiative cannot be scored. Hold it and log what is needed on 05 Open Questions."),
    ("ESTIMATE", "Required", "Each initiative",
     "Cost and effort estimate: line items, hours by role, duration",
     "Cost and duration cells, score per $100k on 02 Portfolio Summary, capacity on 03 Sequencing, Delivery confidence",
     "Line items that add up to the total, hours that imply more than one person per role, consumption charges missing from the estimate",
     "Score cost and confidence conservatively, leave score per $100k empty and log the gap."),
    ("TARGET_STATE", "Required", "Portfolio",
     "Target-state architecture, principles and strategy",
     "Strategic alignment score, section 9 (capability, principles, gap closed), gates 1 and 2",
     "Target state, standards and principles, roadmap, capability map",
     "Score strategic alignment conservatively and log the gap. Do not infer a target state from the proposals."),
    ("LENS", "Required", "Portfolio",
     "Sponsor confirmation of the prioritization lens (email or minutes)",
     "01 Scoring Model weights, Governance record on the Read Me (who agreed the lens, and when)",
     "A named person, a date and the lens chosen: run cost, deadline risk, delivery speed, security or balanced",
     "Do not score. The weights come before the scores. Ask the sponsor first."),
    ("FUNDING", "Required", "Portfolio",
     "Funding envelope or budget statement, or a written 'not yet defined'",
     "03 Sequencing (waves as approval tranches), Delivery confidence weighting",
     "Amount, timing and approval path. If nothing is agreed, get that stated in writing.",
     "Structure the waves as approval tranches and log the open envelope."),
    ("CURRENT_STATE", "Recommended", "Portfolio",
     "Current-state architecture or inventory of the affected systems",
     "Dependencies, section 9 (gap closed), Risk reduction",
     "Systems, owners, interfaces, shared platforms, naming and tenancy conventions",
     "Dependencies are guesses. Mark them as assumptions on 05 Open Questions."),
    ("COST_BASELINE", "Recommended", "Each initiative or portfolio",
     "Current run-cost baseline: cost export, invoices, licence list",
     "06 Savings Case, KPI baselines, Cost optimization score",
     "Twelve months of cost for the affected services, by owner or tag",
     "Never claim a saving. Leave the baseline blank on 06 Savings Case and let the empty payback show it."),
    ("INCIDENTS", "Recommended", "Portfolio",
     "Incident, risk and audit history",
     "Risk reduction and Time criticality scores, 04 Risk Register",
     "Recurring causes, outage counts, open audit findings, vendor end-of-support dates",
     "Score risk reduction and time criticality conservatively and log the gap."),
    ("SECURITY", "Recommended", "Each initiative",
     "Security and compliance requirements, findings and review status",
     "Security and compliance score, gates 3 and 8",
     "Control changes, data classification, regulatory duties, review outcomes",
     "Score security conservatively. A proposal that adds surface scores low until a review says otherwise."),
    ("OWNERSHIP", "Recommended", "Portfolio",
     "Ownership and support model: named people, RACI, org chart",
     "Sponsor, architecture owner, delivery owner, benefits owner, gate 4",
     "Names, not roles. Who runs it after handover.",
     "Owners stay 'TBC' and the check warns."),
    ("CAPACITY", "Recommended", "Portfolio",
     "Resource and capacity plan: named roles and availability",
     "Capacity table on 03 Sequencing, Delivery confidence",
     "The same person committed to two initiatives, peak months, key-person dependence",
     "Capacity conflicts stay invisible. Log it as an open question."),
    ("CONTRACTS", "Recommended", "Portfolio",
     "Vendor contracts, licences and renewal dates",
     "Time criticality, hard deadlines, consumption-billed dependencies",
     "Renewal and notice dates, minimum commitments, price changes",
     "Deadlines may be missed. Log the question."),
    ("PRIOR_DECISIONS", "Optional", "Portfolio",
     "Earlier matrices, ADRs and forum decisions",
     "Consistency of scores, 07 Decision Log, gate 10",
     "Decisions already taken that constrain the portfolio",
     "No effect on the build. Useful for consistency."),
]

# Fixed cell addresses on every project sheet. The summary depends on these.
CELL_NAME, CELL_COST, CELL_MONTHS, CELL_WAVE = "B6", "B10", "B11", "B14"
SCORE_FIRST_ROW, SCORE_LAST_ROW, TOTAL_ROW = 20, 26, 27


# --------------------------------------------------------------- helpers
def add_list(ws, rng, options):
    """Attach a dropdown so users pick a valid value instead of typing one."""
    dv = DataValidation(type="list", formula1='"' + ",".join(options) + '"', allow_blank=True)
    dv.error = "Pick a value from the list."
    ws.add_data_validation(dv)
    dv.add(rng)


def _residual(k):
    """Residual score from an optional 7th and 8th risk item, else None."""
    if len(k) >= 8 and LVL.get(k[6]) and LVL.get(k[7]):
        return LVL[k[6]] * LVL[k[7]]
    return None


def _response_text(k):
    resp = k[5] if len(k) > 5 else ""
    res = _residual(k)
    if resp and res:
        return f"{resp} (residual {res})"
    return resp or (f"residual {res}" if res else "")


def _risk_fill(score):
    amber = max(2, RISK_THRESHOLD - 3)
    return "F8CBAD" if score >= RISK_THRESHOLD else ("FFE699" if score >= amber else "C6E0B4")


def apply_settings(cfg):
    global RISK_THRESHOLD
    RISK_THRESHOLD = int(cfg.get("portfolio", {}).get("risk_threshold", 9))


def validate(cfg):
    """Return a list of (level, where, message). Levels: ERROR breaks the build, WARN breaks a rule of the method."""
    out = []

    def add(level, where, msg):
        out.append((level, where, msg))

    pf = cfg.get("portfolio", {})
    w = cfg.get("weights", {})
    missing = [k for k in DIM_KEYS if k not in w]
    if missing:
        add("ERROR", "weights", "missing weights for: " + ", ".join(missing))
    elif abs(sum(w.values()) - 1.0) > 0.001:
        add("ERROR", "weights", f"weights sum to {sum(w.values()):.3f}, not 1.0")
    try:
        thr = int(pf.get("risk_threshold", 9))
        if not 2 <= thr <= 16:
            raise ValueError
    except (TypeError, ValueError):
        add("ERROR", "portfolio", "risk_threshold must be a whole number from 2 to 16")
        thr = 9
    for key, label in [("lens_approved_by", "who agreed the lens"), ("next_review", "next review date")]:
        v = str(pf.get(key, "")).strip()
        if not v or v.startswith("<") or v == "TBC":
            add("WARN", "portfolio", f"{label} is not recorded (portfolio.{key})")

    scored = []
    for p in cfg.get("projects", []):
        pid = p.get("id", "?")
        sc = p.get("scores")
        if not sc:
            continue
        bad = [k for k in DIM_KEYS
               if isinstance(sc.get(k), bool) or not isinstance(sc.get(k), int) or not 1 <= sc[k] <= 5]
        if bad:
            add("ERROR", pid, "scores must be whole numbers from 1 to 5, check: " + ", ".join(bad))
            continue
        scored.append((pid, tuple(sc[k] for k in DIM_KEYS)))
        ass = p.get("assessments", {})
        noev = [k for k in DIM_KEYS
                if not (str(ass.get(k, {}).get("assessment", "")).strip()
                        and str(ass.get(k, {}).get("evidence", "")).strip())]
        if noev:
            add("WARN", pid, f"{len(noev)} score(s) without an assessment and evidence: " + ", ".join(noev))
        for key, label in [("sponsor", "business sponsor"), ("architecture_owner", "architecture owner"),
                           ("delivery_owner", "delivery owner"), ("benefit_owner", "benefits owner"),
                           ("capability", "business capability")]:
            if str(p.get(key, "")).strip() in ("", "TBC"):
                add("WARN", pid, f"{label} is not named")
        for kpi in p.get("kpis") or []:
            if kpi and kpi[0] and (not str(kpi[1]).strip() or "TBC" in str(kpi[1])):
                add("WARN", pid, f"KPI '{kpi[0][:40]}' has no baseline, log it on Open Questions")
        for k in p.get("risks") or []:
            if not (k and k[0]):
                continue
            L, I = LVL.get(k[1]), LVL.get(k[2])
            if not L or not I:
                add("ERROR", pid, f"risk '{k[0][:40]}' has an invalid likelihood or impact")
                continue
            if L * I >= thr and not str(k[4]).strip():
                add("WARN", pid, f"red risk '{k[0][:40]}' has no owner")
            if len(k) > 5 and k[5] and k[5] not in RISK_RESPONSES:
                add("ERROR", pid, f"risk '{k[0][:40]}': response must be one of " + ", ".join(RISK_RESPONSES))
            if len(k) >= 8 and (k[6] or k[7]):
                res = _residual(k)
                if res is None:
                    add("ERROR", pid, f"risk '{k[0][:40]}': invalid residual likelihood or impact")
                elif res > L * I:
                    add("WARN", pid, f"risk '{k[0][:40]}': residual is higher than the initial score")
        gates = p.get("gates") or []
        if len(gates) > len(GATES):
            add("WARN", pid, f"{len(gates)} gate entries given, the sheet has {len(GATES)}")
        for g in gates:
            if g and g[0] not in GATE_STATUSES:
                add("ERROR", pid, f"gate status '{g[0]}' is not one of " + ", ".join(GATE_STATUSES))

    seen = {}
    for pid, prof in scored:
        seen.setdefault(prof, []).append(pid)
    for ids in seen.values():
        if len(ids) > 1:
            add("WARN", ", ".join(ids), "identical score profile. A matrix that cannot separate these "
                "initiatives is not finished: go back to the sources and find what differs.")
    docs = cfg.get("documents")
    if docs is None:
        add("WARN", "documents", "no document register (documents). The required inputs are unconfirmed, "
            "see references/required-documents.md or run --intake")
    else:
        rows = _doc_rows(cfg)
        codes = {c[0] for c in DOC_CATALOG}
        pids = {p.get("id") for p in cfg.get("projects", [])}
        for d in rows:
            if d["code"] not in codes:
                add("ERROR", "documents", f"unknown document code '{d['code']}', use one of: " + ", ".join(sorted(codes)))
            if d["status"] not in DOC_STATUSES:
                add("ERROR", "documents", f"'{d['code']}': status '{d['status']}' is not one of " + ", ".join(DOC_STATUSES))
            if d["project"] != "ALL" and d["project"] not in pids:
                add("ERROR", "documents", f"'{d['code']}': project '{d['project']}' is not ALL or a project id")
        for code, level, scope, name, *_ in DOC_CATALOG:
            if level != "Required":
                continue
            if scope == "Each initiative":
                for pid, _prof in scored:
                    if not _doc_covered(rows, code, pid):
                        add("WARN", pid, f"required document not received: {name} ({code})")
            elif not _doc_covered(rows, code):
                add("WARN", "documents", f"required document not received: {name} ({code})")
    for d in cfg.get("decisions") or []:
        if d and len(d) > 1 and d[1] and d[1] not in DECISIONS:
            add("ERROR", "decisions", f"'{d[1]}' is not one of " + ", ".join(DECISIONS))
    return out


def report(findings, stream):
    if not findings:
        print("check: no findings", file=stream)
        return
    for lvl, where, msg in sorted(findings, key=lambda f: (f[0] != "ERROR", f[1])):
        print(f"{lvl:5}  {where}: {msg}", file=stream)
    n_err = sum(1 for f in findings if f[0] == "ERROR")
    print(f"check: {n_err} error(s), {len(findings) - n_err} warning(s)", file=stream)


def _doc_rows(cfg):
    out = []
    for d in cfg.get("documents") or []:
        d = (list(d) + [""] * 6)[:6]
        out.append(dict(code=d[0], project=d[1], title=d[2], when=d[3], by=d[4], status=d[5]))
    return out


def _doc_covered(rows, code, pid=None):
    """Per-initiative documents need a Received entry for that project or ALL. Portfolio ones accept Received or Not applicable."""
    for r in rows:
        if r["code"] != code:
            continue
        if pid is None:
            if r["status"] in ("Received", "Not applicable"):
                return True
        elif r["status"] == "Received" and r["project"] in ("ALL", pid):
            return True
    return False


def doc_coverage(cfg):
    """Return {code: (text, ok)} for the Read Me coverage table."""
    rows = _doc_rows(cfg)
    pids = [p["id"] for p in cfg.get("projects", []) if p.get("scores")]
    cov = {}
    for code, level, scope, *_ in DOC_CATALOG:
        if scope == "Each initiative":
            got = [pid for pid in pids if _doc_covered(rows, code, pid)]
            if not pids:
                cov[code] = ("No initiative scored yet", False)
            elif len(got) == len(pids):
                cov[code] = ("Received for all scored initiatives", True)
            elif got:
                cov[code] = ("Received for " + ", ".join(got) + ". Missing for " + ", ".join(p for p in pids if p not in got), False)
            else:
                cov[code] = ("MISSING", False)
        elif scope == "Each initiative or portfolio":
            ok = any(r["code"] == code and r["status"] == "Received" for r in rows)
            cov[code] = ("Received" if ok else "MISSING", ok)
        else:
            if _doc_covered(rows, code):
                st = next((r["status"] for r in rows if r["code"] == code and r["status"] in ("Received", "Not applicable")), "Received")
                cov[code] = (st, True)
            else:
                st = next((r["status"] for r in rows if r["code"] == code), "")
                cov[code] = (st if st in ("Requested", "Not available") else "MISSING", False)
    return cov


def intake_markdown():
    lines = ["# Required documents", "",
             "Generated by `python scripts/build_impact_matrix.py --intake`. The catalog lives in the script, so this file and the checks never disagree.", "",
             "A matrix is only as good as the documents behind it. Five documents are required. The rest raise the quality of a specific sheet. Register every document in the `documents` list of the config with its status (Received, Requested, Not available, Not applicable). A missing required document does not stop the work, it changes it: the third column says how.", ""]
    for level in ("Required", "Recommended", "Optional"):
        lines += [f"## {level}", "",
                  "| Code | Document | Scope | Feeds in the workbook | What to look for | If it is missing |",
                  "|---|---|---|---|---|---|"]
        for code, lv, scope, name, feeds, look, missing in DOC_CATALOG:
            if lv == level:
                lines.append(f"| `{code}` | {name} | {scope} | {feeds} | {look} | {missing} |")
        lines.append("")
    lines += ["## Sufficiency test", "",
              "Before sending the workbook, every score must point to a document and a section, every cost to a line in an estimate, every KPI to a baseline source, and every owner to a name. `--check` tests the first part for you: it warns when a required document is not Received and when a score has no evidence.", "",
              "## Intake request", "",
              "Copy this to the client or sponsor:", "",
              "> To score the proposed initiatives on one consistent basis we need the following. Please send what exists and tell us what does not.",
              ">"]
    n = 1
    for code, lv, scope, name, *_ in DOC_CATALOG:
        if lv == "Required":
            lines.append(f"> {n}. {name} ({scope.lower()}).")
            n += 1
    lines += [">", "> Helpful if available: " + "; ".join(d[3].lower() for d in DOC_CATALOG if d[1] == "Recommended") + ".", ""]
    return "\n".join(lines)


def title_block(ws, title, subtitle, span=7):
    ws["A1"] = title
    ws["A1"].font = Font(name=FONT, size=14, bold=True, color=NAVY)
    ws["A2"] = subtitle
    ws["A2"].font = Font(name=FONT, size=9, italic=True, color=GREY)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=span)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=span)
    ws.row_dimensions[1].height = 22


def section(ws, row, text, span=7):
    c = ws.cell(row=row, column=1, value=text)
    c.font = Font(name=FONT, size=11, bold=True, color="FFFFFF")
    for i in range(1, span + 1):
        ws.cell(row=row, column=i).fill = PatternFill("solid", fgColor=NAVY)
    ws.row_dimensions[row].height = 18


def header_row(ws, row, headers, start_col=1):
    for i, h in enumerate(headers):
        c = ws.cell(row=row, column=start_col + i, value=h)
        c.font = Font(name=FONT, size=10, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=BLUE_HDR)
        c.alignment = Alignment(vertical="center", wrap_text=True)
        c.border = BOX
    ws.row_dimensions[row].height = 30


def body(ws, row, values, start_col=1, band=False, bold=False, color=BLACK):
    for i, v in enumerate(values):
        c = ws.cell(row=row, column=start_col + i, value=v)
        c.font = Font(name=FONT, size=10, color=color, bold=bold)
        c.alignment = Alignment(vertical="top", wrap_text=True)
        c.border = BOX
        if band:
            c.fill = PatternFill("solid", fgColor=BAND)


def widths(ws, spec):
    for col, w in spec.items():
        ws.column_dimensions[col].width = w


def note(ws, row, text, span=7, color=GREY):
    ws.cell(row=row, column=1, value=text).font = Font(
        name=FONT, size=9, italic=True, color=color)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)


def sheet_ref(name):
    return f"'{name}'" if " " in name or "-" in name else name


# --------------------------------------------------------------- builders
def build_readme(wb, cfg):
    ws = wb.create_sheet("00 Read Me")
    p = cfg["portfolio"]
    title_block(ws, "Enterprise Architecture Project Impact Matrix",
                f"{p['name']}  |  Version {p['version']}  |  {p['date']}", span=4)
    widths(ws, {"A": 34, "B": 92, "C": 22, "D": 22})

    r = 4
    section(ws, r, "Purpose", span=4); r += 1
    for line in p.get("purpose", [
        "Scores each proposed initiative on a consistent set of weighted dimensions so that sequencing and funding decisions can be defended in an EA forum.",
        "Every score is a starting position for review, not a decision. The scores are meant to be argued with.",
    ]):
        ws.cell(row=r, column=1, value=line).font = Font(name=FONT, size=10)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
        ws.cell(row=r, column=1).alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 30
        r += 1

    r += 1
    section(ws, r, "Sheets", span=4); r += 1
    header_row(ws, r, ["Sheet", "What it is for", "Who maintains it", "Update cadence"]); r += 1
    for row in [
        ("01 Scoring Model", "Weights and the 1 to 5 anchors. Change a weight and every project score recalculates.", "Enterprise Architecture", "At each portfolio review"),
        ("02 Portfolio Summary", "Roll-up with weighted score, rank, cost and score per $100k. The board view.", "Enterprise Architecture", "Automatic"),
        ("03 Sequencing", "Waves, cross-project dependencies, and resource contention.", "EA and Delivery", "At each portfolio review"),
        ("04 Risk Register", "All project risks with likelihood, impact and mitigation.", "Enterprise Architecture", "Fortnightly"),
        ("05 Open Questions", "What must be answered before approval, and the decision each blocks.", "Engagement lead", "Weekly until closed"),
        ("06 Savings Case", "Payback model. Fill the baseline and expected saving; payback calculates.", "FinOps and EA", "Monthly"),
        ("07 Decision Log", "What the forum decided, on what conditions, and who answers for it.", "EA and Delivery", "After each forum"),
        ("P1 to P8", "One sheet per initiative. Delete the slots you do not use.", "Project architect", "Weekly during delivery"),
    ]:
        body(ws, r, list(row)); r += 1

    r += 1
    section(ws, r, "Cell conventions", span=4); r += 1
    for label, colr, desc in [
        ("Blue text", BLUE_TXT, "Hardcoded input. Edit these."),
        ("Green text", GREEN_TXT, "Link pulled from another sheet. Do not overwrite."),
        ("Black text", BLACK, "Calculated. Do not overwrite."),
    ]:
        ws.cell(row=r, column=1, value=label).font = Font(name=FONT, size=10, bold=True, color=colr)
        ws.cell(row=r, column=2, value=desc).font = Font(name=FONT, size=10)
        r += 1
    c = ws.cell(row=r, column=1, value="Yellow fill")
    c.font = Font(name=FONT, size=10, bold=True)
    c.fill = PatternFill("solid", fgColor=YELLOW)
    ws.cell(row=r, column=2, value="Information still owed. These are the gaps that block approval.").font = Font(name=FONT, size=10)
    r += 2

    section(ws, r, "Rules that keep this useful", span=4); r += 1
    for s in [
        "Never give two initiatives the same score profile. If they score the same on every dimension, the matrix cannot rank them and has no decision value. This is the single most common failure mode.",
        "Score against evidence, and cite the source document and section in the evidence column. A score of 3 with a named source beats a score of 5 with none.",
        "Where no evidence exists, score conservatively and log the gap on '05 Open Questions'. Do not guess upward.",
        "A score can be deliberately low where a proposal is weaker than it claims. Those are the most valuable cells in the workbook.",
        "Set the weights to the prioritization lens the sponsor confirms, before scoring. Do not score first and pick the lens that gives the answer you wanted.",
        "Owners are people, not roles. 'Security' does not approve anything; a named individual does.",
    ]:
        ws.cell(row=r, column=1, value=s).font = Font(name=FONT, size=10)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
        ws.cell(row=r, column=1).alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 28
        r += 1

    r += 1
    section(ws, r, "Governance record", span=4); r += 1
    pf = cfg["portfolio"]
    for label, key, hint in [
        ("Prioritization lens", "lens", "Confirmed with the sponsor before scoring."),
        ("Lens agreed by (named)", "lens_approved_by", "The person who accepted these weights."),
        ("Lens agreed on", "lens_approved_date", ""),
        ("Risk appetite", "risk_appetite", f"Red risks start at a score of {RISK_THRESHOLD}."),
        ("Next review", "next_review", "Re-score on a material change, or on this date."),
    ]:
        v = str(pf.get(key, "") or "").strip()
        ws.cell(row=r, column=1, value=label).font = Font(name=FONT, size=10, bold=True)
        c = ws.cell(row=r, column=2, value=v or "TBC")
        c.font = Font(name=FONT, size=10, color=BLUE_TXT)
        if not v or v == "TBC" or v.startswith("<"):
            c.fill = PatternFill("solid", fgColor=YELLOW)
        ws.cell(row=r, column=3, value=hint).font = Font(name=FONT, size=9, color=GREY)
        r += 1
    ws.cell(row=r, column=1,
            value="Terms follow TOGAF and COBIT: an initiative is a work package, a wave is a transition architecture. "
                  "See references/standards-mapping.md.").font = Font(name=FONT, size=9, color=GREY)
    r += 1

    r += 1
    section(ws, r, "Document coverage", span=4); r += 1
    header_row(ws, r, ["Document", "What it feeds in this workbook", "Level", "Coverage"]); r += 1
    cov = doc_coverage(cfg)
    for code, level, scope, name, feeds, look, missing in DOC_CATALOG:
        text, ok = cov[code]
        if text == "MISSING" and level != "Required":
            text = "Not provided"
        body(ws, r, [f"{name} ({code})", feeds, f"{level}, {scope.lower()}", text], band=False)
        for cc in (1, 2, 3, 4):
            ws.cell(row=r, column=cc).alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(row=r, column=4).fill = PatternFill(
            "solid", fgColor="C6E0B4" if ok else (YELLOW if level == "Required" else "F2F2F2"))
        ws.row_dimensions[r].height = 44
        r += 1
    rows = _doc_rows(cfg)
    if rows:
        r += 1
        section(ws, r, "Document register", span=4); r += 1
        header_row(ws, r, ["Type and initiative", "Document and provider", "Version or date", "Status"]); r += 1
        for i, d in enumerate(rows):
            who = f"{d['title']} ({d['by']})" if d["by"] else d["title"]
            body(ws, r, [f"{d['code']} / {d['project']}", who, d["when"], d["status"]], band=i % 2 == 1)
            ws.cell(row=r, column=4).fill = PatternFill(
                "solid", fgColor=GATE_FILL.get("Met") if d["status"] == "Received" else YELLOW)
            r += 1

    if cfg["portfolio"].get("sources"):
        r += 1
        section(ws, r, "Source documents", span=4); r += 1
        for s in cfg["portfolio"]["sources"]:
            ws.cell(row=r, column=1, value=s).font = Font(name=FONT, size=10)
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
            r += 1


def build_scoring_model(wb, cfg):
    ws = wb.create_sheet("01 Scoring Model")
    title_block(ws, "Scoring model",
                f"Lens: {cfg['portfolio'].get('lens','balanced')}. Changing a weight recalculates every project sheet and the summary.")
    widths(ws, {"A": 36, "B": 12, "C": 46, "D": 42, "E": 42, "F": 42, "G": 42})

    weights = cfg["weights"]
    section(ws, 4, "Weights")
    header_row(ws, 5, ["Dimension", "Weight", "Rationale"])
    rationale = cfg.get("weight_rationale", [""] * 7)
    for i, dim in enumerate(DIMS):
        rr = 6 + i
        band = i % 2 == 1
        body(ws, rr, [dim], band=band)
        c = ws.cell(row=rr, column=2, value=weights[DIM_KEYS[i]])
        c.font = Font(name=FONT, size=10, bold=True, color=BLUE_TXT)
        c.number_format = "0%"; c.border = BOX
        c.alignment = Alignment(horizontal="center")
        if band:
            c.fill = PatternFill("solid", fgColor=BAND)
        body(ws, rr, [rationale[i] if i < len(rationale) else ""], start_col=3, band=band)
        ws.row_dimensions[rr].height = 28

    body(ws, 13, ["Total"], bold=True)
    c = ws.cell(row=13, column=2, value="=SUM(B6:B12)")
    c.font = Font(name=FONT, size=10, bold=True)
    c.number_format = "0%"; c.border = BOX
    c.alignment = Alignment(horizontal="center")
    body(ws, 13, ["Must equal 100 percent."], start_col=3)

    section(ws, 15, "Lens presets")
    header_row(ws, 16, ["Preset"] + DIMS)
    for i, (nm, w) in enumerate(LENS_PRESETS.items()):
        rr = 17 + i
        band = i % 2 == 1
        body(ws, rr, [nm], band=band)
        for j, v in enumerate(w):
            c = ws.cell(row=rr, column=2 + j, value=v)
            c.font = Font(name=FONT, size=10)
            c.number_format = "0%"; c.border = BOX
            c.alignment = Alignment(horizontal="center")
            if band:
                c.fill = PatternFill("solid", fgColor=BAND)
    note(ws, 23, "Copy a preset row into B6:B12 to switch lens. Confirm the lens with the sponsor before scoring, not after.", span=8)

    section(ws, 25, "Score anchors, 1 to 5")
    header_row(ws, 26, ["Dimension", "", "1", "2", "3", "4", "5"])
    for i, row in enumerate(ANCHORS):
        rr = 27 + i
        band = i % 2 == 1
        body(ws, rr, [row[0], ""], band=band)
        body(ws, rr, list(row[1:]), start_col=3, band=band)
        ws.row_dimensions[rr].height = 46


def build_project_sheet(wb, proj, blank=False):
    ws = wb.create_sheet(proj["sheet"])
    title_block(ws, f"{proj['id']}.  {proj.get('name') or '<project name>'}",
                "Enterprise Architecture Impact Matrix")
    widths(ws, {"A": 38, "B": 12, "C": 10, "D": 12, "E": 62, "F": 22, "G": 46})

    section(ws, 4, "1.  Project profile")
    prof = [
        ("Project ID", proj["id"], ""),
        ("Project name", proj.get("name", ""), "Pulled by '02 Portfolio Summary'. Keep it in this cell."),
        ("Business sponsor", proj.get("sponsor", "TBC"), "Named individual required before EA approval."),
        ("Architecture owner", proj.get("architecture_owner", "TBC"), "Named individual required before EA approval."),
        ("Delivery owner", proj.get("delivery_owner", "TBC"), "Named individual required before EA approval."),
        ("Estimated cost (USD)", proj.get("cost", None if blank else 0), proj.get("cost_note", "")),
        ("Estimated duration (months)", proj.get("months", None if blank else 0), proj.get("duration_note", "")),
        ("Hard external deadline", proj.get("deadline", "None identified."), ""),
        ("Current status", proj.get("status", "Proposed"), ""),
        ("Recommended wave", proj.get("wave", "TBC"), "See '03 Sequencing'."),
        ("Basis of estimate", proj.get("basis", "TBC"), "Cite the document and slide or section."),
    ]
    for i, (k, v, n) in enumerate(prof):
        rr = 5 + i
        band = i % 2 == 1
        body(ws, rr, [k], bold=True, band=band)
        c = ws.cell(row=rr, column=2, value=v)
        c.font = Font(name=FONT, size=10, color=BLUE_TXT)
        c.border = BOX
        c.alignment = Alignment(vertical="top", wrap_text=True)
        fill = YELLOW if str(v).strip() in ("TBC", "") else (BAND if band else None)
        ws.merge_cells(start_row=rr, start_column=2, end_row=rr, end_column=4)
        for cc in range(2, 5):
            ws.cell(row=rr, column=cc).border = BOX
            if fill:
                ws.cell(row=rr, column=cc).fill = PatternFill("solid", fgColor=fill)
        if k == "Estimated cost (USD)":
            c.number_format = "$#,##0"
        body(ws, rr, [n], start_col=5, band=band)
        ws.merge_cells(start_row=rr, start_column=5, end_row=rr, end_column=7)

    section(ws, 17, "2.  Weighted impact assessment")
    note(ws, 18, "Enter 1 to 5 in column B. Anchors are on '01 Scoring Model'. Column C is a link; do not edit it.")
    header_row(ws, 19, ["Dimension", "Score (1-5)", "Weight", "Weighted",
                        "Assessment", "Owner", "Evidence and basis"])
    for i, dim in enumerate(DIMS):
        rr = SCORE_FIRST_ROW + i
        band = i % 2 == 1
        key = DIM_KEYS[i]
        sc = proj.get("scores", {}).get(key, None)
        a = proj.get("assessments", {}).get(key, {})
        body(ws, rr, [dim], band=band)
        c = ws.cell(row=rr, column=2, value=sc)
        c.font = Font(name=FONT, size=11, bold=True, color=BLUE_TXT)
        c.alignment = Alignment(horizontal="center", vertical="center"); c.border = BOX
        c.fill = PatternFill("solid", fgColor=YELLOW if sc is None else (BAND if band else "FFFFFF"))
        c = ws.cell(row=rr, column=3, value=f"='01 Scoring Model'!$B${6+i}")
        c.font = Font(name=FONT, size=10, color=GREEN_TXT)
        c.number_format = "0%"; c.border = BOX
        c.alignment = Alignment(horizontal="center", vertical="center")
        if band: c.fill = PatternFill("solid", fgColor=BAND)
        c = ws.cell(row=rr, column=4, value=f"=IF(B{rr}=\"\",\"\",B{rr}*C{rr})")
        c.font = Font(name=FONT, size=10); c.number_format = "0.00"; c.border = BOX
        c.alignment = Alignment(horizontal="center", vertical="center")
        if band: c.fill = PatternFill("solid", fgColor=BAND)
        body(ws, rr, [a.get("assessment", ""), a.get("owner", ""), a.get("evidence", "")],
             start_col=5, band=band)
        ws.row_dimensions[rr].height = 58 if not blank else 30

    rr = TOTAL_ROW
    body(ws, rr, ["Weighted priority score"], bold=True)
    for cc in range(1, 8):
        ws.cell(row=rr, column=cc).fill = PatternFill("solid", fgColor=LIGHT)
        ws.cell(row=rr, column=cc).border = BOX
    c = ws.cell(row=rr, column=4,
                value=f"=IF(COUNT(B{SCORE_FIRST_ROW}:B{SCORE_LAST_ROW})=0,\"\",ROUND(SUM(D{SCORE_FIRST_ROW}:D{SCORE_LAST_ROW}),2))")
    c.font = Font(name=FONT, size=12, bold=True, color=NAVY)
    c.number_format = "0.00"; c.border = BOX
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.fill = PatternFill("solid", fgColor=LIGHT)
    body(ws, rr, ["Out of 5.00. Compare on '02 Portfolio Summary'."], start_col=5, bold=True)
    ws.merge_cells(start_row=rr, start_column=5, end_row=rr, end_column=7)
    for cc in range(5, 8):
        ws.cell(row=rr, column=cc).fill = PatternFill("solid", fgColor=LIGHT)

    r = 29
    section(ws, r, "3.  Scope"); r += 1
    header_row(ws, r, ["In scope", "", "", "", "Explicitly out of scope"])
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=7)
    r += 1
    ins = proj.get("scope_in", [""] * 4) or [""] * 4
    outs = proj.get("scope_out", [""] * 4) or [""] * 4
    for i in range(max(len(ins), len(outs), 4)):
        band = i % 2 == 1
        body(ws, r, [ins[i] if i < len(ins) else ""], band=band)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
        body(ws, r, [outs[i] if i < len(outs) else ""], start_col=5, band=band)
        ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=7)
        for cc in list(range(2, 5)) + [6, 7]:
            ws.cell(row=r, column=cc).border = BOX
            if band: ws.cell(row=r, column=cc).fill = PatternFill("solid", fgColor=BAND)
        r += 1

    r += 1
    section(ws, r, "4.  Dependencies"); r += 1
    header_row(ws, r, ["Type", "Depends on", "", "", "Nature of the dependency", "Owner", "Criticality"]); r += 1
    deps = proj.get("dependencies") or [["", "", "", "", ""] for _ in range(4)]
    for i, d in enumerate(deps):
        band = i % 2 == 1
        body(ws, r, [d[0], d[1]], band=band)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
        for cc in [3, 4]:
            ws.cell(row=r, column=cc).border = BOX
            if band: ws.cell(row=r, column=cc).fill = PatternFill("solid", fgColor=BAND)
        body(ws, r, [d[2], d[3], d[4]], start_col=5, band=band)
        r += 1

    r += 1
    section(ws, r, "5.  Key risks"); r += 1
    header_row(ws, r, ["Risk", "Likelihood", "Impact", "Score", "Mitigation", "Owner", "Response (residual score)"]); r += 1
    risks = proj.get("risks") or [["", "", "", "", ""] for _ in range(4)]
    risk_first = r
    for i, k in enumerate(risks):
        band = i % 2 == 1
        body(ws, r, [k[0]], band=band)
        for j, v in enumerate([k[1], k[2]]):
            c = ws.cell(row=r, column=2 + j, value=v)
            c.font = Font(name=FONT, size=10, color=BLUE_TXT)
            c.alignment = Alignment(horizontal="center", vertical="center"); c.border = BOX
            if band: c.fill = PatternFill("solid", fgColor=BAND)
        c = ws.cell(row=r, column=4, value=LVL.get(k[1], 0) * LVL.get(k[2], 0) or None)
        c.font = Font(name=FONT, size=10, bold=True)
        c.alignment = Alignment(horizontal="center", vertical="center"); c.border = BOX
        if band: c.fill = PatternFill("solid", fgColor=BAND)
        body(ws, r, [k[3], k[4], _response_text(k)], start_col=5, band=band)
        r += 1
    risk_last = r - 1
    note(ws, r, f"Likelihood times impact, each 1 to 4 (Low, Medium, High, Critical). Red at {RISK_THRESHOLD} or above: needs a named owner and a review date. Optional 6th to 8th items per risk: response, residual likelihood, residual impact.")
    r += 2

    section(ws, r, "6.  Success KPIs"); r += 1
    header_row(ws, r, ["KPI", "Baseline", "", "Target", "Measurement source", "Measured by", ""]); r += 1
    kpis = proj.get("kpis") or [["", "", "", "", ""] for _ in range(4)]
    for i, k in enumerate(kpis):
        band = i % 2 == 1
        body(ws, r, [k[0], k[1]], band=band)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        ws.cell(row=r, column=3).border = BOX
        if not str(k[1]).strip() or "TBC" in str(k[1]):
            for cc in [2, 3]:
                ws.cell(row=r, column=cc).fill = PatternFill("solid", fgColor=YELLOW)
        elif band:
            ws.cell(row=r, column=3).fill = PatternFill("solid", fgColor=BAND)
        body(ws, r, [k[2]], start_col=4, band=band)
        body(ws, r, [k[3], k[4], ""], start_col=5, band=band)
        r += 1
    r += 1

    section(ws, r, "7.  EA governance gates"); r += 1
    header_row(ws, r, ["Gate", "Status", "", "", "Evidence or artefact", "Reviewer", "Date cleared"]); r += 1
    gates = proj.get("gates") or [["Not started", "", ""] for _ in GATES]
    gate_first = r
    for i, g in enumerate(GATES):
        st, ev, rev = (gates[i] + ["", "", ""])[:3] if i < len(gates) else ("Not started", "", "")
        body(ws, r, [g], band=(i % 2 == 1))
        c = ws.cell(row=r, column=2, value=st)
        c.font = Font(name=FONT, size=10, bold=True)
        c.alignment = Alignment(horizontal="center", vertical="center"); c.border = BOX
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
        for cc in range(2, 5):
            ws.cell(row=r, column=cc).border = BOX
            ws.cell(row=r, column=cc).fill = PatternFill("solid", fgColor=GATE_FILL.get(st, "F2F2F2"))
        body(ws, r, [ev, rev, ""], start_col=5, band=(i % 2 == 1))
        r += 1
    gate_last = r - 1
    note(ws, r, "Valid statuses: Met, Partial, Not started, Not met, Not applicable.")
    r += 2

    section(ws, r, "8.  Open questions"); r += 1
    qs = proj.get("questions") or [""] * 5
    for i, q in enumerate(qs):
        body(ws, r, [f"Q{i+1}", q], band=(i % 2 == 1))
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
        for cc in range(3, 8):
            ws.cell(row=r, column=cc).border = BOX
            if i % 2 == 1:
                ws.cell(row=r, column=cc).fill = PatternFill("solid", fgColor=BAND)
        r += 1

    r += 1
    section(ws, r, "9.  Traceability and benefits"); r += 1
    for i, (label, key, hint, needed) in enumerate(TRACE_FIELDS):
        band = i % 2 == 1
        v = proj.get(key, "")
        body(ws, r, [label], bold=True, band=band)
        c = ws.cell(row=r, column=2, value=v)
        c.font = Font(name=FONT, size=10, color=BLUE_TXT)
        c.alignment = Alignment(vertical="top", wrap_text=True)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
        fill = YELLOW if (needed and not str(v).strip()) else (BAND if band else None)
        for cc in range(2, 5):
            ws.cell(row=r, column=cc).border = BOX
            if fill:
                ws.cell(row=r, column=cc).fill = PatternFill("solid", fgColor=fill)
        body(ws, r, [hint], start_col=5, band=band)
        ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=7)
        r += 1

    dv = DataValidation(type="whole", operator="between", formula1="1", formula2="5", allow_blank=True)
    dv.error = "Scores are whole numbers from 1 to 5."
    ws.add_data_validation(dv)
    dv.add(f"B{SCORE_FIRST_ROW}:B{SCORE_LAST_ROW}")
    add_list(ws, f"B{gate_first}:B{gate_last}", GATE_STATUSES)
    add_list(ws, f"B{risk_first}:C{risk_last}", list(LVL))

    ws.freeze_panes = "A5"


def build_summary(wb, cfg):
    ws = wb.create_sheet("02 Portfolio Summary")
    title_block(ws, "Portfolio summary",
                "Everything pulls from the project sheets. Weights come from '01 Scoring Model'.", span=11)
    widths(ws, {"A": 7, "B": 46, "C": 15, "D": 12, "E": 15, "F": 8, "G": 16,
                "H": 13, "I": 13, "J": 13, "K": 26})
    section(ws, 4, "Ranked portfolio", span=11)
    header_row(ws, 5, ["ID", "Project", "Estimated cost (USD)", "Duration (months)",
                       "Weighted score", "Rank", "Score per $100k",
                       "Time criticality", "Delivery confidence", "Cost optimization",
                       "Recommended wave"])
    projects = cfg["projects"]
    first, last = 6, 6 + len(projects) - 1
    for i, p in enumerate(projects):
        rr = first + i
        q = sheet_ref(p["sheet"])
        band = i % 2 == 1
        body(ws, rr, [p["id"]], band=band)
        cells = [
            (2, f"=IF({q}!{CELL_NAME}=\"\",\"\",{q}!{CELL_NAME})", "General", GREEN_TXT, False),
            (3, f"=IF({q}!{CELL_COST}=\"\",\"\",{q}!{CELL_COST})", "$#,##0", GREEN_TXT, True),
            (4, f"=IF({q}!{CELL_MONTHS}=\"\",\"\",{q}!{CELL_MONTHS})", "0.0", GREEN_TXT, True),
            (5, f"={q}!$D${TOTAL_ROW}", "0.00", GREEN_TXT, True),
            (6, f"=IF(E{rr}=\"\",\"\",RANK(E{rr},$E${first}:$E${last}))", "0", BLACK, True),
            (7, f"=IFERROR(E{rr}/(C{rr}/100000),\"\")", "0.0", BLACK, True),
            (8, f"=IF({q}!$B${SCORE_FIRST_ROW+2}=\"\",\"\",{q}!$B${SCORE_FIRST_ROW+2})", "0", GREEN_TXT, True),
            (9, f"=IF({q}!$B${SCORE_FIRST_ROW+5}=\"\",\"\",{q}!$B${SCORE_FIRST_ROW+5})", "0", GREEN_TXT, True),
            (10, f"=IF({q}!$B${SCORE_FIRST_ROW+1}=\"\",\"\",{q}!$B${SCORE_FIRST_ROW+1})", "0", GREEN_TXT, True),
            (11, f"=IF({q}!{CELL_WAVE}=\"\",\"\",{q}!{CELL_WAVE})", "General", GREEN_TXT, False),
        ]
        for col, f, fmt, colr, ctr in cells:
            c = ws.cell(row=rr, column=col, value=f)
            c.font = Font(name=FONT, size=10, color=colr, bold=(col in (5, 6)))
            c.number_format = fmt; c.border = BOX
            c.alignment = Alignment(horizontal="center" if ctr else "left", vertical="center")
            if band:
                c.fill = PatternFill("solid", fgColor=BAND)

    rr = last + 1
    body(ws, rr, ["", "Portfolio"], bold=True)
    for cc in range(1, 12):
        ws.cell(row=rr, column=cc).fill = PatternFill("solid", fgColor=LIGHT)
        ws.cell(row=rr, column=cc).border = BOX
    for col, f, fmt in [(3, f"=SUM(C{first}:C{last})", "$#,##0"),
                        (5, f"=IFERROR(AVERAGE(E{first}:E{last}),\"\")", "0.00")]:
        c = ws.cell(row=rr, column=col, value=f)
        c.font = Font(name=FONT, size=10, bold=True)
        c.number_format = fmt; c.border = BOX
        c.alignment = Alignment(horizontal="center")
        c.fill = PatternFill("solid", fgColor=LIGHT)

    r = rr + 3
    section(ws, r, "How to read this", span=11); r += 1
    for line in cfg.get("summary_notes", [
        "Weighted score answers how much this matters. Score per $100k answers how much we get for the money. The two will disagree, and the disagreement is the useful part.",
        "A high-score, low-efficiency initiative is usually a must-do rather than a want-to-do. Fund it on that basis instead of on a business case it cannot win.",
        "A low-score, high-efficiency initiative is a quick win. Use it to fill capacity gaps between larger waves.",
        "Replace these lines with the specific reading for this portfolio.",
    ]):
        ws.cell(row=r, column=1, value=line).font = Font(name=FONT, size=10)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=11)
        ws.cell(row=r, column=1).alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 30
        r += 1
    ws.freeze_panes = "C6"


def build_sequencing(wb, cfg):
    ws = wb.create_sheet("03 Sequencing")
    title_block(ws, "Sequencing, dependencies and capacity",
                "Recommended delivery order and the constraints that shape it", span=6)
    widths(ws, {"A": 12, "B": 40, "C": 16, "D": 14, "E": 60, "F": 44})

    waves = cfg.get("waves") or [["Wave 1", "", "", 0, "", ""] for _ in range(4)]
    section(ws, 4, "Recommended waves", span=6)
    header_row(ws, 5, ["Wave", "Projects", "Indicative window", "Cost (USD)",
                       "Why this grouping", "Entry condition"])
    for i, w in enumerate(waves):
        rr = 6 + i
        band = i % 2 == 1
        body(ws, rr, [w[0], w[1], w[2]], band=band)
        c = ws.cell(row=rr, column=4, value=w[3])
        c.font = Font(name=FONT, size=10, color=BLUE_TXT)
        c.number_format = "$#,##0"; c.border = BOX
        c.alignment = Alignment(horizontal="center", vertical="center")
        if band: c.fill = PatternFill("solid", fgColor=BAND)
        body(ws, rr, [w[4], w[5]], start_col=5, band=band)
        ws.row_dimensions[rr].height = 60
    tr = 6 + len(waves)
    body(ws, tr, ["", "Total", ""], bold=True)
    c = ws.cell(row=tr, column=4, value=f"=SUM(D6:D{tr-1})")
    c.font = Font(name=FONT, size=10, bold=True); c.number_format = "$#,##0"
    c.border = BOX; c.alignment = Alignment(horizontal="center")

    r = tr + 2
    section(ws, r, "Cross-project dependencies", span=6); r += 1
    header_row(ws, r, ["From", "To", "Type", "Strength",
                       "What it means for sequencing", "Owner"]); r += 1
    for i, d in enumerate(cfg.get("cross_dependencies") or [["", "", "", "", "", ""] for _ in range(5)]):
        body(ws, r, list(d[:4]), band=(i % 2 == 1))
        body(ws, r, [d[4], d[5]], start_col=5, band=(i % 2 == 1))
        ws.row_dimensions[r].height = 40
        r += 1

    r += 1
    section(ws, r, "Resource contention", span=6); r += 1
    header_row(ws, r, ["Role", "Committed across the portfolio", "Peak concurrent demand", "", "Constraint", ""]); r += 1
    for i, x in enumerate(cfg.get("resources") or [["", "", "", "", "", ""] for _ in range(4)]):
        body(ws, r, list(x[:3]), band=(i % 2 == 1))
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
        ws.cell(row=r, column=4).border = BOX
        body(ws, r, [x[4]], start_col=5, band=(i % 2 == 1))
        ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=6)
        ws.cell(row=r, column=6).border = BOX
        ws.row_dimensions[r].height = 36
        r += 1


def build_risk_register(wb, cfg):
    ws = wb.create_sheet("04 Risk Register")
    thr = RISK_THRESHOLD
    title_block(ws, "Consolidated risk register",
                "All project risks, ranked by initial likelihood times impact", span=10)
    widths(ws, {"A": 8, "B": 8, "C": 56, "D": 12, "E": 12, "F": 8, "G": 62, "H": 22, "I": 12, "J": 10})
    section(ws, 4, "Risks across the portfolio", span=10)
    header_row(ws, 5, ["ID", "Project", "Risk", "Likelihood", "Impact", "Score",
                       "Mitigation", "Owner", "Response", "Residual"])
    allr = []
    for p in cfg["projects"]:
        for k in (p.get("risks") or []):
            if k and k[0]:
                allr.append((p["id"], k))
    allr.sort(key=lambda x: -(LVL.get(x[1][1], 0) * LVL.get(x[1][2], 0)))
    if not allr:
        allr = [("", ["", "", "", "", ""]) for _ in range(8)]
    for i, (pid, k) in enumerate(allr):
        rr = 6 + i
        band = i % 2 == 1
        body(ws, rr, [f"R{i+1:02d}" if pid else "", pid, k[0]], band=band)
        for j, v in enumerate([k[1], k[2]]):
            c = ws.cell(row=rr, column=4 + j, value=v)
            c.font = Font(name=FONT, size=10, color=BLUE_TXT)
            c.alignment = Alignment(horizontal="center", vertical="center"); c.border = BOX
            if band: c.fill = PatternFill("solid", fgColor=BAND)
        sc = LVL.get(k[1], 0) * LVL.get(k[2], 0)
        c = ws.cell(row=rr, column=6, value=sc or None)
        c.font = Font(name=FONT, size=10, bold=True)
        c.alignment = Alignment(horizontal="center", vertical="center"); c.border = BOX
        if sc:
            c.fill = PatternFill("solid", fgColor=_risk_fill(sc))
        body(ws, rr, [k[3], k[4], k[5] if len(k) > 5 else ""], start_col=7, band=band)
        res = _residual(k)
        c = ws.cell(row=rr, column=10, value=res)
        c.font = Font(name=FONT, size=10, bold=True)
        c.alignment = Alignment(horizontal="center", vertical="center"); c.border = BOX
        if res:
            c.fill = PatternFill("solid", fgColor=_risk_fill(res))
        ws.row_dimensions[rr].height = 32
    add_list(ws, f"I6:I{5 + len(allr)}", RISK_RESPONSES)
    note(ws, 6 + len(allr) + 1,
         f"Score is likelihood times impact, each 1 to 4 (Low, Medium, High, Critical). Red at {thr} or above, amber from {max(2, thr - 3)}. "
         "Every red risk needs a named owner and a review date before its project is approved. "
         "Response is Mitigate, Accept, Transfer or Avoid. Residual is the score once the mitigation is in place.",
         span=10)
    ws.freeze_panes = "C6"


def build_open_questions(wb, cfg):
    ws = wb.create_sheet("05 Open Questions")
    title_block(ws, "Open questions and information gaps",
                "What is still owed, and the decision each answer unblocks", span=7)
    widths(ws, {"A": 8, "B": 10, "C": 26, "D": 66, "E": 56, "F": 24, "G": 14})
    section(ws, 4, "Blocking questions, in priority order", span=7)
    header_row(ws, 5, ["ID", "Project", "Area", "Question",
                       "Why it matters and what it unblocks", "Owner to answer", "Status"])
    qs = cfg.get("open_questions") or [["", "", "", "", "", "Open"] for _ in range(12)]
    fillmap = {"Answered": "C6E0B4", "In progress": "BDD7EE", "Mitigated": "D9D9D9"}
    for i, q in enumerate(qs):
        rr = 6 + i
        band = i % 2 == 1
        body(ws, rr, [f"Q{i+1:02d}"] + list(q[:5]), band=band)
        c = ws.cell(row=rr, column=7, value=q[5] if len(q) > 5 else "Open")
        c.font = Font(name=FONT, size=10, bold=True)
        c.fill = PatternFill("solid", fgColor=fillmap.get(c.value, YELLOW))
        c.alignment = Alignment(horizontal="center", vertical="center"); c.border = BOX
        ws.row_dimensions[rr].height = 40
    ws.freeze_panes = "D6"


def build_decision_log(wb, cfg):
    ws = wb.create_sheet("07 Decision Log")
    title_block(ws, "Decision log",
                "What the forum decided, on what conditions, and who answers for it", span=7)
    widths(ws, {"A": 8, "B": 10, "C": 26, "D": 60, "E": 24, "F": 14, "G": 36})
    section(ws, 4, "Decisions", span=7)
    header_row(ws, 5, ["Ref", "Project", "Decision", "Conditions and rationale",
                       "Decided by (named)", "Date", "Reference (ADR or minutes)"])
    rows = cfg.get("decisions") or [["", "", "", "", "", ""] for _ in range(10)]
    for i, d in enumerate(rows):
        d = (list(d) + [""] * 6)[:6]
        body(ws, 6 + i, [f"D{i+1:02d}" if any(d) else "", d[0], d[1], d[2], d[3], d[4], d[5]],
             band=i % 2 == 1)
        ws.row_dimensions[6 + i].height = 30
    add_list(ws, f"C6:C{5 + len(rows)}", DECISIONS)
    note(ws, 6 + len(rows) + 1,
         "Decisions: Approved, Approved with conditions, Deferred, Held, Rejected. A deferred or held initiative "
         "states what would change the decision. The scores recommend; this sheet records what was decided.",
         span=7)
    ws.freeze_panes = "A6"


def build_savings(wb, cfg):
    ws = wb.create_sheet("06 Savings Case")
    title_block(ws, "Savings case and payback",
                "Yellow cells are the inputs still owed. Everything else calculates.", span=8)
    widths(ws, {"A": 8, "B": 44, "C": 15, "D": 20, "E": 20, "F": 15, "G": 18, "H": 54})
    section(ws, 4, "Payback model", span=8)
    header_row(ws, 5, ["ID", "Project", "Build cost (USD)", "Current annual run cost (USD)",
                       "Expected annual saving (USD)", "Saving %", "Payback (months)",
                       "Evidence position"])
    projects = cfg["projects"]
    first = 6
    for i, p in enumerate(projects):
        rr = first + i
        band = i % 2 == 1
        q = sheet_ref(p["sheet"])
        body(ws, rr, [p["id"]], band=band)
        c = ws.cell(row=rr, column=2, value=f"=IF({q}!{CELL_NAME}=\"\",\"\",{q}!{CELL_NAME})")
        c.font = Font(name=FONT, size=10, color=GREEN_TXT); c.border = BOX
        c.alignment = Alignment(vertical="center")
        if band: c.fill = PatternFill("solid", fgColor=BAND)
        c = ws.cell(row=rr, column=3, value=f"={q}!{CELL_COST}")
        c.font = Font(name=FONT, size=10, color=GREEN_TXT); c.number_format = "$#,##0"
        c.border = BOX; c.alignment = Alignment(horizontal="center", vertical="center")
        if band: c.fill = PatternFill("solid", fgColor=BAND)
        sv = (cfg.get("savings") or {}).get(p["id"], {})
        for col, key in [(4, "annual_run_cost"), (5, "expected_saving")]:
            c = ws.cell(row=rr, column=col, value=sv.get(key, 0))
            c.font = Font(name=FONT, size=10, color=BLUE_TXT); c.number_format = "$#,##0"
            c.border = BOX; c.alignment = Alignment(horizontal="center", vertical="center")
            c.fill = PatternFill("solid", fgColor=YELLOW)
        c = ws.cell(row=rr, column=6, value=f'=IFERROR(E{rr}/D{rr},"")')
        c.font = Font(name=FONT, size=10); c.number_format = "0.0%"; c.border = BOX
        c.alignment = Alignment(horizontal="center", vertical="center")
        if band: c.fill = PatternFill("solid", fgColor=BAND)
        c = ws.cell(row=rr, column=7, value=f'=IFERROR(C{rr}/(E{rr}/12),"Awaiting baseline")')
        c.font = Font(name=FONT, size=10, bold=True); c.number_format = "0.0"; c.border = BOX
        c.alignment = Alignment(horizontal="center", vertical="center")
        if band: c.fill = PatternFill("solid", fgColor=BAND)
        body(ws, rr, [sv.get("note", "")], start_col=8, band=band)
        ws.row_dimensions[rr].height = 40
    last = first + len(projects) - 1
    rr = last + 1
    body(ws, rr, ["", "Portfolio"], bold=True)
    for cc in range(1, 9):
        ws.cell(row=rr, column=cc).fill = PatternFill("solid", fgColor=LIGHT)
        ws.cell(row=rr, column=cc).border = BOX
    for col, f, fmt in [(3, f"=SUM(C{first}:C{last})", "$#,##0"),
                        (4, f"=SUM(D{first}:D{last})", "$#,##0"),
                        (5, f"=SUM(E{first}:E{last})", "$#,##0"),
                        (6, f'=IFERROR(E{rr}/D{rr},"")', "0.0%"),
                        (7, f'=IFERROR(C{rr}/(E{rr}/12),"Awaiting baseline")', "0.0")]:
        c = ws.cell(row=rr, column=col, value=f)
        c.font = Font(name=FONT, size=10, bold=True); c.number_format = fmt
        c.border = BOX; c.alignment = Alignment(horizontal="center")
        c.fill = PatternFill("solid", fgColor=LIGHT)
    note(ws, rr + 2,
         "Columns D and E are the only inputs. An initiative that is a pure enabler should show zero expected saving and be argued on the grounds that it makes the other rows provable.",
         span=8)


# --------------------------------------------------------------- template
def template_config(slots=8):
    projects = []
    for i in range(1, slots + 1):
        p = {"id": f"P{i}", "sheet": f"P{i}", "name": "", "cost": 0, "months": 0,
             "sponsor": "TBC", "architecture_owner": "TBC", "delivery_owner": "TBC",
             "deadline": "", "status": "Proposed", "wave": "", "basis": "TBC"}
        projects.append(p)
    # P1 is a worked example so the expected format is visible
    projects[0].update({
        "name": "EXAMPLE  Replace nightly batch export with event streaming",
        "sponsor": "A. Example, VP Operations",
        "architecture_owner": "B. Example, Cloud Architect",
        "delivery_owner": "C. Example, Delivery Lead",
        "cost": 84000, "months": 3,
        "cost_note": "Proposal deck slide 14. Line items sum correctly.",
        "duration_note": "3 months, two engineers plus part-time architect.",
        "deadline": "None contractual. The nightly window is already overrunning twice a week.",
        "status": "Proposed, awaiting approval",
        "wave": "Wave 1",
        "basis": "Proposal deck slides 6 and 14; batch failure log Jan to Jun 2026.",
        "scores": {"business_value": 4, "cost_optimization": 3, "time_criticality": 3,
                   "risk_reduction": 4, "security_compliance": 3,
                   "delivery_confidence": 4, "strategic_alignment": 4},
        "assessments": {
            "business_value": {"assessment": "Downstream teams currently wait until 06:00 for data that is generated continuously. Removing the wait changes when decisions can be made, not just how fast a job runs.",
                               "owner": "A. Example", "evidence": "Deck slide 6; batch completion times Jan to Jun."},
            "cost_optimization": {"assessment": "Removes an always-on batch VM but adds Event Hubs throughput units. Net saving is real but modest and needs modelling before it is claimed.",
                                  "owner": "FinOps", "evidence": "Deck slide 14. Throughput unit cost not in the estimate."},
            "time_criticality": {"assessment": "No deadline, but the nightly window overruns twice a week and the failure rate is trending up.",
                                 "owner": "Platform Team", "evidence": "Batch failure log Jan to Jun 2026."},
            "risk_reduction": {"assessment": "Removes a single nightly failure point that currently takes the whole downstream chain with it.",
                               "owner": "Platform Team", "evidence": "Incident records, 6 occurrences in H1."},
            "security_compliance": {"assessment": "Replaces a shared service account with managed identity. Modest control improvement.",
                                    "owner": "Security", "evidence": "Current-state review, section 4."},
            "delivery_confidence": {"assessment": "The same pattern is already in production for a smaller feed, so this reuses a proven design.",
                                    "owner": "Architecture", "evidence": "Existing implementation on the orders feed."},
            "strategic_alignment": {"assessment": "Advances the event-driven target state and creates a pattern other feeds can reuse.",
                                    "owner": "Architecture", "evidence": "Target-state architecture, section 3."},
        },
        "scope_in": ["Event Hubs ingestion for the customer feed",
                     "Consumer rewrite with idempotent handling",
                     "Parallel run against the existing batch",
                     "Runbook and alert updates"],
        "scope_out": ["Other feeds beyond the customer feed",
                      "Downstream consumer redesign",
                      "Data model changes",
                      "Decommission of the batch VM (separate change)"],
        "dependencies": [
            ["Platform", "Existing Event Hubs namespace", "Reuse, do not recreate. Upstream diagnostic settings point at it.", "Platform Team", "Critical"],
            ["Client team", "Downstream data owners", "They must confirm parity before the batch is switched off.", "PM", "High"],
            ["Internal", "P2 Cost visibility work", "Supplies the run-cost baseline this project's saving is measured against.", "FinOps", "Medium"],
            ["Resource", "Solutions Architect capacity", "Part time for 3 months. Check against other initiatives.", "Delivery lead", "Medium"],
        ],
        "risks": [
            ["Duplicate or lost records during parallel run", "Medium", "Critical", "Run both paths against a shadow consumer group and reconcile record counts daily before cutover.", "Architecture", "Mitigate", "Low", "High"],
            ["Event Hubs throughput unit cost not modelled", "High", "Medium", "Model expected events per second against the pricing tier before committing.", "FinOps", "Mitigate", "Low", "Medium"],
            ["Downstream consumers assume a nightly batch boundary", "Medium", "High", "Survey each consumer for time-window assumptions during design.", "PM"],
            ["Batch VM decommission deferred indefinitely, so no saving lands", "High", "Medium", "Make the decommission a named deliverable with a date, not a follow-up.", "Delivery lead"],
        ],
        "kpis": [
            ["Data freshness, generation to availability", "8 hours (nightly)", "Under 5 minutes", "Pipeline telemetry", "30 days after cutover"],
            ["Failed nightly runs per month", "2.3 average H1 2026", "0", "Incident records", "90 days after cutover"],
            ["Monthly run cost for the feed", "TBC, batch VM plus storage", "Below baseline including throughput units", "Azure cost export", "60 days after cutover"],
            ["Record parity, new path versus batch", "Batch baseline", "100 percent during parallel run", "Reconciliation report", "During parallel run"],
        ],
        "gates": [
            ["Met", "Target-state architecture section 3", "B. Example"],
            ["Met", "Reuses the orders feed pattern", "B. Example"],
            ["Partial", "Managed identity designed, formal review not recorded", "Security"],
            ["Not started", "Ownership after handover undefined", "Platform Team"],
            ["Partial", "KPIs defined; run-cost baseline outstanding", "FinOps"],
            ["Met", "Batch path retained until parallel run passes", "Architecture"],
            ["Not met", "Throughput unit cost not modelled", "FinOps"],
            ["Not started", "Customer data classification", "Legal"],
            ["Not started", "Change window for cutover", "PMO"],
            ["Partial", "ADR drafted, not yet accepted by the architecture board", "B. Example"],
        ],
        "capability": "Customer data distribution (value stream: order to insight)",
        "principles": "Event-driven by default; managed identity over shared accounts",
        "gap_closed": "Baseline: nightly batch, 8 hour latency. Target: continuous feed, under 5 minutes.",
        "benefit_owner": "A. Example, VP Operations",
        "benefits_review": "90 days after cutover",
        "standards_refs": "TOGAF Phase E/F work package; COBIT APO05, EDM02",
        "questions": [
            "What is the current monthly run cost of the batch VM and its storage?",
            "How many events per second at peak, so throughput units can be sized?",
            "Which downstream consumers assume a nightly batch boundary?",
            "Who owns the pipeline and its alerts after handover?",
            "Can the batch VM decommission be a named deliverable with a date?",
        ],
    })
    return {
        "portfolio": {
            "name": "<client or portfolio name>",
            "version": "1.0",
            "date": "<date>",
            "lens": "balanced",
            "lens_approved_by": "A. Example, VP Operations",
            "lens_approved_date": "2026-01-15",
            "risk_appetite": "No red risk without a named owner and a review date.",
            "risk_threshold": 9,
            "next_review": "2026-04-15",
            "purpose": [
                "Scores each proposed initiative on a consistent set of weighted dimensions so that sequencing and funding decisions can be defended in an EA forum.",
                "Replace this text with the specific purpose for this engagement.",
                "Every score is a starting position for review, not a decision.",
            ],
        },
        "weights": dict(zip(DIM_KEYS, LENS_PRESETS["balanced"])),
        "weight_rationale": [
            "Whether the initiative changes a business outcome, not just a technical one.",
            "Direct savings and the ability to evidence savings. An unevidenced saving cannot be booked.",
            "Deadlines, vendor retirements, and problems that cost something every month they wait.",
            "Reduction in incident frequency, blast radius, or single points of failure.",
            "Net change in security posture. This can be negative.",
            "Confidence the work lands on time and scope. Low confidence discounts a high score.",
            "Whether it advances a target-state standard or is purely tactical.",
        ],
        "projects": projects,
        "waves": [
            ["Wave 1", "<initiatives>", "<window>", 0, "<why this grouping>", "<entry condition>"],
            ["Wave 2", "", "", 0, "", ""],
            ["Wave 3", "", "", 0, "", ""],
            ["Hold", "", "Not scheduled", 0, "<why held rather than dropped>", "<what would unhold it>"],
        ],
        "cross_dependencies": [["", "", "", "", "", ""] for _ in range(6)],
        "resources": [
            ["Solutions Architect", "", "", "", "", ""],
            ["DevOps / Platform Engineer", "", "", "", "", ""],
            ["Project Manager", "", "", "", "", ""],
            ["FinOps", "", "", "", "", ""],
        ],
        "open_questions": [["", "", "", "", "", "Open"] for _ in range(14)],
        "documents": [
            ["PROPOSAL", "P1", "Proposal deck, event streaming for the customer feed", "v3, 2026-01-08", "A. Example", "Received"],
            ["ESTIMATE", "P1", "Cost and effort estimate, deck slide 14", "v3, 2026-01-08", "A. Example", "Received"],
            ["TARGET_STATE", "ALL", "Target-state architecture, section 3", "v2, 2025-11-20", "B. Example", "Received"],
            ["LENS", "ALL", "Sponsor email confirming the balanced lens", "2026-01-15", "A. Example", "Received"],
            ["FUNDING", "ALL", "Funding envelope not yet defined, confirmed in writing", "2026-01-15", "C. Example", "Received"],
            ["INCIDENTS", "ALL", "Batch failure log, January to June", "2026-07-01", "Platform Team", "Received"],
            ["COST_BASELINE", "P1", "Run-cost export for the batch VM and storage", "", "FinOps", "Requested"],
        ],
        "decisions": [["", "", "", "", "", ""] for _ in range(10)],
        "savings": {"P1": {"annual_run_cost": 0, "expected_saving": 0,
                           "note": "EXAMPLE. Fill D and E once the baseline arrives. Payback calculates."}},
        "summary_notes": [
            "Weighted score answers how much this matters. Score per $100k answers how much we get for the money.",
            "Replace these lines with the specific reading for this portfolio, naming initiatives.",
        ],
    }


# --------------------------------------------------------------- main
def build(cfg, out):
    apply_settings(cfg)
    wb = Workbook()
    wb.remove(wb.active)
    build_readme(wb, cfg)
    build_scoring_model(wb, cfg)
    build_summary(wb, cfg)
    build_sequencing(wb, cfg)
    build_risk_register(wb, cfg)
    build_open_questions(wb, cfg)
    build_savings(wb, cfg)
    build_decision_log(wb, cfg)
    for p in cfg["projects"]:
        build_project_sheet(wb, p, blank=not p.get("scores"))
    wb.save(out)
    return out


def main():
    ap = argparse.ArgumentParser(description="Build an EA Project Impact Matrix workbook.")
    ap.add_argument("config", nargs="?", help="Path to a JSON config. Omit with --template.")
    ap.add_argument("-o", "--out", default="Impact_Matrix.xlsx")
    ap.add_argument("--template", action="store_true", help="Emit a blank fill-in template.")
    ap.add_argument("--slots", type=int, default=8, help="Project slots in template mode.")
    ap.add_argument("--dump-config", metavar="PATH", help="Write the template config as JSON and exit.")
    ap.add_argument("--check", action="store_true",
                    help="Validate the config against the method's rules and exit. No workbook is written.")
    ap.add_argument("--strict", action="store_true", help="Treat warnings as errors.")
    ap.add_argument("--intake", action="store_true",
                    help="Print the required-documents list (Markdown) and exit.")
    a = ap.parse_args()

    if a.intake:
        print(intake_markdown())
        return

    if a.dump_config:
        with open(a.dump_config, "w") as f:
            json.dump(template_config(a.slots), f, indent=2)
        print(f"wrote {a.dump_config}")
        return

    if a.template or not a.config:
        cfg = template_config(a.slots)
        findings = []
    else:
        with open(a.config) as f:
            cfg = json.load(f)
        for i, p in enumerate(cfg["projects"], 1):
            p.setdefault("id", f"P{i}")
            p.setdefault("sheet", p["id"])
        apply_settings(cfg)
        findings = validate(cfg)
        report(findings, sys.stdout if a.check else sys.stderr)

    errors = [f for f in findings if f[0] == "ERROR"]
    warns = [f for f in findings if f[0] == "WARN"]
    if a.check:
        sys.exit(2 if errors else (1 if (a.strict and warns) else 0))
    if errors:
        print("not built: fix the errors above.", file=sys.stderr)
        sys.exit(2)
    if a.strict and warns:
        print("not built: --strict treats warnings as errors.", file=sys.stderr)
        sys.exit(1)

    print(build(cfg, a.out))
    print("Now recalculate the file so cached formula values exist "
          "(open and save in Excel, or run LibreOffice headless).", file=sys.stderr)


if __name__ == "__main__":
    main()
