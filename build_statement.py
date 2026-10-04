import os
import sys
import warnings
from datetime import date, datetime, timedelta

import pandas as pd

warnings.filterwarnings("ignore", message="Unknown extension is not supported")
# ============================================================
# DA RATE TABLE  (Effective date -> DA%)
# ============================================================
# To update for a new DA revision, add one tuple to the correct
# CPC block below. Format: ("YYYY-MM-DD", rate)
DEFAULT_INPUT  = "data/employee.xlsx"
DEFAULT_OUTPUT = "output/statement.xlsx"


DA_TABLE = [
    # ---------- 5th CPC period (up to 30-06-2006) ----------
    ("2004-01-01",  61),
    ("2004-04-01",  11),   # 50% merged as Dearness Pay
    ("2004-07-01",  14),
    ("2005-01-01",  17),
    ("2005-07-01",  21),
    ("2006-01-01",  24),

    # ---------- 6th CPC period (01-01-2006 onwards) ----------
    ("2006-01-01",   0),   # reset by 6th CPC
    ("2006-07-01",   2),
    ("2007-01-01",   6),
    ("2007-07-01",   9),
    ("2008-01-01",  12),
    ("2008-07-01",  16),
    ("2009-01-01",  22),
    ("2009-07-01",  27),
    ("2010-01-01",  35),
    ("2010-07-01",  45),
    ("2011-01-01",  51),
    ("2011-07-01",  58),
    ("2012-01-01",  65),
    ("2012-07-01",  72),
    ("2013-01-01",  80),
    ("2013-07-01",  90),
    ("2014-01-01", 100),
    ("2014-07-01", 107),
    ("2015-01-01", 113),
    ("2015-07-01", 119),

    # ---------- 7th CPC period (01-01-2016 onwards) ----------
    ("2016-01-01",   0),   # reset by 7th CPC
    ("2016-07-01",   2),
    ("2017-01-01",   4),
    ("2017-07-01",   5),
    ("2018-01-01",   7),
    ("2018-07-01",   9),
    ("2019-01-01",  12),
    ("2019-07-01",  17),
    ("2020-01-01",  21),
    ("2020-07-01",  24),
    ("2021-01-01",  28),
    ("2021-07-01",  31),
    ("2022-01-01",  34),
    ("2022-07-01",  38),
    ("2023-01-01",  42),
    ("2023-07-01",  46),
    ("2024-01-01",  50),
    ("2024-07-01",  53),
    ("2025-01-01",  55),
    ("2025-07-01",  58),
    ("2026-01-01",  60),
]

# =============================
# HELPERS
# =============================

def parse_date(value):
    # accept a date or string return an object
    if value is None or value=="":
        return None
    if isinstance(value,datetime):
        return value.date()
    if isinstance(value,date):
        return value
    
    return datetime.strptime(str(value).strip(), "%Y-%m-%d").date()

def days_in_month(year,month):
    if month==12:
        return 31
    return (date(year,month +1,1)- date(year,month,1)).days

def month_start(d):
    return date(d.year,d.month,1)

def month_end(d):
    return date(d.year,d.month,days_in_month(d.year,d.month))

def round0(x):
    return int(round(x))

# =============================
# INPUT LOADER
# =============================

def load_workbook(path):
    emp_df= pd.read_excel(path,sheet_name="Employee")
    events_df= pd.read_excel(path,sheet_name="Events")
    hra_df=pd.read_excel(path,sheet_name="HRA_Master") 

    field_col = emp_df["Field"].astype(str).str.strip()
    value_col = emp_df["Value"]

    emp_dict = dict(zip(field_col, value_col))

    return emp_dict, events_df, hra_df


# ==============================
# EVENT EXPANSION
# ==============================

def read_events(df):
    """Convert the events DataFrame into a list of dicts."""
    events = []
    for _, row in df.iterrows():
        from_d = parse_date(row["From"])
        if from_d is None:
            continue  # skip rows without a start date

        to_d = parse_date(row["To"]) if pd.notna(row["To"]) else None

        events.append({
            "from":    from_d,
            "to":      to_d,  # may be None; resolved later
            "bDue":    float(row["BasicDue"])   if pd.notna(row["BasicDue"])   else 0.0,
            "gDue":    float(row["GPDue"])      if pd.notna(row["GPDue"])      else 0.0,
            "bDrawn":  float(row["BasicDrawn"]) if pd.notna(row["BasicDrawn"]) else 0.0,
            "gDrawn":  float(row["GPDrawn"])    if pd.notna(row["GPDrawn"])    else 0.0,
            "station": str(row["Station"]).strip() if pd.notna(row["Station"]) else "",
            "remarks": str(row["Remarks"]).strip() if pd.notna(row["Remarks"]) else "",
        })
    return events


def resolve_event_ranges(events, period_to):
    """Extend each event's end to the day before the next event begins."""
    sorted_events = sorted(events, key=lambda e: e["from"])
    resolved = []

    for i, ev in enumerate(sorted_events):
        start = ev["from"]

        if ev["to"] is not None:
            end = ev["to"]
        elif i < len(sorted_events) - 1:
            next_start = sorted_events[i + 1]["from"]
            end = next_start - timedelta(days=1)
        else:
            end = period_to

        # Cap at day-before-next if explicit 'to' overshoots
        if i < len(sorted_events) - 1:
            next_start = sorted_events[i + 1]["from"]
            cap = next_start - timedelta(days=1)
            if end > cap:
                end = cap

        resolved.append({**ev, "from": start, "to": end})

    return resolved


def expand_event_to_months(event):
    """Turn one resolved event into one dict per month covered."""
    out = []
    start = event["from"]
    end   = event["to"]

    # Start cursor at the first of the start month
    cursor = date(start.year, start.month, 1)

    while cursor <= end:
        y, m = cursor.year, cursor.month
        dim  = days_in_month(y, m)
        mstart = date(y, m, 1)
        mend   = date(y, m, dim)

        seg_start = start if start > mstart else mstart
        seg_end   = end   if end   < mend   else mend

        days_covered = (seg_end - seg_start).days + 1
        fraction     = days_covered / dim

        seg_label = ""
        if days_covered < dim:
            seg_label = (
                f"{seg_start.day:02d}-{seg_start.strftime('%b')} "
                f"to {seg_end.day:02d}-{seg_end.strftime('%b')}"
            )

        out.append({
            "month":         date(y, m, 1),
            "month_label":   mstart.strftime("%b %Y"),
            "segment_label": seg_label,
            "days_covered":  days_covered,
            "days_in_month": dim,
            "fraction":      fraction,
            "seg_start":     seg_start,
            "bDue":          event["bDue"],
            "gDue":          event["gDue"],
            "bDrawn":        event["bDrawn"],
            "gDrawn":        event["gDrawn"],
            "station":       event["station"],
            "remarks":       event["remarks"],
        })

        # Advance to first of next month
        if m == 12:
            cursor = date(y + 1, 1, 1)
        else:
            cursor = date(y, m + 1, 1)

    return out


def expand_events(df, period_to):
    """Top-level: events DataFrame -> full list of monthly dicts."""
    events = read_events(df)
    resolved = resolve_event_ranges(events, period_to)

    months = []
    for ev in resolved:
        months.extend(expand_event_to_months(ev))

    months.sort(key=lambda x: (x["month"], x["seg_start"]))
    return months


# ============================================================
# LOOKUPS
# ============================================================

def build_da_master():
    """Return a sorted list of (date, rate) from DA_TABLE.
    If two entries share a date, the later one wins."""
    by_date = {}
    for eff, rate in DA_TABLE:
        by_date[parse_date(eff)] = rate
    return sorted(by_date.items())


def get_da(month, da_master):
    """Return DA% for the given month (latest effective <= month)."""
    rate = 0
    for eff, r in da_master:
        if eff <= month:
            rate = r
        else:
            break
    return rate


def get_hra(station, hra_df):
    """Return HRA% for the given station, case-insensitive."""
    if not station:
        return 0
    target = station.strip().lower()
    for _, row in hra_df.iterrows():
        if str(row["Station"]).strip().lower() == target:
            return float(row["HRA %"])
    return 0

# ============================================================
# COMPUTE ROWS
# ============================================================

def compute_rows(months, da_master, hra_df):
    """For each month, look up rates and compute rupee amounts."""
    rows = []

    for m in months:
        da_pct  = get_da(m["month"], da_master)
        hra_pct = get_hra(m["station"], hra_df)

        # Due side (what should have been paid)
        b_due   = round0(m["bDue"]   * m["fraction"])
        g_due   = round0(m["gDue"]   * m["fraction"])
        d_da    = round0((b_due + g_due) * da_pct  / 100)
        d_hra   = round0((b_due + g_due) * hra_pct / 100)
        d_total = b_due + g_due + d_da + d_hra

        # Drawn side (what was actually paid)
        b_drawn = round0(m["bDrawn"] * m["fraction"])
        g_drawn = round0(m["gDrawn"] * m["fraction"])
        w_da    = round0((b_drawn + g_drawn) * da_pct  / 100)
        w_hra   = round0((b_drawn + g_drawn) * hra_pct / 100)
        w_total = b_drawn + g_drawn + w_da + w_hra

        rows.append({
            "month":         m["month"],
            "month_label":   m["month_label"],
            "segment_label": m["segment_label"],
            "da_pct":        da_pct,
            "hra_pct":       hra_pct,

            # Due side
            "b_due":   b_due,
            "g_due":   g_due,
            "d_da":    d_da,
            "d_hra":   d_hra,
            "d_total": d_total,

            # Drawn side
            "b_drawn": b_drawn,
            "g_drawn": g_drawn,
            "w_da":    w_da,
            "w_hra":   w_hra,
            "w_total": w_total,

            # Difference
            "diff": d_total - w_total,
        })

    return rows

# ============================================================
# EXCEL WRITER
# ============================================================

def write_excel(rows, emp, output_path):
    """Write the statement to a formatted .xlsx file."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "Statement"

    # ---- Style objects ----
    thin     = Side(style="thin", color="888888")
    border   = Border(left=thin, right=thin, top=thin, bottom=thin)
    hdr_fill = PatternFill("solid", fgColor="DDDDDD")
    bold     = Font(bold=True, size=14)
    normal   = Font(size=14)
    center   = Alignment(horizontal="center", vertical="center", wrap_text=True)
    right    = Alignment(horizontal="right",  vertical="center")
    left     = Alignment(horizontal="left",   vertical="center")

    # ---- Title block (rows 1-2) ----
    title = (f"{emp['StatementTitle']} OF {str(emp['Name']).upper()}, "
             f"{str(emp['Designation']).upper()}")
    ws.merge_cells("A1:N1")
    ws["A1"] = title
    ws["A1"].font = Font(bold=True, size=14)
    ws["A1"].alignment = center

    pfrom = parse_date(emp["PeriodFrom"])
    pto   = parse_date(emp["PeriodTo"])
    period = (f"FOR THE PERIOD FROM {pfrom.strftime('%d/%m/%Y')} "
              f"TO {pto.strftime('%d/%m/%Y')}")
    ws.merge_cells("A2:N2")
    ws["A2"] = period
    ws["A2"].font = Font(size=12)
    ws["A2"].alignment = center

    # ---- Header (rows 4-5) ----
    row1 = ["MONTH", "AMOUNT DUE", None, None, None, None,
                     "AMOUNT DRAWN", None, None, None, None,
                     "DIFFERENCE", "DA %", "HRA %"]
    row2 = ["", "Basic", "GP", "DA", "HRA", "Total",
                "Basic", "GP", "DA", "HRA", "Total", "", "", ""]

    for col, val in enumerate(row1, start=1):
        cell = ws.cell(row=4, column=col, value=val)
        cell.font = bold
        cell.fill = hdr_fill
        cell.alignment = center
        cell.border = border

    for col, val in enumerate(row2, start=1):
        cell = ws.cell(row=5, column=col, value=val)
        cell.font = bold
        cell.fill = hdr_fill
        cell.alignment = center
        cell.border = border

    ws.merge_cells("B4:F4")
    ws.merge_cells("G4:K4")

    # Re-apply borders and fill to merged cells
    for r in (4, 5):
        for c in range(1, 15):
            ws.cell(row=r, column=c).border = border
            ws.cell(row=r, column=c).fill   = hdr_fill

    # ---- Data rows (start row 6) ----
    row_idx = 6
    for r in rows:
        label = r["month_label"]
        if r["segment_label"]:
            label += f" ({r['segment_label']})"

        values = [
            label,
            r["b_due"], r["g_due"], r["d_da"], r["d_hra"], r["d_total"],
            r["b_drawn"], r["g_drawn"], r["w_da"], r["w_hra"], r["w_total"],
            r["diff"], r["da_pct"], r["hra_pct"],
        ]

        for col, val in enumerate(values, start=1):
            cell = ws.cell(row=row_idx, column=col, value=val)
            cell.border = border
            cell.font = normal
            cell.alignment = left if col == 1 else right

        row_idx += 1

    # ---- Totals row ----
    def total(key):
        return sum(r[key] for r in rows)

    totals = [
        "TOTAL",
        total("b_due"), total("g_due"), total("d_da"), total("d_hra"), total("d_total"),
        total("b_drawn"), total("g_drawn"), total("w_da"), total("w_hra"), total("w_total"),
        total("diff"), "", ""
    ]
    for col, val in enumerate(totals, start=1):
        cell = ws.cell(row=row_idx, column=col, value=val)
        cell.border = border
        cell.font = bold
        cell.fill = hdr_fill
        cell.alignment = left if col == 1 else right

    # ---- Column widths ----
    widths = [32, 10, 9, 10, 10, 11, 10, 9, 10, 10, 11, 12, 8, 8]
    for col, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(col)].width = w

    # ---- Summary block (stops at Gross Difference) ----
    row_idx += 2

    diff_basic = total("b_due") - total("b_drawn")
    diff_gp    = total("g_due") - total("g_drawn")
    diff_da    = total("d_da")  - total("w_da")
    diff_hra   = total("d_hra") - total("w_hra")
    gross      = total("diff")

    summary = [
        ("Basic Difference",     diff_basic),
        ("Grade Pay Difference", diff_gp),
        ("DA Difference",        diff_da),
        ("HRA Difference",       diff_hra),
        ("Gross Difference",     gross),
    ]
    for label, val in summary:
        ws.cell(row=row_idx, column=1, value=label).font = bold
        c = ws.cell(row=row_idx, column=2, value=val)
        c.alignment = right
        c.font = normal
        row_idx += 1

    # ---- Certification line ----
    row_idx += 1
    ws.merge_cells(start_row=row_idx, start_column=1,
                   end_row=row_idx, end_column=14)
    cert = ws.cell(row=row_idx, column=1,
                   value="Certified that the above amount has not been drawn and paid previously.")
    cert.font = normal
    cert.alignment = left

    # ---- Save ----
    wb.save(output_path)
    return {
        "rows":       len(rows),
        "diff_basic": diff_basic,
        "diff_gp":    diff_gp,
        "diff_da":    diff_da,
        "diff_hra":   diff_hra,
        "gross":      gross,
    }

# ============================================================
# PIPELINE
# ============================================================

def run(input_path, output_path):
    """Full pipeline: load -> expand -> compute -> write."""
    print(f"Loading: {input_path}")
    emp, events_df, hra_df = load_workbook(input_path)

    period_to = parse_date(emp["PeriodTo"])
    if period_to is None:
        raise ValueError("PeriodTo in the Employee sheet is missing or invalid.")

    print("Building DA master...")
    da_master = build_da_master()

    print(f"Expanding {len(events_df)} events...")
    months = expand_events(events_df, period_to)
    print(f"  -> {len(months)} monthly rows")

    print("Computing amounts...")
    rows = compute_rows(months, da_master, hra_df)

    print(f"Writing: {output_path}")
    result = write_excel(rows, emp, output_path)

    # Final summary
    print()
    print("=" * 54)
    print(f"  Employee           : {emp['Name']}")
    print(f"  Period             : {parse_date(emp['PeriodFrom'])} to {period_to}")
    print(f"  Monthly rows       : {result['rows']}")
    print("-" * 54)
    print(f"  Basic Difference   : {result['diff_basic']:>12,}")
    print(f"  Grade Pay Diff     : {result['diff_gp']:>12,}")
    print(f"  DA Difference      : {result['diff_da']:>12,}")
    print(f"  HRA Difference     : {result['diff_hra']:>12,}")
    print(f"  Gross Difference   : {result['gross']:>12,}")
    print("=" * 54)

    return result

# ============================================================
# CONFIG LOADER
# ============================================================

def load_config(path="config.json"):
    """Load config.json if it exists; otherwise return defaults."""
    defaults = {
        "input_path":       DEFAULT_INPUT,
        "output_path":      DEFAULT_OUTPUT,
        "batch_input_dir":  "data",
        "batch_output_dir": "output",
    }
    if not os.path.exists(path):
        return defaults

    with open(path, "r") as f:
        user = json.load(f)

    return {**defaults, **user}


# ============================================================
# MAIN
# ============================================================

def main():
    cfg = load_config()

    # ---- Batch mode ----
    if len(sys.argv) > 1 and sys.argv[1] == "--batch":
        os.makedirs(cfg["batch_output_dir"], exist_ok=True)
        run_batch(cfg["batch_input_dir"], cfg["batch_output_dir"])
        return

    # ---- Single-file mode ----
    input_path  = sys.argv[1] if len(sys.argv) > 1 else cfg["input_path"]

    # Derive output filename from the input file's basename
    name = os.path.splitext(os.path.basename(input_path))[0]
    output_dir = os.path.dirname(cfg["output_path"]) or "output"
    output_path = os.path.join(output_dir, f"{name}_statement.xlsx")

    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)

    if not os.path.exists(input_path):
        print(f"Error: input file not found: {input_path}", file=sys.stderr)
        sys.exit(1)

    try:
        run(input_path, output_path)
    except FileNotFoundError as e:
        print(f"Error: file not found: {e}", file=sys.stderr)
        sys.exit(1)
    except KeyError as e:
        print(f"Error: missing column or field in workbook: {e}", file=sys.stderr)
        sys.exit(1)
    except ValueError as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)
   

if __name__ == "__main__":
    main()
    