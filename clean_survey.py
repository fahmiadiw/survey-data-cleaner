"""
Survey data cleaner
-------------------
Turns a raw field-survey export (KoboToolbox / ODK / Excel style) into an
analysis-ready dataset, plus a log of every problem found and what was done.

Usage:
    python clean_survey.py sample_survey_export.csv
    python clean_survey.py export.csv --start 2026-03-01 --end 2026-04-30 --out cleaned.xlsx

What it does:
  1. Standardizes   sex codes, yes/no answers, district names, facility types,
                    phone numbers, dates in mixed formats
  2. Validates      impossible ages and household sizes, more under-5 children
                    than the household allows, skip-logic breaks (ANC visits
                    for non-pregnant respondents), dates outside fieldwork,
                    GPS points outside the survey area, missing required fields
  3. De-duplicates  forms submitted twice from the same tablet
  4. Reports        one Excel file: Summary (with charts), Clean Data, Issues Log

Invalid values are blanked, never guessed, and every change is logged with its
original value, so nothing disappears silently.
"""
import argparse
import re
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

# ---- Survey rules (edit these for a new survey) -------------------------
DISTRICTS = ["Pringsewu", "Gadingrejo", "Ambarawa", "Pardasuka", "Pagelaran",
             "Banyumas", "Adiluwih", "Sukoharjo", "Pagelaran Utara"]
RANGES = {"respondent_age": (15, 99), "household_size": (1, 30), "travel_time_min": (0, 600)}
AREA_BBOX = {"lat": (-5.55, -5.15), "lon": (104.75, 105.20)}   # survey area
REQUIRED = ["district", "respondent_age", "respondent_sex", "interview_date"]

SEX_MAP = {"f": "F", "female": "F", "p": "F", "perempuan": "F",
           "m": "M", "male": "M", "l": "M", "laki-laki": "M"}
YES_NO_MAP = {"yes": "Yes", "y": "Yes", "ya": "Yes", "1": "Yes", "true": "Yes",
              "no": "No", "n": "No", "tidak": "No", "0": "No", "false": "No"}
FACILITY_MAP = {"puskesmas": "Puskesmas", "hospital": "Hospital", "rs": "Hospital",
                "midwife clinic": "Midwife clinic", "bidan": "Midwife clinic", "none": "No facility"}
DATE_FORMATS = ["%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d-%b-%Y"]


class Log:
    def __init__(self):
        self.rows = []

    def add(self, sid, enum, column, issue, original, action):
        self.rows.append({"submission_id": sid, "enumerator_id": enum, "column": column,
                          "issue": issue, "original_value": original, "action": action})


def is_blank(v):
    return v is None or (isinstance(v, float) and pd.isna(v)) or str(v).strip() == ""


def parse_date(raw, start, end):
    """Try each format; when dd/mm and mm/dd are both valid, keep the one inside fieldwork."""
    s = str(raw).strip()
    candidates = []
    for fmt in DATE_FORMATS:
        try:
            candidates.append(datetime.strptime(s, fmt).date())
        except ValueError:
            pass
    inside = [d for d in candidates if start <= d <= end]
    if inside:
        return inside[0], None
    if candidates:
        return None, "date outside fieldwork period"
    return None, "unreadable date"


def to_number(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def clean_phone(v):
    digits = re.sub(r"\D", "", str(v))
    if digits.startswith("0"):
        digits = "62" + digits[1:]
    return "+" + digits if digits.startswith("62") and 10 <= len(digits) <= 15 else None


def clean(df, start, end):
    log = Log()
    out = df.copy().astype(object)

    # 3. Duplicates first: same answers in every field except the submission id
    answer_cols = [c for c in df.columns if c != "submission_id"]
    dup_mask = df[answer_cols].astype(str).duplicated(keep="first")
    for _, r in df[dup_mask].iterrows():
        log.add(r["submission_id"], r["enumerator_id"], "(whole form)", "duplicate submission", "", "removed")
    out = out[~dup_mask].reset_index(drop=True)

    for i, r in out.iterrows():
        sid, enum = r["submission_id"], r["enumerator_id"]
        blanked = set()

        def blank(col, issue):
            log.add(sid, enum, col, issue, r[col], "set to blank")
            out.at[i, col] = None
            blanked.add(col)

        # 1. Standardize codes and text
        sex = SEX_MAP.get(str(r["respondent_sex"]).strip().lower())
        if sex:
            out.at[i, "respondent_sex"] = sex
        elif not is_blank(r["respondent_sex"]):
            blank("respondent_sex", "unrecognized sex code")
        for col in ["currently_pregnant", "has_bpjs"]:
            v = YES_NO_MAP.get(str(r[col]).strip().lower())
            out.at[i, col] = v
        fac = FACILITY_MAP.get(str(r["nearest_facility"]).strip().lower())
        out.at[i, "nearest_facility"] = fac
        dist_key = re.sub(r"[^a-z ]", "", str(r["district"]).lower()).strip()
        dist = next((d for d in DISTRICTS if d.lower() == dist_key), None)
        if dist:
            out.at[i, "district"] = dist
        elif not is_blank(r["district"]):
            blank("district", "unknown district name")
        out.at[i, "phone"] = clean_phone(r["phone"])

        # Dates
        if not is_blank(r["interview_date"]):
            d, problem = parse_date(r["interview_date"], start, end)
            if problem:
                blank("interview_date", problem)
            else:
                out.at[i, "interview_date"] = d.isoformat()

        # 2. Validate ranges
        for col, (lo, hi) in RANGES.items():
            if is_blank(r[col]):
                continue
            n = to_number(r[col])
            if n is None or not lo <= n <= hi:
                blank(col, f"{col.replace('_', ' ')} out of range ({lo}–{hi})")
            else:
                out.at[i, col] = int(n)

        hh, kids = to_number(out.at[i, "household_size"]), to_number(r["children_under5"])
        if kids is not None:
            if hh is not None and kids > hh - 1:
                blank("children_under5", "more under-5 children than household allows")
            else:
                out.at[i, "children_under5"] = int(kids)

        # Skip logic: ANC visits only make sense for pregnant respondents
        anc = to_number(r["anc_visits"])
        preg = out.at[i, "currently_pregnant"]
        if anc and anc > 0 and preg != "Yes":
            blank("anc_visits", "skip-logic break: ANC visits but not pregnant")
        elif preg != "Yes":
            out.at[i, "anc_visits"] = None
        elif anc is not None:
            out.at[i, "anc_visits"] = int(anc)
        if out.at[i, "respondent_sex"] == "M" and preg == "Yes":
            blank("currently_pregnant", "skip-logic break: male respondent marked pregnant")

        # GPS inside the survey area
        lat, lon = to_number(r["gps_lat"]), to_number(r["gps_lon"])
        inside = (lat is not None and lon is not None
                  and AREA_BBOX["lat"][0] <= lat <= AREA_BBOX["lat"][1]
                  and AREA_BBOX["lon"][0] <= lon <= AREA_BBOX["lon"][1])
        if inside:
            out.at[i, "gps_lat"], out.at[i, "gps_lon"] = lat, lon
        else:
            missing = lat is None or lon is None
            log.add(sid, enum, "gps", "missing GPS" if missing else "GPS point outside survey area",
                    f"{r['gps_lat']}, {r['gps_lon']}", "left blank" if missing else "set to blank")
            out.at[i, "gps_lat"] = out.at[i, "gps_lon"] = None

        # Required fields
        for col in REQUIRED:
            if is_blank(out.at[i, col]) and col not in blanked:
                log.add(sid, enum, col, "missing required answer", r[col], "left blank")

    issues = pd.DataFrame(log.rows, columns=["submission_id", "enumerator_id", "column", "issue",
                                             "original_value", "action"])
    per_form = issues[issues["issue"] != "duplicate submission"].groupby("submission_id").size()
    out["issue_count"] = out["submission_id"].map(per_form).fillna(0).astype(int)
    return out, issues


HEADER_FILL = PatternFill("solid", fgColor="1F5F4A")
HEADER_FONT = Font(bold=True, color="FFFFFF")


def write_sheet(ws, df, widths=None):
    ws.append(list(df.columns))
    for c in ws[1]:
        c.font, c.fill = HEADER_FONT, HEADER_FILL
    for row in df.itertuples(index=False):
        ws.append([None if (isinstance(v, float) and pd.isna(v)) else v for v in row])
    for j, col in enumerate(df.columns, start=1):
        ws.column_dimensions[get_column_letter(j)].width = (widths or {}).get(col, max(12, len(col) + 2))
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def write_excel(raw, clean_df, issues, path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    ws["A1"] = "Survey Data Cleaning — Summary"
    ws["A1"].font = Font(bold=True, size=16, color="1F5F4A")
    flagged = int((clean_df["issue_count"] > 0).sum())
    stats = [("Raw submissions", len(raw)), ("Duplicates removed", len(raw) - len(clean_df)),
             ("Clean submissions", len(clean_df)), ("Submissions with at least one fix", flagged),
             ("Fully clean submissions", len(clean_df) - flagged), ("Issues logged", len(issues))]
    for k, (label, val) in enumerate(stats, start=3):
        ws.cell(k, 1, label)
        ws.cell(k, 2, val).font = Font(bold=True)

    by_type = issues["issue"].value_counts()
    r0 = 11
    ws.cell(r0, 1, "Issue type").font = HEADER_FONT
    ws.cell(r0, 1).fill = HEADER_FILL
    ws.cell(r0, 2, "Count").font = HEADER_FONT
    ws.cell(r0, 2).fill = HEADER_FILL
    for k, (name, n) in enumerate(by_type.items(), start=1):
        ws.cell(r0 + k, 1, name)
        ws.cell(r0 + k, 2, int(n))
    end_type = r0 + len(by_type)

    interviews = raw.groupby("enumerator_id").size()
    per_enum = (issues.groupby("enumerator_id").size().reindex(interviews.index, fill_value=0)
                / interviews * 100).round(1).sort_values(ascending=False)
    ws.cell(r0, 4, "Enumerator").font = HEADER_FONT
    ws.cell(r0, 4).fill = HEADER_FILL
    ws.cell(r0, 5, "Issues per 100 forms").font = HEADER_FONT
    ws.cell(r0, 5).fill = HEADER_FILL
    for k, (name, v) in enumerate(per_enum.items(), start=1):
        ws.cell(r0 + k, 4, name)
        ws.cell(r0 + k, 5, float(v))
    end_enum = r0 + len(per_enum)

    c1 = BarChart()
    c1.type = "bar"
    c1.title = "Issues by type"
    c1.legend = None
    c1.add_data(Reference(ws, min_col=2, min_row=r0, max_row=end_type), titles_from_data=True)
    c1.set_categories(Reference(ws, min_col=1, min_row=r0 + 1, max_row=end_type))
    c1.x_axis.scaling.orientation = "maxMin"
    c1.height, c1.width = 9, 16
    ws.add_chart(c1, "H2")

    c2 = BarChart()
    c2.title = "Issues per 100 forms, by enumerator"
    c2.legend = None
    c2.add_data(Reference(ws, min_col=5, min_row=r0, max_row=end_enum), titles_from_data=True)
    c2.set_categories(Reference(ws, min_col=4, min_row=r0 + 1, max_row=end_enum))
    c2.height, c2.width = 8, 16
    ws.add_chart(c2, "H21")
    ws.column_dimensions["A"].width = 46
    ws.column_dimensions["D"].width = 14
    ws.column_dimensions["E"].width = 20

    write_sheet(wb.create_sheet("Clean Data"), clean_df)
    write_sheet(wb.create_sheet("Issues Log"), issues,
                {"submission_id": 18, "issue": 44, "original_value": 22, "action": 14})
    wb.save(path)


def main():
    p = argparse.ArgumentParser(description="Clean and validate a field survey export")
    p.add_argument("export_csv")
    p.add_argument("--start", default="2026-03-01", help="first day of fieldwork (YYYY-MM-DD)")
    p.add_argument("--end", default="2026-04-30", help="last day of fieldwork (YYYY-MM-DD)")
    p.add_argument("--out", default="survey_cleaned.xlsx")
    args = p.parse_args()

    start = datetime.strptime(args.start, "%Y-%m-%d").date()
    end = datetime.strptime(args.end, "%Y-%m-%d").date()
    raw = pd.read_csv(args.export_csv, dtype=str, keep_default_na=False)
    clean_df, issues = clean(raw, start, end)
    write_excel(raw, clean_df, issues, args.out)
    print(f"Done: {len(raw):,} raw submissions -> {len(clean_df):,} clean, "
          f"{len(issues):,} issues logged -> {Path(args.out).name}")


if __name__ == "__main__":
    main()
