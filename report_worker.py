"""
report_worker.py
================
Runs as a subprocess: reads sys.argv[1] (output xlsx path),
computes the dashboard summary, and prints the result as JSON to stdout.

Called by app.py to isolate pandas/openpyxl from Flask's WSGI threads,
which avoids a Python 3.14 segfault when openpyxl is used in a non-main thread.
"""
import json
import sys
from pathlib import Path

import pandas as pd


def _safe_float(val) -> float:
    try:
        return float(val)
    except Exception:
        return 0.0


def summarise(output_path: Path) -> dict:
    if not output_path.exists():
        return {"available": False}

    xl     = pd.ExcelFile(output_path)
    sheets = xl.sheet_names
    result: dict = {"available": True, "sheets": sheets}

    if "Details" not in sheets:
        return result

    df = pd.read_excel(output_path, sheet_name="Details")
    df.columns = [str(c).strip() for c in df.columns]

    # ── KPIs ──────────────────────────────────────────────
    total_hours  = 0.0
    dept_count   = 0
    staff_count  = 0
    completion   = None

    if "Hours Worked" in df.columns:
        total_hours = _safe_float(pd.to_numeric(df["Hours Worked"], errors="coerce").sum())
    if "Department" in df.columns:
        dept_count = int(df["Department"].nunique())
    if "Staff Name" in df.columns:
        staff_count = int(df["Staff Name"].nunique())
    if "Status" in df.columns:
        done = df["Status"].astype(str).str.lower().str.contains("complet").mean()
        completion = round(_safe_float(done) * 100, 1)

    result["kpis"] = {
        "total_rows":          int(len(df)),
        "total_hours":         round(total_hours, 1),
        "dept_count":          dept_count,
        "staff_count":         staff_count,
        "completion_pct":      completion,
        "avg_hours_per_person": round(total_hours / staff_count, 1) if staff_count else None,
    }

    # ── Department bar chart ──────────────────────────────
    if "Department" in df.columns and "Hours Worked" in df.columns:
        # Filter out artifact rows (empty Department from array-formula cells)
        dff = df[df["Department"].notna() & (df["Department"].astype(str).str.strip() != "")]
        dg = (
            dff.groupby("Department")["Hours Worked"]
            .apply(lambda s: pd.to_numeric(s, errors="coerce").sum())
            .sort_values(ascending=False).round(1)
        )
        # Only include departments that actually logged hours
        dg = dg[dg > 0]
        result["dept_hours"] = {
            "labels": dg.index.tolist(),
            "values": [_safe_float(v) for v in dg.values],
        }

    # ── Status donut ──────────────────────────────────────
    if "Status" in df.columns:
        sg = df["Status"].astype(str).str.strip().value_counts()
        result["status_dist"] = {
            "labels": sg.index.tolist(),
            "values": [int(v) for v in sg.values],
        }

    # ── Project bar chart ─────────────────────────────────
    if "Project" in df.columns and "Hours Worked" in df.columns:
        pg = (
            df.groupby("Project")["Hours Worked"]
            .apply(lambda s: pd.to_numeric(s, errors="coerce").sum())
            .sort_values(ascending=False).head(10).round(1)
        )
        result["project_hours"] = {
            "labels": pg.index.tolist(),
            "values": [_safe_float(v) for v in pg.values],
        }

    # ── Hours trend by date ───────────────────────────────
    if "Date" in df.columns and "Hours Worked" in df.columns:
        dfc = df.copy()
        dfc["_date"] = pd.to_datetime(dfc["Date"], errors="coerce")
        trend = (
            dfc.dropna(subset=["_date"])
            .groupby("_date")["Hours Worked"]
            .apply(lambda s: pd.to_numeric(s, errors="coerce").sum())
            .sort_index().round(1)
        )
        result["hours_trend"] = {
            "labels": [d.strftime("%d %b") for d in trend.index],
            "values": [_safe_float(v) for v in trend.values],
        }

    # ── Department table ──────────────────────────────────
    if "Department" in df.columns:
        dff = df[df["Department"].notna() & (df["Department"].astype(str).str.strip() != "")]
        agg: dict = {}
        if "Hours Worked" in dff.columns:
            agg["Total Hours"] = pd.NamedAgg(
                column="Hours Worked",
                aggfunc=lambda s: round(_safe_float(pd.to_numeric(s, errors="coerce").sum()), 1),
            )
        if "Staff Name" in dff.columns:
            agg["Staff"] = pd.NamedAgg(column="Staff Name", aggfunc="nunique")
        agg["Rows"] = pd.NamedAgg(column="Department", aggfunc="count")

        dt = dff.groupby("Department").agg(**agg).reset_index()
        if "Total Hours" in dt.columns:
            dt = dt.sort_values("Total Hours", ascending=False)
        result["dept_table"] = json.loads(dt.to_json(orient="records", default_handler=str))

    # ── Recent activity (last 50 rows) ────────────────────
    keep = [c for c in ["Date","Staff Name","Department","Activity","Hours Worked","Status"]
            if c in df.columns]
    recent = df[keep].tail(50).copy()

    # Normalise Date column: convert timestamps/epoch-ms to readable strings
    if "Date" in recent.columns:
        def _fmt_date(v):
            if v is None or (isinstance(v, float) and pd.isna(v)) or v == "":
                return ""
            try:
                return pd.to_datetime(v, unit="ms" if isinstance(v, (int, float)) and v > 1e10 else None).strftime("%Y-%m-%d")
            except Exception:
                return str(v)
        recent["Date"] = recent["Date"].apply(_fmt_date)

    # Drop rows that are entirely empty (artifact rows from array-formula cells)
    recent = recent[recent.apply(
        lambda r: any(str(x).strip() not in ("", "nan") for x in r), axis=1
    )]

    result["recent_rows"] = json.loads(recent.fillna("").to_json(orient="records", default_handler=str))

    return result


if __name__ == "__main__":
    path = Path(sys.argv[1]) if len(sys.argv) > 1 else Path()
    try:
        data = summarise(path)
    except Exception as exc:
        data = {"available": False, "error": str(exc)}
    # Write JSON to stdout; Flask reads it back
    print(json.dumps(data, default=str))
