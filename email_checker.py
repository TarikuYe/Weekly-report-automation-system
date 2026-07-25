"""
email_checker.py
================
Scans department input folders at the *per-employee* level and orchestrates
email delivery.

Data model
----------
config.json  →  departments[]
                  name        : display name  (e.g. "Contract")
                  folder      : sub-folder inside input_dir  (e.g. "Contract")
                  cc_email    : optional dept-head CC address
                  employees[]
                    name          : employee display name
                    email         : employee email address
                    expected_file : filename to look for (blank = any .xlsx counts)

Submission logic
----------------
If  expected_file is set   → employee is "submitted" only if that exact filename
                              exists in the department folder.
If  expected_file is blank → employee is "submitted" if ANY .xlsx file exists
                              in the department folder that is not already claimed
                              by another employee who has an expected_file set.

Key functions
-------------
check_submissions(cfg)
    Returns submitted / missing lists at the employee level.

run_email_cycle(cfg, ...)
    Full automated cycle: scan → ack submitters → remind missing.

is_past_deadline(cfg) / next_deadline(cfg)
    Deadline helpers.

schedule_auto_check(cfg_path, interval_minutes)
    Background scheduler thread.
"""

from __future__ import annotations

import json
import logging
import threading
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any

from email_sender import send_acknowledgment, send_reminder, week_label_for

log = logging.getLogger("email_checker")

_WEEKDAYS = {
    "monday": 0, "tuesday": 1, "wednesday": 2, "thursday": 3,
    "friday": 4, "saturday": 5, "sunday": 6,
}

# Per-process dedup: keyed by  week_label + "|" + employee_email
_sent_keys: set[str] = set()
_sent_lock = threading.Lock()


# ---------------------------------------------------------------------------#
# Config
# ---------------------------------------------------------------------------#
def _load_config(cfg_path: Path | str | None = None) -> dict:
    if cfg_path is None:
        return {}
    p = Path(cfg_path)
    if not p.exists():
        return {}
    with open(p, "r", encoding="utf-8") as f:
        return json.load(f)


# ---------------------------------------------------------------------------#
# Path helpers
# ---------------------------------------------------------------------------#
def _resolve_dept_folder(input_dir: Path, folder_value: str) -> Path | None:
    """Return the absolute Path for a department's input folder.

    Rules (backward-compatible):
    - If *folder_value* is empty → return None (department has no folder set).
    - If *folder_value* is an absolute path → use it directly.
    - Otherwise treat it as a sub-folder name under *input_dir*.

    This lets existing configs (where ``folder`` is just "Contract") keep
    working, while new configs can point to any location on the server
    (e.g. "\\\\SERVER\\shares\\Contract").
    """
    if not folder_value:
        return None
    p = Path(folder_value)
    if p.is_absolute():
        return p
    return input_dir / folder_value


# ---------------------------------------------------------------------------#
# File helpers
# ---------------------------------------------------------------------------#
def _xlsx_files_in(folder: Path) -> list[Path]:
    """All non-temp .xlsx files directly inside *folder* (not recursive)."""
    if not folder.is_dir():
        return []
    return sorted(
        p for p in folder.glob("*.xlsx")
        if not p.name.startswith("~$")
    )


def _match_file(
    files: list[Path],
    expected_file: str,
    claimed: set[str],
) -> Path | None:
    """Return the matched Path or None.

    - If expected_file is set: look for that exact name (case-insensitive).
    - If expected_file is blank: return the first unclaimed file.
    """
    if expected_file:
        target = expected_file.strip().lower()
        for f in files:
            if f.name.lower() == target:
                return f
        return None
    # blank expected_file — first unclaimed file
    for f in files:
        if f.name not in claimed:
            return f
    return None


# ---------------------------------------------------------------------------#
# Core scan
# ---------------------------------------------------------------------------#
def check_submissions(cfg: dict) -> dict[str, Any]:
    """Scan folders and return per-employee submission status.

    Return schema
    -------------
    {
      "week_label":    str,
      "checked_at":   str,
      "departments": [
        {
          "name":       str,
          "folder":     str,
          "cc_email":   str,
          "submitted":  [ EmployeeResult, ... ],
          "missing":    [ EmployeeResult, ... ],
        }, ...
      ],
      "all_submitted": [ EmployeeResult, ... ],   # flat list
      "all_missing":   [ EmployeeResult, ... ],   # flat list
      "submit_count":  int,
      "missing_count": int,
      "total":         int,
      "input_dir":     str,
    }

    EmployeeResult (submitted)
    --------------------------
    { dept, dept_folder, name, email, expected_file,
      matched_file, modified }

    EmployeeResult (missing)
    ------------------------
    { dept, dept_folder, name, email, expected_file }
    """
    input_dir   = Path(cfg.get("input_dir", "input"))
    departments = cfg.get("departments", [])
    week_label  = week_label_for()

    dept_results: list[dict] = []
    all_submitted: list[dict] = []
    all_missing:   list[dict] = []

    for dept in departments:
        dept_name   = dept.get("name", "Unknown")
        folder_name = dept.get("folder", "")
        cc_email    = dept.get("cc_email", "")
        employees   = dept.get("employees", [])

        dept_folder = _resolve_dept_folder(input_dir, folder_name)
        files       = _xlsx_files_in(dept_folder) if dept_folder else []

        # Files claimed by employees who have an explicit expected_file
        claimed: set[str] = {
            e["expected_file"].strip()
            for e in employees
            if e.get("expected_file", "").strip()
        }

        dept_submitted: list[dict] = []
        dept_missing:   list[dict] = []

        for emp in employees:
            emp_name  = emp.get("name", "").strip()
            emp_email = emp.get("email", "").strip()
            exp_file  = emp.get("expected_file", "").strip()

            matched = _match_file(files, exp_file, claimed)

            if matched:
                ts = datetime.fromtimestamp(matched.stat().st_mtime)
                entry: dict[str, Any] = {
                    "dept":         dept_name,
                    "dept_folder":  folder_name,
                    "cc_email":     cc_email,
                    "name":         emp_name,
                    "email":        emp_email,
                    "expected_file": exp_file,
                    "matched_file": matched.name,
                    "modified":     ts.strftime("%Y-%m-%d %H:%M"),
                }
                dept_submitted.append(entry)
                all_submitted.append(entry)
                log.info("  OK  %-22s  %-25s  %s", dept_name, emp_name, matched.name)
            else:
                entry = {
                    "dept":          dept_name,
                    "dept_folder":   folder_name,
                    "cc_email":      cc_email,
                    "name":          emp_name,
                    "email":         emp_email,
                    "expected_file": exp_file,
                }
                dept_missing.append(entry)
                all_missing.append(entry)
                log.info("  --  %-22s  %-25s  (not found)", dept_name, emp_name)

        dept_results.append({
            "name":      dept_name,
            "folder":    str(dept_folder) if dept_folder else "",
            "cc_email":  cc_email,
            "submitted": dept_submitted,
            "missing":   dept_missing,
        })

    total = sum(len(d.get("employees", [])) for d in departments)

    return {
        "week_label":    week_label,
        "checked_at":    datetime.now().isoformat(timespec="seconds"),
        "departments":   dept_results,
        "all_submitted": all_submitted,
        "all_missing":   all_missing,
        "submit_count":  len(all_submitted),
        "missing_count": len(all_missing),
        "total":         total,
        "input_dir":     str(input_dir),
    }


# ---------------------------------------------------------------------------#
# Deadline helpers
# ---------------------------------------------------------------------------#
def is_past_deadline(cfg: dict) -> bool:
    email_cfg      = cfg.get("email", {})
    deadline_day   = email_cfg.get("deadline_day", "Thursday").lower()
    deadline_hr    = int(email_cfg.get("deadline_hour", 17))
    deadline_min   = int(email_cfg.get("deadline_minute", 0))
    target_weekday = _WEEKDAYS.get(deadline_day, 3)
    now            = datetime.now()
    same_day       = now.weekday() == target_weekday
    past_time      = (now.hour > deadline_hr) or \
                     (now.hour == deadline_hr and now.minute >= deadline_min)
    return now.weekday() > target_weekday or (same_day and past_time)


def next_deadline(cfg: dict) -> datetime:
    email_cfg      = cfg.get("email", {})
    deadline_day   = email_cfg.get("deadline_day", "Thursday").lower()
    deadline_hr    = int(email_cfg.get("deadline_hour", 17))
    deadline_min   = int(email_cfg.get("deadline_minute", 0))
    target_weekday = _WEEKDAYS.get(deadline_day, 3)
    now            = datetime.now()
    same_day       = now.weekday() == target_weekday
    past_time      = (now.hour > deadline_hr) or \
                     (now.hour == deadline_hr and now.minute >= deadline_min)
    days_ahead     = target_weekday - now.weekday()
    if days_ahead < 0 or (days_ahead == 0 and past_time):
        days_ahead += 7
    target_date = (now + timedelta(days=days_ahead)).date()
    return datetime.combine(
        target_date,
        datetime.min.time().replace(hour=deadline_hr, minute=deadline_min),
    )


# ---------------------------------------------------------------------------#
# Full email cycle
# ---------------------------------------------------------------------------#
def run_email_cycle(
    cfg: dict,
    send_acks: bool = True,
    send_reminders: bool = True,
    week_label: str | None = None,
    force: bool = False,
) -> dict[str, Any]:
    """Scan folders and send per-employee emails.

    Returns
    -------
    {
      week_label, ran_at,
      departments,          # from check_submissions
      all_submitted,
      all_missing,
      emails_sent:  [ {type, dept, name, to, status, detail}, ... ],
      emails_failed: [ ... ],
      errors:       [ str, ... ],
      skipped_reason: str | None,
    }
    """
    email_cfg     = cfg.get("email", {})
    email_enabled = email_cfg.get("enabled", False)
    wl            = week_label or week_label_for()

    result: dict[str, Any] = {
        "week_label":    wl,
        "ran_at":        datetime.now().isoformat(timespec="seconds"),
        "departments":   [],
        "all_submitted": [],
        "all_missing":   [],
        "emails_sent":   [],
        "emails_failed": [],
        "errors":        [],
        "skipped_reason": None,
    }

    # Step 1 — scan
    scan = check_submissions(cfg)
    result["departments"]   = scan["departments"]
    result["all_submitted"] = scan["all_submitted"]
    result["all_missing"]   = scan["all_missing"]

    if not email_enabled:
        result["skipped_reason"] = "Email is disabled in config (email.enabled = false)."
        log.info(
            "Email cycle dry-run — %d submitted, %d missing",
            scan["submit_count"], scan["missing_count"],
        )
        return result

    from_email = email_cfg.get("sender_email", "") or email_cfg.get("smtp_user", "")
    if not from_email:
        result["skipped_reason"] = "No sender email configured."
        result["errors"].append("sender_email not configured")
        return result

    # Step 2 — acknowledgments
    if send_acks:
        for emp in scan["all_submitted"]:
            dedup_key = f"{wl}|ack|{emp['email']}"
            with _sent_lock:
                if dedup_key in _sent_keys and not force:
                    continue

            ok, msg = send_acknowledgment(
                cfg=cfg,
                dept_name=emp["dept"],
                employee_name=emp["name"],
                employee_email=emp["email"],
                cc_email=emp.get("cc_email", ""),
                file_name=emp["matched_file"],
                week_label=wl,
                submitted_at=emp["modified"],
            )
            entry = {
                "type":   "acknowledgment",
                "dept":   emp["dept"],
                "name":   emp["name"],
                "to":     emp["email"],
                "status": "sent" if ok else "failed",
                "detail": msg,
            }
            if ok:
                result["emails_sent"].append(entry)
                with _sent_lock:
                    _sent_keys.add(dedup_key)
            else:
                result["emails_failed"].append(entry)
                result["errors"].append(f"Ack to {emp['name']} ({emp['dept']}): {msg}")

    # Step 3 — reminders
    if send_reminders:
        for emp in scan["all_missing"]:
            dedup_key = f"{wl}|rem|{emp['email']}"
            with _sent_lock:
                if dedup_key in _sent_keys and not force:
                    continue

            ok, msg = send_reminder(
                cfg=cfg,
                dept_name=emp["dept"],
                employee_name=emp["name"],
                employee_email=emp["email"],
                cc_email=emp.get("cc_email", ""),
                expected_file=emp.get("expected_file", ""),
                week_label=wl,
            )
            entry = {
                "type":   "reminder",
                "dept":   emp["dept"],
                "name":   emp["name"],
                "to":     emp["email"],
                "status": "sent" if ok else "failed",
                "detail": msg,
            }
            if ok:
                result["emails_sent"].append(entry)
                with _sent_lock:
                    _sent_keys.add(dedup_key)
            else:
                result["emails_failed"].append(entry)
                result["errors"].append(f"Reminder to {emp['name']} ({emp['dept']}): {msg}")

    log.info(
        "Email cycle complete — %d sent, %d failed",
        len(result["emails_sent"]), len(result["emails_failed"]),
    )
    return result


# ---------------------------------------------------------------------------#
# Background scheduler
# ---------------------------------------------------------------------------#
_scheduler_thread: threading.Thread | None = None
_scheduler_stop   = threading.Event()


def _within_window(cfg: dict) -> bool:
    """Return True if the current hour is inside the configured active window."""
    ac = cfg.get("auto_check", {})
    start = int(ac.get("window_start_hour", 7))
    end   = int(ac.get("window_end_hour",  20))
    hour  = datetime.now().hour
    return start <= hour < end


def _write_last_run(cfg_path: Path, status: str) -> None:
    """Persist last_run timestamp and last_result to config.json on disk."""
    try:
        p = Path(cfg_path)
        with open(p, "r", encoding="utf-8") as f:
            data = json.load(f)
        data.setdefault("auto_check", {})
        data["auto_check"]["last_run"]    = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        data["auto_check"]["last_result"] = status
        with open(p, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=4)
    except Exception as exc:
        log.warning("Could not write last_run to config: %s", exc)


def schedule_auto_check(cfg_path: str | Path) -> None:
    """Start the background auto-check scheduler.

    All timing settings are read fresh from config.json on every tick so
    changes take effect without restarting the app.

    Behaviour
    ---------
    - Sleeps until the active window (window_start_hour … window_end_hour).
    - After the deadline has passed, fires run_email_cycle() once per week.
    - Reminders re-fire every ``reminder_repeat_hours`` hours for employees
      who still haven't submitted (force=True bypasses the per-week dedup
      on subsequent passes — dedup is per-email per type per week on the
      first run; repeat runs use force so missing employees keep getting
      nudged until they submit).
    - Writes last_run / last_result back to config.json after each cycle.
    - Interval between checks: ``interval_minutes`` from config.
    """
    global _scheduler_thread, _scheduler_stop
    if _scheduler_thread and _scheduler_thread.is_alive():
        log.info("Auto-check scheduler already running")
        return
    _scheduler_stop.clear()
    cfg_path = Path(cfg_path)

    def _loop():
        log.info("Auto-check scheduler started — config: %s", cfg_path)
        _last_cycle_wl: str = ""
        _last_cycle_ts: datetime | None = None

        while not _scheduler_stop.is_set():
            # Always read interval fresh so GUI changes take effect immediately
            try:
                cfg          = _load_config(cfg_path)
                ac           = cfg.get("auto_check", {})
                interval_sec = max(60, int(ac.get("interval_minutes", 30)) * 60)
            except Exception:
                _scheduler_stop.wait(60)
                continue

            try:
                if not ac.get("enabled", False):
                    log.debug("Auto-check disabled, sleeping %d s", interval_sec)
                    _scheduler_stop.wait(interval_sec)
                    continue

                repeat_hours   = float(ac.get("reminder_repeat_hours", 24))
                send_acks      = bool(ac.get("send_acks",      True))
                send_reminders = bool(ac.get("send_reminders", True))
                now            = datetime.now()
                wl             = week_label_for()

                # ── Sleep if outside active window ──
                if not _within_window(cfg):
                    log.info(
                        "Auto-check: outside active window (%s–%s), sleeping %d min",
                        ac.get("window_start_hour", 7),
                        ac.get("window_end_hour",  20),
                        interval_sec // 60,
                    )
                    _scheduler_stop.wait(interval_sec)
                    continue

                # ── Only act after the deadline ──
                if not is_past_deadline(cfg):
                    nd = next_deadline(cfg)
                    # Sleep only until deadline if it's closer than one full interval
                    secs_to_deadline = max(0, (nd - now).total_seconds())
                    sleep_secs = min(interval_sec, secs_to_deadline + 5)
                    log.info(
                        "Auto-check: deadline not reached — next: %s (sleeping %.0f min)",
                        nd.strftime("%A %d %b %Y at %H:%M"),
                        sleep_secs / 60,
                    )
                    _scheduler_stop.wait(sleep_secs)
                    continue

                # ── Decide whether to fire this tick ──
                new_week   = wl != _last_cycle_wl
                repeat_due = (
                    _last_cycle_ts is None
                    or (now - _last_cycle_ts).total_seconds() >= repeat_hours * 3600
                )

                if not (new_week or repeat_due):
                    remaining_min = (
                        repeat_hours * 3600
                        - (now - _last_cycle_ts).total_seconds()
                    ) / 60
                    log.info(
                        "Auto-check: already ran for %s, next repeat in %.0f min",
                        wl, remaining_min,
                    )
                    _scheduler_stop.wait(interval_sec)
                    continue

                # ── Fire email cycle ──
                log.info("Auto-check: firing email cycle for %s", wl)
                result = run_email_cycle(
                    cfg,
                    send_acks=send_acks,
                    send_reminders=send_reminders,
                    week_label=wl,
                    force=not new_week,
                )

                _last_cycle_wl = wl
                _last_cycle_ts = now

                sent   = len(result.get("emails_sent",   []))
                failed = len(result.get("emails_failed", []))
                status = f"sent={sent} failed={failed} @ {now.strftime('%H:%M')}"
                log.info("Auto-check cycle done — %s", status)
                _write_last_run(cfg_path, status)

            except Exception as exc:
                log.error("Auto-check error: %s", exc)
                _write_last_run(cfg_path, f"ERROR: {exc}")

            _scheduler_stop.wait(interval_sec)

        log.info("Auto-check scheduler stopped")

    _scheduler_thread = threading.Thread(
        target=_loop, name="email-auto-check", daemon=True
    )
    _scheduler_thread.start()


def stop_auto_check() -> None:
    """Signal the background scheduler to stop."""
    _scheduler_stop.set()


def scheduler_is_running() -> bool:
    """Return True if the background scheduler thread is alive."""
    return _scheduler_thread is not None and _scheduler_thread.is_alive()


def reset_sent_cache() -> None:
    """Clear the in-memory dedup record (useful for testing)."""
    with _sent_lock:
        _sent_keys.clear()


# ---------------------------------------------------------------------------#
# CLI
# ---------------------------------------------------------------------------#
if __name__ == "__main__":
    import sys
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)-8s | %(message)s",
        datefmt="%H:%M:%S",
    )
    cfg_file = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).parent / "config.json"
    cfg = _load_config(cfg_file)

    scan = check_submissions(cfg)
    print(f"\nWeek : {scan['week_label']}")
    print(f"Total: {scan['total']}  Submitted: {scan['submit_count']}  Missing: {scan['missing_count']}\n")
    for d in scan["departments"]:
        print(f"  [{d['name']}]")
        for e in d["submitted"]:
            print(f"    ✓ {e['name']:<30} {e['matched_file']}")
        for e in d["missing"]:
            exp = f"  (expects: {e['expected_file']})" if e["expected_file"] else ""
            print(f"    ✗ {e['name']:<30}{exp}")

    print(f"\nPast deadline : {is_past_deadline(cfg)}")
    print(f"Next deadline : {next_deadline(cfg).strftime('%A %d %b %Y at %H:%M')}")

    if "--send" in sys.argv:
        result = run_email_cycle(cfg, force=True)
        print(f"\nSent: {len(result['emails_sent'])}  Failed: {len(result['emails_failed'])}")
        for e in result["errors"]:
            print(f"  ERR: {e}")
