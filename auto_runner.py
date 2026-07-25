"""
auto_runner.py
==============
Headless one-shot email cycle runner.

Designed to be called by Windows Task Scheduler (or any cron-like tool)
with no GUI and no human interaction.  It loads config.json, runs one
email cycle, writes the result back to config (last_run / last_result),
logs everything to logs/auto_runner.log, and exits.

Usage
-----
    python auto_runner.py
    python auto_runner.py --config "C:/path/to/config.json"
    python auto_runner.py --force          # ignore past-deadline check
    python auto_runner.py --reminders-only
    python auto_runner.py --acks-only

Exit codes
----------
    0   success (even if 0 emails sent — no recipients missing)
    1   partial failure (some emails failed)
    2   fatal error (config missing, SMTP not configured, etc.)
"""

from __future__ import annotations

import argparse
import json
import logging
import sys
import io

# Force UTF-8 on the stdout stream so log lines with special chars never crash
if hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass
_utf8_stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace') \
    if hasattr(sys.stdout, 'buffer') else sys.stdout
from datetime import datetime
from pathlib import Path


# ---------------------------------------------------------------------------#
# Resolve base directory (works as .py script or PyInstaller .exe)
# ---------------------------------------------------------------------------#
if getattr(sys, "frozen", False):
    BASE = Path(sys.executable).resolve().parent
else:
    BASE = Path(__file__).resolve().parent


# ---------------------------------------------------------------------------#
# Logging — always write to logs/auto_runner.log
# ---------------------------------------------------------------------------#
LOG_DIR = BASE / "logs"
LOG_DIR.mkdir(exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)-8s | %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[
        logging.FileHandler(LOG_DIR / "auto_runner.log", encoding="utf-8"),
        logging.StreamHandler(
            stream=open(sys.stdout.fileno(), mode='w', encoding='utf-8', errors='replace', closefd=False)
            if hasattr(sys.stdout, 'fileno') else sys.stdout
        ),
    ],
)
log = logging.getLogger("auto_runner")


# ---------------------------------------------------------------------------#
# Helpers
# ---------------------------------------------------------------------------#
def _load_config(path: Path) -> dict:
    if not path.exists():
        log.error("Config not found: %s", path)
        sys.exit(2)
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def _write_last_run(path: Path, status: str) -> None:
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        data.setdefault("auto_check", {})
        data["auto_check"]["last_run"]    = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        data["auto_check"]["last_result"] = status
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=4)
    except Exception as exc:
        log.warning("Could not write last_run: %s", exc)


# ---------------------------------------------------------------------------#
# Main
# ---------------------------------------------------------------------------#
def main() -> int:
    parser = argparse.ArgumentParser(description="Weekly Report — headless email runner")
    parser.add_argument(
        "--config", default=str(BASE / "config.json"),
        help="Path to config.json (default: same folder as script)",
    )
    parser.add_argument(
        "--force", action="store_true",
        help="Send even if past-deadline check fails or emails already sent this week",
    )
    parser.add_argument(
        "--reminders-only", action="store_true",
        help="Only send reminder emails (skip acknowledgments)",
    )
    parser.add_argument(
        "--acks-only", action="store_true",
        help="Only send acknowledgment emails (skip reminders)",
    )
    parser.add_argument(
        "--ignore-window", action="store_true",
        help="Run even if outside the configured active window hours",
    )
    args = parser.parse_args()

    cfg_path = Path(args.config)
    cfg      = _load_config(cfg_path)

    log.info("=" * 60)
    log.info("Auto-runner started")
    log.info("Config : %s", cfg_path)

    # ── Check email is enabled ──
    email_cfg = cfg.get("email", {})
    if not email_cfg.get("enabled", False):
        log.warning("Email is disabled in config (email.enabled = false). Exiting.")
        _write_last_run(cfg_path, "skipped: email disabled")
        return 0

    # ── Validate sender ──
    from_email = email_cfg.get("sender_email", "") or email_cfg.get("smtp_user", "")
    if not from_email:
        log.error("No sender email configured. Exiting.")
        _write_last_run(cfg_path, "ERROR: sender_email not configured")
        return 2

    # ── Import after logging is set up ──
    import email_checker

    # ── Active window check ──
    ac = cfg.get("auto_check", {})
    if not args.ignore_window and not email_checker._within_window(cfg):
        start = ac.get("window_start_hour", 7)
        end   = ac.get("window_end_hour",  20)
        log.info("Outside active window (%02d:00–%02d:00). Use --ignore-window to override.", start, end)
        _write_last_run(cfg_path, f"skipped: outside window {start:02d}:00–{end:02d}:00")
        return 0

    # ── Deadline check ──
    if not args.force and not email_checker.is_past_deadline(cfg):
        nd = email_checker.next_deadline(cfg)
        log.info(
            "Deadline not reached — next: %s. Use --force to override.",
            nd.strftime("%A %d %b %Y at %H:%M"),
        )
        _write_last_run(cfg_path, f"skipped: deadline not reached (next: {nd.strftime('%d %b %H:%M')})")
        return 0

    # ── Determine what to send ──
    send_acks      = not args.reminders_only
    send_reminders = not args.acks_only

    # Honour auto_check overrides from config
    if not args.reminders_only and not args.acks_only:
        send_acks      = bool(ac.get("send_acks",      True))
        send_reminders = bool(ac.get("send_reminders", True))

    log.info(
        "Running email cycle -- acks=%s reminders=%s force=%s",
        send_acks, send_reminders, args.force,
    )

    # ── Run cycle ──
    result = email_checker.run_email_cycle(
        cfg,
        send_acks=send_acks,
        send_reminders=send_reminders,
        force=args.force,
    )

    sent   = len(result.get("emails_sent",   []))
    failed = len(result.get("emails_failed", []))
    sub    = len(result.get("all_submitted", []))
    miss   = len(result.get("all_missing",   []))
    skip   = result.get("skipped_reason")

    if skip:
        log.info("Cycle skipped: %s", skip)
        _write_last_run(cfg_path, f"skipped: {skip}")
        return 0

    log.info(
        "Cycle complete — submitted: %d  missing: %d  sent: %d  failed: %d",
        sub, miss, sent, failed,
    )

    for entry in result.get("emails_sent", []):
        log.info("  SENT %-14s -> %-30s %s", entry["type"], entry["name"], entry["to"])

    for entry in result.get("emails_failed", []):
        log.warning("  FAIL %-14s -> %-30s %s -- %s",
                    entry["type"], entry["name"], entry["to"], entry["detail"])

    status = f"sent={sent} failed={failed} sub={sub} miss={miss}"
    _write_last_run(cfg_path, status)

    return 1 if failed > 0 else 0


if __name__ == "__main__":
    sys.exit(main())
