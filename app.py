"""
Weekly Report Automator — Web Dashboard
========================================
Flask backend that exposes REST + SSE endpoints consumed by the
single-page dashboard (templates/index.html).

Run with:
    python app.py
    python app.py --port 8080
"""

from __future__ import annotations

import json
import logging
import queue
import subprocess
import sys
import threading
import time
from dataclasses import dataclass
from pathlib import Path

from flask import (
    Flask,
    Response,
    jsonify,
    render_template,
    request,
    send_file,
    stream_with_context,
)

# --------------------------------------------------------------------------- #
# BASE directory — resolved here, independently of main.py, so importing
# app.py never triggers main.py's module-level side-effects (logging.basicConfig
# + FileHandler) which conflict with our SSE handler on Python 3.14.
# --------------------------------------------------------------------------- #
if getattr(sys, "frozen", False):
    BASE = Path(sys.executable).resolve().parent
else:
    BASE = Path(__file__).resolve().parent

# --------------------------------------------------------------------------- #
# App bootstrap
# --------------------------------------------------------------------------- #
app = Flask(__name__, template_folder="templates", static_folder="static")
app.config["MAX_CONTENT_LENGTH"] = 64 * 1024 * 1024  # 64 MB upload cap

# --------------------------------------------------------------------------- #
# Logging — set up BEFORE importing main so basicConfig is already done
# and main's call becomes a no-op.
# --------------------------------------------------------------------------- #
LOG_DIR = BASE / "logs"
LOG_DIR.mkdir(exist_ok=True)

_log_formatter = logging.Formatter(
    "%(asctime)s | %(levelname)-8s | %(message)s", "%Y-%m-%d %H:%M:%S"
)
_file_handler = logging.FileHandler(LOG_DIR / "app.log", encoding="utf-8")
_file_handler.setFormatter(_log_formatter)

_root_logger = logging.getLogger()
_root_logger.setLevel(logging.INFO)
_root_logger.addHandler(_file_handler)
if sys.stdout is not None:
    _stdout_handler = logging.StreamHandler(sys.stdout)
    _stdout_handler.setFormatter(_log_formatter)
    _root_logger.addHandler(_stdout_handler)

# Prevent duplicate handlers if Flask reloads the module
logging.basicConfig()   # marks basicConfig as "already called" — main.py's call is now a no-op

# --------------------------------------------------------------------------- #
# SSE log queue — ring buffer + per-client queues
# --------------------------------------------------------------------------- #
_log_queues: list[queue.Queue] = []
_log_ring: list[dict] = []
_RING_SIZE = 200
_ring_lock = threading.Lock()


class _SSELogHandler(logging.Handler):
    """Pushes log records to every active SSE subscriber + the ring buffer."""

    def emit(self, record: logging.LogRecord) -> None:
        entry = {
            "ts": time.strftime("%H:%M:%S"),
            "level": record.levelname,
            "msg": record.getMessage(),
        }
        with _ring_lock:
            _log_ring.append(entry)
            if len(_log_ring) > _RING_SIZE:
                _log_ring.pop(0)
        for q in list(_log_queues):
            try:
                q.put_nowait(entry)
            except queue.Full:
                pass


_sse_handler = _SSELogHandler()
_root_logger.addHandler(_sse_handler)

# --------------------------------------------------------------------------- #
# Lazy pipeline import — deferred until first use so module-level
# side-effects in main.py don't interfere with our logging setup.
# --------------------------------------------------------------------------- #
_pipeline = None
_pipeline_lock = threading.Lock()


def _get_pipeline():
    global _pipeline
    if _pipeline is None:
        with _pipeline_lock:
            if _pipeline is None:
                import main as _m   # noqa: PLC0415
                _pipeline = _m
    return _pipeline


# --------------------------------------------------------------------------- #
# Pipeline state
# --------------------------------------------------------------------------- #
@dataclass
class _PipelineState:
    running: bool = False
    success: bool | None = None
    error: str = ""
    started_at: float = 0.0
    finished_at: float = 0.0


_state = _PipelineState()
_state_lock = threading.Lock()


# --------------------------------------------------------------------------- #
# Config helpers
# --------------------------------------------------------------------------- #
def _load_config() -> dict:
    cfg_path = BASE / "config.json"
    defaults = {
        "input_dir":   str(BASE / "input"),
        "output_file": str(BASE / "output" / "master_weekly_report.xlsx"),
    }
    if cfg_path.exists():
        try:
            with open(cfg_path, "r", encoding="utf-8") as f:
                defaults.update(json.load(f))
        except Exception:
            pass
    return defaults


def _save_config(data: dict) -> None:
    with open(BASE / "config.json", "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4)


# --------------------------------------------------------------------------- #
# Data helpers — delegate to report_worker.py subprocess to avoid a
# Python 3.14 segfault when pandas/openpyxl run inside Flask's WSGI threads.
# --------------------------------------------------------------------------- #
_WORKER = BASE / "report_worker.py"


def _output_summary() -> dict:
    """Spawn report_worker.py, capture its JSON output, return as dict."""
    cfg = _load_config()
    output_path = Path(cfg.get("output_file", ""))

    if not output_path.exists():
        return {"available": False}

    try:
        result = subprocess.run(
            [sys.executable, str(_WORKER), str(output_path)],
            capture_output=True, text=True, timeout=60,
        )
        if result.returncode != 0 or not result.stdout.strip():
            err = result.stderr.strip() or "worker exited with no output"
            logging.getLogger(__name__).error("report_worker error: %s", err)
            return {"available": False, "error": err}
        return json.loads(result.stdout)
    except Exception as exc:
        logging.getLogger(__name__).exception("_output_summary failed")
        return {"available": False, "error": str(exc)}


def _resolve_dept_folder(input_dir: Path, folder_value: str) -> Path | None:
    """Mirror of the same helper in main.py / email_checker.py.

    - Empty          → None
    - Absolute path  → used directly (external / shared network folder)
    - Relative name  → joined to input_dir  (legacy behaviour)
    """
    if not folder_value:
        return None
    p = Path(folder_value)
    if p.is_absolute():
        return p
    return input_dir / folder_value


def _list_input_files() -> list[dict]:
    """Return metadata for every .xlsx visible to the pipeline.

    Covers:
    - Root-level files directly under input_dir (legacy flat layout).
    - Each department's configured folder — which may be an absolute path
      on the server (external shared folder) or a sub-folder of input_dir.

    The ``archive`` sub-folder is always excluded.
    Each entry carries a ``dept`` field so the UI can group by department.
    """
    cfg          = _load_config()
    input_dir    = Path(cfg.get("input_dir", BASE / "input"))
    departments  = cfg.get("departments", [])

    files: list[dict] = []
    seen:  set[Path]  = set()   # de-duplicate in case two depts point to the same folder

    def _add_file(p: Path, dept_label: str) -> None:
        if p in seen:
            return
        seen.add(p)
        st = p.stat()
        # Display path: prefer relative-to-input_dir when possible,
        # otherwise just show  <dept_name>/<filename>  so the UI stays readable.
        try:
            rel = p.relative_to(input_dir)
            display = str(rel).replace("\\", "/")
        except ValueError:
            display = f"{dept_label}/{p.name}"
        files.append({
            "name":     display,
            "dept":     dept_label,
            "size_kb":  round(st.st_size / 1024, 1),
            "modified": time.strftime("%Y-%m-%d %H:%M", time.localtime(st.st_mtime)),
        })

    # ── Root-level files (no department) ────────────────────────────────
    if input_dir.exists():
        for p in sorted(input_dir.glob("*.xlsx")):
            if not p.name.startswith("~$"):
                _add_file(p, "")

    # ── Per-department folders ───────────────────────────────────────────
    for dept in departments:
        dept_name    = dept.get("name", "")
        folder_value = dept.get("folder", "").strip()
        dept_folder  = _resolve_dept_folder(input_dir, folder_value)
        if dept_folder is None or not dept_folder.is_dir():
            continue

        for p in sorted(dept_folder.rglob("*.xlsx")):
            if p.name.startswith("~$"):
                continue
            # Exclude archive directories
            try:
                rel_parts = p.relative_to(dept_folder).parts[:-1]
            except ValueError:
                rel_parts = ()
            if "archive" in [part.lower() for part in rel_parts]:
                continue
            _add_file(p, dept_name)

    return files


# --------------------------------------------------------------------------- #
# Routes — Pages
# --------------------------------------------------------------------------- #
@app.route("/")
def index():
    return render_template("index.html")


# --------------------------------------------------------------------------- #
# Routes — REST API
# --------------------------------------------------------------------------- #
@app.route("/api/config", methods=["GET"])
def get_config():
    return jsonify(_load_config())


@app.route("/api/config", methods=["POST"])
def save_config():
    data = request.get_json(force=True)
    if not data:
        return jsonify({"error": "No JSON body"}), 400
    cfg = _load_config()
    if "input_dir"   in data: cfg["input_dir"]   = data["input_dir"]
    if "output_file" in data: cfg["output_file"]  = data["output_file"]
    try:
        _save_config(cfg)
        Path(cfg["input_dir"]).mkdir(parents=True, exist_ok=True)
        Path(cfg["output_file"]).parent.mkdir(parents=True, exist_ok=True)
        return jsonify({"ok": True, "config": cfg})
    except Exception as exc:
        return jsonify({"error": str(exc)}), 500


@app.route("/api/status", methods=["GET"])
def get_status():
    with _state_lock:
        elapsed = None
        if _state.running and _state.started_at:
            elapsed = round(time.time() - _state.started_at, 1)
        elif _state.finished_at:
            elapsed = round(_state.finished_at - _state.started_at, 1)
        s = {
            "running": _state.running,
            "success": _state.success,
            "error":   _state.error,
            "elapsed": elapsed,
        }
    return jsonify(s)


@app.route("/api/pipeline/run", methods=["POST"])
def run_pipeline():
    with _state_lock:
        if _state.running:
            return jsonify({"error": "Pipeline already running"}), 409
        _state.running    = True
        _state.success    = None
        _state.error      = ""
        _state.started_at = time.time()
        _state.finished_at = 0.0

    cfg         = _load_config()
    input_dir   = cfg.get("input_dir",   str(BASE / "input"))
    output_file = cfg.get("output_file", str(BASE / "output" / "master_weekly_report.xlsx"))
    departments = cfg.get("departments", None)

    def worker():
        try:
            _get_pipeline().run_pipeline_from_paths(input_dir, output_file, departments)
            with _state_lock:
                _state.success     = True
                _state.running     = False
                _state.finished_at = time.time()
        except Exception as exc:
            logging.getLogger(__name__).error("Pipeline error: %s", exc)
            with _state_lock:
                _state.success     = False
                _state.error       = str(exc)
                _state.running     = False
                _state.finished_at = time.time()

    threading.Thread(target=worker, daemon=True).start()
    return jsonify({"ok": True, "message": "Pipeline started"})


@app.route("/api/data/summary", methods=["GET"])
def data_summary():
    return jsonify(_output_summary())


@app.route("/api/files/input", methods=["GET"])
def list_input_files():
    return jsonify(_list_input_files())


@app.route("/api/files/upload", methods=["POST"])
def upload_files():
    """Upload one or more .xlsx files.

    Optional query / form param ``dept``: the department *name* (as stored
    in config) to upload into.  When supplied the file is placed in that
    department's configured folder (which may be an external absolute path).
    When omitted the file lands in the root of input_dir (legacy behaviour).
    """
    cfg          = _load_config()
    input_dir    = Path(cfg.get("input_dir", BASE / "input"))
    departments  = cfg.get("departments", [])

    # Resolve target folder
    dept_name   = (request.args.get("dept") or request.form.get("dept") or "").strip()
    target_dir  = input_dir   # default: root

    if dept_name:
        for dept in departments:
            if dept.get("name", "").strip().lower() == dept_name.lower():
                folder_value = dept.get("folder", "").strip()
                resolved     = _resolve_dept_folder(input_dir, folder_value)
                if resolved is not None:
                    target_dir = resolved
                break

    target_dir.mkdir(parents=True, exist_ok=True)

    uploaded, errors = [], []
    for f in request.files.getlist("files"):
        if not f.filename:
            continue
        fname = Path(f.filename).name
        if not fname.lower().endswith(".xlsx"):
            errors.append(f"{fname}: only .xlsx files are accepted")
            continue
        try:
            f.save(str(target_dir / fname))
            uploaded.append(fname)
        except Exception as exc:
            errors.append(f"{fname}: {exc}")

    return jsonify({"uploaded": uploaded, "errors": errors})


@app.route("/api/files/delete/<path:filename>", methods=["DELETE"])
def delete_input_file(filename: str):
    """Delete an input file by its display name (as returned by /api/files/input).

    The display name is either:
    - A path relative to input_dir  (e.g. "Contract/report.xlsx")  — legacy
    - "<DeptName>/<filename>"       — used when the dept folder is external

    Security: the resolved target must be inside either input_dir OR one of
    the configured department folders.  Anything else is rejected.
    """
    cfg         = _load_config()
    input_dir   = Path(cfg.get("input_dir", BASE / "input")).resolve()
    departments = cfg.get("departments", [])

    # Build the set of trusted root directories
    trusted_roots: set[Path] = {input_dir}
    for dept in departments:
        folder_value = dept.get("folder", "").strip()
        resolved     = _resolve_dept_folder(input_dir, folder_value)
        if resolved is not None:
            trusted_roots.add(resolved.resolve())

    # Try resolving as relative to input_dir first (covers legacy paths and
    # paths like "Contract/report.xlsx" for internal sub-folders).
    candidate = (input_dir / Path(filename)).resolve()

    # If that candidate isn't under any trusted root, the display name may use
    # the "<DeptName>/<file>" format for an external folder.  In that case strip
    # the first path component (dept name) and search all dept folders.
    def _is_trusted(p: Path) -> bool:
        p_str = str(p)
        return any(p_str.startswith(str(root)) for root in trusted_roots)

    if not _is_trusted(candidate):
        # Try matching by filename across all department folders
        fname = Path(filename).name
        candidate = None
        for dept in departments:
            folder_value = dept.get("folder", "").strip()
            resolved     = _resolve_dept_folder(input_dir, folder_value)
            if resolved is None or not resolved.is_dir():
                continue
            for p in resolved.rglob(fname):
                if p.is_file():
                    candidate = p.resolve()
                    break
            if candidate:
                break

    if candidate is None or not _is_trusted(candidate):
        return jsonify({"error": "Invalid path"}), 400

    try:
        candidate.unlink()
        return jsonify({"ok": True})
    except FileNotFoundError:
        return jsonify({"error": "File not found"}), 404
    except Exception as exc:
        return jsonify({"error": str(exc)}), 500


@app.route("/api/files/archive", methods=["GET"])
def list_archive():
    """Return a list of archived batches (timestamped sub-folders under input/archive)."""
    cfg = _load_config()
    archive_root = Path(cfg.get("input_dir", BASE / "input")) / "archive"
    if not archive_root.exists():
        return jsonify([])

    batches = []
    for folder in sorted(archive_root.iterdir(), reverse=True):
        if not folder.is_dir():
            continue
        files = sorted(folder.glob("*.xlsx"))
        batches.append({
            "batch":    folder.name,
            "files":    [f.name for f in files],
            "count":    len(files),
            "path":     str(folder),
        })
    return jsonify(batches)


@app.route("/api/email/scheduler", methods=["GET"])
def scheduler_status():
    """Return whether the background email scheduler is running."""
    try:
        import email_checker as _ec
        cfg = _load_config()
        ac  = cfg.get("auto_check", {})
        return jsonify({
            "running":        _ec.scheduler_is_running(),
            "enabled":        ac.get("enabled", False),
            "interval_min":   ac.get("interval_minutes", 30),
            "window":         f"{ac.get('window_start_hour',7):02d}:00 – {ac.get('window_end_hour',20):02d}:00",
            "last_run":       ac.get("last_run",    ""),
            "last_result":    ac.get("last_result", ""),
        })
    except Exception as exc:
        return jsonify({"error": str(exc)}), 500


@app.route("/api/email/scheduler", methods=["POST"])
def scheduler_control():
    """Start or stop the background scheduler.  Body: { "action": "start"|"stop" }"""
    try:
        import email_checker as _ec
        body   = request.get_json(force=True) or {}
        action = body.get("action", "start")
        if action == "stop":
            _ec.stop_auto_check()
            return jsonify({"ok": True, "running": False})
        else:
            _ec.schedule_auto_check(BASE / "config.json")
            return jsonify({"ok": True, "running": _ec.scheduler_is_running()})
    except Exception as exc:
        return jsonify({"error": str(exc)}), 500


@app.route("/api/email/status", methods=["GET"])
def email_status():
    """Scan department folders and return submission status + deadline info."""
    try:
        import email_checker as _ec
        cfg  = _load_config()
        scan = _ec.check_submissions(cfg)
        scan["past_deadline"]  = _ec.is_past_deadline(cfg)
        scan["next_deadline"]  = _ec.next_deadline(cfg).strftime("%A %d %b %Y at %H:%M")
        scan["email_enabled"]  = cfg.get("email", {}).get("enabled", False)
        return jsonify(scan)
    except Exception as exc:
        logging.getLogger(__name__).exception("email_status failed")
        return jsonify({"error": str(exc)}), 500


@app.route("/api/email/send-reminders", methods=["POST"])
def send_reminders():
    """Trigger the full email cycle (reminders + acks).

    Body (all optional):
        { "send_acks": bool, "send_reminders": bool, "force": bool }
    """
    try:
        import email_checker as _ec
        body          = request.get_json(force=True) or {}
        send_acks     = bool(body.get("send_acks",     True))
        send_rems     = bool(body.get("send_reminders", True))
        force         = bool(body.get("force",         False))
        cfg           = _load_config()
        result        = _ec.run_email_cycle(
            cfg,
            send_acks=send_acks,
            send_reminders=send_rems,
            force=force,
        )
        return jsonify(result)
    except Exception as exc:
        logging.getLogger(__name__).exception("send_reminders failed")
        return jsonify({"error": str(exc)}), 500


@app.route("/api/email/test", methods=["POST"])
def send_test_email():
    """Send a test SMTP email.  Body: { "to": "address@example.com" }"""
    try:
        import email_sender as _es
        body = request.get_json(force=True) or {}
        to   = body.get("to", "").strip()
        if not to:
            return jsonify({"error": "Missing 'to' field"}), 400
        cfg     = _load_config()
        ok, msg = _es.send_test_email(cfg, to)
        return jsonify({"ok": ok, "message": msg})
    except Exception as exc:
        logging.getLogger(__name__).exception("send_test_email failed")
        return jsonify({"error": str(exc)}), 500


@app.route("/api/email/config", methods=["GET"])
def get_email_config():
    """Return current email + departments config.

    Sensitive fields are redacted:
    - email.smtp_password  → omitted entirely
    - dept.dept_head_password → replaced with a boolean 'has_password' flag
      so the UI can show whether a password is set without exposing the value.
    """
    cfg  = _load_config()
    ecfg = {k: v for k, v in cfg.get("email", {}).items() if k != "smtp_password"}

    # Redact dept_head_password — send a has_password flag instead
    depts_safe = []
    for dept in cfg.get("departments", []):
        d = dict(dept)
        raw_pass = d.pop("dept_head_password", "")
        d["has_password"] = bool(raw_pass)   # True / False — never the value
        depts_safe.append(d)

    return jsonify({"email": ecfg, "departments": depts_safe})


@app.route("/api/email/config", methods=["POST"])
def save_email_config():
    """Save email and/or departments config.

    Body keys (all optional): email (dict), departments (list)

    Department password handling
    ----------------------------
    The GET endpoint never returns the actual dept_head_password — it sends
    a ``has_password`` boolean flag instead.  Therefore when the client POSTs
    departments back we must NOT blindly overwrite passwords:

    - If ``dept_head_password`` is present and non-empty in the POST body
      → use the new value (user explicitly set / changed it).
    - If ``dept_head_password`` is absent or empty in the POST body
      → preserve the existing password from the current config (no change).

    This prevents the web UI from accidentally erasing saved passwords.
    """
    data = request.get_json(force=True) or {}
    cfg  = _load_config()

    if "email" in data:
        cfg.setdefault("email", {}).update(data["email"])

    if "departments" in data:
        # Build a lookup of existing passwords keyed by dept name
        existing_passwords: dict[str, str] = {
            d.get("name", ""): d.get("dept_head_password", "")
            for d in cfg.get("departments", [])
        }

        merged_depts = []
        for dept in data["departments"]:
            dept_name     = dept.get("name", "")
            incoming_pass = dept.get("dept_head_password", "")

            if incoming_pass:
                # New / changed password supplied — use it
                final_pass = incoming_pass
            else:
                # No password in request — keep whatever was stored
                final_pass = existing_passwords.get(dept_name, "")

            merged_dept = dict(dept)
            merged_dept["dept_head_password"] = final_pass
            # Remove the UI-only flag that the GET adds; it must not be stored
            merged_dept.pop("has_password", None)
            merged_depts.append(merged_dept)

        cfg["departments"] = merged_depts

    try:
        _save_config(cfg)
        return jsonify({"ok": True})
    except Exception as exc:
        return jsonify({"error": str(exc)}), 500


@app.route("/api/output/download", methods=["GET"])
def download_output():
    output_path = Path(_load_config().get("output_file", ""))
    if not output_path.exists():
        return jsonify({"error": "Output file not found"}), 404
    return send_file(
        str(output_path),
        as_attachment=True,
        download_name=output_path.name,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/api/logs/history", methods=["GET"])
def log_history():
    with _ring_lock:
        return jsonify(list(_log_ring))


# --------------------------------------------------------------------------- #
# Server-Sent Events — live log stream
# --------------------------------------------------------------------------- #
@app.route("/api/logs/stream")
def log_stream():
    client_q: queue.Queue = queue.Queue(maxsize=500)
    _log_queues.append(client_q)

    with _ring_lock:
        backlog = list(_log_ring)

    @stream_with_context
    def generate():
        try:
            for entry in backlog:
                yield f"data: {json.dumps(entry)}\n\n"
            while True:
                try:
                    entry = client_q.get(timeout=5)
                    yield f"data: {json.dumps(entry)}\n\n"
                except queue.Empty:
                    yield ": heartbeat\n\n"
        except GeneratorExit:
            pass
        finally:
            try:
                _log_queues.remove(client_q)
            except ValueError:
                pass

    return Response(
        generate(),
        mimetype="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


# --------------------------------------------------------------------------- #
# Entry point
# --------------------------------------------------------------------------- #
if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Weekly Report Web Dashboard")
    parser.add_argument("--port", type=int, default=5000)
    parser.add_argument("--host", default="127.0.0.1")
    args = parser.parse_args()

    for d in [BASE / "input", BASE / "output",
              BASE / "templates", BASE / "static" / "css", BASE / "static" / "js"]:
        d.mkdir(parents=True, exist_ok=True)

    # Start background email scheduler
    try:
        import email_checker as _ec
        _ec.schedule_auto_check(BASE / "config.json")
        logging.getLogger(__name__).info("Email auto-check scheduler started")
    except Exception as _e:
        logging.getLogger(__name__).warning("Could not start email scheduler: %s", _e)

    print(f"\n  Weekly Report Dashboard  →  http://{args.host}:{args.port}\n")
    print("  Server: Flask threaded\n")
    app.run(host=args.host, port=args.port, debug=False,
            threaded=True, use_reloader=False)
