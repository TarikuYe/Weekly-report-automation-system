"""
email_sender.py
===============
Core email delivery module for the Weekly Report Automator.

All public functions accept an explicit ``cfg`` dict (parsed config.json) so
they are fully stateless and safe to call from threads.

Public API
----------
send_reminder(cfg, dept_name, employee_name, employee_email,
              cc_email, expected_file, week_label)

send_acknowledgment(cfg, dept_name, employee_name, employee_email,
                    cc_email, file_name, week_label, submitted_at)

send_dept_master(cfg, dept_master_result)

send_test_email(cfg, to_email)

week_label_for(dt=None)
"""

from __future__ import annotations

import logging
import smtplib
import ssl
from datetime import datetime
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
from typing import Any

log = logging.getLogger("email_sender")

# ── Design tokens (match app theme) ──────────────────────────────────────────
_ACCENT_DARK = "#1B2A5C"
_SUCCESS     = "#1B8A5A"
_WARNING     = "#B7791F"
_DANGER      = "#C0392B"
_MUTED       = "#6B7280"
_BG          = "#F5F6FA"
_WHITE       = "#FFFFFF"


# ── HTML shell ────────────────────────────────────────────────────────────────
def _base_html(body_content: str, header_color: str = _ACCENT_DARK) -> str:
    year = datetime.now().year
    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1.0"/>
<title>Weekly Report System</title>
</head>
<body style="margin:0;padding:0;background:{_BG};
             font-family:'Segoe UI',Arial,sans-serif;">
<table width="100%" cellpadding="0" cellspacing="0"
       style="background:{_BG};padding:32px 16px;">
  <tr><td align="center">
    <table width="600" cellpadding="0" cellspacing="0"
           style="max-width:600px;width:100%;background:{_WHITE};
                  border-radius:12px;overflow:hidden;
                  box-shadow:0 4px 24px rgba(0,0,0,.08);">

      <!-- Header -->
      <tr>
        <td style="background:{header_color};padding:28px 36px;">
          <p style="margin:0;font-size:11px;color:rgba(255,255,255,.65);
                    letter-spacing:1.5px;text-transform:uppercase;">
            Weekly Report System
          </p>
          <h1 style="margin:8px 0 0;font-size:22px;font-weight:700;color:#fff;">
            Report Status Notification
          </h1>
        </td>
      </tr>

      <!-- Body -->
      <tr>
        <td style="padding:32px 36px;">{body_content}</td>
      </tr>

      <!-- Footer -->
      <tr>
        <td style="background:{_BG};padding:20px 36px;
                   border-top:1px solid #E3E6EE;">
          <p style="margin:0;font-size:11px;color:{_MUTED};line-height:1.6;">
            This is an automated message from the Weekly Report Automator.
            Please do not reply directly to this email.<br/>
            &copy; {year} Weekly Report System
          </p>
        </td>
      </tr>

    </table>
  </td></tr>
</table>
</body>
</html>"""


def _divider() -> str:
    return '<hr style="border:none;border-top:1px solid #E3E6EE;margin:24px 0;"/>'


def _info_card(color_bg: str, color_border: str, color_label: str,
               icon: str, label: str, rows: list[tuple[str, str]]) -> str:
    row_html = "".join(
        f'<p style="margin:4px 0;font-size:14px;color:#374151;line-height:1.6;">'
        f'{k}: <strong>{v}</strong></p>'
        for k, v in rows
    )
    return f"""
<table width="100%" cellpadding="0" cellspacing="0"
       style="background:{color_bg};border:1px solid {color_border};
              border-radius:8px;margin:24px 0;">
  <tr>
    <td style="padding:20px 24px;">
      <p style="margin:0 0 8px;font-size:13px;font-weight:700;
                color:{color_label};text-transform:uppercase;letter-spacing:.5px;">
        {icon} {label}
      </p>
      {row_html}
    </td>
  </tr>
</table>"""


# ── Reminder ──────────────────────────────────────────────────────────────────
def _reminder_html(
    dept_name: str,
    employee_name: str,
    expected_file: str,
    deadline_day: str,
    deadline_hour: int,
    week_label: str,
) -> str:
    h = deadline_hour % 12 or 12
    ampm = f"{h}:00 {'AM' if deadline_hour < 12 else 'PM'}"
    file_note = (
        f"Please submit the file named <strong>{expected_file}</strong>."
        if expected_file
        else "Please submit your weekly report file (.xlsx)."
    )
    card = _info_card(
        "#FFF8ED", "#F6D860", _WARNING,
        "⚠", "Submission Overdue",
        [("Deadline", f"{deadline_day} at {ampm}"),
         ("Department", dept_name),
         ("Report Week", week_label)],
    )
    body = f"""
    <p style="margin:0 0 8px;font-size:13px;color:{_MUTED};font-weight:600;
              letter-spacing:.5px;text-transform:uppercase;">Action Required</p>
    <h2 style="margin:0 0 20px;font-size:20px;font-weight:700;color:{_ACCENT_DARK};">
      Weekly Report Not Yet Submitted
    </h2>
    <p style="margin:0 0 16px;font-size:15px;color:#374151;line-height:1.7;">
      Dear <strong>{employee_name}</strong>,
    </p>
    <p style="margin:0 0 16px;font-size:15px;color:#374151;line-height:1.7;">
      This is a friendly reminder that your weekly report for the
      <strong>{dept_name}</strong> department has <em>not yet been received</em>
      for <strong>{week_label}</strong>.
    </p>
    {card}
    <p style="margin:0 0 16px;font-size:15px;color:#374151;line-height:1.7;">
      {file_note} Please place it in the <strong>{dept_name}</strong> input
      folder as soon as possible. If you have already submitted and believe this
      message was sent in error, please contact your system administrator.
    </p>
    {_divider()}
    <p style="margin:0;font-size:13px;color:{_MUTED};line-height:1.6;">
      Report Week: <strong>{week_label}</strong> &nbsp;|&nbsp;
      Department: <strong>{dept_name}</strong>
    </p>"""
    return _base_html(body, header_color=_WARNING)


def _reminder_plain(
    dept_name: str,
    employee_name: str,
    expected_file: str,
    deadline_day: str,
    deadline_hour: int,
    week_label: str,
) -> str:
    h = deadline_hour % 12 or 12
    ampm = f"{h}:00 {'AM' if deadline_hour < 12 else 'PM'}"
    file_note = (
        f"Please submit the file named '{expected_file}'."
        if expected_file
        else "Please submit your weekly report (.xlsx) file."
    )
    return (
        f"Dear {employee_name},\n\n"
        f"This is a reminder that your weekly report for the {dept_name} department "
        f"has not yet been received for {week_label}.\n\n"
        f"Deadline : {deadline_day} at {ampm}\n"
        f"Department: {dept_name}\n\n"
        f"{file_note}\n\n"
        f"---\nWeekly Report System (automated message)"
    )


# ── Acknowledgment ────────────────────────────────────────────────────────────
def _ack_html(
    dept_name: str,
    employee_name: str,
    file_name: str,
    week_label: str,
    submitted_at: str,
) -> str:
    card = _info_card(
        "#EDFAF4", "#6EE7B7", _SUCCESS,
        "✓", "Report Received",
        [("File", file_name),
         ("Received", submitted_at),
         ("Department", dept_name),
         ("Report Week", week_label)],
    )
    body = f"""
    <p style="margin:0 0 8px;font-size:13px;color:{_MUTED};font-weight:600;
              letter-spacing:.5px;text-transform:uppercase;">Confirmation</p>
    <h2 style="margin:0 0 20px;font-size:20px;font-weight:700;color:{_ACCENT_DARK};">
      Weekly Report Received — Thank You
    </h2>
    <p style="margin:0 0 16px;font-size:15px;color:#374151;line-height:1.7;">
      Dear <strong>{employee_name}</strong>,
    </p>
    <p style="margin:0 0 16px;font-size:15px;color:#374151;line-height:1.7;">
      We have successfully received your weekly report for the
      <strong>{dept_name}</strong> department for <strong>{week_label}</strong>.
      Thank you for submitting on time!
    </p>
    {card}
    <p style="margin:0 0 16px;font-size:15px;color:#374151;line-height:1.7;">
      Your report will be processed and included in the consolidated master
      weekly report. No further action is needed.
    </p>
    {_divider()}
    <p style="margin:0;font-size:13px;color:{_MUTED};line-height:1.6;">
      Report Week: <strong>{week_label}</strong> &nbsp;|&nbsp;
      Department: <strong>{dept_name}</strong>
    </p>"""
    return _base_html(body, header_color=_SUCCESS)


def _ack_plain(
    dept_name: str,
    employee_name: str,
    file_name: str,
    week_label: str,
    submitted_at: str,
) -> str:
    return (
        f"Dear {employee_name},\n\n"
        f"We have successfully received your weekly report for the {dept_name} "
        f"department for {week_label}. Thank you for submitting on time!\n\n"
        f"File     : {file_name}\n"
        f"Received : {submitted_at}\n\n"
        f"Your report will be included in the master weekly report.\n\n"
        f"---\nWeekly Report System (automated message)"
    )


# ── Test email ────────────────────────────────────────────────────────────────
def _test_html() -> str:
    card = _info_card(
        "#EDFAF4", "#6EE7B7", _SUCCESS,
        "✓", "Connection Successful",
        [("Status", "Email delivery is working"),
         ("Sent", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))],
    )
    body = f"""
    <p style="margin:0 0 8px;font-size:13px;color:{_MUTED};font-weight:600;
              letter-spacing:.5px;text-transform:uppercase;">SMTP Test</p>
    <h2 style="margin:0 0 20px;font-size:20px;font-weight:700;color:{_ACCENT_DARK};">
      Test Email — Settings Verified
    </h2>
    <p style="margin:0 0 16px;font-size:15px;color:#374151;line-height:1.7;">
      Your SMTP settings are correctly configured. The Weekly Report Automator
      can send emails from this account.
    </p>
    {card}"""
    return _base_html(body)


# ── SMTP helpers ──────────────────────────────────────────────────────────────
def _make_smtp(cfg: dict) -> smtplib.SMTP:
    ecfg     = cfg.get("email", {})
    host     = ecfg.get("smtp_host", "smtp.gmail.com")
    port     = int(ecfg.get("smtp_port", 587))
    use_tls  = bool(ecfg.get("use_tls", True))
    user     = ecfg.get("smtp_user", "")
    password = ecfg.get("smtp_password", "")

    if port == 465:
        ctx    = ssl.create_default_context()
        server = smtplib.SMTP_SSL(host, port, context=ctx, timeout=15)
    else:
        server = smtplib.SMTP(host, port, timeout=15)
        if use_tls:
            server.ehlo()
            server.starttls(context=ssl.create_default_context())
            server.ehlo()

    if user and password:
        server.login(user, password)
    return server


def _build_message(
    cfg: dict,
    to_email: str,
    to_name: str,
    subject: str,
    html_body: str,
    plain_body: str,
    cc_email: str = "",
) -> MIMEMultipart:
    ecfg        = cfg.get("email", {})
    sender_name = ecfg.get("sender_name", "Weekly Report System")
    from_email  = ecfg.get("sender_email", "") or ecfg.get("smtp_user", "")

    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"]    = f"{sender_name} <{from_email}>"
    msg["To"]      = f"{to_name} <{to_email}>" if to_name else to_email
    if cc_email:
        msg["Cc"] = cc_email

    msg.attach(MIMEText(plain_body, "plain", "utf-8"))
    msg.attach(MIMEText(html_body,  "html",  "utf-8"))
    return msg


def _recipients(to_email: str, cc_email: str) -> list[str]:
    recips = [to_email]
    if cc_email:
        recips.append(cc_email)
    return recips


# ── Public API ────────────────────────────────────────────────────────────────
def send_reminder(
    cfg: dict,
    dept_name: str,
    employee_name: str,
    employee_email: str,
    cc_email: str = "",
    expected_file: str = "",
    week_label: str | None = None,
) -> tuple[bool, str]:
    """Send a reminder to an individual employee who hasn't submitted yet."""
    ecfg         = cfg.get("email", {})
    subject      = ecfg.get("reminder_subject", "Reminder: Weekly Report Not Yet Submitted")
    deadline_day = ecfg.get("deadline_day", "Thursday")
    deadline_hr  = int(ecfg.get("deadline_hour", 17))
    wl           = week_label or _current_week_label()

    html  = _reminder_html(dept_name, employee_name, expected_file,
                            deadline_day, deadline_hr, wl)
    plain = _reminder_plain(dept_name, employee_name, expected_file,
                             deadline_day, deadline_hr, wl)
    msg   = _build_message(cfg, employee_email, employee_name,
                            subject, html, plain, cc_email)

    try:
        from_email = (cfg.get("email", {}).get("sender_email", "")
                      or cfg.get("email", {}).get("smtp_user", ""))
        with _make_smtp(cfg) as server:
            server.sendmail(from_email, _recipients(employee_email, cc_email),
                            msg.as_string())
        log.info("Reminder  → %s (%s / %s)", employee_name, dept_name, employee_email)
        return True, f"Reminder sent to {employee_email}"
    except Exception as exc:
        log.error("Reminder FAILED → %s: %s", employee_email, exc)
        return False, str(exc)


def send_acknowledgment(
    cfg: dict,
    dept_name: str,
    employee_name: str,
    employee_email: str,
    cc_email: str = "",
    file_name: str = "",
    week_label: str | None = None,
    submitted_at: str | None = None,
) -> tuple[bool, str]:
    """Send an acknowledgment to an individual employee who submitted on time."""
    ecfg         = cfg.get("email", {})
    subject      = ecfg.get("ack_subject", "Weekly Report Received — Thank You")
    wl           = week_label or _current_week_label()
    submitted_at = submitted_at or datetime.now().strftime("%Y-%m-%d %H:%M")

    html  = _ack_html(dept_name, employee_name, file_name, wl, submitted_at)
    plain = _ack_plain(dept_name, employee_name, file_name, wl, submitted_at)
    msg   = _build_message(cfg, employee_email, employee_name,
                            subject, html, plain, cc_email)

    try:
        from_email = (cfg.get("email", {}).get("sender_email", "")
                      or cfg.get("email", {}).get("smtp_user", ""))
        with _make_smtp(cfg) as server:
            server.sendmail(from_email, _recipients(employee_email, cc_email),
                            msg.as_string())
        log.info("Ack       → %s (%s / %s)", employee_name, dept_name, employee_email)
        return True, f"Acknowledgment sent to {employee_email}"
    except Exception as exc:
        log.error("Ack FAILED → %s: %s", employee_email, exc)
        return False, str(exc)


def send_test_email(cfg: dict, to_email: str) -> tuple[bool, str]:
    """Send a test email to verify SMTP settings."""
    html  = _test_html()
    plain = (
        "SMTP Test — Connection Successful\n\n"
        "Your Weekly Report Automator email settings are correctly configured.\n"
        f"Sent: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
    )
    msg = _build_message(
        cfg, to_email, "",
        "[Test] Weekly Report Automator — SMTP Check",
        html, plain,
    )
    try:
        from_email = (cfg.get("email", {}).get("sender_email", "")
                      or cfg.get("email", {}).get("smtp_user", ""))
        with _make_smtp(cfg) as server:
            server.sendmail(from_email, [to_email], msg.as_string())
        log.info("Test email → %s", to_email)
        return True, f"Test email sent to {to_email}"
    except Exception as exc:
        log.error("Test email FAILED: %s", exc)
        return False, str(exc)


# ── Helpers ───────────────────────────────────────────────────────────────────
def _current_week_label() -> str:
    from datetime import timedelta
    today  = datetime.now().date()
    monday = today - timedelta(days=today.weekday())
    return f"Week of {monday.strftime('%d %b %Y')}"


def week_label_for(dt: datetime | None = None) -> str:
    from datetime import timedelta
    d      = (dt or datetime.now()).date()
    monday = d - timedelta(days=d.weekday())
    return f"Week of {monday.strftime('%d %b %Y')}"


# ── Dept-master HTML / plain helpers ─────────────────────────────────────────

def _dept_master_html(
    dept_name: str,
    week_label: str,
    file_name: str,
    encrypted: bool,
    password_hint: str,
) -> str:
    lock_icon = "🔒" if encrypted else "📄"
    enc_note  = (
        f'<p style="margin:8px 0 0;font-size:13px;color:{_WARNING};line-height:1.6;">'
        f"<strong>This file is password-protected.</strong> "
        f"Your access password has been shared with you separately by the system administrator."
        f"</p>"
        if encrypted else
        f'<p style="margin:8px 0 0;font-size:13px;color:{_MUTED};line-height:1.6;">'
        f"This file is not password-protected. "
        f"Please store it securely."
        f"</p>"
    )

    card = _info_card(
        "#EEF2FF", "#93C5FD", _ACCENT_DARK,
        lock_icon, f"{dept_name} — Department Master Report",
        [
            ("Week",       week_label),
            ("File",       file_name),
            ("Encrypted",  "Yes — password required to open" if encrypted else "No"),
        ],
    )

    body = f"""
    <p style="margin:0 0 8px;font-size:13px;color:{_MUTED};font-weight:600;
              letter-spacing:.5px;text-transform:uppercase;">
      Department Head Report — Confidential
    </p>
    <h2 style="margin:0 0 20px;font-size:20px;font-weight:700;color:{_ACCENT_DARK};">
      {lock_icon} Your Weekly Department Report
    </h2>
    <p style="margin:0 0 16px;font-size:15px;color:#374151;line-height:1.7;">
      Please find attached the consolidated master report for the
      <strong>{dept_name}</strong> department for <strong>{week_label}</strong>.
      This report contains all submitted weekly data from your team and is
      intended for your eyes only.
    </p>
    {card}
    {enc_note}
    {_divider()}
    <p style="margin:0;font-size:13px;color:{_MUTED};line-height:1.7;">
      If you have any questions about the data or cannot open the attachment,
      please contact your system administrator.
    </p>"""
    return _base_html(body, header_color=_ACCENT_DARK)


def _dept_master_plain(
    dept_name: str,
    week_label: str,
    file_name: str,
    encrypted: bool,
) -> str:
    enc_line = (
        "This file is PASSWORD-PROTECTED. Your password was shared separately."
        if encrypted
        else "This file is not password-protected. Please store it securely."
    )
    return (
        f"Department Head Report — CONFIDENTIAL\n"
        f"{'=' * 44}\n\n"
        f"Department : {dept_name}\n"
        f"Week       : {week_label}\n"
        f"File       : {file_name}\n\n"
        f"{enc_line}\n\n"
        f"Please find the consolidated master report for your department attached.\n"
        f"This report is intended for department heads only.\n\n"
        f"If you have questions, contact your system administrator."
    )


# ── Public: send dept master to department head ───────────────────────────────

def send_dept_master(
    cfg: dict,
    dept_master_result: "Any",  # dept_master.DeptMasterResult
    week_label: str | None = None,
) -> tuple[bool, str]:
    """Send the password-protected department master Excel to the department head.

    Parameters
    ----------
    cfg:
        Parsed config.json dict (must contain ``email`` section with SMTP settings).
    dept_master_result:
        A ``DeptMasterResult`` named-tuple from ``dept_master.generate_dept_masters()``.
        Must have: dept_name, output_path, encrypted, dept_head_email, error.
    week_label:
        Human-readable week string; auto-derived when omitted.

    Returns
    -------
    (True, success_message) on success, (False, error_message) on failure.
    """
    # Guard: skip if result has an error or no recipient
    if dept_master_result.error:
        return False, f"Skipped — generation error: {dept_master_result.error}"
    if not dept_master_result.dept_head_email:
        return False, "Skipped — no dept_head_email configured"
    if not dept_master_result.output_path or not Path(dept_master_result.output_path).exists():
        return False, f"Skipped — output file not found: {dept_master_result.output_path}"

    dept_name  = dept_master_result.dept_name
    to_email   = dept_master_result.dept_head_email
    file_path  = Path(dept_master_result.output_path)
    encrypted  = dept_master_result.encrypted
    wl         = week_label or week_label_for()
    subject    = f"[{dept_name}] Weekly Department Master Report — {wl}"

    html  = _dept_master_html(dept_name, wl, file_path.name, encrypted, "")
    plain = _dept_master_plain(dept_name, wl, file_path.name, encrypted)

    # Build a multipart/mixed message so we can attach the file
    ecfg        = cfg.get("email", {})
    sender_name = ecfg.get("sender_name", "Weekly Report System")
    from_email  = ecfg.get("sender_email", "") or ecfg.get("smtp_user", "")

    msg = MIMEMultipart("mixed")
    msg["Subject"] = subject
    msg["From"]    = f"{sender_name} <{from_email}>"
    msg["To"]      = to_email

    # Attach alternative text/html body
    body_part = MIMEMultipart("alternative")
    body_part.attach(MIMEText(plain, "plain", "utf-8"))
    body_part.attach(MIMEText(html,  "html",  "utf-8"))
    msg.attach(body_part)

    # Attach the Excel file
    try:
        with open(file_path, "rb") as fp:
            attachment = MIMEBase(
                "application",
                "vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            attachment.set_payload(fp.read())
        encoders.encode_base64(attachment)
        attachment.add_header(
            "Content-Disposition",
            "attachment",
            filename=file_path.name,
        )
        msg.attach(attachment)
    except Exception as exc:
        return False, f"Could not read attachment {file_path.name}: {exc}"

    # Send
    try:
        with _make_smtp(cfg) as server:
            server.sendmail(from_email, [to_email], msg.as_string())
        log.info(
            "Dept master → %s (%s, encrypted=%s, file=%s)",
            to_email, dept_name, encrypted, file_path.name,
        )
        return True, f"Dept master sent to {to_email}"
    except Exception as exc:
        log.error(
            "Dept master FAILED → %s (%s): %s",
            to_email, dept_name, exc,
        )
        return False, str(exc)
