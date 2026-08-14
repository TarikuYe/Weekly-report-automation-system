# Weekly Report Automator

[![Python Version](https://img.shields.io/badge/python-3.8%2B-blue)](https://www.python.org/downloads/)
[![License](https://img.shields.io/badge/license-MIT-green)](LICENSE)
[![Platform](https://img.shields.io/badge/platform-Windows-lightgrey)](https://www.microsoft.com/windows)

A comprehensive enterprise solution for automating weekly departmental report collection, consolidation, and distribution. Designed for engineering and consulting firms managing multiple departments with recurring weekly reporting requirements.

---

## 🎯 Overview

The Weekly Report Automator streamlines the entire lifecycle of weekly report management:

- **Automated Collection**: Monitors department folders for incoming report files
- **Intelligent Consolidation**: Merges multiple Excel formats into a unified master workbook
- **Email Automation**: Sends reminders to non-submitters and acknowledgments to submitters
- **Department Dashboards**: Generates password-protected per-department summary reports
- **Web Dashboard**: Modern single-page application for monitoring and manual operations
- **Scheduled Execution**: Background scheduler for hands-free weekly operation

---

## ✨ Key Features

### 📊 **Multi-Format Report Processing**
- Supports flat-table spreadsheets and per-employee weekly templates
- Automatic header detection and column normalization
- Intelligent schema matching with graceful degradation
- Preserves unmatched data in separate sections (no data loss)

### 📧 **Smart Email Automation**
- **Deadline-aware reminders** for employees who haven't submitted
- **Acknowledgment emails** confirming successful submission
- **Per-department CC routing** to department heads
- **HTML-formatted emails** with modern, professional design
- **Configurable schedules** (deadline day, active hours, repeat intervals)

### 🔐 **Security & Access Control**
- Password-protected department master files
- Department heads receive only their team's consolidated data
- Email credentials stored locally (never transmitted to third parties)
- Per-department folder isolation (supports network shares)

### 🎨 **Modern Web Dashboard**
- Real-time log streaming via Server-Sent Events
- Drag-and-drop file upload
- Archive browser with batch restoration
- Email test utility with SMTP verification
- Department configuration editor
- Mobile-responsive design

### 📁 **Flexible Folder Architecture**
- Department folders can be local subdirectories or external network shares
- Automatic archiving preserves submission history per department
- Legacy flat-layout support for backward compatibility

### 🚀 **Deployment Options**
- **Flask web app** (`python app.py`)
- **GUI launcher** (`pythonw gui.pyw`)
- **CLI pipeline** (`python main.py --config config.json`)
- **Headless scheduler** (`python auto_runner.py`)
- **PyInstaller builds** for distribution without Python

---

## 📋 Requirements

### Software Dependencies
```
Python 3.8 or higher
pandas >= 3.0
openpyxl >= 3.1.5
flask >= 3.0
pyinstaller >= 6.0
msoffcrypto-tool >= 5.0.0
```

### System Requirements
- **OS**: Windows 10/11 or Windows Server 2016+
- **RAM**: 2 GB minimum (4 GB recommended for large reports)
- **Disk**: 500 MB for application + space for reports and archives

### Optional
- SMTP server access (Gmail, Office 365, or corporate mail server)
- Network share access for departmental folder routing

---

## 🛠️ Installation

### Quick Start

1. **Clone or download the repository**
   ```bash
   git clone https://github.com/your-org/weekly-report-automator.git
   cd weekly-report-automator
   ```

2. **Install Python dependencies**
   ```bash
   pip install -r requirements.txt
   ```

3. **Configure the application**
   ```bash
   # Copy the example config
   copy config.example.json config.json
   
   # Edit config.json with your settings
   # IMPORTANT: Never commit config.json - it contains credentials!
   ```
   - Edit `config.json` to set input/output paths and department structure
   - Configure SMTP settings for email automation (optional)
   - **⚠️ SECURITY**: `config.json` is gitignored and should NEVER be committed

4. **Launch the web dashboard**
   ```bash
   python app.py
   ```
   Open your browser to `http://localhost:5000`

### Building Executables (Optional)

Create standalone `.exe` files for deployment without Python:

```bash
# Build the web dashboard
pyinstaller --onefile --windowed ^
    --add-data "templates;templates" ^
    --add-data "static;static" ^
    --name "WeeklyReportDashboard" ^
    app.py

# Build the GUI launcher
pyinstaller --onefile --windowed ^
    --icon=icon.ico ^
    --name "WeeklyReportGUI" ^
    gui.pyw
```

Executables will appear in the `dist/` folder.

---

## 📖 Usage

### Web Dashboard (Recommended)

**Start the server:**
```bash
python app.py --port 5000 --host 127.0.0.1
```

**Features:**
- **Pipeline tab**: Run consolidation manually, view real-time logs
- **Files tab**: Upload reports, manage input files, browse archives
- **Email tab**: Configure departments, send test emails, trigger reminders
- **Config tab**: Edit paths and SMTP settings

### CLI Mode

**Run the consolidation pipeline:**
```bash
python main.py --config config.json
```

**Check submission status:**
```bash
python email_checker.py config.json
```

**Send email cycle manually:**
```bash
python email_checker.py config.json --send
```

**Headless scheduled run:**
```bash
python auto_runner.py --config config.json --force
```

### GUI Mode

**Launch the graphical interface:**
```bash
pythonw gui.pyw
```

---

## ⚙️ Configuration

### config.json Structure

⚠️ **IMPORTANT**: Copy `config.example.json` to `config.json` before editing. Never commit `config.json`!

```json
{
  "input_dir": "C:/Reports/input",
  "output_file": "C:/Reports/output/master_weekly_report.xlsx",
  
  "email": {
    "enabled": true,
    "smtp_host": "smtp.gmail.com",
    "smtp_port": 587,
    "smtp_user": "your-email@company.com",
    "smtp_password": "your-app-password",
    "use_tls": true,
    "sender_name": "Weekly Report System",
    "sender_email": "reports@company.com",
    "deadline_day": "Thursday",
    "deadline_hour": 17,
    "deadline_minute": 0,
    "reminder_subject": "Reminder: Weekly Report Not Yet Submitted",
    "ack_subject": "Weekly Report Received — Thank You"
  },
  
  "auto_check": {
    "enabled": true,
    "interval_minutes": 30,
    "window_start_hour": 7,
    "window_end_hour": 20,
    "send_acks": true,
    "send_reminders": true,
    "reminder_repeat_hours": 24
  },
  
  "departments": [
    {
      "name": "Engineering",
      "folder": "C:/Reports/input/Engineering",
      "cc_email": "eng-head@company.com",
      "dept_head_email": "eng-head@company.com",
      "dept_head_password": "secure123",
      "employees": [
        {
          "name": "John Doe",
          "email": "john.doe@company.com",
          "expected_file": "John_Weekly_Report.xlsx"
        }
      ]
    }
  ]
}
```

### Department Folder Modes

**Mode 1: Subdirectory (legacy)**
```json
"folder": "Engineering"
```
Resolves to `{input_dir}/Engineering/`

**Mode 2: Absolute path (network share)**
```json
"folder": "//SERVER/shared/Engineering"
```
Allows departments to submit to external locations

**Mode 3: Disabled**
```json
"folder": ""
```
Department is tracked but no folder is scanned

---

## 🔄 Workflow

### Typical Weekly Cycle

```
Monday 00:00   → New week begins, archive folders are empty
Tuesday 14:00  → Employees start submitting reports to their dept folders
Thursday 17:00 → Deadline passes
Thursday 17:30 → Auto-checker sends reminders to non-submitters
                 Sends acknowledgments to submitters
Friday 09:00   → All reports received
Friday 09:30   → Administrator runs pipeline (manual or scheduled)
                 → Consolidates all department reports
                 → Generates master report + dept dashboards
                 → Archives input files
                 → Emails dept masters to department heads
Friday 10:00   → Department heads open their password-protected reports
```

### Archive Structure

After each run, input files are moved to timestamped archive folders:

```
input/
├── archive/
│   └── 2026-08-15_09-30-00/
│       └── root_level_file.xlsx
├── Engineering/
│   └── archive/
│       └── 2026-08-15_09-30-00/
│           └── eng_report.xlsx
└── Finance/
    └── archive/
        └── 2026-08-15_09-30-00/
            └── finance_report.xlsx
```

---

## 📊 Output Structure

### Master Report Excel File

**Sheet 1: Summary**
- Executive KPI cards (total hours, departments, completion rate)
- Department comparison table with charts
- Status breakdown across all departments
- Project-level hour rollups
- Data quality notes

**Sheet 2: Details**
- Unified table with all matched (rollup-eligible) rows
- Excel AutoFilter and freeze panes enabled
- Status-based conditional formatting
- Followed by unmatched data sections (separate visual blocks)

### Department Master Files

Located in `output/dept_masters/`:
```
output/
└── dept_masters/
    ├── Engineering_Week_of_12_Aug_2026.xlsx  (password-protected)
    ├── Finance_Week_of_12_Aug_2026.xlsx      (password-protected)
    └── ...
```

Each file contains:
- Only that department's data
- Same Summary + Details layout as master report
- Password-protected with per-department password
- Automatically emailed to dept_head_email

---

## 🔧 Troubleshooting

### Common Issues

**Issue**: *"No .xlsx files found in input directory"*
- **Solution**: Verify `input_dir` path in `config.json`. Ensure department folders exist and contain `.xlsx` files.

**Issue**: *"SMTP authentication failed"*
- **Solution**: 
  - For Gmail: Generate an [App Password](https://support.google.com/accounts/answer/185833)
  - For Office 365: Enable SMTP AUTH in admin portal
  - Verify `smtp_host`, `smtp_port`, and credentials

**Issue**: *"PermissionError: [Errno 13] Permission denied"*
- **Solution**: Close the output Excel file before running the pipeline. The app will create a timestamped alternative if the file is locked.

**Issue**: *"Department master contains no data"*
- **Solution**: Verify the `Department` column in input files matches the `name` field in `config.json` exactly (case-sensitive).

**Issue**: *"Email sent but not received"*
- **Solution**:
  - Check spam/junk folders
  - Verify recipient email addresses in config
  - Send a test email via the dashboard to verify SMTP settings
  - Check `logs/app.log` for delivery errors

### Logs

All operations are logged to the `logs/` directory:

- `app.log` — Web dashboard and pipeline operations
- `auto_runner.log` — Scheduled execution logs

Enable verbose logging by setting `logging.basicConfig(level=logging.DEBUG)` in the source files.

---

## 🔐 Security Considerations

### ⚠️ CRITICAL: Never Commit Credentials

**The `config.json` file is gitignored for security reasons.** It contains:
- SMTP passwords
- Email addresses  
- Department passwords

**Setup checklist:**
1. ✅ Copy `config.example.json` to `config.json`
2. ✅ Edit `config.json` with your credentials
3. ✅ Verify `config.json` is in `.gitignore`
4. ❌ **NEVER** run `git add config.json`
5. ❌ **NEVER** commit or push `config.json`

If you accidentally commit credentials:
1. **Immediately revoke** the exposed credential (e.g., delete Gmail app password)
2. Generate a new credential
3. Update `config.json` locally
4. Use `git filter-branch` or BFG to remove from history
5. Force push: `git push --force origin main`

### Email Credentials
- Stored in `config.json` on local disk (never transmitted)
- Use application-specific passwords (not your main account password)
- Restrict file permissions on `config.json` (Windows: Right-click → Properties → Security)

### Department Passwords
- Dept master Excel files are encrypted with `msoffcrypto-tool`
- Passwords are configured per-department in `config.json`
- Passwords are NOT sent via email (communicate separately)

### Network Shares
- Ensure department folders have appropriate NTFS permissions
- Service account running the app needs read access to all dept folders
- Consider using a dedicated service account (not a personal account)

---

## 🚀 Deployment

### Option 1: Manual Execution
Best for small teams or infrequent use.
- Administrator runs `python app.py` and triggers pipeline via web dashboard

### Option 2: Windows Task Scheduler
Best for fully automated weekly operation.

**Steps:**
1. Create a scheduled task (weekly trigger, Friday 09:00)
2. Action: `python.exe`
3. Arguments: `C:\path\to\auto_runner.py --config C:\path\to\config.json`
4. Run whether user is logged on or not
5. Use a service account with appropriate permissions

### Option 3: Windows Service
Best for always-on background operation with web dashboard.
- Use [NSSM](https://nssm.cc/) to wrap `python app.py` as a Windows service
- Configure startup type as Automatic
- Web dashboard remains accessible 24/7

---

## 📚 Advanced Topics

### Custom Column Mappings

Edit `COLUMN_ALIASES` in `main.py` to add support for alternate column names:

```python
COLUMN_ALIASES = {
    "Date": {"date", "report date"},
    "Staff Name": {"staff name", "employee", "engineer"},
    # Add your custom mappings here
}
```

### Email Template Customization

Edit HTML templates in `email_sender.py`:
- `_reminder_html()` — Reminder email design
- `_ack_html()` — Acknowledgment email design
- `_base_html()` — Overall email shell and branding

### Adding New Departments

Via Web Dashboard:
1. Navigate to **Email** tab → **Department Configuration**
2. Click **Add Department**
3. Fill in name, folder path, emails, and employee list
4. Click **Save**

Via config file:
```json
{
  "name": "New Department",
  "folder": "C:/Reports/input/NewDept",
  "cc_email": "newdept-head@company.com",
  "dept_head_email": "newdept-head@company.com",
  "dept_head_password": "password123",
  "employees": [
    {"name": "Jane Smith", "email": "jane@company.com", "expected_file": ""}
  ]
}
```

---

## 🤝 Contributing

Contributions are welcome! Please follow these guidelines:

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

---

## 📄 License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

---

## 👥 Support

For questions, issues, or feature requests:

- **Email**: tarikunegesa19@gmail.com

## 🎓 Acknowledgments

Built with:
- [Pandas](https://pandas.pydata.org/) — Data manipulation
- [OpenPyXL](https://openpyxl.readthedocs.io/) — Excel file handling
- [Flask](https://flask.palletsprojects.com/) — Web framework
- [msoffcrypto-tool](https://github.com/nolze/msoffcrypto-tool) — Excel encryption

---

## 📈 Roadmap

Planned features for future releases:

- [ ] Multi-language support (internationalization)
- [ ] PDF export option for master reports
- [ ] Integration with Microsoft Teams / Slack notifications
- [ ] REST API for external system integration
- [ ] PowerBI connector for live dashboard analytics
- [ ] Mobile app for submission status tracking
- [ ] AI-powered anomaly detection in reported hours

---

<div align="center">

**Weekly Report Automator** — Streamlining enterprise reporting workflows

Made with ❤️ for engineering teams

</div>
