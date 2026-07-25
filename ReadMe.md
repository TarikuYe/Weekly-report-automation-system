# Making the app double-clickable on Windows

Two options below. **Option A** takes 2 minutes and needs no extra tools.
**Option B** builds a real standalone `.exe` — the better long-term choice,
especially if you'll share this with coworkers who don't have Python installed.

Both require these 4 files to sit together in one folder:
`main.py`, `gui.py`, `icon.ico`, `requirements.txt` (and `build_exe.bat` for Option B).

---

## Option A — Quick: no console window, no build step

1. Install Python normally if you haven't (python.org — check "Add to PATH" during install).
2. Open a terminal **once** and run:
   ```
   pip install pandas openpyxl
   ```
3. Rename `gui.py` to `gui.pyw` (same content, different extension — the `.pyw`
   extension is what tells Windows to launch it with `pythonw.exe`, which has
   no console window at all).
4. Right-click `gui.pyw` → **Create shortcut**. Drag that shortcut to your
   Desktop or pin it to Start/Taskbar.
5. (Optional) Right-click the shortcut → Properties → **Change Icon…** →
   browse to `icon.ico`.

Double-clicking the shortcut now opens the app directly — no terminal ever appears.

**Downside:** the machine still needs Python + the two packages installed.
Fine for your own PC; not great for sharing with others.

---

## Option B — Proper standalone `.exe` (recommended)

This bundles Python itself into one `.exe`. Anyone can double-click it, even
on a machine with no Python installed at all.

1. Put `main.py`, `gui.py`, `icon.ico`, `requirements.txt`, and
   `build_exe.bat` in the same folder.
2. Double-click **`build_exe.bat`**.
   - First run installs `pandas`, `openpyxl`, and `pyinstaller`.
   - Then it builds the app. This takes a minute or two — a terminal
     window will show progress and close (or pause) when done. This is a
     one-time build step, not something end users ever see.
3. Find your app at `dist\WeeklyReportAutomator.exe`.
4. That single file is the whole app. Copy it anywhere — Desktop, a shared
   drive, a USB stick — and double-click it. No terminal, no Python
   installation required on the target machine.
5. (Optional) Right-click the `.exe` → **Send to → Desktop (create shortcut)**
   for a proper desktop icon.

### Notes
- The first launch of a freshly built `.exe` can take a few seconds longer
  than normal (Windows/antivirus scanning an unfamiliar file) — after that
  it opens quickly.
- `config.json`, the `logs\` folder, and default `input\`/`output\` folders
  are all created next to wherever the `.exe` lives, so it's self-contained —
  safe to move the whole folder around.
- If Windows SmartScreen warns "Windows protected your PC" the first time
  you run a freshly built exe, that's normal for unsigned apps — click
  **More info → Run anyway**. To avoid this warning entirely you'd need a
  code-signing certificate, which isn't necessary for personal/internal use.
- To rebuild after editing `gui.py` or `main.py`, just double-click
  `build_exe.bat` again.