# risk-python — Epic Implementation Plan

Generated: 2026-04-20  
Project: risk-python (LH Fund Risk Automation)  
Repo: https://github.com/lhfund/risk-python  
Notion Board: https://www.notion.so/348d9f1fcda181688020dff322287c79

---

## Status Overview

| Epic | Name | Tasks | Status |
|------|------|-------|--------|
| A | Foundation | A1, A2, A3 | ✅ DONE |
| B | Reliability & Error Handling | B1, B2 | ✅ DONE |
| C | Security | C1, C2 | ⬜ TODO |
| D | CI & Developer Experience | D1, D2 | ⬜ TODO |

---

## ✅ Epic A — Foundation (DONE — commit on develop)

**Goal:** Make scripts reproducible and observable on any Windows machine.

### A1 — Create requirements.txt ✅
- **File:** `requirements.txt` (project root)
- **Content:** pinned versions for selenium, webdriver-manager, pandas, openpyxl, python-dotenv, pywin32
- **Why:** Without this, new developers cannot set up the environment

### A2 — Fix Credential Loading Path ✅
- **Files changed:** `Run_morning_ThaiBMA.py`, `Run_morning_ThaiBMA_part_2.py`, `GPO.py`
- **Change:** Replace hardcoded `~/Desktop/PP.env` with:
  ```python
  env_path = os.environ.get('RISK_ENV_PATH') or os.path.join(os.path.expanduser('~'), 'Desktop', 'PP.env')
  ```
- **How to use on new machine:**
  ```bat
  set RISK_ENV_PATH=C:\path\to\your\credentials.env
  python Run_morning_ThaiBMA.py
  ```
- **Backwards compat:** Still falls back to Desktop/PP.env if env var not set

### A3 — Structured Logging ✅
- **File created:** `risk_logger.py` (shared logging module)
- **Log format:** `YYYY-MM-DD HH:MM:SS | LEVEL | script_name | message`
- **Log location:** `logs/` directory (auto-created), one file per script per day
  - `logs/Run_morning_ThaiBMA_YYYYMMDD.log`
  - `logs/Run_morning_ThaiBMA_part2_YYYYMMDD.log`
  - `logs/GPO_YYYYMMDD.log`
- **Usage in new scripts:**
  ```python
  from risk_logger import get_logger
  log = get_logger("script_name")
  log.info("Step completed")
  log.error("Failed", exc_info=True)
  ```

---

## ⬜ Epic B — Reliability & Error Handling

**Goal:** Ensure failures are detected immediately and reported to the team.

### B1 — Failure Email Alert When Morning Script Fails
- **File to create/modify:** `Run_morning_ThaiBMA.py`, `Run_morning_ThaiBMA_part_2.py`
- **Approach:** Wrap entire script in try/except at top level. On exception, send failure email via Outlook (reuse win32com pattern from GPO.py)
- **Email target:** `risk@lhfund.co.th` (or a dedicated ops alias)
- **Pattern:**
  ```python
  try:
      # ... all existing script logic ...
  except Exception as e:
      log.critical(f"Script failed: {e}", exc_info=True)
      send_failure_alert("Run_morning_ThaiBMA", str(e))
  ```
- **`send_failure_alert` function** (add to `risk_logger.py` or a new `risk_alerts.py`):
  ```python
  def send_failure_alert(script_name: str, error_msg: str):
      import win32com.client
      ol = win32com.client.Dispatch('Outlook.Application')
      mail = ol.CreateItem(0)
      mail.Subject = f"[ALERT] {script_name} FAILED — {datetime.now().strftime('%Y-%m-%d %H:%M')}"
      mail.To = 'risk@lhfund.co.th'
      mail.Body = f"Script: {script_name}\nError: {error_msg}\nCheck log: logs/{script_name}_*.log"
      mail.Send()
  ```
- **Effort:** M | **Priority:** High

### B2 — Consolidate GPO.py + run_evening.py
- **Files:** `GPO.py` (keep), `file dib/run_evening.py` (delete after review)
- **Approach:**
  1. Diff the two files to confirm GPO.py is the superset
  2. Apply any unique fixes from run_evening.py into GPO.py
  3. Delete `file dib/run_evening.py`
  4. Update `file dib/run_evening.bat` to point to `GPO.py` instead
- **Key difference found:** run_evening.py has `excel_app.Visible = False` vs GPO.py's `True` — decide which is correct
- **Effort:** S | **Priority:** Medium

---

## ⬜ Epic C — Security

**Goal:** Prevent credential leaks and add GitHub-level protection.

### C1 — Audit .gitignore and Verify No Credentials Are Staged
- **Files:** `.gitignore` (already created), `scripts/.env`
- **Steps:**
  1. Run `git status` to confirm .env files are NOT tracked
  2. Run `git log --all --full-history -- "*.env"` to verify no env files in git history
  3. If found in history → use `git filter-repo` or BFG to purge
  4. Add `logs/` to .gitignore (log files should not be committed)
  5. Confirm `*.xlsm` is also excluded
- **Effort:** S | **Priority:** High

### C2 — Add GitHub Actions Secret Scanning Workflow
- **File to create:** `.github/workflows/secret-scan.yml`
- **Tool:** `gitleaks` (free, open source) OR GitHub native secret scanning (requires GitHub Advanced Security)
- **Pattern using gitleaks:**
  ```yaml
  name: Secret Scan
  on: [push, pull_request]
  jobs:
    gitleaks:
      runs-on: ubuntu-latest
      steps:
        - uses: actions/checkout@v4
          with:
            fetch-depth: 0
        - uses: gitleaks/gitleaks-action@v2
          env:
            GITHUB_TOKEN: ${{ secrets.GITHUB_TOKEN }}
  ```
- **Blocks merge** if any secret pattern detected in diff
- **Effort:** S | **Priority:** High

---

## ⬜ Epic D — CI & Developer Experience

**Goal:** Make it safe to contribute changes and easy for new developers to onboard.

### D1 — GitHub Actions Lint Workflow (flake8)
- **File to create:** `.github/workflows/lint.yml`
- **Pattern:**
  ```yaml
  name: Lint
  on: [push, pull_request]
  jobs:
    lint:
      runs-on: windows-latest
      steps:
        - uses: actions/checkout@v4
        - uses: actions/setup-python@v5
          with:
            python-version: '3.11'
        - run: pip install flake8
        - run: flake8 *.py --max-line-length=120 --ignore=E501,W503
  ```
- **Note:** Use `windows-latest` because pywin32 is Windows-only. Alternatively lint on ubuntu and skip win32 imports.
- **Effort:** S | **Priority:** Medium

### D2 — Create README.md with Setup Guide
- **File to create:** `README.md`
- **Sections to include:**
  1. Project overview (what this automates)
  2. Prerequisites (Windows, Office, Anaconda, Chrome, network access)
  3. Installation steps:
     ```bat
     git clone https://github.com/lhfund/risk-python.git
     cd risk-python
     pip install -r requirements.txt
     ```
  4. Credential setup (create PP.env, set RISK_ENV_PATH)
  5. Running the scripts (morning, evening, standalone)
  6. Log location (`logs/` directory)
  7. Network paths reference
  8. Troubleshooting common issues (ThaiBMA login fails, network share not found, Excel macro error)
- **Effort:** M | **Priority:** Medium

---

## Implementation Order (recommended)

```
Epic A ✅ → Epic B (B1 first, then B2) → Epic C (C1 first) → Epic D
                ↑ most impactful for reliability        ↑ security    ↑ DX
```

## How to Continue

When ready to work on the next epic, tell Claude Code:
- `"start Epic B"` — implement B1 then B2
- `"start Epic C"` — implement C1 then C2
- `"start Epic D"` — implement D1 then D2
- `"start all remaining epics"` — run B → C → D in sequence

## File Locations

| File | Purpose |
|------|---------|
| `requirements.txt` | Python dependencies (A1) |
| `risk_logger.py` | Shared logging module (A3) |
| `logs/` | Runtime log files (auto-created, gitignored) |
| `.gitignore` | Protects credentials and large files |
| `CLAUDE.md` | Project spec for Claude Code |
| `devstarter-config.yml` | DevStarter project config |
| `docs/brd.html` | Business requirements |
| `docs/schema.html` | Database schema |
| `docs/scripts.html` | Scripts reference |
