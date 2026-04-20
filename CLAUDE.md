# CLAUDE.md — risk-python (LH Fund Risk Automation)

## Project Overview

**risk-python** is a Windows-based Python automation suite for the Risk team at LH Fund Management (LHFund).
It automates daily morning and evening data workflows: scraping bond index data from ThaiBMA, running Excel VBA macros,
querying SQL Server for NAV/Holdings/VaR data, and sending automated email reports via Outlook.

---

## Tech Stack

| Layer | Technology |
|-------|-----------|
| Language | Python 3 (Anaconda) |
| Automation | Windows .bat files, win32com (Excel + Outlook) |
| Web Scraping | Selenium + webdriver_manager (Chrome) |
| Data | pandas, openpyxl |
| Database | SQL Server (pyodbc / raw SQL) |
| Config | python-dotenv (.env files) |
| VCS | Git → GitHub (`https://github.com/lhfund/risk-python.git`) |

---

## Repository Structure

```
risk-python/
├── Run_morning_ThaiBMA.py          # Morning: Selenium scrape ThaiBMA bond indexes (Part 1)
├── Run_morning_ThaiBMA_part_2.py   # Morning: ThaiBMA scrape Part 2 (ZRR, Corp Bond, etc.)
├── GPO.py                          # Evening: run Excel VBA macros + send Benchmark email
├── GPO_only_send_email.py          # Standalone: send GPO performance email via Outlook
├── PP.env                          # ⚠️ CREDENTIALS — ThaiBMA username/password (gitignored)
├── two.env                         # ⚠️ CREDENTIALS — additional credentials (gitignored)
├── power_automate_for_Afternoon.xlsm  # Excel macro file (afternoon workflow)
├── NAV_Complete.xlsm               # Excel NAV workbook
├── SQL/
│   └── SQL Code/
│       ├── Bloomberg.sql           # Bloomberg MTM FX/EQ pivot query
│       ├── Bloomberg_raw.sql       # Bloomberg raw data query
│       ├── Final_VaR.sql           # VaR (Value at Risk) parametric calculation
│       ├── Holding_Daily.sql       # Daily portfolio holdings query
│       ├── Holding_Daily_2.sql
│       ├── Holding_Daily_3.sql
│       ├── Call_stock_price_create_bench.sql
│       ├── Tier.sql
│       └── Call NAV/
│           ├── MF_NAV.sql          # Mutual Fund NAV
│           ├── MF_TotalNAV.sql
│           ├── PF_NAV.sql          # Private Fund NAV
│           ├── PF_TotalNAV.sql
│           ├── PVD_NAV.sql         # Provident Fund NAV
│           └── PVD_TotalNAV.sql
├── file dib/
│   ├── run_evening.py              # Evening workflow (alternative/dev version)
│   ├── run_evening.bat             # BAT runner for evening via Anaconda
│   ├── Run_Morning - Copy.bat      # BAT runner for morning via Anaconda
│   ├── Combine_all_code.ipynb      # Combined workflow notebook (dev)
│   ├── Part2_complete.ipynb        # Part 2 notebook
│   └── True_After_part1.ipynb     # Post-part1 notebook
├── logfile_formorning/             # Downloaded morning data files (ThaiBMA, Bloomberg)
├── Logfile_forevening/             # Downloaded evening data files
├── docs/                           # Project documentation (HTML)
└── memory/                         # Claude Code progress tracking
```

---

## Workflows

### Morning Workflow (Run daily before market open)

1. **`Run_morning_ThaiBMA.py`** — Logs into ThaiBMA website via Selenium, downloads:
   - Morningstar Benchmark index
   - ZRR index, Short-term index, Bond index, MTM Corp index, Composite index, Corp ZRR index
   - Copies files to network share `\\w2fsspho101.lhfund.net\FM-RI$\risk\`
2. **`Run_morning_ThaiBMA_part_2.py`** — Downloads remaining ThaiBMA indexes

### Evening Workflow (Run after market close)

1. **`GPO.py`** — Runs Excel VBA macro `Create_Afternoon` on `power_automate_for_Afternoon.xlsm`
   - Runs `evening` macro on `Benchmark.xlsm`
   - Sends GPO update email via Outlook to `risk@lhfund.co.th; operation@lhfund.co.th`

### Standalone Email

- **`GPO_only_send_email.py`** — Reads GPO-FIXED and GPO-EQ Excel files from Bloomberg share,
  formats HTML table, sends via Outlook

---

## SQL Databases

| Database | Description |
|----------|-------------|
| `LHF_SYSTEM.DBO` | Bloomberg MTM FX/EQ morning data (`LHF_BBG_DL_MTM_FX_EQ_MORNING`) |
| `LHF_PERFORMANCE.dbo` | Fund NAV view (`ViewFundNavAll`) |
| `INV_LHF.INVEST` | Portfolio holdings (`HOLDING`) |
| `FIN_REG_LHF.dbo` | Holiday calendar (`holiday`) |
| `192.168.102.7\DB2008` | Legacy DB2 instance for holiday data |

---

## Environment / Credentials

Credentials are loaded from `.env` files (never committed to git):

| File | Contents | Loaded By |
|------|----------|-----------|
| `PP.env` | ThaiBMA `user` + `pass` | `Run_morning_ThaiBMA.py` |
| `two.env` | Additional credentials | Other scripts |

Both files are loaded from `~/Desktop/PP.env` path via `python-dotenv`.

**⚠️ CRITICAL: Never commit `.env` files. Always check `.gitignore` before pushing.**

---

## Network Paths (Windows SMB shares)

| Path | Purpose |
|------|---------|
| `\\w2fsspho101.lhfund.net\FM-RI$\risk\` | Primary risk data share |
| `\\172.16.21.100\Risk$\Morningstar Benchmark` | Morningstar source files |
| `P:\Bloomberg\Management Fee for PVD\` | Bloomberg GPO files |

---

## Known Issues / Tech Debt

1. **No requirements.txt** — dependencies are implicit; new dev must guess what to install
2. **No error handling / retry logic** in some scripts (Selenium loops exist but swallow exceptions)
3. **Hardcoded network paths** — breaks if run on a machine without mapped drives
4. **No tests** — zero test coverage
5. **Credentials in PP.env on Desktop** — non-standard path; breaks if Desktop path differs
6. **Duplicate code** — `GPO.py` and `file dib/run_evening.py` have near-identical logic

---

## Development Notes

- **Runtime: Windows only** (win32com requires Windows + Office installed)
- **Python: Anaconda** (path hardcoded in .bat files: `C:\ProgramData\anaconda3\python.exe`)
- **Browser: Chrome** (webdriver_manager auto-installs ChromeDriver)
- **Run order**: Morning scripts → SQL queries → Evening scripts
- **Logs**: Downloaded files stored in `logfile_formorning/` and `Logfile_forevening/`

---

## Git & GitHub

- Remote: `https://github.com/lhfund/risk-python.git`
- Branch strategy: `main` (production) + feature branches
- **Must .gitignore**: `*.env`, `*.xlsm`, `Logfile_*`, `logfile_*`, `__pycache__/`, `*.zip`
