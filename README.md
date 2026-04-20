# risk-python — LH Fund Risk Automation

Python automation suite for the Risk team at LH Fund Management.
Runs daily morning and evening workflows: scraping bond index data from ThaiBMA,
triggering Excel VBA macros, and sending GPO update emails via Outlook.

---

## Prerequisites

| Requirement | Notes |
|-------------|-------|
| Windows 10/11 | win32com requires Windows + Office |
| Microsoft Office (Excel + Outlook) | Must be installed and activated |
| Python 3.x (Anaconda) | Default path: `C:\ProgramData\anaconda3\python.exe` |
| Microsoft Edge | Selenium uses Edge via EdgeChromiumDriverManager |
| Network access | Must be on LHFund LAN or VPN to reach SMB shares and `ibond.thaibma.or.th` |

---

## Installation

```bat
git clone https://github.com/lhfund/risk-python.git
cd risk-python
pip install -r requirements.txt
```

---

## Credential Setup

Scripts load ThaiBMA username and password from a `.env` file.

1. Copy the example file:
   ```bat
   copy scripts\.env.example scripts\.env
   ```
2. Open `scripts\.env` and fill in your credentials:
   ```
   user=your_thaibma_username
   pass=your_thaibma_password
   ```

To use a custom path (e.g. on a shared machine):
```bat
set RISK_ENV_PATH=C:\path\to\your\credentials.env
```

**Never commit `scripts\.env`.** It is gitignored.

---

## Running the Scripts

### Morning workflow (run before market open)

```bat
runners\run_morning.bat
```

Runs Part 1 then Part 2 in sequence:

| Script | What it does |
|--------|-------------|
| `scripts\morning\run_morning_part1.py` | Login to ThaiBMA, download Yield Gov / MTM Gov / CP / MTM Corp / ESG indexes; copy Morningstar Benchmark; run `UpdateData` macro |
| `scripts\morning\run_morning_part2.py` | Download ZRR / Short-term / Bond / Composite / Corp ZRR indexes; run `UpdateData2` macro |

### Evening workflow (run after market close)

```bat
runners\run_evening.bat
```

Runs `scripts\evening\run_evening.py`:
- Runs `Create_Afternoon` macro on `power_automate_for_Afternoon.xlsm`
- Downloads ST/Bond/Corp indexes from ThaiBMA and transfers them
- Runs `finish_BMA` and `new_finish_set` macros
- Sends GPO update email via Outlook to `risk@lhfund.co.th` and `operation@lhfund.co.th`

### Standalone GPO email

```bat
python scripts\evening\send_gpo_email.py
```

Reads GPO-FIXED and GPO-EQ Excel files from the Bloomberg share and sends the email without running any macros.

---

## Logs

Runtime logs are written to `logs\` (auto-created, gitignored):

```
logs\Run_morning_ThaiBMA_YYYYMMDD.log
logs\Run_morning_ThaiBMA_part2_YYYYMMDD.log
logs\GPO_YYYYMMDD.log
```

Format: `YYYY-MM-DD HH:MM:SS | LEVEL | script | message`

If a script fails, an alert email is sent automatically to `risk@lhfund.co.th`
with the error message and log file path.

---

## Configuration

All paths, URLs, macro names, and email recipients are in `scripts\config.py`.
Edit that file when network paths or macro names change — do not hardcode values in scripts.

| Setting | Location |
|---------|----------|
| Paths / URLs / macros | `scripts\config.py` |
| ThaiBMA credentials | `scripts\.env` |
| Email recipients | `scripts\config.py` → `EMAIL_RECIPIENTS` |

---

## Network Paths

| Path | Purpose |
|------|---------|
| `\\w2fsspho101.lhfund.net\FM-RI$\risk\` | Primary risk data share |
| `\\172.16.21.100\Risk$\Morningstar Benchmark` | Morningstar source files |
| `P:\Bloomberg\Management Fee for PVD\` | Bloomberg GPO files |
| `P:\Fund_EQ\` | LH Report source |
| `P:\###RISK###\SET TRI\` | SET TRI CSV files |

---

## Troubleshooting

**ThaiBMA login fails**
- Confirm `user` and `pass` in `scripts\.env` are correct
- Check that the machine has internet access to `ibond.thaibma.or.th`
- The login loop retries for up to 3 minutes — wait for timeout before investigating

**Network share not found / file copy fails**
- Confirm you are on the LHFund LAN or VPN
- Verify the `P:` drive is mapped and accessible in Explorer

**Excel macro error**
- Office must be fully licensed and activated
- Check that the target `.xlsm` file is not already open by another user
- Look in the log file for the exact COM error message

**Edge driver version mismatch**
- `webdriver_manager` auto-installs the matching EdgeChromiumDriver on first run
- If Edge was recently updated, delete the cached driver: `%USERPROFILE%\.wdm\`

---

## Repository Structure

```
risk-python/
├── scripts/
│   ├── config.py                   # All paths, URLs, macros, recipients
│   ├── .env                        # Credentials (gitignored — copy from .env.example)
│   ├── .env.example                # Credential template
│   ├── morning/
│   │   ├── run_morning_part1.py
│   │   └── run_morning_part2.py
│   └── evening/
│       ├── run_evening.py
│       └── send_gpo_email.py
├── runners/
│   ├── run_morning.bat
│   └── run_evening.bat
├── risk_logger.py                  # Shared logging + failure alert module
├── requirements.txt
├── sql/                            # SQL queries (bloomberg, holdings, nav, var)
├── logs/                           # Runtime logs (gitignored)
└── docs/                           # Project documentation (HTML)
```
