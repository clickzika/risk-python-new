# =============================================================================
# config.py — Single source of truth for all environment-specific values.
# Edit this file when paths, URLs, macros, or recipients change.
# Credentials (user/pass) stay in scripts/.env — never put them here.
# =============================================================================

# === NETWORK SHARE BASE ===
_RISK_SHARE = r'\\w2fsspho101.lhfund.net\FM-RI$\risk'

# === MORNING PATHS ===
MORNING_DL_DIR   = _RISK_SHARE + r'\Amornsiri\logfile_formorning\From_load'
MORNING_DATA_DIR = _RISK_SHARE + r'\Amornsiri\logfile_formorning'
MORNINGSTAR_SRC  = r'\\172.16.21.100\Risk$\Morningstar Benchmark'
DATA_FILE_DIR    = _RISK_SHARE + r'\98.Data File'
LH_REPORT_DEST   = DATA_FILE_DIR + r'\D_LHReport.xlsx'
LH_REPORT_GLOB   = r'P:\Fund_EQ\**\*PortVal*'
POWER_AUTOMATE   = _RISK_SHARE + r'\Amornsiri\power_automate.xlsm'

# === EVENING PATHS ===
EVENING_DL_DIR        = _RISK_SHARE + r'\Amornsiri\Logfile_forevening\From_load'
EVENING_DATA_DIR      = _RISK_SHARE + r'\Amornsiri\Logfile_forevening'
POWER_AUTOMATE_PM     = _RISK_SHARE + r'\Amornsiri\power_automate_for_Afternoon.xlsm'
BENCHMARK_XLSM        = _RISK_SHARE + r'\1.Risk Report\1.Daily Risk Report\Benchmark.xlsm'
BENCHMARK_XLSM_NAME   = 'Benchmark.xlsm'
SET_TRI_GLOB          = r'P:\###RISK###\SET TRI\202*\*\*'

# === P: DRIVE — BLOOMBERG GPO FILES ===
GPO_FIXED_FILE = r'P:\Bloomberg\Management Fee for PVD\Management Fee for PVD_GPO-FIXED - LHFUND_REVISED_RATE.xls'
GPO_EQ_FILE    = r'P:\Bloomberg\Management Fee for PVD\Management Fee for PVD_GPO-EQ - LHFUND  REVISED_RATE.xls'

# === THAIBMA URLs ===
THAIBMA_LOGIN      = 'https://www.ibond.thaibma.or.th/login?page=/bondsearch/bondsearchpage'
THAIBMA_YIELD_GOV  = 'https://www.ibond.thaibma.or.th/yield-curve/government'
THAIBMA_MTM_GOV    = 'https://www.ibond.thaibma.or.th/mtm-gov-index'
THAIBMA_CP         = 'https://www.ibond.thaibma.or.th/cp-index'
THAIBMA_MTM_CORP   = 'https://www.ibond.thaibma.or.th/mtm-corp-index'
THAIBMA_ESG        = 'https://www.ibond.thaibma.or.th/esg-index'
THAIBMA_ZRR        = 'https://www.ibond.thaibma.or.th/zrr-index'
THAIBMA_SHORTTERM  = 'https://www.ibond.thaibma.or.th/shortterm-index'
THAIBMA_BOND       = 'https://www.ibond.thaibma.or.th/bond-index'
THAIBMA_COMPOSITE  = 'https://www.ibond.thaibma.or.th/composite-index'
THAIBMA_CORP_ZRR   = 'https://www.ibond.thaibma.or.th/corp-zrr-index'
THAIBMA_LOGIN_ST   = 'https://www.ibond.thaibma.or.th/login?page=/shortterm-index'
THAIBMA_LOGIN_BOND = 'https://www.ibond.thaibma.or.th/login?page=/bond-index'
THAIBMA_LOGIN_CORP = 'https://www.ibond.thaibma.or.th/login?page=/mtm-corp-index'

# === EXCEL MACROS ===
MACRO_UPDATE_DATA    = 'UpdateData'
MACRO_UPDATE_DATA2   = 'UpdateData2'
MACRO_FINISH_BMA     = 'finish_BMA'
MACRO_NEW_FINISH_SET = 'new_finish_set'
MACRO_CREATE_PM      = 'Create_Afternoon'
MACRO_EVENING_BENCH  = 'evening'

# === EMAIL ===
EMAIL_RECIPIENTS = 'risk@lhfund.co.th ; operation@lhfund.co.th'

# === SQL SERVER CONNECTIONS ===
# Primary server hosts: LHF_SYSTEM, LHF_PERFORMANCE, INV_LHF, FIN_REG_LHF
DB_MAIN_SERVER  = r'YOUR_MAIN_SQL_SERVER'  # e.g. r'192.168.1.10\SQLEXPRESS'
DB_LEGACY_SERVER = r'192.168.102.7\DB2008'  # Legacy — holiday calendar

# === SQL QUERY OUTPUT ===
SQL_OUTPUT_DIR = _RISK_SHARE + r'\Amornsiri\SQL_Output'

# === VAR DEFAULT PARAMETERS ===
VAR_LOOKBACK_DAYS  = 250   # trading days for historical VaR
VAR_CONFIDENCE_PCT = 95    # confidence level (%)

# === FILE MAPPINGS — MORNING PART 1 ===
MORNING_PART1_FILE_MAPPINGS = [
    ('MTMGov_G1_Index_',        'MTMGov_G1'),
    ('MTMGov_G2_Index_',        'MTMGov_G2'),
    ('MTMGov_Index_202',        'Gov_Index'),
    ('YieldTTM_202',            'D_YieldTTM'),
    ('CP_Aminusup_Index',       'Commercial Paper Index'),
    ('MTMCorp_Aminusup_G1_Index', 'CorBond_G1'),
    ('ESGGOV_Index',            'ESGGovBond'),
]

# === FILE MAPPINGS — MORNING PART 2 ===
MORNING_PART2_FILE_MAPPINGS = [
    ('ZRRIndexAll',                  'ZRR_All'),
    ('STGov_Index_',                 'ST_Gov'),
    ('GOV_Index_20',                 'Gov'),
    ('GOV_Index_G1_',                'Gov_G1'),
    ('GOV_Index_G2_',                'Gov_G2'),
    ('MTMCorp_BBBplusup_Index_',     'MTMCorp_BBBplus'),
    ('MTMCorp_BBBplusup_G1_Index_',  'MTMCorp_BBBplus_G1'),
    ('MTMCorp_BBBplusup_G2_Index_',  'MTMCorp_BBBplus_G2'),
    ('Composite_Index_',             'Composite'),
    ('CorpZRR_A_1Y_Index_',          'CorpZRR_A_1Y'),
]

# === FILE MAPPINGS — EVENING ===
EVENING_FILE_MAPPINGS = [
    ('GOV_Index_G1',        'GOV_Index_G1_after'),
    ('MTMCorp_BBBplusup_G1', 'MTMCorp_BBBplusup_G1'),
    ('STGov_Index',          'STGov_Index'),
]
