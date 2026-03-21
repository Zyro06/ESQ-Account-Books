import sqlite3
from datetime import datetime
from typing import List, Dict, Any, Optional

# ============================================================
# Numeric suffixes for automatic journal posting.
# Matched against the numeric part of account codes only
# (e.g. '1110' matches 'COA-1110', 'ABC-1110', 'XYZ-1110').
# This makes the app work with any COA prefix the user imports.
# ============================================================
_AR_NUM            = '1110'
_OUTPUT_VAT_NUM    = '2210'
_INPUT_VAT_NUM     = '1320'
_AP_NUM            = '2010'
_SALES_GOODS_NUM   = '4010'
_SALES_SERVICE_NUM = '4020'


def _numeric_suffix(code: str) -> str:
    """Return the digits-only suffix of an account code.
    'COA-1110' -> '1110',  'ABC-1110' -> '1110',  '1110' -> '1110'
    """
    if '-' in code:
        return code.rsplit('-', 1)[-1].strip()
    return code.strip()


class DatabaseManager:
    """Manages all database operations for the accounting system"""

    def __init__(self, db_path: str):
        self.db_path = db_path
        self.connection = None
        self.current_year = datetime.now().year

    def get_connection(self):
        if self.connection is None:
            self.connection = sqlite3.connect(self.db_path)
            self.connection.row_factory = sqlite3.Row
        return self.connection

    def set_current_year(self, year: int):
        self.current_year = year

    def get_current_year(self) -> int:
        return self.current_year

    # ------------------------------------------------------------------
    # Schema creation + seeding
    # ------------------------------------------------------------------

    def initialize_database(self, coa_xlsx_path: str = None, use_default_coa: bool = True):
        """Create all tables if they don't exist, then seed COA if empty."""
        conn = self.get_connection()
        cursor = conn.cursor()

        # Chart of Accounts
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS chart_of_accounts (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                account_code TEXT UNIQUE NOT NULL,
                account_description TEXT NOT NULL,
                normal_balance TEXT DEFAULT 'Debit',
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')

        # Alphalist
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS alphalist (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                tin TEXT UNIQUE NOT NULL,
                company_name TEXT,
                first_name TEXT,
                middle_name TEXT,
                last_name TEXT,
                address1 TEXT,
                address2 TEXT,
                vat TEXT DEFAULT 'VAT Regular',
                entry_type TEXT DEFAULT 'Customer',
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')

        # Migrations for existing databases
        cursor.execute("PRAGMA table_info(alphalist)")
        columns = [c[1] for c in cursor.fetchall()]
        if 'vat' not in columns:
            cursor.execute("ALTER TABLE alphalist ADD COLUMN vat TEXT DEFAULT 'VAT Regular'")
        if 'entry_type' not in columns:
            cursor.execute("ALTER TABLE alphalist ADD COLUMN entry_type TEXT DEFAULT 'Customer'")

        cursor.execute("PRAGMA table_info(chart_of_accounts)")
        coa_cols = [c[1] for c in cursor.fetchall()]
        if 'normal_balance' not in coa_cols:
            cursor.execute(
                "ALTER TABLE chart_of_accounts ADD COLUMN normal_balance TEXT DEFAULT 'Debit'"
            )
            cursor.execute("""
                UPDATE chart_of_accounts SET normal_balance = 'Credit'
                WHERE account_code LIKE 'COA-2%'
                   OR account_code LIKE 'COA-3%'
                   OR account_code LIKE 'COA-4%'
                   OR account_code LIKE 'COA-7%'
            """)
            cursor.execute("""
                UPDATE chart_of_accounts SET normal_balance = 'Credit'
                WHERE account_description LIKE 'ACCUM%DEP%'
                   OR account_description LIKE 'ACCUM%AMORT%'
                   OR account_description LIKE 'ACCUMULATED DEP%'
            """)
            cursor.execute("""
                UPDATE chart_of_accounts SET normal_balance = 'Debit'
                WHERE account_description LIKE '%DRAWING%'
                   OR account_description LIKE '%DRAWINGS%'
            """)

        # Sales Journal
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS sales_journal (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                date DATE NOT NULL,
                customer_name TEXT NOT NULL,
                reference_no TEXT NOT NULL,
                tin TEXT,
                net_amount REAL NOT NULL,
                output_vat REAL NOT NULL,
                gross_amount REAL NOT NULL,
                goods REAL DEFAULT 0,
                services REAL DEFAULT 0,
                particulars TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')

        # Purchase Journal
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS purchase_journal (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                date DATE NOT NULL,
                payee_name TEXT,
                reference_no TEXT NOT NULL,
                tin TEXT,
                branch_code TEXT,
                net_amount REAL NOT NULL,
                input_vat REAL NOT NULL,
                gross_amount REAL NOT NULL,
                account_description TEXT,
                account_code TEXT,
                debit REAL DEFAULT 0,
                credit REAL DEFAULT 0,
                particulars TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')

        # Cash Disbursement Journal
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS cash_disbursement_journal (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                date DATE NOT NULL,
                reference_no TEXT NOT NULL,
                particulars TEXT,
                account_code TEXT NOT NULL,
                account_description TEXT,
                debit REAL DEFAULT 0,
                credit REAL DEFAULT 0,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')

        # Cash Receipts Journal
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS cash_receipts_journal (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                date DATE NOT NULL,
                reference_no TEXT NOT NULL,
                particulars TEXT,
                account_code TEXT NOT NULL,
                account_description TEXT,
                debit REAL DEFAULT 0,
                credit REAL DEFAULT 0,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')

        # General Journal
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS general_journal (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                date DATE NOT NULL,
                reference_no TEXT NOT NULL,
                particulars TEXT,
                account_code TEXT NOT NULL,
                account_description TEXT,
                debit REAL DEFAULT 0,
                credit REAL DEFAULT 0,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')

        conn.commit()

        # Seed COA only when the table is empty
        cursor.execute('SELECT COUNT(*) FROM chart_of_accounts')
        if cursor.fetchone()[0] == 0:
            if coa_xlsx_path:
                self._import_coa_from_xlsx(coa_xlsx_path)
            elif use_default_coa:
                self._seed_default_coa()

    # ------------------------------------------------------------------
    # COA seeding helpers
    # ------------------------------------------------------------------

    def _seed_default_coa(self):
        """Insert the default COA (COA-xxxx codes from COA_-_FIXED)."""
        conn = self.get_connection()
        cursor = conn.cursor()

        default_accounts = [
            # ASSETS
            ('COA-1010', 'CASH IN BANK',                                  'Debit'),
            ('COA-1020', 'CASH ON HAND',                                  'Debit'),
            ('COA-1021', 'PETTY CASH FUND',                               'Debit'),
            ('COA-1110', 'ACCOUNTS RECEIVABLE',                           'Debit'),
            ('COA-1120', 'OTHER RECEIVABLES',                             'Debit'),
            ('COA-1130', 'ADVANCES TO SUPPLIERS',                         'Debit'),
            ('COA-1140', 'ADVANCES TO EMPLOYEES',                         'Debit'),
            ('COA-1150', 'ADVANCES FOR LIQUIDATION',                      'Debit'),
            ('COA-1210', 'INVENTORIES',                                   'Debit'),
            ('COA-1310', 'PREPAID TAXES WITHHELD',                        'Debit'),
            ('COA-1320', 'INPUT VAT',                                     'Debit'),
            ('COA-1330', 'PREPAID EXPENSES',                              'Debit'),
            ('COA-1610', 'COMMUNICATION EQUIPMENT',                       'Debit'),
            ('COA-1611', 'ACCUM DEP - COMMUNICATION EQUIPMENT',           'Credit'),
            ('COA-1620', 'OFFICE EQUIPMENT',                              'Debit'),
            ('COA-1621', 'ACCUM DEP - OFFICE EQUIPMENT',                  'Credit'),
            ('COA-1630', 'TRANSPORTATION EQUIPMENT',                      'Debit'),
            ('COA-1631', 'ACCUM DEP - TRANSPORTATION EQUIPMENT',          'Credit'),
            ('COA-1640', 'COMPUTER EQUIPMENT',                            'Debit'),
            ('COA-1641', 'ACCUM DEP - COMPUTER EQUIPMENT',                'Credit'),
            ('COA-1650', 'FURNITURE AND FIXTURE',                         'Debit'),
            ('COA-1651', 'ACCUM DEP - FURNITURE AND FIXTURE',             'Credit'),
            ('COA-1660', 'LEASEHOLD IMPROVEMENT',                         'Debit'),
            ('COA-1661', 'ACCUM DEP - LEASEHOLD IMPROVEMENT',             'Credit'),
            ('COA-1670', 'COMPUTER SOFTWARE',                             'Debit'),
            ('COA-1671', 'ACCUM AMORT - COMPUTER SOFTWARE',               'Credit'),
            ('COA-1710', 'RENT DEPOSIT',                                  'Debit'),
            ('COA-1720', 'SECURITY DEPOSIT',                              'Debit'),
            # LIABILITIES
            ('COA-2010', 'ACCOUNTS PAYABLE',                              'Credit'),
            ('COA-2020', 'ACCOUNTS PAYABLE - OTHERS',                     'Credit'),
            ('COA-2110', 'PHILHEALTH PREMIUM PAYABLE',                    'Credit'),
            ('COA-2120', 'SSS PREMIUM PAYABLE',                           'Credit'),
            ('COA-2130', 'HDMF PREMIUM PAYABLE',                          'Credit'),
            ('COA-2140', 'SSS LOAN PAYABLE',                              'Credit'),
            ('COA-2150', 'HDMF LOAN PAYABLE',                             'Credit'),
            ('COA-2210', 'OUTPUT VAT PAYABLE',                            'Credit'),
            ('COA-2220', 'WITHHOLDING TAX PAYABLE - COMPENSATION',        'Credit'),
            ('COA-2230', 'INCOME TAX PAYABLE',                            'Credit'),
            ('COA-2310', 'ACCRUED EXPENSES',                              'Credit'),
            ('COA-2320', 'ACCRUED SALARIES',                              'Credit'),
            # EQUITY
            ('COA-3000', 'OWNER, CAPITAL',                                'Credit'),
            ('COA-3100', 'OWNER, DRAWING',                                'Debit'),
            ('COA-3999', 'INCOME SUMMARY',                                'Credit'),
            # REVENUE
            ('COA-4010', 'SALES - GOODS',                                 'Credit'),
            ('COA-4020', 'SALSE - SERVICES',                              'Credit'),
            ('COA-4610', 'SALES DISCOUNT',                                'Debit'),
            ('COA-4620', 'SALES RETURNS',                                 'Debit'),
            ('COA-4630', 'SALES ALLOWANCES',                              'Debit'),
            # COST OF SALES
            ('COA-5010', 'COST OF SALES',                                 'Debit'),
            ('COA-5020', 'PURCHASES',                                     'Debit'),
            ('COA-5030', 'DIRECT LABOR',                                  'Debit'),
            ('COA-5040', 'FACTORY OVERHEAD',                              'Debit'),
            # OPERATING EXPENSES
            ('COA-6010', 'SALARIES EXPENSES - BASIC PAY',                 'Debit'),
            ('COA-6011', 'SALARIES EXPENSES - OVERTIME',                  'Debit'),
            ('COA-6012', 'SALARIES EXPENSES - DE MINIMIS',                'Debit'),
            ('COA-6020', 'GOVERNMENT CONTRIBUTION EXPENSES',              'Debit'),
            ('COA-6021', 'PHILHEALTH EXPESNE - EMPLOYER SHARE',           'Debit'),
            ('COA-6022', 'SSS EXPESNE - EMPLOYER SHARE',                  'Debit'),
            ('COA-6023', 'HDMF EXPESNE - EMPLOYER SHARE',                 'Debit'),
            ('COA-6041', 'UNIFORM',                                       'Debit'),
            ('COA-6060', 'MARKETING EXPENSES',                            'Debit'),
            ('COA-6110', 'DEPRECIATION EXPENSE',                          'Debit'),
            ('COA-6160', 'TRANSPORTATION AND TRAVEL',                     'Debit'),
            ('COA-6161', 'GASOLINE AND OIL',                              'Debit'),
            ('COA-6210', 'INSURANCE EXPENSE',                             'Debit'),
            ('COA-6260', 'PROFESSIONAL FEES',                             'Debit'),
            ('COA-6310', 'RENTAL EXPENSES',                               'Debit'),
            ('COA-6320', 'CUSA EXPENSE',                                  'Debit'),
            ('COA-6360', 'COMMUNICATION EXPENSE',                         'Debit'),
            ('COA-6410', 'UTILITIES',                                     'Debit'),
            ('COA-6460', 'STATIONERY AND SUPPLIES',                       'Debit'),
            ('COA-6461', 'TOOLS AND EQUIPMENT',                           'Debit'),
            ('COA-6510', 'TAXES AND LICENSES',                            'Debit'),
            ('COA-6540', 'FINES AND PENALTIES',                           'Debit'),
            ('COA-6560', 'SUBSCRIPTION EXPENSE',                          'Debit'),
            ('COA-6610', 'REPRESENTATION EXPENSE',                        'Debit'),
            ('COA-6660', 'REPAIRS AND MAINTENANCE',                       'Debit'),
            ('COA-6710', 'TRAININGS AND SEMINARS',                        'Debit'),
            ('COA-6910', 'MISCELLANEOUS EXPENSE',                         'Debit'),
            ('COA-6911', 'SERVICE FEES',                                  'Debit'),
            ('COA-6912', 'MEALS',                                         'Debit'),
            ('COA-6913', 'NOTARIAL EXPENSES',                             'Debit'),
            # OTHER INCOME
            ('COA-7010', 'OTHER INCOME',                                  'Credit'),
            ('COA-7020', 'MISCELLANEOUS INCOME',                          'Credit'),
            ('COA-7030', 'COMMISSION INCOME',                             'Credit'),
            ('COA-7040', 'INTEREST INCOME - BANK DEPOSIT',                'Credit'),
            # INCOME TAX
            ('COA-9010', 'PROVISION FOR INCOME TAX - CURRENT',            'Debit'),
            ('COA-9020', 'PROVISION FOR INCOME TAX - DEFERRED',           'Debit'),
        ]
        cursor.executemany(
            'INSERT OR IGNORE INTO chart_of_accounts '
            '(account_code, account_description, normal_balance) VALUES (?, ?, ?)',
            default_accounts
        )
        conn.commit()

    def _import_coa_from_xlsx(self, xlsx_path: str) -> tuple:
        """
        Import COA from an xlsx file.

        Supported formats:
          - With  DEBIT/CREDIT column: uses that value directly.
          - Without DEBIT/CREDIT column: infers normal balance from the account
            code's numeric prefix using standard accounting conventions:
              1xxx → Asset       → Debit   (except Accumulated Dep/Amort → Credit)
              2xxx → Liability   → Credit
              3xxx → Equity      → Credit  (except Drawing/Drawings → Debit)
              4xxx → Revenue     → Credit
              5xxx → COGS        → Debit
              6xxx → Expenses    → Debit
              7xxx → Other Income→ Credit
              8xxx → Other Exp   → Debit
              9xxx → Tax/Other   → Debit
              anything else      → Debit   (safe default)

        Header row is auto-detected within the first 10 rows.
        Returns (imported_count, error_list).
        """
        try:
            from openpyxl import load_workbook
        except ImportError:
            return 0, ["openpyxl not installed — cannot import xlsx"]

        try:
            wb = load_workbook(xlsx_path, read_only=True, data_only=True)
        except Exception as e:
            return 0, [f"Cannot open file: {e}"]

        ws = wb.active
        DATA_START = 2  # fallback if no header found

        # Auto-detect header row (look for a cell containing "account code")
        col_map = {}
        for ri, row in enumerate(ws.iter_rows(min_row=1, max_row=10, values_only=True), 1):
            row_strs = [str(v).strip().lower() if v else '' for v in row]
            if any('account code' in s or 'account_code' in s for s in row_strs):
                for ci, val in enumerate(row):
                    if val is None:
                        continue
                    s = str(val).strip().lower()
                    if 'account code' in s or 'account_code' in s:
                        col_map['code'] = ci
                    elif 'description' in s:
                        col_map['desc'] = ci
                    elif 'debit' in s or 'credit' in s or 'normal' in s:
                        col_map['nb'] = ci
                DATA_START = ri + 1
                break

        if 'code' not in col_map or 'desc' not in col_map:
            wb.close()
            return 0, [
                "Header row not found in the first 10 rows.\n"
                "Expected a row containing 'Account Code' and 'Account Description'."
            ]

        has_nb_column = 'nb' in col_map

        def _infer_normal_balance(code: str, desc: str) -> str:
            """Infer Debit/Credit from account code numeric prefix and description."""
            # Strip any alphabetic prefix (e.g. 'COA-', 'ABC-') to get digits
            digits = ''
            for ch in code:
                if ch.isdigit():
                    digits += ch
            if not digits:
                return 'Debit'

            first_digit = digits[0]

            # Description-based overrides (highest priority)
            desc_upper = desc.upper()
            if any(kw in desc_upper for kw in ('ACCUM DEP', 'ACCUM. DEP', 'ACCUMULATED DEP',
                                                'ACCUM AMORT', 'ACCUM. AMORT')):
                return 'Credit'
            if any(kw in desc_upper for kw in ('DRAWING', 'DRAWINGS')):
                return 'Debit'
            if 'DISCOUNT' in desc_upper and first_digit == '4':
                return 'Debit'   # Sales Discount / contra-revenue
            if 'RETURN' in desc_upper and first_digit == '4':
                return 'Debit'   # Sales Returns / contra-revenue
            if 'ALLOWANCE' in desc_upper and first_digit == '4':
                return 'Debit'   # Sales Allowances / contra-revenue

            # Numeric-prefix rules
            credit_prefixes = {'2', '3', '4', '7'}
            if first_digit in credit_prefixes:
                return 'Credit'
            return 'Debit'

        conn = self.get_connection()
        cursor = conn.cursor()
        imported = 0
        errors = []

        for rn, row in enumerate(ws.iter_rows(min_row=DATA_START, values_only=True), DATA_START):
            if all(v is None for v in row):
                continue
            try:
                code = str(row[col_map['code']]).strip() if row[col_map['code']] else ''
                desc = str(row[col_map['desc']]).strip() if row[col_map['desc']] else ''

                if not code or not desc:
                    continue
                # Skip obviously non-data rows (e.g. section headings with no code)
                if code.lower() in ('none', 'nan', 'account code'):
                    continue

                # Determine normal balance
                if has_nb_column and col_map['nb'] < len(row) and row[col_map['nb']]:
                    nb_raw = str(row[col_map['nb']]).strip().upper()
                    nb = 'Credit' if 'CREDIT' in nb_raw else 'Debit'
                else:
                    # No column — infer from code + description
                    nb = _infer_normal_balance(code, desc)

                cursor.execute(
                    'INSERT OR IGNORE INTO chart_of_accounts '
                    '(account_code, account_description, normal_balance) VALUES (?, ?, ?)',
                    (code, desc, nb)
                )
                imported += 1
            except Exception as e:
                errors.append(f"Row {rn}: {e}")

        conn.commit()
        wb.close()
        return imported, errors

    # ------------------------------------------------------------------
    # Chart of Accounts CRUD
    # ------------------------------------------------------------------

    def get_all_accounts(self) -> List[Dict[str, Any]]:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('SELECT * FROM chart_of_accounts ORDER BY account_code')
        return [dict(row) for row in cursor.fetchall()]

    def add_account(self, data: Dict[str, Any]) -> bool:
        try:
            conn = self.get_connection()
            cursor = conn.cursor()
            cursor.execute(
                'INSERT INTO chart_of_accounts '
                '(account_code, account_description, normal_balance) VALUES (?, ?, ?)',
                (data.get('account_code'), data.get('account_description'),
                 data.get('normal_balance', 'Debit'))
            )
            conn.commit()
            return True
        except:
            return False

    def update_account(self, account_id: int, data: Dict[str, Any]) -> bool:
        try:
            conn = self.get_connection()
            cursor = conn.cursor()
            cursor.execute(
                'UPDATE chart_of_accounts '
                'SET account_code = ?, account_description = ?, normal_balance = ? '
                'WHERE id = ?',
                (data.get('account_code'), data.get('account_description'),
                 data.get('normal_balance', 'Debit'), account_id)
            )
            conn.commit()
            return True
        except:
            return False

    def delete_account(self, account_id: int) -> bool:
        try:
            conn = self.get_connection()
            cursor = conn.cursor()
            cursor.execute('DELETE FROM chart_of_accounts WHERE id = ?', (account_id,))
            conn.commit()
            return True
        except:
            return False

    def import_coa_from_xlsx(self, xlsx_path: str) -> tuple:
        """Public wrapper for xlsx COA import (used by COAWidget)."""
        return self._import_coa_from_xlsx(xlsx_path)

    def export_coa_to_xlsx(self, xlsx_path: str) -> tuple:
        """Export the current COA to a formatted xlsx file. Returns (count, error)."""
        try:
            from openpyxl import Workbook
            from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
            from openpyxl.utils import get_column_letter
            from datetime import datetime as _dt
        except ImportError:
            return 0, "openpyxl not installed"

        accounts = self.get_all_accounts()
        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "Chart of Accounts"

            # Styles
            hdr_font  = Font(name='Arial', bold=True, color='FFFFFF', size=11)
            hdr_fill  = PatternFill('solid', start_color='2F5496')
            hdr_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell_font = Font(name='Arial', size=10)
            alt_fill  = PatternFill('solid', start_color='DCE6F1')
            thin      = Side(style='thin', color='B0B0B0')
            border    = Border(left=thin, right=thin, top=thin, bottom=thin)

            # Title rows
            ws.merge_cells('A2:C2')
            ws['A2'].value     = 'CHART OF ACCOUNTS'
            ws['A2'].font      = Font(name='Arial', bold=True, size=14)
            ws['A2'].alignment = Alignment(horizontal='left', vertical='center')
            ws.row_dimensions[2].height = 22

            ws.merge_cells('A3:C3')
            ws['A3'].value     = f'For the Year {_dt.now().year}'
            ws['A3'].font      = Font(name='Arial', italic=True, size=11)
            ws['A3'].alignment = Alignment(horizontal='left', vertical='center')

            # Header row (row 5)
            HEADER_ROW = 5
            headers = ['Account Code', 'Account Description', 'DEBIT/CREDIT']
            ws.row_dimensions[HEADER_ROW].height = 28
            for ci, hdr in enumerate(headers, 1):
                cell = ws.cell(row=HEADER_ROW, column=ci, value=hdr)
                cell.font = hdr_font; cell.fill = hdr_fill
                cell.alignment = hdr_align; cell.border = border

            # Data rows
            for ri, acct in enumerate(accounts):
                row_idx = 6 + ri
                ws.row_dimensions[row_idx].height = 18
                fill = alt_fill if ri % 2 == 0 else None
                vals = [
                    acct['account_code'],
                    acct['account_description'],
                    acct.get('normal_balance', 'Debit').upper(),
                ]
                for ci, val in enumerate(vals, 1):
                    cell = ws.cell(row=row_idx, column=ci, value=val)
                    cell.font   = cell_font
                    cell.border = border
                    cell.alignment = Alignment(
                        horizontal='center' if ci != 2 else 'left',
                        vertical='center'
                    )
                    if fill:
                        cell.fill = fill

            ws.column_dimensions['A'].width = 18
            ws.column_dimensions['B'].width = 45
            ws.column_dimensions['C'].width = 16
            ws.freeze_panes = 'A6'
            ws.auto_filter.ref = f'A{HEADER_ROW}:C{HEADER_ROW}'

            wb.save(xlsx_path)
            return len(accounts), ''
        except Exception as e:
            return 0, str(e)

    # ------------------------------------------------------------------
    # Alphalist
    # ------------------------------------------------------------------

    def get_all_alphalist(self, entry_type: str = None) -> List[Dict[str, Any]]:
        conn = self.get_connection()
        cursor = conn.cursor()
        if entry_type and entry_type != 'All List':
            cursor.execute(
                "SELECT * FROM alphalist WHERE entry_type = ? "
                "OR entry_type = 'Customer&Vendor' ORDER BY company_name, last_name",
                (entry_type,)
            )
        else:
            cursor.execute('SELECT * FROM alphalist ORDER BY company_name, last_name')
        return [dict(row) for row in cursor.fetchall()]

    def add_alphalist(self, data: Dict[str, Any]) -> bool:
        try:
            conn = self.get_connection()
            cursor = conn.cursor()
            cursor.execute('''
                INSERT INTO alphalist
                (tin, company_name, first_name, middle_name, last_name,
                 address1, address2, vat, entry_type)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                data.get('tin'), data.get('company_name'), data.get('first_name'),
                data.get('middle_name'), data.get('last_name'), data.get('address1'),
                data.get('address2'), data.get('vat', 'VAT Regular'),
                data.get('entry_type', 'Customer')
            ))
            conn.commit()
            return True
        except:
            return False

    def update_alphalist(self, entry_id: int, data: Dict[str, Any]) -> bool:
        try:
            conn = self.get_connection()
            cursor = conn.cursor()
            cursor.execute('''
                UPDATE alphalist
                SET tin=?, company_name=?, first_name=?, middle_name=?,
                    last_name=?, address1=?, address2=?, vat=?, entry_type=?
                WHERE id=?
            ''', (
                data.get('tin'), data.get('company_name'), data.get('first_name'),
                data.get('middle_name'), data.get('last_name'), data.get('address1'),
                data.get('address2'), data.get('vat', 'VAT Regular'),
                data.get('entry_type', 'Customer'), entry_id
            ))
            conn.commit()
            return True
        except:
            return False

    def delete_alphalist(self, entry_id: int) -> bool:
        try:
            conn = self.get_connection()
            cursor = conn.cursor()
            cursor.execute('DELETE FROM alphalist WHERE id = ?', (entry_id,))
            conn.commit()
            return True
        except:
            return False

    def get_alphalist_vat_for_name(self, name: str) -> str:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            SELECT vat FROM alphalist
            WHERE company_name = ?
               OR (trim(first_name || ' ' || last_name) = ?)
            LIMIT 1
        ''', (name, name))
        row = cursor.fetchone()
        return row['vat'] if row else 'VAT Regular'

    # ------------------------------------------------------------------
    # Sales Journal
    # ------------------------------------------------------------------

    def get_sales_journal(self, year: int = None) -> List[Dict[str, Any]]:
        conn = self.get_connection()
        cursor = conn.cursor()
        if year is None:
            cursor.execute('SELECT * FROM sales_journal ORDER BY date DESC')
        else:
            cursor.execute(
                "SELECT * FROM sales_journal WHERE strftime('%Y', date) = ? ORDER BY date DESC",
                (str(year),)
            )
        return [dict(row) for row in cursor.fetchall()]

    def add_sales_entry(self, data: Dict[str, Any]) -> bool:
        try:
            conn = self.get_connection()
            cursor = conn.cursor()
            cursor.execute('''
                INSERT INTO sales_journal
                (date, customer_name, reference_no, tin, net_amount, output_vat,
                 gross_amount, goods, services, particulars)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                data.get('date'), data.get('customer_name'), data.get('reference_no'),
                data.get('tin'), data.get('net_amount'), data.get('output_vat'),
                data.get('gross_amount'), data.get('goods', 0), data.get('services', 0),
                data.get('particulars')
            ))
            conn.commit()
            return True
        except:
            return False

    def update_sales_entry(self, entry_id: int, data: Dict[str, Any]) -> bool:
        try:
            conn = self.get_connection()
            cursor = conn.cursor()
            cursor.execute('''
                UPDATE sales_journal
                SET date=?, customer_name=?, reference_no=?, tin=?,
                    net_amount=?, output_vat=?, gross_amount=?,
                    goods=?, services=?, particulars=?
                WHERE id=?
            ''', (
                data.get('date'), data.get('customer_name'), data.get('reference_no'),
                data.get('tin'), data.get('net_amount'), data.get('output_vat'),
                data.get('gross_amount'), data.get('goods', 0), data.get('services', 0),
                data.get('particulars'), entry_id
            ))
            conn.commit()
            return True
        except:
            return False

    def delete_sales_entry(self, entry_id: int) -> bool:
        try:
            conn = self.get_connection()
            cursor = conn.cursor()
            cursor.execute('DELETE FROM sales_journal WHERE id = ?', (entry_id,))
            conn.commit()
            return True
        except:
            return False

    # ------------------------------------------------------------------
    # Purchase Journal
    # ------------------------------------------------------------------

    def get_purchase_journal(self, year: int = None) -> List[Dict[str, Any]]:
        conn = self.get_connection()
        cursor = conn.cursor()
        if year is None:
            cursor.execute('SELECT * FROM purchase_journal ORDER BY date DESC')
        else:
            cursor.execute(
                "SELECT * FROM purchase_journal WHERE strftime('%Y', date) = ? ORDER BY date DESC",
                (str(year),)
            )
        return [dict(row) for row in cursor.fetchall()]

    def add_purchase_entry(self, data: Dict[str, Any]) -> bool:
        try:
            conn = self.get_connection()
            cursor = conn.cursor()
            cursor.execute('''
                INSERT INTO purchase_journal
                (date, payee_name, reference_no, tin, branch_code, net_amount,
                 input_vat, gross_amount, account_description, account_code,
                 debit, credit, particulars)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                data.get('date'), data.get('payee_name'), data.get('reference_no'),
                data.get('tin'), data.get('branch_code'), data.get('net_amount'),
                data.get('input_vat'), data.get('gross_amount'),
                data.get('account_description'), data.get('account_code'),
                data.get('debit', 0), data.get('credit', 0), data.get('particulars')
            ))
            conn.commit()
            return True
        except:
            return False

    def update_purchase_entry(self, entry_id: int, data: Dict[str, Any]) -> bool:
        try:
            conn = self.get_connection()
            cursor = conn.cursor()
            cursor.execute('''
                UPDATE purchase_journal
                SET date=?, payee_name=?, reference_no=?, tin=?, branch_code=?,
                    net_amount=?, input_vat=?, gross_amount=?,
                    account_description=?, account_code=?, debit=?, credit=?, particulars=?
                WHERE id=?
            ''', (
                data.get('date'), data.get('payee_name'), data.get('reference_no'),
                data.get('tin'), data.get('branch_code'), data.get('net_amount'),
                data.get('input_vat'), data.get('gross_amount'),
                data.get('account_description'), data.get('account_code'),
                data.get('debit', 0), data.get('credit', 0),
                data.get('particulars'), entry_id
            ))
            conn.commit()
            return True
        except:
            return False

    def delete_purchase_entry(self, entry_id: int) -> bool:
        try:
            conn = self.get_connection()
            cursor = conn.cursor()
            cursor.execute('DELETE FROM purchase_journal WHERE id = ?', (entry_id,))
            conn.commit()
            return True
        except:
            return False

    # ------------------------------------------------------------------
    # Cash Disbursement Journal
    # ------------------------------------------------------------------

    def get_cash_disbursement_journal(self, year: int = None) -> List[Dict[str, Any]]:
        conn = self.get_connection()
        cursor = conn.cursor()
        if year is None:
            cursor.execute('SELECT * FROM cash_disbursement_journal ORDER BY date DESC')
        else:
            cursor.execute(
                "SELECT * FROM cash_disbursement_journal WHERE strftime('%Y', date) = ? ORDER BY date DESC",
                (str(year),)
            )
        return [dict(row) for row in cursor.fetchall()]

    def add_cash_disbursement_entry(self, data: Dict[str, Any]) -> bool:
        try:
            conn = self.get_connection()
            cursor = conn.cursor()
            cursor.execute('''
                INSERT INTO cash_disbursement_journal
                (date, reference_no, particulars, account_code, account_description, debit, credit)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            ''', (
                data.get('date'), data.get('reference_no'), data.get('particulars'),
                data.get('account_code'), data.get('account_description'),
                data.get('debit', 0), data.get('credit', 0)
            ))
            conn.commit()
            return True
        except Exception as e:
            print(f"Error: {e}")
            return False

    def update_cash_disbursement_entry(self, entry_id: int, data: Dict[str, Any]) -> bool:
        try:
            conn = self.get_connection()
            cursor = conn.cursor()
            cursor.execute('''
                UPDATE cash_disbursement_journal
                SET date=?, reference_no=?, particulars=?, account_description=?,
                    account_code=?, debit=?, credit=?
                WHERE id=?
            ''', (
                data.get('date'), data.get('reference_no'), data.get('particulars'),
                data.get('account_description'), data.get('account_code'),
                data.get('debit', 0), data.get('credit', 0), entry_id
            ))
            conn.commit()
            return True
        except:
            return False

    def delete_cash_disbursement_entry(self, entry_id: int) -> bool:
        try:
            conn = self.get_connection()
            cursor = conn.cursor()
            cursor.execute('DELETE FROM cash_disbursement_journal WHERE id = ?', (entry_id,))
            conn.commit()
            return True
        except:
            return False

    # ------------------------------------------------------------------
    # Cash Receipts Journal
    # ------------------------------------------------------------------

    def get_cash_receipts_journal(self, year: int = None) -> List[Dict[str, Any]]:
        conn = self.get_connection()
        cursor = conn.cursor()
        if year is None:
            cursor.execute('SELECT * FROM cash_receipts_journal ORDER BY date DESC')
        else:
            cursor.execute(
                "SELECT * FROM cash_receipts_journal WHERE strftime('%Y', date) = ? ORDER BY date DESC",
                (str(year),)
            )
        return [dict(row) for row in cursor.fetchall()]

    def add_cash_receipts_entry(self, data: Dict[str, Any]) -> bool:
        try:
            conn = self.get_connection()
            cursor = conn.cursor()
            cursor.execute('''
                INSERT INTO cash_receipts_journal
                (date, reference_no, particulars, account_code, account_description, debit, credit)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            ''', (
                data.get('date'), data.get('reference_no'), data.get('particulars'),
                data.get('account_code'), data.get('account_description'),
                data.get('debit', 0), data.get('credit', 0)
            ))
            conn.commit()
            return True
        except:
            return False

    def update_cash_receipts_entry(self, entry_id: int, data: Dict[str, Any]) -> bool:
        try:
            conn = self.get_connection()
            cursor = conn.cursor()
            cursor.execute('''
                UPDATE cash_receipts_journal
                SET date=?, reference_no=?, particulars=?, account_description=?,
                    account_code=?, debit=?, credit=?
                WHERE id=?
            ''', (
                data.get('date'), data.get('reference_no'), data.get('particulars'),
                data.get('account_description'), data.get('account_code'),
                data.get('debit', 0), data.get('credit', 0), entry_id
            ))
            conn.commit()
            return True
        except:
            return False

    def delete_cash_receipts_entry(self, entry_id: int) -> bool:
        try:
            conn = self.get_connection()
            cursor = conn.cursor()
            cursor.execute('DELETE FROM cash_receipts_journal WHERE id = ?', (entry_id,))
            conn.commit()
            return True
        except:
            return False

    # ------------------------------------------------------------------
    # General Journal
    # ------------------------------------------------------------------

    def get_general_journal(self, year: int = None) -> List[Dict[str, Any]]:
        conn = self.get_connection()
        cursor = conn.cursor()
        if year is None:
            cursor.execute('SELECT * FROM general_journal ORDER BY date DESC')
        else:
            cursor.execute(
                "SELECT * FROM general_journal WHERE strftime('%Y', date) = ? ORDER BY date DESC",
                (str(year),)
            )
        return [dict(row) for row in cursor.fetchall()]

    def add_general_journal_entry(self, data: Dict[str, Any]) -> bool:
        try:
            conn = self.get_connection()
            cursor = conn.cursor()
            cursor.execute('''
                INSERT INTO general_journal
                (date, reference_no, particulars, account_code, account_description, debit, credit)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            ''', (
                data.get('date'), data.get('reference_no'), data.get('particulars'),
                data.get('account_code'), data.get('account_description'),
                data.get('debit', 0), data.get('credit', 0)
            ))
            conn.commit()
            return True
        except:
            return False

    def update_general_journal_entry(self, entry_id: int, data: Dict[str, Any]) -> bool:
        try:
            conn = self.get_connection()
            cursor = conn.cursor()
            cursor.execute('''
                UPDATE general_journal
                SET date=?, reference_no=?, particulars=?, account_description=?,
                    account_code=?, debit=?, credit=?
                WHERE id=?
            ''', (
                data.get('date'), data.get('reference_no'), data.get('particulars'),
                data.get('account_description'), data.get('account_code'),
                data.get('debit', 0), data.get('credit', 0), entry_id
            ))
            conn.commit()
            return True
        except:
            return False

    def delete_general_journal_entry(self, entry_id: int) -> bool:
        try:
            conn = self.get_connection()
            cursor = conn.cursor()
            cursor.execute('DELETE FROM general_journal WHERE id = ?', (entry_id,))
            conn.commit()
            return True
        except:
            return False

    # ------------------------------------------------------------------
    # General Ledger
    # ------------------------------------------------------------------

    def get_general_ledger(self, account_code: str,
                           date_from: str = None, date_to: str = None) -> Dict[str, Any]:
        """
        Aggregate all journal sources for a given account_code.

        The account_code is looked up dynamically against the COA — this means
        the GL works regardless of which prefix (COA-, ABC-, custom) is in use,
        as long as the code constants at the top of this file match the actual
        COA loaded in the database.
        """
        conn = self.get_connection()
        cursor = conn.cursor()

        entries = []

        def _to_sql_date(mmddyyyy: str) -> str:
            try:
                from datetime import datetime
                return datetime.strptime(mmddyyyy, "%m/%d/%Y").strftime("%Y-%m-%d")
            except:
                return mmddyyyy

        if date_from and date_to:
            sql_from = _to_sql_date(date_from)
            sql_to   = _to_sql_date(date_to)
            date_filter = (
                "AND strftime('%Y-%m-%d', "
                "substr(date,7,4)||'-'||substr(date,1,2)||'-'||substr(date,4,2)) "
                "BETWEEN ? AND ?"
            )
            date_params = (sql_from, sql_to)
        elif date_to:
            sql_to = _to_sql_date(date_to)
            date_filter = (
                "AND strftime('%Y-%m-%d', "
                "substr(date,7,4)||'-'||substr(date,1,2)||'-'||substr(date,4,2)) <= ?"
            )
            date_params = (sql_to,)
        elif date_from:
            sql_from = _to_sql_date(date_from)
            date_filter = (
                "AND strftime('%Y-%m-%d', "
                "substr(date,7,4)||'-'||substr(date,1,2)||'-'||substr(date,4,2)) >= ?"
            )
            date_params = (sql_from,)
        else:
            date_filter = ""
            date_params = ()

        # ── Sales Journal postings ────────────────────────────────────────
        if _numeric_suffix(account_code) == _AR_NUM:
            cursor.execute(f'''
                SELECT date, reference_no, customer_name as particulars,
                       gross_amount as debit, 0 as credit, 'SJ' as source
                FROM sales_journal WHERE 1=1 {date_filter} ORDER BY date
            ''', date_params)
            entries.extend(cursor.fetchall())

        if _numeric_suffix(account_code) == _SALES_GOODS_NUM:
            cursor.execute(f'''
                SELECT date, reference_no, customer_name as particulars,
                       0 as debit, goods as credit, 'SJ' as source
                FROM sales_journal WHERE goods > 0 {date_filter} ORDER BY date
            ''', date_params)
            entries.extend(cursor.fetchall())

        if _numeric_suffix(account_code) == _SALES_SERVICE_NUM:
            cursor.execute(f'''
                SELECT date, reference_no, customer_name as particulars,
                       0 as debit, services as credit, 'SJ' as source
                FROM sales_journal WHERE services > 0 {date_filter} ORDER BY date
            ''', date_params)
            entries.extend(cursor.fetchall())

        if _numeric_suffix(account_code) == _OUTPUT_VAT_NUM:
            cursor.execute(f'''
                SELECT date, reference_no, customer_name as particulars,
                       0 as debit, output_vat as credit, 'SJ' as source
                FROM sales_journal WHERE output_vat > 0 {date_filter} ORDER BY date
            ''', date_params)
            entries.extend(cursor.fetchall())

        # ── Purchase Journal postings ─────────────────────────────────────
        if _numeric_suffix(account_code) == _AP_NUM:
            cursor.execute(f'''
                SELECT date, reference_no, payee_name as particulars,
                       0 as debit, gross_amount as credit, 'PJ' as source
                FROM purchase_journal WHERE 1=1 {date_filter} ORDER BY date
            ''', date_params)
            entries.extend(cursor.fetchall())

        if _numeric_suffix(account_code) == _INPUT_VAT_NUM:
            cursor.execute(f'''
                SELECT date, reference_no, payee_name as particulars,
                       input_vat as debit, 0 as credit, 'PJ' as source
                FROM purchase_journal WHERE input_vat > 0 {date_filter} ORDER BY date
            ''', date_params)
            entries.extend(cursor.fetchall())

        # PJ user-selected debit account
        cursor.execute(f'''
            SELECT date, reference_no, payee_name as particulars,
                   debit, 0 as credit, 'PJ' as source
            FROM purchase_journal
            WHERE account_code = ? AND debit > 0 {date_filter} ORDER BY date
        ''', (account_code,) + date_params)
        entries.extend(cursor.fetchall())

        # ── Other journals ────────────────────────────────────────────────
        for table, tag in [
            ('cash_disbursement_journal', 'CDJ'),
            ('cash_receipts_journal',     'CRJ'),
            ('general_journal',           'GJ'),
        ]:
            cursor.execute(f'''
                SELECT date, reference_no, particulars, debit, credit, '{tag}' as source
                FROM {table}
                WHERE account_code = ? {date_filter} ORDER BY date
            ''', (account_code,) + date_params)
            entries.extend(cursor.fetchall())

        result = sorted([dict(row) for row in entries], key=lambda x: x.get('date', ''))

        # Fetch normal_balance from the live COA (works for any prefix)
        cursor.execute(
            "SELECT normal_balance FROM chart_of_accounts WHERE account_code = ?",
            (account_code,)
        )
        row = cursor.fetchone()
        normal_balance = row['normal_balance'] if row else 'Debit'

        return {
            'account_code':   account_code,
            'normal_balance': normal_balance,
            'entries':        result,
        }

    # ------------------------------------------------------------------
    # Trial Balance
    # ------------------------------------------------------------------

    def get_trial_balance(self, date_from: str = None, date_to: str = None) -> List[Dict[str, Any]]:
        conn = self.get_connection()
        cursor = conn.cursor()

        cursor.execute(
            'SELECT account_code, account_description, normal_balance '
            'FROM chart_of_accounts ORDER BY account_code'
        )
        accounts = cursor.fetchall()

        def _to_sql_date(mmddyyyy: str) -> str:
            try:
                from datetime import datetime
                return datetime.strptime(mmddyyyy, "%m/%d/%Y").strftime("%Y-%m-%d")
            except:
                return mmddyyyy

        if date_from and date_to:
            sql_from = _to_sql_date(date_from)
            sql_to   = _to_sql_date(date_to)
            date_filter = (
                "AND strftime('%Y-%m-%d', "
                "substr(date,7,4)||'-'||substr(date,1,2)||'-'||substr(date,4,2)) "
                "BETWEEN ? AND ?"
            )
            date_params = (sql_from, sql_to)
        elif date_to:
            sql_to = _to_sql_date(date_to)
            date_filter = (
                "AND strftime('%Y-%m-%d', "
                "substr(date,7,4)||'-'||substr(date,1,2)||'-'||substr(date,4,2)) <= ?"
            )
            date_params = (sql_to,)
        elif date_from:
            sql_from = _to_sql_date(date_from)
            date_filter = (
                "AND strftime('%Y-%m-%d', "
                "substr(date,7,4)||'-'||substr(date,1,2)||'-'||substr(date,4,2)) >= ?"
            )
            date_params = (sql_from,)
        else:
            date_filter = ""
            date_params = ()

        trial_balance = []

        for account in accounts:
            account_code = account['account_code']
            account_desc = account['account_description']
            normal_bal   = account['normal_balance'] or 'Debit'
            total_debit  = 0
            total_credit = 0

            if _numeric_suffix(account_code) == _AR_NUM:
                cursor.execute(f"SELECT COALESCE(SUM(gross_amount),0) FROM sales_journal WHERE 1=1 {date_filter}", date_params)
                total_debit += cursor.fetchone()[0]
            if _numeric_suffix(account_code) == _SALES_GOODS_NUM:
                cursor.execute(f"SELECT COALESCE(SUM(goods),0) FROM sales_journal WHERE 1=1 {date_filter}", date_params)
                total_credit += cursor.fetchone()[0]
            if _numeric_suffix(account_code) == _SALES_SERVICE_NUM:
                cursor.execute(f"SELECT COALESCE(SUM(services),0) FROM sales_journal WHERE 1=1 {date_filter}", date_params)
                total_credit += cursor.fetchone()[0]
            if _numeric_suffix(account_code) == _OUTPUT_VAT_NUM:
                cursor.execute(f"SELECT COALESCE(SUM(output_vat),0) FROM sales_journal WHERE 1=1 {date_filter}", date_params)
                total_credit += cursor.fetchone()[0]

            if _numeric_suffix(account_code) == _AP_NUM:
                cursor.execute(f"SELECT COALESCE(SUM(gross_amount),0) FROM purchase_journal WHERE 1=1 {date_filter}", date_params)
                total_credit += cursor.fetchone()[0]
            if _numeric_suffix(account_code) == _INPUT_VAT_NUM:
                cursor.execute(f"SELECT COALESCE(SUM(input_vat),0) FROM purchase_journal WHERE 1=1 {date_filter}", date_params)
                total_debit += cursor.fetchone()[0]

            cursor.execute(f"SELECT COALESCE(SUM(debit),0) FROM purchase_journal WHERE account_code=? AND debit>0 {date_filter}", (account_code,)+date_params)
            total_debit += cursor.fetchone()[0]

            for tbl in ('cash_disbursement_journal', 'cash_receipts_journal', 'general_journal'):
                cursor.execute(f"SELECT COALESCE(SUM(debit),0), COALESCE(SUM(credit),0) FROM {tbl} WHERE account_code=? {date_filter}", (account_code,)+date_params)
                row = cursor.fetchone()
                total_debit  += row[0]
                total_credit += row[1]

            balance = total_debit - total_credit
            if balance != 0:
                trial_balance.append({
                    'account_code':        account_code,
                    'account_description': account_desc,
                    'normal_balance':      normal_bal,
                    'amount':              balance,
                })

        return trial_balance

    # ------------------------------------------------------------------

    def close(self):
        if self.connection:
            self.connection.close()
            self.connection = None