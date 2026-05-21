import json
import re
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Moniepoint MIS Automation", layout="wide")

APP_DIR = Path(__file__).parent
DATA_DIR = APP_DIR / "data"
DATA_DIR.mkdir(exist_ok=True)
MAPPING_FILE = DATA_DIR / "mapping_master.json"
AUDIT_FILE = DATA_DIR / "mapping_audit_trail.csv"


def load_mapping_master():
    if MAPPING_FILE.exists():
        return json.loads(MAPPING_FILE.read_text())
    seed = {
        "SALES": {"account_type": "Income", "head": "Revenue from Operations"},
        "SALARIES & WAGES": {"account_type": "Expense", "head": "Salaries & wages"},
        "SALARIES & WAGES- MPUK": {"account_type": "Expense", "head": "Salaries & wages"},
        "SALARIES & WAGES - MPUK": {"account_type": "Expense", "head": "Salaries & wages"},
        "SALARIES & WAGES - MPN": {"account_type": "Expense", "head": "Salaries & wages"},
        "SALARIES & WAGES - MPIN": {"account_type": "Expense", "head": "Salaries & wages"},
        "PERFORMANCE BONUSES AND INCENTIVES": {"account_type": "Expense", "head": "Performance bonuses"},
        "PERFORMANCE BONUS": {"account_type": "Expense", "head": "Performance bonuses"},
        "PERFORMANCE BONUSES": {"account_type": "Expense", "head": "Performance bonuses"},
        "OTHER EMPLOYEE COSTS": {"account_type": "Expense", "head": "Other employee costs"},
        "LEAVE ENCASHMENT": {"account_type": "Expense", "head": "Other employee costs", "sub_head": "Leave Encashment"},
        "PF ADMIN CHARGES": {"account_type": "Expense", "head": "Other employee costs", "sub_head": "PF Admin Charges"},
        "EMPLOYER PF CONTRIBUTION": {"account_type": "Expense", "head": "Other employee costs", "sub_head": "Employers Contribution to Social Security Funds"},
        "EMPLOYEES PF CONTRIBUTION": {"account_type": "Expense", "head": "Other employee costs", "sub_head": "Employees PF Contribution"},
        "EMPLOYEE BENEFITS (PRIVATE MEDICAL INSURANCE)": {"account_type": "Expense", "head": "Insurance"},
        "INSURANCE": {"account_type": "Expense", "head": "Insurance"},
        "ACCOUNTING AND TAX CONSULTING FEES": {"account_type": "Expense", "head": "Legal and professional costs", "sub_head": "Accounting & Bookkeeping"},
        "LEGAL AND PROFESSIONAL COSTS": {"account_type": "Expense", "head": "Legal and professional costs"},
        "DEPRECIATION": {"account_type": "Expense", "head": "Depreciation"},
        "RENT AND RATES": {"account_type": "Expense", "head": "Rent and other office expenses"},
        "RENT AND OTHER OFFICE EXPENSES": {"account_type": "Expense", "head": "Rent and other office expenses"},
        "SOFTWARE SUBSCRIPTIONS - G&A": {"account_type": "Expense", "head": "Other operating expenses", "sub_head": "Software Subscriptions"},
        "SOFTWARE SUBSCRIPTIONS": {"account_type": "Expense", "head": "Other operating expenses", "sub_head": "Software Subscriptions"},
        "AUDIT FEES": {"account_type": "Expense", "head": "Other operating expenses", "sub_head": "Audit Fees"},
        "BANK TRANSACTION FEES AND SERVICE CHARGES": {"account_type": "Expense", "head": "Other operating expenses", "sub_head": "Bank Charges"},
        "STATUTORY FEES": {"account_type": "Expense", "head": "Other operating expenses", "sub_head": "Statutory Fees"},
        "SHIPPING & POSTAGE": {"account_type": "Expense", "head": "Other operating expenses", "sub_head": "Shipping & Postage"},
        "SHOPS AND ESTABLISHMENT": {"account_type": "Expense", "head": "Other operating expenses", "sub_head": "Shops and Establishment"},
        "EXCHANGE GAIN OR LOSS": {"account_type": "Expense", "head": "Unrealised FX gains or losses"},
        "UNREALISED FX GAINS OR LOSSES": {"account_type": "Expense", "head": "Unrealised FX gains or losses"},
        "CORPORATE INCOME TAX": {"account_type": "Expense", "head": "Income tax"},
        "ACCOUNTS PAYABLE": {"account_type": "Liability", "head": "Trade Payables"},
        "TRADE PAYABLES": {"account_type": "Liability", "head": "Trade Payables"},
        "ACCRUED LEGAL AND PROFESSIONAL EXPENSES": {"account_type": "Liability", "head": "Accrued Legal and Professional Expenses"},
        "COMPUTERS AND LAPTOPS": {"account_type": "Asset", "head": "Property, Plant & Equipment and Intangible Assets", "sub_head": "COMPUTERS AND LAPTOPS"},
        "PREPAID EXPENSES": {"account_type": "Asset", "head": "Prepaid Expenses"},
        "ACCOUNTS RECEIVABLE": {"account_type": "Asset", "head": "Trade Receivables"},
        "RAZORPAY WALLET": {"account_type": "Asset", "head": "Other Current Assets", "sub_head": "Razorpay Wallet"},
        "TDS RECEIVABLE": {"account_type": "Asset", "head": "Other Current Assets", "sub_head": "TDS Receivable"},
        "ADVANCE TAX": {"account_type": "Asset", "head": "Other Current Assets", "sub_head": "Advance Tax"},
        "INTERCOMPANY RECEIVABLES - MONIEPOINT UK": {"account_type": "Asset", "head": "Other Assets", "sub_head": "Intercompany Receivables - Moniepoint UK"},
        "INTERCOMPANY RECEIVABLES - MONIEPOINT INC": {"account_type": "Asset", "head": "Other Assets", "sub_head": "Intercompany Receivables - Moniepoint Inc"},
        "INTERCOMPANY RECEIVABLES - MONIEPOINT MFB": {"account_type": "Asset", "head": "Other Assets", "sub_head": "Intercompany Receivables - Moniepoint MFB"},
        "INTERCOMPANY PAYABLES - MONIEPOINT UK": {"account_type": "Liability", "head": "Other Current Liabilities", "sub_head": "Intercompany Payables - Moniepoint UK"},
        "INTERCOMPANY PAYABLES - MONIEPOINT INC": {"account_type": "Liability", "head": "Other Current Liabilities", "sub_head": "Intercompany Payables - Moniepoint Inc"},
        "INTERCOMPANY PAYABLES - MONIEPOINT GB": {"account_type": "Liability", "head": "Other Current Liabilities", "sub_head": "Intercompany Payables - Moniepoint GB"},
        "ACCRUED PAYROLL EXPENSE": {"account_type": "Liability", "head": "Accrued expenses and other payables", "sub_head": "Accrued Payroll Expense"},
        "ACCUMULATED DEPRECIATION - COMPUTERS": {"account_type": "Asset", "head": "Property, Plant & Equipment and Intangible Assets", "sub_head": "Accumulated Depreciation - Computers"},
        "DEFERRED TAX ASSET": {"account_type": "Asset", "head": "Other Current Assets", "sub_head": "Deferred Tax Asset"},
        "CORPORATE INCOME TAX PAYABLE": {"account_type": "Liability", "head": "Short Term Provisions", "sub_head": "Corporate Income Tax Payable"},
        "PROVISION FOR PROFESSIONAL TAX": {"account_type": "Liability", "head": "Short Term Provisions", "sub_head": "Provision for Professional Tax"},
    }
    MAPPING_FILE.write_text(json.dumps(seed, indent=2))
    return seed


def save_mapping_master(mapping):
    MAPPING_FILE.write_text(json.dumps(mapping, indent=2))


def append_audit_trail(rows):
    if not rows:
        return
    df = pd.DataFrame(rows)
    if AUDIT_FILE.exists():
        old = pd.read_csv(AUDIT_FILE)
        df = pd.concat([old, df], ignore_index=True)
    df.to_csv(AUDIT_FILE, index=False)


def normalize_name(v):
    return re.sub(r"\s+", " ", str(v).strip().upper())


def read_tb(upload):
    xls = pd.ExcelFile(upload)
    sheets = xls.sheet_names
    st.write("Detected TB sheets:", sheets)
    if len(sheets) == 1:
        return pd.read_excel(xls, sheet_name=sheets[0]), pd.DataFrame()
    bs_sheet = next((s for s in sheets if any(k in s.lower() for k in ["bs", "balance"])), sheets[0])
    pl_sheet = next((s for s in sheets if any(k in s.lower() for k in ["p&l", "pl", "profit"])), sheets[1] if len(sheets) > 1 else sheets[0])
    return pd.read_excel(xls, sheet_name=bs_sheet), pd.read_excel(xls, sheet_name=pl_sheet)


def read_salary(upload):
    xls = pd.ExcelFile(upload)
    return pd.read_excel(xls, sheet_name=xls.sheet_names[0])


def read_reference(upload):
    xls = pd.ExcelFile(upload)
    return pd.read_excel(xls, sheet_name=xls.sheet_names[0])


def extract_rate(ref_df):
    cols = {str(c).strip().lower(): c for c in ref_df.columns}
    pair_col = next((cols[k] for k in cols if "currency" in k or "pair" in k), ref_df.columns[0])
    rate_col = next((cols[k] for k in cols if "rate" in k), ref_df.columns[min(1, len(ref_df.columns)-1)])
    temp = ref_df.copy()
    temp = temp[temp[pair_col].astype(str).str.contains("USD", case=False, na=False)]
    return round(pd.to_numeric(temp[rate_col], errors="coerce").dropna().mean(), 4)


def salary_summary(df):
    cols = [str(c).strip() for c in df.columns]
    canon = {c.upper(): c for c in cols}
    gross_col = canon.get("GROSS SALARY") or canon.get("GROSS PAY") or canon.get("GROSS")
    if gross_col is None:
        raise ValueError("Gross Salary column not found")

    def n(c):
        if c in df.columns:
            return pd.to_numeric(df[c], errors="coerce").fillna(0)
        return pd.Series([0] * len(df), index=df.index)

    gross = n(gross_col)
    pfee = n(canon.get("PFEE", "PFEE"))
    pfer = n(canon.get("PFER", "PFER"))
    perf = n(canon.get("PERFORMANCE BONUS", "Performance Bonus"))
    leave = n(canon.get("LEAVE ENCASHMENT", "Leave Encashment"))
    basic = n(canon.get("BASIC SALARY", "Basic Salary"))
    da = n(canon.get("DA", "DA"))

    return {
        "rows": int(len(df)),
        "gross_total": float(gross.sum()),
        "pf_ee_total": float(pfee.sum()),
        "pf_er_total": float(pfer.sum()),
        "performance_bonus_total": float(perf.sum()),
        "leave_encashment_total": float(leave.sum()),
        "basic_da_total": float((basic + da).sum()),
        "pf_admin_0_5pct_basic_da": float(round(((basic + da).sum()) * 0.005, 0)),
    }


def ledger_amount_frame(df):
    work = df.copy()
    work.columns = [str(c).strip() for c in work.columns]
    ledger_col = next((c for c in work.columns if "ledger" in c.lower() or "account" in c.lower()), work.columns[0])
    amount_col = next((c for c in work.columns if "net" in c.lower() or "amount" in c.lower() or "debit" in c.lower()), work.columns[-1])
    out = work[[ledger_col, amount_col]].copy()
    out.columns = ["ledger", "amount"]
    out["ledger"] = out["ledger"].astype(str).map(normalize_name)
    out["amount"] = pd.to_numeric(out["amount"], errors="coerce").fillna(0)
    return out.groupby("ledger", as_index=False)["amount"].sum()


def match_mapping(ledger, mapping):
    key = normalize_name(ledger)
    if key in mapping:
        return key, mapping[key]
    for k, v in mapping.items():
        if k in key or key in k:
            return k, v
    return None, None


def classify(df, mapping, source_name, audit_rows):
    out = df.copy()
    out["mapping_key"] = ""
    out["account_type"] = ""
    out["head"] = ""
    out["sub_head"] = ""
    out["status"] = "Mapped"
    for i, row in out.iterrows():
        key, info = match_mapping(row["ledger"], mapping)
        if info:
            out.at[i, "mapping_key"] = key
            out.at[i, "account_type"] = info.get("account_type", "")
            out.at[i, "head"] = info.get("head", "")
            out.at[i, "sub_head"] = info.get("sub_head", "")
        else:
            out.at[i, "status"] = "Unmapped"
            audit_rows.append({"source": source_name, "ledger": row["ledger"], "amount": row["amount"], "status": "Unmapped"})
    return out


def make_output_file(bs_mapped, pl_mapped, salary_summary_dict, rate):
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        bs_mapped.to_excel(writer, index=False, sheet_name="BS Mapped")
        pl_mapped.to_excel(writer, index=False, sheet_name="P&L Mapped")
        pd.DataFrame([salary_summary_dict]).to_excel(writer, index=False, sheet_name="Salary Summary")
        pd.DataFrame([{"USD_rate": rate}]).to_excel(writer, index=False, sheet_name="Reference Rate")
    buffer.seek(0)
    return buffer


def main():
    st.title("Moniepoint MIS Automation")
    mapping = load_mapping_master()
    audit_rows = []

    tb_file = st.file_uploader("Upload TB file", type=["xlsx"])
    salary_file = st.file_uploader("Upload Salary Register", type=["xlsx"])
    ref_file = st.file_uploader("Upload Reference Rates", type=["xlsx"])

    if tb_file and salary_file and ref_file:
        tb_bs, tb_pl = read_tb(tb_file)
        salary_df = read_salary(salary_file)
        ref_df = read_reference(ref_file)

        st.subheader("Sheet preview")
        st.write("TB BS rows:", len(tb_bs), "TB P&L rows:", len(tb_pl))
        st.write("Salary rows:", len(salary_df), "Reference rows:", len(ref_df))

        rate = extract_rate(ref_df)
        summary = salary_summary(salary_df)

        st.subheader("Salary summary")
        st.json(summary)
        st.subheader("USD rate")
        st.write(rate)

        bs_mapped = pd.DataFrame()
        pl_mapped = pd.DataFrame()
        if not tb_bs.empty:
            bs = ledger_amount_frame(tb_bs)
            bs_mapped = classify(bs, mapping, "BS", audit_rows)
            st.subheader("BS mapped")
            st.dataframe(bs_mapped, use_container_width=True)

        if not tb_pl.empty:
            pl = ledger_amount_frame(tb_pl)
            pl_mapped = classify(pl, mapping, "P&L", audit_rows)
            st.subheader("P&L mapped")
            st.dataframe(pl_mapped, use_container_width=True)

        append_audit_trail(audit_rows)
        save_mapping_master(mapping)

        output_file = make_output_file(bs_mapped, pl_mapped, summary, rate)
        st.download_button(
            label="Download output Excel",
            data=output_file,
            file_name="moniepoint_mis_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        if audit_rows:
            st.warning(f"Unmapped ledgers found: {len(audit_rows)}")
        else:
            st.success("All ledgers mapped.")
    else:
        st.info("Upload the 3 files to continue.")

if __name__ == "__main__":
    main()
