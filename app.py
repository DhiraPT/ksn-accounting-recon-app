import streamlit as st
import pandas as pd
import pdfplumber
import re
from datetime import datetime, timedelta

# ==========================================
# 1. BANKING-GRADE PARSERS
# ==========================================

@st.cache_data
def extract_text_from_pdf(uploaded_file):
    full_text = ""
    try:
        with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text: full_text += text + "\n"
    except Exception as e:
        st.error(f"Error reading PDF: {e}")
        return None
    return full_text

def parse_amount_str(raw_amt, type_grp):
    clean_amt = raw_amt.replace(',', '')
    txn_type = type_grp if type_grp else "CR"
    try:
        return float(clean_amt), txn_type
    except ValueError:
        return 0.0, txn_type

@st.cache_data
def parse_bank_statement(text_content):
    if not text_content: return pd.DataFrame()
    transactions = []
    lines = text_content.split('\n')
    current_txn = {}
    buffer_desc = []
    
    # 1. Default to current year as a fallback
    found_year = str(datetime.now().year)

    # 2. Try to find the year in the text using Regex
    # Looks for "PERIODE", ignores case, matches any characters until it finds 4 digits
    # Example match: "PERIODE : AGUSTUS 2025" -> captures "2025"
    year_match = re.search(r"PERIODE.*?(\d{4})", text_content, re.IGNORECASE)

    if year_match:
        found_year = year_match.group(1)

    DEFAULT_YEAR = found_year
    date_pattern = re.compile(r"^(\d{2}/\d{2})")
    # Matches amounts like 50,000.00 or 1,234.56
    amount_pattern = re.compile(r"([\d,]+\.\d{2})\s*(DB|CR)?")

    for line in lines:
        line = line.strip()
        if not line: continue
        date_match = date_pattern.match(line)
        if date_match:
            if current_txn:
                desc_text = " ".join(buffer_desc).strip()
                if "SALDO" not in desc_text.upper(): 
                    current_txn['description'] = desc_text
                    transactions.append(current_txn)
                current_txn = {}
                buffer_desc = []
            date_str = date_match.group(1)
            current_txn['date'] = f"{date_str}/{DEFAULT_YEAR}"
            amount_match = amount_pattern.search(line)
            if amount_match:
                amt, type_str = parse_amount_str(amount_match.group(1), amount_match.group(2))
                current_txn['amount'] = amt
                current_txn['type'] = type_str
                clean_line = line.replace(date_str, '').replace(amount_match.group(0), '').strip()
                buffer_desc.append(clean_line)
            else:
                buffer_desc.append(line[len(date_str):].strip())
        elif current_txn and 'amount' not in current_txn:
            amount_match = amount_pattern.search(line)
            if amount_match:
                amt, type_str = parse_amount_str(amount_match.group(1), amount_match.group(2))
                current_txn['amount'] = amt
                current_txn['type'] = type_str
            else:
                buffer_desc.append(line)
        else:
            buffer_desc.append(line)
            
    if current_txn:
        desc_text = " ".join(buffer_desc).strip()
        if "SALDO" not in desc_text.upper():
            current_txn['description'] = desc_text
            transactions.append(current_txn)

    df = pd.DataFrame(transactions)
    if not df.empty:
        df['id'] = df.index.astype(str) + "_bank"
        df['matched'] = False
        df['note'] = ""
    return df

def clean_excel_date(val):
    if pd.isna(val): return ""
    if isinstance(val, (int, float)):
        try:
            return (datetime(1899, 12, 30) + timedelta(days=val)).strftime("%d/%m/%Y")
        except: return str(val)
    if isinstance(val, pd.Timestamp): return val.strftime("%d/%m/%Y")
    val_str = str(val).strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%Y%m%d"):
        try: return datetime.strptime(val_str, fmt).strftime("%d/%m/%Y")
        except ValueError: continue
    return val_str

@st.cache_data
def parse_excel_ledger(uploaded_file):
    try:
        # READ FIRST SHEET ONLY (Default behavior)
        # Optimization: Read only first 20 rows to find header
        raw_df = pd.read_excel(uploaded_file, header=None, nrows=20)
        
        header_row_idx = 0
        found = False
        for i, row in raw_df.iterrows():
            row_str = " ".join([str(x) for x in row.values])
            if "Description" in row_str and ("Date" in row_str or "Tgl" in row_str):
                header_row_idx = i
                found = True
                break
        
        uploaded_file.seek(0)
        if found:
            df = pd.read_excel(uploaded_file, header=header_row_idx)
            start_row = header_row_idx + 2 
        else:
            df = pd.read_excel(uploaded_file)
            start_row = 2

        col_map = {
            'No': 'ref_code', 'No ': 'ref_code', 'No.': 'ref_code',
            'Date(NT)': 'date', 'Tgl Nota': 'date', 'Date': 'date',
            'Description': 'description',
            'DB': 'debit', 'CR': 'credit',
            'IN': 'debit', 'OUT': 'credit'
        }
        df = df.rename(columns=col_map)
        df['excel_row'] = range(start_row, start_row + len(df))

        # Ensure columns exist
        if 'debit' not in df.columns: df['debit'] = 0
        if 'credit' not in df.columns: df['credit'] = 0
        if 'description' not in df.columns: df['description'] = "No Desc"
        if 'date' not in df.columns: df['date'] = ""
        if 'ref_code' not in df.columns: df['ref_code'] = ""
        
        df['date'] = df['date'].apply(clean_excel_date)
        df['debit'] = pd.to_numeric(df['debit'], errors='coerce').fillna(0)
        df['credit'] = pd.to_numeric(df['credit'], errors='coerce').fillna(0)
        df['amount'] = df['debit'] + df['credit']
        
        # Filter valid rows
        df = df[df['amount'] > 0.01].copy()

        df['id'] = df.index.astype(str) + "_ledger"
        df['matched'] = False
        df['note'] = ""

        return df[['id', 'excel_row', 'date', 'ref_code', 'description', 'amount', 'matched', 'note']]

    except Exception as e:
        st.error(f"Error parsing Excel: {e}")
        return pd.DataFrame()

def format_currency(amount):
    return f"{amount:,.2f}"

def generate_group_name(l_rows, b_rows):
    """Generates a group name based on Ledger Ref Code."""
    if l_rows.empty:
        return ""

    # Extract codes from Ledger 'ref_code' (e.g. "KKM045-1" -> "KKM045")
    codes = l_rows['ref_code'].astype(str).str.upper().apply(lambda x: x.split('-')[0].strip())
    
    # Get unique values, excluding 'NAN'
    unique_codes = sorted([c for c in codes.unique() if c and c != 'NAN'])
    
    if unique_codes:
        return f"({', '.join(unique_codes)})"

    return ""

def auto_match_logic():
    """
    New Logic: 
    1. Get unique codes from Ledger 'No' column (stripping suffix like -1).
    2. Check if that code exists in any Bank Description.
    3. If yes, check if the sums balance.
    """
    u_l = st.session_state['ledger_df'][~st.session_state['ledger_df']['matched']]
    u_b = st.session_state['bank_df'][~st.session_state['bank_df']['matched']]
    
    matches_found = 0
    
    # 1. Prepare Ledger Codes (Normalize: KKM045-1 -> KKM045)
    # We create a temporary column for matching
    u_l = u_l.copy()
    u_l['base_code'] = u_l['ref_code'].astype(str).str.upper().apply(lambda x: x.split('-')[0].strip())
    
    # Filter out empty or short codes to avoid false positives
    valid_codes = [c for c in u_l['base_code'].unique() if len(c) >= 3 and c != "NAN"]
    
    for code in valid_codes:
        # 2. Find Bank items containing this code
        # We search the code inside the Bank Description
        potential_bank = u_b[u_b['description'].astype(str).str.upper().str.contains(code, regex=False)]
        
        if not potential_bank.empty:
            # 3. Get all Ledger items with this specific base code
            potential_ledger = u_l[u_l['base_code'] == code]
            
            # 4. Check every bank item found (usually 1-to-Many)
            for b_idx, b_row in potential_bank.iterrows():
                b_amt = b_row['amount']
                l_sum = potential_ledger['amount'].sum()
                
                # Check Balance
                if abs(l_sum - b_amt) < 1.0:
                    l_ids = potential_ledger['id'].tolist()
                    b_ids = [b_row['id']]
                    
                    # Apply Match
                    st.session_state['ledger_df'].loc[st.session_state['ledger_df']['id'].isin(l_ids), 'matched'] = True
                    st.session_state['bank_df'].loc[st.session_state['bank_df']['id'].isin(b_ids), 'matched'] = True
                    
                    st.session_state['matches'].append({
                        "ledger_ids": l_ids,
                        "bank_ids": b_ids
                    })
                    matches_found += 1
                    
                    # Update pools to prevent double matching
                    u_l = u_l[~u_l['id'].isin(l_ids)]
                    u_b = u_b[~u_b['id'].isin(b_ids)]
                    break # Move to next code

    return matches_found

# ==========================================
# 2. APP UI
# ==========================================

st.set_page_config(layout="wide", page_title="KSN Accounting Recon App")

st.markdown("""
<style>
    /* Compact the layout */
    .block-container { padding-top: 1rem; padding-bottom: 5rem; }
    /* Dense tables */
    .stDataFrame { font-size: 0.85rem; }
    /* Expander header styling */
    div[data-testid="stExpander"] details summary p { 
        font-weight: 600; font-size: 1.05rem; 
    }
    /* Trash button styling */
    div[data-testid="column"] button {
        border-color: #ff4b4b;
        color: #ff4b4b;
    }
</style>
""", unsafe_allow_html=True)

st.title("🏦 KSN Accounting Recon App")
st.markdown("Made by DhiraPT")

# --- SIDEBAR ---
with st.sidebar:
    st.header("1. Upload Data")
    bank_file = st.file_uploader("Bank Statement (PDF)", type=["pdf"])
    ledger_file = st.file_uploader("Ledger (Excel)", type=["xlsx", "xls"])
    if st.button("🔄 Reset / Clear All", type="secondary"):
        st.session_state.clear()
        st.rerun()

# --- STATE ---
if 'bank_df' not in st.session_state: st.session_state['bank_df'] = pd.DataFrame()
if 'ledger_df' not in st.session_state: st.session_state['ledger_df'] = pd.DataFrame()
if 'matches' not in st.session_state: st.session_state['matches'] = []

# --- PARSING ---
if bank_file and st.session_state['bank_df'].empty:
    text = extract_text_from_pdf(bank_file)
    if text: st.session_state['bank_df'] = parse_bank_statement(text)

if ledger_file and st.session_state['ledger_df'].empty:
    st.session_state['ledger_df'] = parse_excel_ledger(ledger_file)

# --- HELPER: UNMATCH ---
def unmatch_items(id_list, side):
    if side == 'ledger':
        st.session_state['ledger_df'].loc[st.session_state['ledger_df']['id'].isin(id_list), 'matched'] = False
    else:
        st.session_state['bank_df'].loc[st.session_state['bank_df']['id'].isin(id_list), 'matched'] = False

# --- MAIN WORKSPACE ---
if not st.session_state['bank_df'].empty and not st.session_state['ledger_df'].empty:

    # ==========================================
    # SECTION 1: MATCHED GROUPS (NESTED ACCORDIONS)
    # ==========================================
    # Outer Accordion to hide the whole section
    with st.expander("✅ Matched Groups", expanded=False):
        
        matches = st.session_state['matches']
        groups_to_delete = []

        if not matches:
            st.info("No matches yet. Link items below.")
        
        # Iterate through groups (Inner Accordions)
        for i, match in enumerate(matches):
            group_num = i + 1
            
            # Filter for currently valid matched IDs
            l_ids = [uid for uid in match['ledger_ids'] if st.session_state['ledger_df'].loc[st.session_state['ledger_df']['id']==uid, 'matched'].values[0]]
            b_ids = [uid for uid in match['bank_ids'] if st.session_state['bank_df'].loc[st.session_state['bank_df']['id']==uid, 'matched'].values[0]]
            
            # Sync
            match['ledger_ids'] = l_ids
            match['bank_ids'] = b_ids

            if not l_ids and not b_ids:
                groups_to_delete.append(i)
                continue

            # Data
            l_rows = st.session_state['ledger_df'][st.session_state['ledger_df']['id'].isin(l_ids)].copy()
            b_rows = st.session_state['bank_df'][st.session_state['bank_df']['id'].isin(b_ids)].copy()

            l_sum = l_rows['amount'].sum() if not l_rows.empty else 0
            b_sum = b_rows['amount'].sum() if not b_rows.empty else 0
            diff = l_sum - b_sum

            smart_tag = generate_group_name(l_rows, b_rows)
            
            status = "🟢" if abs(diff) < 0.01 else "🔴"
            title = f"Group #{group_num} {smart_tag} {status} | Ledger: {format_currency(l_sum)} | Bank: {format_currency(b_sum)} | Diff: {format_currency(diff)}"

            # --- INNER ACCORDION ---
            with st.expander(title, expanded=False):
                # Header with Trash Icon
                c_act1, c_act2 = st.columns([0.85, 0.15])
                with c_act1:
                    st.caption("Review or unmatch specific items.")
                with c_act2:
                    if st.button("🗑️ Delete Group", key=f"del_grp_{i}"):
                        unmatch_items(l_ids, 'ledger')
                        unmatch_items(b_ids, 'bank')
                        groups_to_delete.append(i)
                        st.rerun()
                
                # Grids
                gc1, gc2 = st.columns(2)
                
                # LEDGER GRID
                with gc1:
                    st.markdown("**📖 Ledger**")
                    if not l_rows.empty:
                        l_rows.insert(0, "Unmatch", False)
                        edited_l = st.data_editor(
                            l_rows[['Unmatch', 'excel_row', 'date', 'description', 'amount', 'note']],
                            disabled=["excel_row", "date", "description", "amount"],
                            key=f"ge_l_{i}", hide_index=True,
                            column_config={
                                "Unmatch": st.column_config.CheckboxColumn(width="small"),
                                "amount": st.column_config.NumberColumn(format="accounting"),
                                "excel_row": st.column_config.NumberColumn("Row", width="small")
                            }
                        )
                        if edited_l['Unmatch'].any():
                            to_drop = edited_l[edited_l['Unmatch']].index
                            for idx, row in edited_l.iterrows():
                                 orig_id = l_rows.loc[idx, 'id']
                                 main_idx = st.session_state['ledger_df'][st.session_state['ledger_df']['id'] == orig_id].index[0]
                                 st.session_state['ledger_df'].at[main_idx, 'note'] = row['note']
                            ids_to_drop = [l_rows.loc[x, 'id'] for x in to_drop]
                            unmatch_items(ids_to_drop, 'ledger')
                            st.rerun()
                
                # BANK GRID
                with gc2:
                    st.markdown("**🏦 Bank**")
                    if not b_rows.empty:
                        b_rows.insert(0, "Unmatch", False)
                        edited_b = st.data_editor(
                            b_rows[['Unmatch', 'date', 'description', 'amount', 'note']],
                            disabled=["date", "description", "amount"],
                            key=f"ge_b_{i}", hide_index=True,
                            column_config={
                                "Unmatch": st.column_config.CheckboxColumn(width="small"),
                                "amount": st.column_config.NumberColumn(format="accounting")
                            }
                        )
                        if edited_b['Unmatch'].any():
                            to_drop = edited_b[edited_b['Unmatch']].index
                            for idx, row in edited_b.iterrows():
                                 orig_id = b_rows.loc[idx, 'id']
                                 main_idx = st.session_state['bank_df'][st.session_state['bank_df']['id'] == orig_id].index[0]
                                 st.session_state['bank_df'].at[main_idx, 'note'] = row['note']
                            ids_to_drop = [b_rows.loc[x, 'id'] for x in to_drop]
                            unmatch_items(ids_to_drop, 'bank')
                            st.rerun()

    # Process deletions
    if groups_to_delete:
        for idx in sorted(groups_to_delete, reverse=True):
            del st.session_state['matches'][idx]
        st.rerun()

    st.markdown("---")

    # Auto-Match Button
    c_auto1, c_auto2 = st.columns([1, 4])
    with c_auto1:
        if st.button("✨ Auto-Match by Code", help="Matches Bank items to Ledger items if they share a code (e.g. KKM045) and the amount sums up."):
            count = auto_match_logic()
            if count > 0:
                st.success(f"Successfully auto-matched {count} groups!")
                st.rerun()
            else:
                st.warning("No code-based matches found.")

    # ==========================================
    # SECTION 2: UNMATCHED WORKSPACE
    # ==========================================
    st.subheader("🧩 Unmatched Workspace")
    
    unmatched_l = st.session_state['ledger_df'][~st.session_state['ledger_df']['matched']].copy()
    unmatched_b = st.session_state['bank_df'][~st.session_state['bank_df']['matched']].copy()

    unmatched_l.insert(0, "Select", False)
    unmatched_b.insert(0, "Select", False)

    col_l, col_r = st.columns(2)

    with col_l:
        st.markdown("**📖 Ledger Items**")
        edited_ul = st.data_editor(
            unmatched_l[['Select', 'excel_row', 'date', 'description', 'amount', 'note']],
            disabled=["excel_row", "date", "description", "amount"],
            key="ul_editor", hide_index=True, height=500, width='stretch',
            column_config={
                "Select": st.column_config.CheckboxColumn(width="small"),
                "excel_row": st.column_config.NumberColumn("Row", width="small"),
                "amount": st.column_config.NumberColumn(format="accounting"),
                "note": st.column_config.TextColumn("Note", width="medium"),
                "description": st.column_config.TextColumn(width="large")
            }
        )

    with col_r:
        st.markdown("**🏦 Bank Items**")
        edited_ub = st.data_editor(
            unmatched_b[['Select', 'date', 'description', 'amount', 'note']],
            disabled=["date", "description", "amount"],
            key="ub_editor", hide_index=True, height=500, width='stretch',
            column_config={
                "Select": st.column_config.CheckboxColumn(width="small"),
                "amount": st.column_config.NumberColumn(format="accounting"),
                "note": st.column_config.TextColumn("Note", width="medium"),
                "description": st.column_config.TextColumn(width="large")
            }
        )

    # --- CALCULATION BAR ---
    sel_l = edited_ul[edited_ul['Select']]
    sel_b = edited_ub[edited_ub['Select']]

    l_sum = sel_l['amount'].sum()
    b_sum = sel_b['amount'].sum()
    diff = l_sum - b_sum

    st.markdown("### Action")
    m1, m2, m3, btn = st.columns([2, 2, 2, 2])
    
    m1.metric(f"Selected Ledger ({len(sel_l)})", format_currency(l_sum))
    m2.metric(f"Selected Bank ({len(sel_b)})", format_currency(b_sum))
    
    if diff == 0 and (l_sum > 0 or b_sum > 0):
        m3.metric("Difference", "0.00", delta="Balanced!")
    else:
        m3.metric("Difference", format_currency(diff), delta_color="inverse")

    if btn.button("🔗 Link Selected", type="primary", width='stretch'):
        if sel_l.empty and sel_b.empty:
            st.warning("Please select items to link.")
        else:
            # Sync Notes
            for idx, row in sel_l.iterrows():
                orig_id = unmatched_l.loc[idx, 'id']
                main_idx = st.session_state['ledger_df'][st.session_state['ledger_df']['id'] == orig_id].index[0]
                st.session_state['ledger_df'].at[main_idx, 'note'] = row['note']
            
            for idx, row in sel_b.iterrows():
                orig_id = unmatched_b.loc[idx, 'id']
                main_idx = st.session_state['bank_df'][st.session_state['bank_df']['id'] == orig_id].index[0]
                st.session_state['bank_df'].at[main_idx, 'note'] = row['note']

            # Link
            l_ids = [unmatched_l.loc[i, 'id'] for i in sel_l.index]
            b_ids = [unmatched_b.loc[i, 'id'] for i in sel_b.index]
            
            st.session_state['ledger_df'].loc[st.session_state['ledger_df']['id'].isin(l_ids), 'matched'] = True
            st.session_state['bank_df'].loc[st.session_state['bank_df']['id'].isin(b_ids), 'matched'] = True
            
            st.session_state['matches'].append({
                "ledger_ids": l_ids,
                "bank_ids": b_ids
            })
            st.success("Linked successfully!")
            st.rerun()

else:
    st.info("👋 Upload your Bank Statement and Ledger to begin.")
