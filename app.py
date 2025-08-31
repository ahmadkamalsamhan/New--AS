import streamlit as st
import pandas as pd
import re
from collections import defaultdict
from io import BytesIO
import time

st.set_page_config(page_title="ITP ↔ WIR Matching (Fast, with Disp.)", layout="wide")
st.title("📊 ITP ↔ WIR Title Matching + Activity Status (Optimized, with Disp.)")

# -------------------------------
# Text preprocessing
# -------------------------------
def preprocess_text(text):
    if pd.isna(text):
        return set()
    text = str(text).lower()
    text = re.sub(r'[^a-z0-9\s]', ' ', text)
    text = re.sub(r'\s+', ' ', text).strip()
    return set(text.split())

def assign_status(code):
    if pd.isna(code):
        return 0
    code = str(code).strip().upper()
    if code in ['A', 'B']:
        return 1
    if code in ['C', 'D']:
        return 2
    return 0

def same_disp(a, b):
    # case-insensitive + trimmed comparison; keeps the logic but robust to case/spaces
    return str(a).strip().lower() == str(b).strip().lower()

# -------------------------------
# Tabs
# -------------------------------
tab1, tab2 = st.tabs(["Part 1: Title Matching (ITP ↔ WIR)", "Part 2: Activity Matching"])

# ===============================
# Part 1: Title Matching (Token/Inverted Index) + Disp. gate
# ===============================
with tab1:
    st.header("🔹 Part 1: WIR ↔ ITP Title Matching (Token-based) + Disp. must match")

    wir_file = st.file_uploader("Upload WIR Log (Document Control Log)", type=["xlsx"], key="wir1")
    itp_file = st.file_uploader("Upload ITP Log", type=["xlsx"], key="itp1")

    if wir_file and itp_file:
        wir_log = pd.read_excel(wir_file)
        itp_log = pd.read_excel(itp_file)

        wir_log.columns = wir_log.columns.str.strip()
        itp_log.columns = itp_log.columns.str.strip()

        # Column selection
        st.subheader("Select columns")
        col1, col2 = st.columns(2)
        with col1:
            wir_doc_col   = st.selectbox("WIR Document No.", wir_log.columns, key="wir_doc")
            wir_title_col = st.selectbox("WIR Title (Title / Description2)", wir_log.columns, key="wir_title")
            wir_pm_col    = st.selectbox("WIR PM Web Code", wir_log.columns, key="wir_pm")
            wir_disp_col  = st.selectbox("WIR Disp.", wir_log.columns, key="wir_disp")
        with col2:
            itp_no_col    = st.selectbox("ITP No.", itp_log.columns, key="itp_no")
            itp_title_col = st.selectbox("ITP Title (Title / Description)", itp_log.columns, key="itp_title")
            itp_disp_col  = st.selectbox("ITP Disp.", itp_log.columns, key="itp_disp")

        if st.button("🚀 Start Title Matching"):
            st.info("Building tokens and inverted index…")
            t0 = time.time()

            # Tokens
            wir_log['WIR_Tokens'] = wir_log[wir_title_col].apply(preprocess_text)
            itp_log['ITP_Tokens'] = itp_log[itp_title_col].apply(preprocess_text)

            # Inverted index: token -> set(ITP indices)
            token_to_itp = defaultdict(set)
            for idx, tokens in enumerate(itp_log['ITP_Tokens']):
                for tok in tokens:
                    token_to_itp[tok].add(idx)

            st.info("Matching WIR titles to ITP titles (candidate-only)…")
            matched_rows = []
            total = len(wir_log)
            pbar = st.progress(0)
            status_text = st.empty()

            for i, wrow in wir_log.iterrows():
                w_tokens = wrow['WIR_Tokens']
                if not w_tokens:
                    # No tokens -> skip quickly
                    if i % 20 == 0 or i == total - 1:
                        pbar.progress((i + 1) / total)
                        status_text.text(f"{i+1}/{total}")
                    continue

                # Collect candidate ITP rows via shared tokens
                candidate_idxs = set()
                for t in w_tokens:
                    candidate_idxs |= token_to_itp.get(t, set())

                best_idx = None
                best_score = 0.0

                # Evaluate only candidates; require Disp. equality
                for idx in candidate_idxs:
                    # Disp must match BEFORE scoring
                    if not same_disp(wrow[wir_disp_col], itp_log.at[idx, itp_disp_col]):
                        continue

                    itp_tokens = itp_log.at[idx, 'ITP_Tokens']
                    if not itp_tokens:
                        continue

                    # Score = Jaccard-like using WIR as denominator (as you had)
                    score = len(w_tokens & itp_tokens) / max(len(w_tokens), 1)
                    if score > best_score:
                        best_score = score
                        best_idx = idx

                if best_idx is not None:
                    itp_row = itp_log.loc[best_idx]
                    matched_rows.append({
                        "WIR Document No": wrow[wir_doc_col],
                        "WIR Title": wrow[wir_title_col],
                        "WIR Disp.": wrow[wir_disp_col],
                        "PM Web Code": wrow[wir_pm_col],
                        "ITP No": itp_row[itp_no_col],
                        "ITP Title": itp_row[itp_title_col],
                        "ITP Disp.": itp_row[itp_disp_col],
                        "Match Score (%)": round(best_score * 100, 1),
                    })

                if i % 20 == 0 or i == total - 1:
                    pbar.progress((i + 1) / total)
                    status_text.text(f"{i+1}/{total}")

            part1_df = pd.DataFrame(matched_rows)
            st.success(f"✅ Part 1 complete — {len(part1_df)} matches in {time.time()-t0:.2f}s")
            st.dataframe(part1_df.head(100), use_container_width=True)

            # Keep in session for Part 2
            st.session_state['part1_result'] = part1_df.copy()

            # Download
            out = BytesIO()
            part1_df.to_excel(out, index=False, engine='openpyxl')
            out.seek(0)
            st.download_button(
                "📥 Download Part 1 Result (Excel)",
                data=out,
                file_name="Part1_WIR_ITP_Title_Match.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

# ===============================
# Part 2: Activity Matching (reuses Part 1) — no Disp in Activities; rely on Part 1’s Disp gate
# ===============================
with tab2:
    st.header("🔹 Part 2: Activity Matching (uses Part 1 result)")

    if 'part1_result' not in st.session_state:
        st.info("Please finish Part 1 first (the results are used here).")
    else:
        part1_df = st.session_state['part1_result']

        activity_file = st.file_uploader("Upload ITP Activities Log", type=["xlsx"], key="activity2")
        if activity_file:
            activity_log = pd.read_excel(activity_file)
            activity_log.columns = activity_log.columns.str.strip()
            part1_df.columns = part1_df.columns.str.strip()

            st.success(f"✅ Loaded Activities ({len(activity_log)} rows) + Part 1 ({len(part1_df)} rows).")

            # Column selection in Activities
            activity_desc_col = st.selectbox("Activity Description column", activity_log.columns, key="act_desc")
            itp_ref_col      = st.selectbox("ITP Reference column (in Activities)", activity_log.columns, key="act_itp_ref")
            activity_no_col  = st.selectbox("Activity No. column", activity_log.columns, key="act_no")

            if st.button("🚀 Start Activity Matching"):
                st.info("Scoring activities vs ITP title tokens (from Part 1)…")
                t0 = time.time()

                # Precompute tokens
                activity_log['Activity_Tokens'] = activity_log[activity_desc_col].apply(preprocess_text)
                # One set of ITP tokens per ITP No (from Part 1 output)
                # (ITP title tokens identical for multiple rows of same ITP, but we compute once)
                itp_tokens_map = (
                    part1_df[['ITP No', 'ITP Title']]
                    .drop_duplicates('ITP No')
                    .assign(ITP_Tokens=lambda d: d['ITP Title'].apply(preprocess_text))
                    .set_index('ITP No')['ITP_Tokens']
                    .to_dict()
                )

                status_codes = []
                match_scores = []

                total = len(activity_log)
                pbar = st.progress(0)
                status_text = st.empty()

                # For each activity, find its ITP No → tokens, compute overlap score with activity tokens,
                # and take a PM Web Code from *any* Part 1 row for that ITP (Disp already matched in Part 1).
                for i, arow in activity_log.iterrows():
                    itp_no = str(arow[itp_ref_col]).strip()
                    a_tokens = arow['Activity_Tokens']

                    itp_tokens = itp_tokens_map.get(itp_no, set())
                    score = 0.0
                    if itp_tokens:
                        score = len(a_tokens & itp_tokens) / max(len(itp_tokens), 1)

                    # Choose the first PM Web Code for that ITP from Part 1 (keeping your previous flow)
                    subset = part1_df[part1_df['ITP No'].astype(str).str.strip() == itp_no]
                    pm_code = subset.iloc[0]['PM Web Code'] if not subset.empty else None

                    status_codes.append(assign_status(pm_code))
                    match_scores.append(round(score * 100, 1))

                    if i % 20 == 0 or i == total - 1:
                        pbar.progress((i + 1) / total)
                        status_text.text(f"{i+1}/{total}")

                # Append outputs to the SAME activities shape
                activity_log['WIR Status Code'] = status_codes
                activity_log['Match Score (%)'] = match_scores

                st.success(f"✅ Part 2 complete in {time.time()-t0:.2f}s")
                st.dataframe(activity_log.head(100), use_container_width=True)

                # Download
                out2 = BytesIO()
                activity_log.to_excel(out2, index=False, engine='openpyxl')
                out2.seek(0)
                st.download_button(
                    "📥 Download Part 2 Result (Excel)",
                    data=out2,
                    file_name="Part2_Activity_Match.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
