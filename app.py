import streamlit as st
import pandas as pd
import re
import time

st.set_page_config(page_title="ITP ↔ WIR ↔ Activity Matching", layout="wide")
st.title("📊 ITP ↔ WIR ↔ Activity Matching Tool")

# ------------------------
# Helper: normalize text
# ------------------------
def normalize(text):
    if pd.isna(text): 
        return ""
    text = str(text).lower()
    text = re.sub(r'[^a-z0-9\s]', ' ', text)
    text = re.sub(r'\s+', ' ', text).strip()
    return text

# ------------------------
# PART 1 – ITP ↔ WIR
# ------------------------
st.header("🔹 Part 1: Match ITP Log with WIR Log (Title + Disp.)")

itp_file = st.file_uploader("Upload ITP Log (Excel)", type="xlsx", key="itp")
wir_file = st.file_uploader("Upload WIR Log (Excel)", type="xlsx", key="wir")

if itp_file and wir_file:
    itp_df = pd.read_excel(itp_file)
    wir_df = pd.read_excel(wir_file)

    st.success(f"✅ Loaded ITP Log ({len(itp_df)} rows), WIR Log ({len(wir_df)} rows)")

    itp_title_col = st.selectbox("Select ITP Title Column", itp_df.columns)
    itp_no_col = st.selectbox("Select ITP Number Column", itp_df.columns)
    itp_disp_col = st.selectbox("Select ITP Disp. Column", itp_df.columns)

    wir_title_col = st.selectbox("Select WIR Title Column", wir_df.columns)
    wir_doc_col = st.selectbox("Select WIR Document Number Column", wir_df.columns)
    wir_disp_col = st.selectbox("Select WIR Disp. Column", wir_df.columns)

    if st.button("Start Part 1 Matching"):
        st.info("⏳ Matching ITP ↔ WIR by Title + Disp...")
        start = time.time()

        itp_df["_norm_title"] = itp_df[itp_title_col].apply(normalize)
        wir_df["_norm_title"] = wir_df[wir_title_col].apply(normalize)

        results = []
        total = len(wir_df)
        progress = st.progress(0)

        for idx, wir in wir_df.iterrows():
            wir_title = wir["_norm_title"]
            wir_disp = wir[wir_disp_col]
            matches = itp_df[itp_df["_norm_title"] == wir_title]

            for _, itp in matches.iterrows():
                # ✅ check Disp. condition
                if str(itp[itp_disp_col]) == str(wir_disp):
                    results.append({
                        "WIR Document Number": wir[wir_doc_col],
                        "WIR Title": wir[wir_title_col],
                        "WIR Disp.": wir[wir_disp_col],
                        "ITP Number": itp[itp_no_col],
                        "ITP Title": itp[itp_title_col],
                        "ITP Disp.": itp[itp_disp_col]
                    })
            progress.progress((idx + 1) / total)

        part1_df = pd.DataFrame(results)
        st.success(f"✅ Done! Found {len(part1_df)} matches in {time.time()-start:.2f}s")
        st.dataframe(part1_df.head(50))

        st.session_state["part1_result"] = part1_df

        csv = part1_df.to_csv(index=False).encode("utf-8")
        st.download_button("💾 Download Part 1 Results", data=csv, file_name="part1_results.csv")

# ------------------------
# PART 2 – Activity ↔ ITP
# ------------------------
st.header("🔹 Part 2: Match Activities with ITP (check Disp. consistency)")

if "part1_result" not in st.session_state:
    st.info("⚠️ Please complete Part 1 first and keep results in session.")
else:
    activity_file = st.file_uploader("Upload Activity Log (Excel)", type="xlsx", key="activity")
    if activity_file:
        act_df = pd.read_excel(activity_file)
        part1_df = st.session_state["part1_result"]

        st.success(f"✅ Loaded Activity Log ({len(act_df)} rows), Part 1 Results ({len(part1_df)} rows)")

        act_no_col = st.selectbox("Select Activity Number Column", act_df.columns)
        itp_no_col = "ITP Number"   # from part1 result
        itp_disp_col = "ITP Disp."
        wir_disp_col = "WIR Disp."

        if st.button("Start Part 2 Matching"):
            st.info("⏳ Matching Activities ↔ ITP (with Disp. check)...")
            start = time.time()

            results2 = []
            total2 = len(act_df)
            progress2 = st.progress(0)

            for idx, act in act_df.iterrows():
                act_no = str(act[act_no_col]).strip()
                # Find candidate ITP from Part 1
                candidates = part1_df[part1_df[itp_no_col].astype(str).str.strip() == act_no]

                for _, row in candidates.iterrows():
                    # ✅ Disp. must be consistent between ITP and WIR (already filtered in Part 1)
                    if row[itp_disp_col] == row[wir_disp_col]:
                        results2.append({
                            "Activity Number": act[act_no_col],
                            "ITP Number": row[itp_no_col],
                            "ITP Disp.": row[itp_disp_col],
                            "WIR Document Number": row["WIR Document Number"],
                            "WIR Disp.": row[wir_disp_col]
                        })

                progress2.progress((idx + 1) / total2)

            final_df = pd.DataFrame(results2)
            st.success(f"✅ Done! Found {len(final_df)} activity matches in {time.time()-start:.2f}s")
            st.dataframe(final_df.head(50))

            csv2 = final_df.to_csv(index=False).encode("utf-8")
            st.download_button("💾 Download Part 2 Results", data=csv2, file_name="part2_results.csv")
