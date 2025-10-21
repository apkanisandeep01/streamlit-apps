import streamlit as st
import pandas as pd
from io import BytesIO
import time

# =========================
# Page Setup
# =========================
st.set_page_config(
    page_title="🔗 Excel Join Tool",
    page_icon="🔗",
    layout="wide"
)

# =========================
# Title & Description
# =========================
st.markdown(
    """
    <style>
    .title {
        text-align: center;
        font-size: 42px;
        font-weight: 800;
        color: #2E86C1;
    }
    .subtitle {
        text-align: center;
        font-size: 18px;
        color: #566573;
    }
    .stButton>button {
        background: linear-gradient(90deg, #2E86C1, #00BCD4);
        color: white;
        border-radius: 10px;
        padding: 0.6em 1.4em;
        border: none;
        font-weight: 600;
        transition: 0.3s ease;
    }
    .stButton>button:hover {
        background: linear-gradient(90deg, #00BCD4, #2E86C1);
        transform: scale(1.03);
    }
    </style>
    <h1 class="title">🔗 Excel Join Tool</h1>
    <p class="subtitle">Easily merge two Excel files using <b>Inner</b>, <b>Left</b>, <b>Right</b>, or <b>Full (Outer)</b> joins — now with smart guidance.</p>
    """,
    unsafe_allow_html=True
)

st.divider()

# =========================
# User Guide for Beginners
# =========================
with st.expander("📘 Quick Start: How to Use This Tool", expanded=False):
    st.markdown("""
    **Step-by-step guide for first-time users:**

    1️⃣ **Upload two Excel files** — one as your *primary dataset*, one as the *reference dataset*.

    2️⃣ **Select header rows** — adjust if column names are not in the first row.

    3️⃣ **Choose join columns** — select the *common field* (like Farmer ID, Aadhar No, etc.).

    4️⃣ **Select join type:**
    - 🟩 **Inner Join:** Keeps only matching rows.
    - 🟦 **Left Join:** Keeps all from the first file, matching from the second.
    - 🟨 **Right Join:** Keeps all from the second file, matching from the first.
    - 🟪 **Outer Join:** Keeps all from both files.

    5️⃣ **Download the result** after joining.

    💡 *Tip:* Always ensure the join columns have the same format (text vs number) and identical spelling.
    """)

# =========================
# Helper Functions
# =========================
def read_excel_with_header(file, header_row):
    """Read Excel with custom header row."""
    return pd.read_excel(file, header=header_row)

def convert_df_to_excel(df):
    """Convert DataFrame to downloadable Excel file."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='MergedData')
    return output.getvalue()

# =========================
# File Upload Section
# =========================
st.markdown("### 📂 Step 1: Upload Your Excel Files")
col1, col2 = st.columns(2)

with col1:
    file1 = st.file_uploader("📘 Upload First Excel File", type=["xls", "xlsx", "ods"], key="file1")
    if file1:
        st.toast("✅ First file uploaded successfully!", icon="📘")

with col2:
    file2 = st.file_uploader("📗 Upload Second Excel File", type=["xls", "xlsx", "ods"], key="file2")
    if file2:
        st.toast("✅ Second file uploaded successfully!", icon="📗")

st.divider()

# =========================
# Processing Section
# =========================
if file1 and file2:
    st.markdown("### ⚙️ Step 2: Configure Join Settings")

    col3, col4 = st.columns(2)
    with col3:
        header1 = st.number_input("Header row for First File (starting from 1)", min_value=1, value=1, step=1) - 1
    with col4:
        header2 = st.number_input("Header row for Second File (starting from 1)", min_value=1, value=1, step=1) - 1

    try:
        df1 = read_excel_with_header(file1, header1)
        df2 = read_excel_with_header(file2, header2)
        st.toast("📊 Files loaded successfully!", icon="📊")

        with st.expander("📘 Preview First File (Top 5 Rows)"):
            st.dataframe(df1.head(), use_container_width=True)
        with st.expander("📗 Preview Second File (Top 5 Rows)"):
            st.dataframe(df2.head(), use_container_width=True)

        st.divider()
        st.markdown("### 🔗 Step 3: Select Join Columns & Type")

        col5, col6, col7 = st.columns(3)
        with col5:
            join_col1 = st.selectbox("🔸 Join Column (File 1)", options=df1.columns)
        with col6:
            join_col2 = st.selectbox("🔹 Join Column (File 2)", options=df2.columns)
        with col7:
            join_type = st.selectbox(
                "⚖️ Join Type",
                options=["inner", "left", "right", "outer"],
                format_func=str.capitalize,
            )

        # Join type description
        join_descriptions = {
            "inner": "🟩 **Inner Join:** Returns only matching rows between both files.",
            "left": "🟦 **Left Join:** Keeps all rows from the first file and matches from the second.",
            "right": "🟨 **Right Join:** Keeps all rows from the second file and matches from the first.",
            "outer": "🟪 **Full Outer Join:** Returns all rows from both files, even if they don’t match."
        }
        st.info(join_descriptions[join_type])

        # =========================
        # Smart Validation Before Join
        # =========================
        st.divider()
        st.markdown("### 🧠 Step 4: Smart Validation Check")

        # Check if both columns exist and have overlap
        df1_keys = df1[join_col1].dropna().astype(str).unique()
        df2_keys = df2[join_col2].dropna().astype(str).unique()

        common = set(df1_keys) & set(df2_keys)
        if len(common) == 0:
            st.error(
                f"⚠️ No matching values found between **'{join_col1}'** (File 1) "
                f"and **'{join_col2}'** (File 2)."
            )
            st.warning("""
            👉 **Possible Causes & Fixes:**
            - The column names may look same but contain spaces or typos.
            - Data types may differ (e.g., one column is text, other is number).
            - Check for extra spaces or leading zeros.
            - Try cleaning both Excel columns using Excel’s “Trim” or convert to same type before uploading.
            """)
            st.stop()
        else:
            st.success(f"✅ Found {len(common)} matching records. Proceeding with join...")

        # =========================
        # Perform Join
        # =========================
        with st.spinner(f"🔄 Performing {join_type.capitalize()} Join..."):
            time.sleep(1)
            merged_df = pd.merge(df1, df2, how=join_type, left_on=join_col1, right_on=join_col2)
        st.toast(f"✅ {join_type.capitalize()} Join completed successfully!", icon="✅")

        # =========================
        # Display Result
        # =========================
        st.markdown(f"### 📊 Step 5: {join_type.capitalize()} Join Result Preview")
        st.dataframe(merged_df.head(10), use_container_width=True)
        st.info(f"🔹 Total Rows in Result: {len(merged_df)}")

        # =========================
        # Download Button
        # =========================
        st.divider()
        st.markdown("### 💾 Step 6: Download Merged File")
        st.download_button(
            label="📥 Download Excel File",
            data=convert_df_to_excel(merged_df),
            file_name=f"{join_type}_joined_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        st.toast("📁 Merged Excel file ready for download!", icon="📁")

    except Exception as e:
        st.error(f"❌ Something went wrong while processing: {e}")
        st.toast("⚠️ Error occurred. Please verify your files or column selections.", icon="⚠️")
