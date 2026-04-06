import streamlit as st
import pandas as pd
import datetime

# -------------------------------
# 🎨 UI STYLE
# -------------------------------
st.markdown("""
<style>
body {
    background-color: #0e1117;
    color: white;
}
.kpi-card {
    background-color: #1c1f26;
    padding: 15px;
    border-radius: 10px;
    text-align: center;
}
.kpi-title {
    font-size: 14px;
    color: #aaa;
}
.kpi-value {
    font-size: 22px;
    font-weight: bold;
    color: #00c8ff;
}
</style>
""", unsafe_allow_html=True)

# -------------------------------
# 🚀 CACHE LOADING (VERY IMPORTANT)
# -------------------------------
@st.cache_data
def load_file(file):
    if file.name.endswith(".xlsb"):
        return pd.read_excel(file, engine="pyxlsb", header=None)
    else:
        return pd.read_excel(file, engine="openpyxl", header=None)

# -------------------------------
# Title
# -------------------------------
st.markdown("<h1 style='text-align:center;'>📊 DGR Dashboard</h1><hr>", unsafe_allow_html=True)

# -------------------------------
# 📁 Upload Files
# -------------------------------
uploaded_files = st.file_uploader(
    "📁 Upload DGR Files",
    type=["xlsx", "xlsb"],
    accept_multiple_files=True
)

file_dict = {}

if uploaded_files:
    for file in uploaded_files:
        name = file.name
        date_part = name.split("_")[-1].replace(".xlsb", "").replace(".xlsx", "")
        file_dict[date_part] = file

# -------------------------------
# Date Selection
# -------------------------------
if file_dict:
    selected_date = st.selectbox("📅 Select Date", sorted(file_dict.keys()))
    submit = st.button("Submit")
else:
    st.warning("👉 Upload DGR files to continue")
    submit = False

# -------------------------------
# Load Data
# -------------------------------
if submit:
    file = file_dict[selected_date]
    df = load_file(file)

    # 🔥 Convert once for faster searching
    text_df = df.astype(str)

    st.session_state["df"] = df
    st.session_state["text_df"] = text_df

# -------------------------------
# MAIN LOGIC
# -------------------------------
if "df" in st.session_state:

    df = st.session_state["df"]
    text_df = st.session_state["text_df"]

    # -------------------------------
    # 📅 Extract Info (FAST SEARCH)
    # -------------------------------
    date_idx = text_df[text_df.iloc[:, :].apply(lambda r: r.str.contains("Date", case=False).any(), axis=1)].index
    month_idx = text_df[text_df.iloc[:, :].apply(lambda r: r.str.contains("Month", case=False).any(), axis=1)].index

    st.subheader("📅 Extracted Info")

    if len(date_idx) > 0:
        raw_date = df.iloc[date_idx[0], 4]
        try:
            real_date = datetime.datetime(1899, 12, 30) + datetime.timedelta(days=float(raw_date))
            st.write("Date:", real_date.strftime("%d-%b-%Y"))
        except:
            st.write("Date:", raw_date)

    if len(month_idx) > 0:
        st.write("Month:", df.iloc[month_idx[0], 4])

    # -------------------------------
    # SECTION SELECT
    # -------------------------------
    st.subheader("📊 Select Section")

    section = st.selectbox(
        "Choose Section",
        ["Key Performance Indicators", "Breakdown"]
    )

    # ============================================================
    # KPI SECTION
    # ============================================================
    if section == "Key Performance Indicators":

        view_type = st.selectbox("Select View", ["Daily", "MTD", "YTD"])
        st.subheader(f"📈 KPI Summary ({view_type})")

        kpi_idx = text_df[text_df.apply(
            lambda r: r.str.contains("Key Performance Indicators", case=False).any(), axis=1
        )].index

        if len(kpi_idx) > 0:

            start = kpi_idx[0]
            header_row = df.iloc[start]

            col_index = None
            for i, val in enumerate(header_row):
                if str(val).strip().lower() == view_type.lower():
                    col_index = i
                    break

            if col_index is None:
                col_index = 5

            kpi_table = df.iloc[start+1:start+15]

            cols = st.columns(3)
            i = 0
            chart_data = []

            for _, row in kpi_table.iterrows():

                name = str(row[3]).strip() if len(row) > 3 else ""
                if not name or name.lower() == "none":
                    continue

                value = row[col_index] if col_index < len(row) else None

                if pd.isna(value):
                    value = row[5] if len(row) > 5 else None

                display_value = value

                try:
                    numeric_value = float(value)
                    display_value = round(numeric_value, 2)
                except:
                    numeric_value = None

                if "Variance" in name and numeric_value is not None:
                    display_value = f"{numeric_value * 100:.1f}%"

                with cols[i % 3]:
                    st.markdown(f"""
                    <div class="kpi-card">
                        <div class="kpi-title">{name}</div>
                        <div class="kpi-value">{display_value}</div>
                    </div>
                    """, unsafe_allow_html=True)

                i += 1

                if numeric_value is not None:
                    if any(x in name for x in ["WTG", "Day Max", "Day Min"]):
                        continue
                    if numeric_value > 10000:
                        continue

                    chart_data.append({
                        "KPI": name,
                        "Value": numeric_value
                    })

            if chart_data:
                st.markdown("### 📊 KPI Chart")
                chart_df = pd.DataFrame(chart_data)
                chart_df = chart_df.sort_values("Value", ascending=False)
                st.bar_chart(chart_df.set_index("KPI"))

    # ============================================================
    # BREAKDOWN
    # ============================================================
    elif section == "Breakdown":

        st.subheader("🔧 Breakdown")

        clean_data = []

        breakdown_idx = text_df[text_df.apply(
            lambda r: r.str.contains("Breakdown", case=False).any(), axis=1
        )].index

        if len(breakdown_idx) > 0:

            start = breakdown_idx[0]
            breakdown_table = df.iloc[start+2:start+20]

            for _, row in breakdown_table.iterrows():
                values = row.dropna().values
                if len(values) >= 3:
                    clean_data.append({
                        "WTG": values[0],
                        "Remark": values[1],
                        "Down Time": values[2]
                    })

        clean_df = pd.DataFrame(clean_data)

        if clean_df.empty:
            st.error("❌ No Breakdown detected")
        else:
            st.success(f"✅ Found {len(clean_df)} breakdown records")

            search = st.text_input("🔍 Search WTG")

            if search:
                clean_df = clean_df[
                    clean_df["WTG"].astype(str).str.contains(search, case=False)
                ]

            def to_minutes(t):
                try:
                    h, m = str(t).split(":")[:2]
                    return int(h) * 60 + int(m)
                except:
                    return 0

            clean_df["Minutes"] = clean_df["Down Time"].apply(to_minutes)

            if not clean_df.empty:
                worst = clean_df.sort_values("Minutes", ascending=False).iloc[0]

                st.error(f"""
🚨 Highest Breakdown  
WTG: {worst['WTG']}  
Issue: {worst['Remark']}  
Down Time: {worst['Down Time']}
""")

            st.dataframe(clean_df.drop(columns=["Minutes"]))

            st.download_button(
                label="⬇️ Download Breakdown",
                data=clean_df.to_csv(index=False),
                file_name="breakdown.csv",
                mime="text/csv"
            )

else:
    st.warning("👉 Upload files and click Submit")
