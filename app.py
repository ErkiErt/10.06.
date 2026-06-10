import streamlit as st
import pandas as pd
import glob
import os

st.set_page_config(
    page_title="Zenith Kataloogi Tootebaas",
    page_icon="🔍",
    layout="wide"
)

# --- Dünaamiline failitee: leiab uusima .xlsx faili outputs/ alt ---
def find_latest_xlsx(base_dir="outputs"):
    pattern = os.path.join(base_dir, "**", "*.xlsx")
    files = glob.glob(pattern, recursive=True)
    if not files:
        return None
    return max(files, key=os.path.getmtime)

XLSX_PATH = find_latest_xlsx()

@st.cache_data
def load_sheet(path, sheet_name):
    return pd.read_excel(path, sheet_name=sheet_name, dtype=str).fillna("")

@st.cache_data
def get_sheet_names(path):
    xl = pd.ExcelFile(path)
    return xl.sheet_names

st.title("🔍 Zenith Kataloogi Tootebaas")
st.caption("Plastok OÜ — Zenith 2026 materjalibaas")

if XLSX_PATH is None:
    st.error("❌ Exceli faili ei leitud kaustas 'outputs/'. Kontrolli, kas fail on üles laaditud.")
    st.stop()

st.caption(f"📂 Aktiivne fail: `{XLSX_PATH}`")

try:
    sheet_names = get_sheet_names(XLSX_PATH)
except Exception as e:
    st.error(f"Exceli faili laadimine ebaõnnestus: {e}")
    st.stop()

# Sidebar
with st.sidebar:
    st.header("⚙️ Seaded")
    leht = st.selectbox("Vali leht", sheet_names)
    st.divider()
    st.markdown("**Otsing**")
    otsing = st.text_input("Otsi kõigist veergudest", placeholder="nt. NBR, 70 Shore...")
    st.divider()
    st.markdown("**Allalaadimine**")
    with open(XLSX_PATH, "rb") as f:
        st.download_button(
            label="⬇️ Laadi Excel alla",
            data=f,
            file_name=os.path.basename(XLSX_PATH),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# Lae valitud leht
try:
    df = load_sheet(XLSX_PATH, leht)
except Exception as e:
    st.error(f"Lehe '{leht}' laadimine ebaõnnestus: {e}")
    st.stop()

# Rakenda otsing
if otsing.strip():
    mask = df.apply(
        lambda col: col.str.contains(otsing, case=False, na=False, regex=False)
    ).any(axis=1)
    df_filtered = df[mask]
    st.info(f"Otsing '{otsing}' — leiti {len(df_filtered)} rida (kokku {len(df)})")
else:
    df_filtered = df

# Veeru filter
if not df_filtered.empty:
    with st.expander("🎛️ Veeru filter", expanded=False):
        cols = st.multiselect("Näita veerge", df_filtered.columns.tolist(), default=df_filtered.columns.tolist())
        if cols:
            df_filtered = df_filtered[cols]

# Näita tabel
st.subheader(f"📋 {leht}  —  {len(df_filtered)} rida")
st.dataframe(df_filtered, use_container_width=True, height=600)
