import streamlit as st
import pandas as pd

st.set_page_config(
    page_title="Zenith Kataloogi Tootebaas",
    page_icon="🔍",
    layout="wide"
)

XLSX_PATH = "outputs/zenith_catalog_database_20260609/Zenith_Kataloogi_Tootebaas_2026-06-09.xlsx"

@st.cache_data
def load_sheet(sheet_name):
    return pd.read_excel(XLSX_PATH, sheet_name=sheet_name, dtype=str).fillna("")

@st.cache_data
def get_sheet_names():
    xl = pd.ExcelFile(XLSX_PATH)
    return xl.sheet_names

st.title("🔍 Zenith Kataloogi Tootebaas")
st.caption("Plastok OÜ — Zenith 2026 materjalibaas")

try:
    sheet_names = get_sheet_names()
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
            file_name="Zenith_Kataloogi_Tootebaas_2026-06-09.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# Lae valitud leht
df = load_sheet(leht)

# Rakenda otsing
if otsing.strip():
    mask = df.apply(lambda col: col.str.contains(otsing, case=False, na=False)).any(axis=1)
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
