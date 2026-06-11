import streamlit as st
import pandas as pd
import glob
import os
import re

st.set_page_config(
    page_title="Zenith Materjali Assistent",
    page_icon="🧱",
    layout="wide"
)

# ── Materjali kirjeldused ──────────────────────────────────────────────
MATERJALI_INFO = {
    "NBR": {
        "nimetus": "Nitriilkummi (NBR)",
        "omadused": "Kõrge õli-, kütuse- ja rasvakindlus. Töötemperatuur −40 … +120 °C.",
        "kasutus": "Hüdraulikatiivikud, õlitorud, kütusesüsteemid, tööstustiivikud.",
    },
    "EPDM": {
        "nimetus": "Etüleen-propüleenkummi (EPDM)",
        "omadused": "Suurepärane ilmastiku-, osooni- ja UVkindlus. Töötemp −40 … +150 °C.",
        "kasutus": "Katusetiivikud, aknaraamid, jõeveesüsteemid, autouksed.",
    },
    "SBR": {
        "nimetus": "Stüreen-butadieenkummi (SBR)",
        "omadused": "Hea kulumiskindlus ja mehaanilised omadused. Töötemp −40 … +100 °C.",
        "kasutus": "Tiivikuplaadid, põrandakatted, lintkonveierid.",
    },
    "NR": {
        "nimetus": "Looduslik kummi (NR)",
        "omadused": "Väga hea rebimis- ja venivuskindlus. Töötemp −50 … +80 °C.",
        "kasutus": "Viôratsioonikummid, dünaamilised tiivikud.",
    },
    "CR": {
        "nimetus": "Kloropreenkummi (CR / Neopreen)",
        "omadused": "Hea ilmastiku-, osooni- ja leegikindlus. Töötemp −40 … +120 °C.",
        "kasutus": "Mereveesüsteemid, kaablite isolatsioon, välistiivikud.",
    },
    "FKM": {
        "nimetus": "Fluorokummi (FKM / Viton)",
        "omadused": "Erakordselt kõrge keemia- ja kuumakindlus. Töötemp −20 … +200 °C.",
        "kasutus": "Keemiatööstus, automootorid, kõrgtemperatuur tiivikud.",
    },
    "SILICONE": {
        "nimetus": "Silikoonkummi",
        "omadused": "Väga lai temperatuurivahemik −60 … +230 °C. Töidukontaktiks sobiv.",
        "kasutus": "Töidutööstus, meditsiiniseadmed, ahjud ja köögiseadmed.",
    },
    "SILIKON": {
        "nimetus": "Silikoonkummi",
        "omadused": "Väga lai temperatuurivahemik −60 … +230 °C. Töidukontaktiks sobiv.",
        "kasutus": "Töidutööstus, meditsiiniseadmed, ahjud ja köögiseadmed.",
    },
    "NEOPRENE": {
        "nimetus": "Neopreen (CR)",
        "omadused": "Hea ilmastiku-, osooni- ja leegikindlus. Töötemp −40 … +120 °C.",
        "kasutus": "Mereveesüsteemid, kaablite isolatsioon, välistiivikud.",
    },
    "IIR": {
        "nimetus": "Butüülkummi (IIR)",
        "omadused": "Väga madal gaasiläbilaskvus, hea happekindlus. Töötemp −40 … +120 °C.",
        "kasutus": "Pneumaat. süsteemid, gaasikindlad tiivikud, akupatarei.",
    },
}

# ── FDA märge (eemaldatakse nimest, kuvatakse eraldi) ──────────────────
FDA_PATTERN = re.compile(r"\bFDA\b", re.IGNORECASE)

def eralda_fda(nimi: str):
    """Tagastab (puhas_nimi, on_fda)"""
    on_fda = bool(FDA_PATTERN.search(nimi))
    puhas = FDA_PATTERN.sub("", nimi).strip().strip(",").strip()
    puhas = re.sub(r"\s{2,}", " ", puhas)
    return puhas, on_fda

# ── Faili leidmine ────────────────────────────────────────────────────
def leia_uusim_xlsx(base_dir="outputs"):
    files = glob.glob(os.path.join(base_dir, "**", "*.xlsx"), recursive=True)
    return max(files, key=os.path.getmtime) if files else None

XLSX_PATH = leia_uusim_xlsx()

# ── Andmete laadimine ─────────────────────────────────────────────────
@st.cache_data
def laadi_tooted(path):
    df = pd.read_excel(path, sheet_name=0, dtype=str).fillna("")
    # Leia nimetus-veerg
    nimi_veerud = [c for c in df.columns if any(k in c.upper() for k in ["NAME", "NIMETUS", "NIMI"])]
    if nimi_veerud:
        vcol = nimi_veerud[0]
        df["_PUHAS_NIMI"] = df[vcol].apply(lambda x: eralda_fda(x)[0])
        df["FDA ⚠️"] = df[vcol].apply(lambda x: "✅ FDA" if eralda_fda(x)[1] else "")
    return df

@st.cache_data
def laadi_leht(path, sheet):
    df = pd.read_excel(path, sheet_name=sheet, dtype=str).fillna("")
    return df

@st.cache_data
def lehtede_nimed(path):
    return pd.ExcelFile(path).sheet_names

# ── Veeru otsija ──────────────────────────────────────────────────────
def leia_veerg(df, *candidates):
    for c in candidates:
        if c in df.columns:
            return c
    return None

# ── Pealeht ───────────────────────────────────────────────────────────
st.title("🧱 Zenith Materjali Assistent")
st.caption("Plastok OÜ — Zenith 2026 materjalibaas")

if not XLSX_PATH:
    st.error("❌ Exceli faili ei leitud kaustas 'outputs/'. Kontrolli, kas fail on üles laaditud.")
    st.stop()

try:
    df_raw = laadi_tooted(XLSX_PATH)
    sheet_names = lehtede_nimed(XLSX_PATH)
except Exception as e:
    st.error(f"Andmete laadimine ebaõnnestus: {e}")
    st.stop()

# Veergude tuvastamine
col_art  = leia_veerg(df_raw, "Article nr.", "ARTIKKEL", "ARTICLE_NR")
col_nimi = leia_veerg(df_raw, "Name", "NIMETUS", "NIMI", "NAME")
col_mat  = leia_veerg(df_raw, "Material", "MATERJAL", "MATERIAL")
col_grp  = leia_veerg(df_raw, "Group", "GRUPP", "GROUP")
col_kov  = leia_veerg(df_raw, "Hardness", "KÕVADUS", "HARDNESS")
col_pak  = leia_veerg(df_raw, "Thickness", "PAKSUS", "THICKNESS")
col_lai  = leia_veerg(df_raw, "Width", "LAIUS", "WIDTH")
col_pik  = leia_veerg(df_raw, "Length", "PIKKUS", "LENGTH")
col_tmin = leia_veerg(df_raw, "Min \u00b0C", "MIN_TEMP", "TEMP_MIN")
col_tmax = leia_veerg(df_raw, "Max \u00b0C", "MAX_TEMP", "TEMP_MAX")
col_nett = leia_veerg(df_raw, "Nett price, EUR", "NETT_HIND", "NETT_PRICE_EUR")
col_gros = leia_veerg(df_raw, "Gross price, EUR", "BRUTO_HIND", "GROSS_PRICE_EUR")
col_unit = leia_veerg(df_raw, "Unit", "Ü HIKM", "UNIT")
col_qty  = leia_veerg(df_raw, "Quantity", "KOGUS", "QUANTITY")

# ── Sidebar: filtrid ─────────────────────────────────────────────────
def unikaalsed(df, col):
    if not col:
        return []
    return sorted(set(v.strip() for v in df[col].unique() if v.strip()))

with st.sidebar:
    st.header("🔎 Otsi materjali")
    vaba = st.text_input("Vaba otsing", placeholder="nt. NBR, tihendriba, 3mm...")

    st.divider()
    st.subheader("🎛️ Filtrid")

    mat_valik = st.selectbox("Materjal", ["— kõik —"] + unikaalsed(df_raw, col_mat))
    grp_valik = st.selectbox("Grupp / tooteliik", ["— kõik —"] + unikaalsed(df_raw, col_grp))
    kov_valik = st.selectbox("Kõvadus (Shore A)", ["— kõik —"] + unikaalsed(df_raw, col_kov))

    st.markdown("**Temperatuur (°C)**")
    t_min = st.number_input("Min temp", value=-60, step=5)
    t_max = st.number_input("Max temp", value=250, step=5)
    kasuta_temp = st.checkbox("Rakenda temp filter")

    st.divider()
    st.subheader("🧮 Koguse arvutus")
    st.caption("Sisesta soovitud kogus tellimuse jaoks")
    tell_kogus = st.number_input("Rullide / tükkide arv", min_value=1, value=1, step=1)

    st.divider()
    with open(XLSX_PATH, "rb") as fh:
        st.download_button(
            "⬇️ Laadi kataloog alla",
            data=fh,
            file_name=os.path.basename(XLSX_PATH),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# ── Filtreerimine ─────────────────────────────────────────────────────
df = df_raw.copy()

if vaba.strip():
    mask = df.apply(lambda c: c.str.contains(vaba.strip(), case=False, na=False, regex=False)).any(axis=1)
    df = df[mask]

if mat_valik != "— kõik —" and col_mat:
    df = df[df[col_mat].str.strip() == mat_valik]

if grp_valik != "— kõik —" and col_grp:
    df = df[df[col_grp].str.strip() == grp_valik]

if kov_valik != "— kõik —" and col_kov:
    df = df[df[col_kov].str.strip() == kov_valik]

if kasuta_temp:
    if col_tmin:
        tmin_s = pd.to_numeric(df[col_tmin], errors="coerce")
        df = df[tmin_s.isna() | (tmin_s >= t_min)]
    if col_tmax:
        tmax_s = pd.to_numeric(df[col_tmax], errors="coerce")
        df = df[tmax_s.isna() | (tmax_s <= t_max)]

# ── Materjali infokaart ───────────────────────────────────────────────
if mat_valik != "— kõik —":
    info = MATERJALIDE_INFO = MATERJALI_INFO.get(mat_valik.upper())
    if info:
        with st.expander(f"📖 {info['nimetus']} — materjali kirjeldus", expanded=True):
            c1, c2 = st.columns(2)
            c1.markdown(f"**⚙️ Omadused**\n\n{info['omadused']}")
            c2.markdown(f"**🏭 Kasutuseesmärk**\n\n{info['kasutus']}")

# ── FDA hoiatus ───────────────────────────────────────────────────────
fda_arv = (df.get("FDA ⚠️", pd.Series(dtype=str)) != "").sum() if "FDA ⚠️" in df.columns else 0
if fda_arv > 0:
    st.warning(
        f"⚠️ {fda_arv} tootel on FDA sertifikaat. "
        "FDA-tooteid saab tellida ainult eraldi sertifikaaditaotlusega — "
        "märgi tellimusel selgelt, et FDA sertifikaat on vajalik."
    )

# ── Tulemuste tabel ───────────────────────────────────────────────────
st.subheader(f"📋 Leitud: {len(df)} toodet")

if df.empty:
    st.info("🔍 Tulemusi ei leitud. Proovi muuta filtreid või otsinguterminit.")
else:
    # Mòõtude arvutus m² ja koguhind
    disp = df.copy()

    # Puhasta nimi (FDA eemaldatud)
    if "_PUHAS_NIMI" in disp.columns and col_nimi:
        disp[col_nimi] = disp["_PUHAS_NIMI"]
    disp = disp.drop(columns=[c for c in ["_PUHAS_NIMI"] if c in disp.columns])

    # m² arvutus (rull: laius mm × pikkus m)
    if col_lai and col_pik:
        laius_m = pd.to_numeric(disp[col_lai], errors="coerce") / 1000  # mm -> m
        pikkus_m = pd.to_numeric(disp[col_pik], errors="coerce")          # eeldame et juba meetrites
        m2_rullis = laius_m * pikkus_m
        disp["m² rullis"] = m2_rullis.round(2).where(m2_rullis.notna(), "")
        disp["m² kokku"] = (m2_rullis * tell_kogus).round(2).where(m2_rullis.notna(), "")

    # Koguhind arvutus
    hind_col = col_nett or col_gros
    if hind_col and col_lai and col_pik:
        hind = pd.to_numeric(disp[hind_col], errors="coerce")
        koguhind = hind * (m2_rullis * tell_kogus)
        disp["Koguhind (€)"] = koguhind.round(2).where(koguhind.notna(), "")
    elif hind_col:
        hind = pd.to_numeric(disp[hind_col], errors="coerce")
        koguhind = hind * tell_kogus
        disp["Koguhind (€)"] = koguhind.round(2).where(koguhind.notna(), "")

    # Järjesta veerud loogiliselt
    eelis = [c for c in [
        col_art, col_nimi, "FDA ⚠️",
        col_mat, col_grp, col_kov,
        col_pak, col_lai, col_pik,
        "m² rullis", "m² kokku",
        col_unit, col_qty,
        col_nett, col_gros, "Koguhind (€)",
        col_tmin, col_tmax
    ] if c and c in disp.columns]
    muud = [c for c in disp.columns if c not in eelis]
    disp = disp[eelis + muud]

    st.dataframe(disp, use_container_width=True, height=580)

    # Tellimuse kokkuvõte
    if "Koguhind (€)" in disp.columns:
        kogusumma = pd.to_numeric(disp["Koguhind (€)"], errors="coerce").sum()
        m2_kokku_sum = pd.to_numeric(disp.get("m² kokku", pd.Series(dtype=str)), errors="coerce").sum()
        st.divider()
        k1, k2, k3 = st.columns(3)
        k1.metric("📦 Tooteid", f"{len(disp)} tk")
        if m2_kokku_sum > 0:
            k2.metric("📏 Kokku m²", f"{m2_kokku_sum:.2f} m²")
        k3.metric("💶 Ligikaudne koguhind", f"{kogusumma:.2f} €")

# ── Kõik lehed (lisavaade) ────────────────────────────────────────────
with st.expander("📂 Vaata kõiki andmelehtede sisu", expanded=False):
    leht_valik = st.selectbox("Leht", sheet_names, key="leht_valik")
    df_leht = laadi_leht(XLSX_PATH, leht_valik)
    st.caption(f"{len(df_leht)} rida · {len(df_leht.columns)} veergu")
    st.dataframe(df_leht, use_container_width=True, height=400)
