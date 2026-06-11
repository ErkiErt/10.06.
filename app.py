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

# ── Materjali kirjeldused ────────────────────────────────────────────
MATERJALI_INFO = {
    "NBR": {
        "nimetus": "Nitriilkummi (NBR)",
        "omadused": "Kõrge õli-, kütuse- ja rasvakindlus. Töötemperatuur −40…+120 °C.",
        "kasutus": "Hüdraulikatiivikud, õlitorud, kütusesüsteemid, tööstustiivikud.",
    },
    "EPDM": {
        "nimetus": "Etüleen-propüleenkummi (EPDM)",
        "omadused": "Suurepärane ilmastiku-, osooni- ja UV-kindlus. Töötemp −40…+150 °C.",
        "kasutus": "Katusetiivikud, aknaraamid, jõeveesüsteemid, autouksed.",
    },
    "SBR": {
        "nimetus": "Stüreen-butadieenkummi (SBR)",
        "omadused": "Hea kulumiskindlus ja mehaanilised omadused. Töötemp −40…+100 °C.",
        "kasutus": "Tiivikuplaadid, põrandakatted, lintkonveierid.",
    },
    "NR": {
        "nimetus": "Looduslik kummi (NR)",
        "omadused": "Väga hea rebimis- ja venivuskindlus. Töötemp −50…+80 °C.",
        "kasutus": "Vibratsioonikummid, dünaamilised tiivikud.",
    },
    "CR": {
        "nimetus": "Kloropreenkummi (CR / Neopreen)",
        "omadused": "Hea ilmastiku-, osooni- ja leegikindlus. Töötemp −40…+120 °C.",
        "kasutus": "Mereveesüsteemid, kaablite isolatsioon, välistiivikud.",
    },
    "FKM": {
        "nimetus": "Fluorokummi (FKM / Viton)",
        "omadused": "Erakordselt kõrge keemia- ja kuumakindlus. Töötemp −20…+200 °C.",
        "kasutus": "Keemiatööstus, automootorid, kõrgtemperatuur-tiivikud.",
    },
    "SILICONE": {
        "nimetus": "Silikoonkummi",
        "omadused": "Lai temperatuurivahemik −60…+230 °C. Töidukontaktiks sobiv.",
        "kasutus": "Töidutööstus, meditsiiniseadmed, ahjud ja köögiseadmed.",
    },
    "SILIKON": {
        "nimetus": "Silikoonkummi",
        "omadused": "Lai temperatuurivahemik −60…+230 °C. Töidukontaktiks sobiv.",
        "kasutus": "Töidutööstus, meditsiiniseadmed, ahjud ja köögiseadmed.",
    },
    "NEOPRENE": {
        "nimetus": "Neopreen (CR)",
        "omadused": "Hea ilmastiku-, osooni- ja leegikindlus. Töötemp −40…+120 °C.",
        "kasutus": "Mereveesüsteemid, kaablite isolatsioon, välistiivikud.",
    },
    "IIR": {
        "nimetus": "Butüülkummi (IIR)",
        "omadused": "Väga madal gaasiläbilaskvus, hea happekindlus. Töötemp −40…+120 °C.",
        "kasutus": "Pneumaatilised süsteemid, gaasikindlad tiivikud.",
    },
}

# ── FDA abifunktsioon ──────────────────────────────────────────────
FDA_PATTERN = re.compile(r"\bFDA\b", re.IGNORECASE)

def eralda_fda(nimi: str):
    on_fda = bool(FDA_PATTERN.search(nimi))
    puhas = FDA_PATTERN.sub("", nimi).strip().strip(",").strip()
    puhas = re.sub(r"\s{2,}", " ", puhas)
    return puhas, on_fda

# ── Veeru otsija ──────────────────────────────────────────────────
def leia_veerg(df, *candidates):
    for c in candidates:
        if c in df.columns:
            return c
    return None

# ── Automaatne lehtede tuvastus ja liitmine ─────────────────────────
# Toote-/hinnaandmeid sisaldavad lehed tuvastatakse võtmesõnade järgi.
# Kasutaja EI pea lehte valima.
TOODE_LEHED   = ["01_TOOTED", "TOOTED", "Products", "ALL", "Sheet1"]
HINNA_LEHED   = ["09_ZENITH_HINNAD", "HINNAD", "Prices", "Pricelist"]
ARTIKLI_LEHED = ["02_ARTIKLID_MOODUD", "ARTIKLID", "Articles"]

def leia_uusim_xlsx(base_dir="outputs"):
    files = glob.glob(os.path.join(base_dir, "**", "*.xlsx"), recursive=True)
    return max(files, key=os.path.getmtime) if files else None

@st.cache_data
def laadi_parim_leht(path, kandidaadid):
    """Laadib esimese leidu võtmesõnade nimekirjast; kui ühtegi ei leidu, laadib esimese lehe."""
    lehed = pd.ExcelFile(path).sheet_names
    for k in kandidaadid:
        for l in lehed:
            if k.lower() in l.lower():
                df = pd.read_excel(path, sheet_name=l, dtype=str).fillna("")
                if len(df) > 1:   # ignoreerime tühja lehe
                    return df, l
    # varuvariandina esimene leht
    df = pd.read_excel(path, sheet_name=0, dtype=str).fillna("")
    return df, lehed[0]

@st.cache_data
def laadi_koik_toodetelehed(path):
    """Liidab kõik lehed, mis sisaldavad artiklinumbrit ja nimetust (automaatne tuvastus)."""
    lehed = pd.ExcelFile(path).sheet_names
    tulemused = []
    for l in lehed:
        try:
            df = pd.read_excel(path, sheet_name=l, dtype=str).fillna("")
            cols_upper = [c.upper() for c in df.columns]
            # Leht loetakse tooteleheküljeks kui sisaldab nii artiklit kui nimetust
            on_art  = any(k in " ".join(cols_upper) for k in ["ARTICLE", "ARTIKKEL", "ARTIKLI"])
            on_nimi = any(k in " ".join(cols_upper) for k in ["NAME", "NIMETUS", "NIMI"])
            if on_art and on_nimi and len(df) > 2:
                df["_allikas_leht"] = l
                tulemused.append(df)
        except Exception:
            continue
    if not tulemused:
        # fallback: esimene leht
        df = pd.read_excel(path, sheet_name=0, dtype=str).fillna("")
        df["_allikas_leht"] = "leht1"
        return df
    return pd.concat(tulemused, ignore_index=True)

@st.cache_data
def lehtede_nimed(path):
    return pd.ExcelFile(path).sheet_names

# ── Laadimine ─────────────────────────────────────────────────────────
st.title("🧱 Zenith Materjali Assistent")
st.caption("Plastok OÜ — Zenith 2026 materjalibaas")

XLSX_PATH = leia_uusim_xlsx()
if not XLSX_PATH:
    st.error("❌ Exceli faili ei leitud kaustas ‘outputs/’. Kontrolli, kas fail on üles laaditud.")
    st.stop()

try:
    df_raw = laadi_koik_toodetelehed(XLSX_PATH)
    sheet_names = lehtede_nimed(XLSX_PATH)
except Exception as e:
    st.error(f"Andmete laadimine ebaõnnestus: {e}")
    st.stop()

# FDA töötlus
nimi_veerud = [c for c in df_raw.columns if any(k in c.upper() for k in ["NAME", "NIMETUS", "NIMI"])]
if nimi_veerud:
    vcol = nimi_veerud[0]
    df_raw["_PUHAS_NIMI"] = df_raw[vcol].apply(lambda x: eralda_fda(str(x))[0])
    df_raw["FDA ⚠️"]     = df_raw[vcol].apply(lambda x: "⚠️ FDA" if eralda_fda(str(x))[1] else "")

# Veergude tuvastamine
col_art  = leia_veerg(df_raw, "Article nr.", "ARTIKKEL", "ARTICLE_NR", "Article")
col_nimi = leia_veerg(df_raw, "Name", "NIMETUS", "NIMI", "NAME")
col_mat  = leia_veerg(df_raw, "Material", "MATERJAL", "MATERIAL")
col_grp  = leia_veerg(df_raw, "Group", "GRUPP", "GROUP")
col_kov  = leia_veerg(df_raw, "Hardness", "KÕVADUS", "HARDNESS")
col_pak  = leia_veerg(df_raw, "Thickness", "PAKSUS", "THICKNESS")
col_lai  = leia_veerg(df_raw, "Width", "LAIUS", "WIDTH")
col_pik  = leia_veerg(df_raw, "Length", "PIKKUS", "LENGTH")
col_tmin = leia_veerg(df_raw, "Min °C", "MIN_TEMP", "TEMP_MIN")
col_tmax = leia_veerg(df_raw, "Max °C", "MAX_TEMP", "TEMP_MAX")
col_nett = leia_veerg(df_raw, "Nett price, EUR", "NETT_HIND", "NETT_PRICE_EUR")
col_gros = leia_veerg(df_raw, "Gross price, EUR", "BRUTO_HIND", "GROSS_PRICE_EUR")
col_unit = leia_veerg(df_raw, "Unit", "Ühik", "UNIT")
col_qty  = leia_veerg(df_raw, "Quantity", "KOGUS", "QUANTITY")

# ── Sidebar ──────────────────────────────────────────────────────────
def unikaalsed(df, col):
    if not col or col not in df.columns:
        return []
    return sorted(set(v.strip() for v in df[col].unique() if v.strip() and v.strip() != "nan"))

with st.sidebar:
    st.header("🔎 Otsi materjali")
    vaba = st.text_input("Vaba otsing", placeholder="nt. NBR, tihendriba, 3 mm…")

    st.divider()
    st.subheader("🎛️ Filtrid")
    mat_valik = st.selectbox("Materjal",       ["— kõik —"] + unikaalsed(df_raw, col_mat))
    grp_valik = st.selectbox("Grupp",          ["— kõik —"] + unikaalsed(df_raw, col_grp))
    kov_valik = st.selectbox("Kõvadus Shore A",["— kõik —"] + unikaalsed(df_raw, col_kov))

    st.markdown("**Temperatuur (°C)**")
    t_min = st.number_input("Min °C", value=-60, step=5)
    t_max = st.number_input("Max °C", value=250, step=5)
    kasuta_temp = st.checkbox("Rakenda temperatuurifilter")

    st.divider()
    st.subheader("🧮 Koguse arvutus")
    tell_kogus = st.number_input("Rullide / tükkide arv", min_value=1, value=1, step=1)

    st.divider()
    with open(XLSX_PATH, "rb") as fh:
        st.download_button(
            "⬇️ Laadi kataloog alla",
            data=fh,
            file_name=os.path.basename(XLSX_PATH),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# ── Filtreerimine ───────────────────────────────────────────────────
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
        s = pd.to_numeric(df[col_tmin], errors="coerce")
        df = df[s.isna() | (s >= t_min)]
    if col_tmax:
        s = pd.to_numeric(df[col_tmax], errors="coerce")
        df = df[s.isna() | (s <= t_max)]

# ── Materjali infokaart ─────────────────────────────────────────────
if mat_valik != "— kõik —":
    info = MATERJALI_INFO.get(mat_valik.upper())
    if info:
        with st.expander(f"📖 {info['nimetus']} — materjali kirjeldus", expanded=True):
            c1, c2 = st.columns(2)
            c1.markdown(f"**⚙️ Omadused**\n\n{info['omadused']}")
            c2.markdown(f"**🏭 Kasutuseesmärk**\n\n{info['kasutus']}")

# ── FDA hoiatus ────────────────────────────────────────────────────
fda_arv = (df.get("FDA ⚠️", pd.Series(dtype=str)).str.strip() != "").sum()
if fda_arv > 0:
    st.warning(
        f"⚠️ {fda_arv} tootel on FDA-sertifikaat. "
        "FDA-tooteid saab tellida ainult koos eraldi sertifikaaditaotlusega — "
        "märgi tellimusel selgelt, et FDA-sertifikaat on vajalik."
    )

# ── Tulemuste tabel ──────────────────────────────────────────────────
st.subheader(f"📋 Leitud: {len(df)} toodet")

if df.empty:
    st.info("🔍 Tulemusi ei leitud. Proovi muuta filtreid või otsinguterminit.")
else:
    disp = df.copy()

    # Puhasta nimi (FDA eemaldatud)
    if "_PUHAS_NIMI" in disp.columns and col_nimi and col_nimi in disp.columns:
        disp[col_nimi] = disp["_PUHAS_NIMI"]
    disp.drop(columns=[c for c in ["_PUHAS_NIMI", "_allikas_leht"] if c in disp.columns], inplace=True)

    # m² arvutus (rull: laius mm ÷ 1000 × pikkus meetrites)
    m2_rullis = None
    if col_lai and col_pik and col_lai in disp.columns and col_pik in disp.columns:
        laius_m  = pd.to_numeric(disp[col_lai], errors="coerce") / 1000
        pikkus_m = pd.to_numeric(disp[col_pik], errors="coerce")
        m2_rullis = laius_m * pikkus_m
        disp["m² / rull"]  = m2_rullis.round(3).where(m2_rullis.notna(), other="")
        disp["m² kokku"]   = (m2_rullis * tell_kogus).round(3).where(m2_rullis.notna(), other="")

    # Hind üle m² või üle tk/rulli
    hind_col = col_nett if col_nett and col_nett in disp.columns else (
               col_gros if col_gros and col_gros in disp.columns else None)
    if hind_col:
        hind = pd.to_numeric(disp[hind_col], errors="coerce")
        if m2_rullis is not None:
            koguhind = hind * (m2_rullis * tell_kogus)
        else:
            koguhind = hind * tell_kogus
        disp["Koguhind (€)"] = koguhind.round(2).where(koguhind.notna(), other="")

    # Veergude järjestus
    eelis = [c for c in [
        col_art, col_nimi, "FDA ⚠️",
        col_mat, col_grp, col_kov,
        col_pak, col_lai, col_pik,
        "m² / rull", "m² kokku",
        col_unit, col_qty,
        col_nett, col_gros, "Koguhind (€)",
        col_tmin, col_tmax,
    ] if c and c in disp.columns]
    muud = [c for c in disp.columns if c not in eelis]
    disp = disp[eelis + muud]

    st.dataframe(disp, use_container_width=True, height=560)

    # Tellimuse kokkuvõte
    st.divider()
    k1, k2, k3 = st.columns(3)
    k1.metric("📦 Tooteid", f"{len(disp)}")
    if "m² kokku" in disp.columns:
        m2_summa = pd.to_numeric(disp["m² kokku"], errors="coerce").sum()
        k2.metric("📏 Kokku m²", f"{m2_summa:.2f} m²" if m2_summa > 0 else "—")
    if "Koguhind (€)" in disp.columns:
        summa = pd.to_numeric(disp["Koguhind (€)"], errors="coerce").sum()
        k3.metric("💶 Ligikaudne koguhind", f"{summa:,.2f} €" if summa > 0 else "—")

# ── Lisavaade: kõik lehed (täpsemaks uurimiseks) ───────────────────
with st.expander("📂 Täpsem andmevaade (vali leht käsitsi)", expanded=False):
    leht_valik = st.selectbox("Leht", sheet_names, key="leht_valik")
    df_leht = pd.read_excel(XLSX_PATH, sheet_name=leht_valik, dtype=str).fillna("")
    st.caption(f"{len(df_leht)} rida · {len(df_leht.columns)} veergu · leht: {leht_valik}")
    st.dataframe(df_leht, use_container_width=True, height=400)
