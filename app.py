import re
from datetime import datetime, timedelta
from io import BytesIO

import pandas as pd
import streamlit as st
from rapidfuzz import process, fuzz
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.comments import Comment

st.set_page_config(page_title="FBO Hours Control tool", layout="wide")

st.markdown("""
<style>

/* KPI kaarten */
.kpi-card {
    background-color: #f5f5f5;
    border-radius: 20px;
    padding: 25px;
    color: #145a32;  /* 🔥 DONKER GROEN voor tekst */
}

/* Titel in kaart */
.kpi-title {
    color: #666666;  /* grijs i.p.v. wit */
    font-size: 14px;
}

/* Grote cijfers */
.kpi-value {
    color: #145a32;  /* donker groen */
    font-size: 48px;
    font-weight: bold;
}

</style>
""", unsafe_allow_html=True)
    .block-container { padding-top: 1.6rem; padding-bottom: 2rem; max-width: 1400px; }
    .fbo-hero {
        background: linear-gradient(135deg, #0B5D3B 0%, #0F7A4F 60%, #18A866 100%);
        border-radius: 24px; padding: 30px 34px; color: white;
        box-shadow: 0 18px 45px rgba(11, 93, 59, 0.22); margin-bottom: 22px;
    }
    .fbo-kicker { text-transform: uppercase; letter-spacing: 0.14em; font-size: 0.78rem; opacity: 0.86; font-weight: 700; margin-bottom: 6px; }
    .fbo-title { font-size: 2.45rem; line-height: 1.05; font-weight: 900; margin: 0; }
    .fbo-subtitle { font-size: 1rem; margin-top: 10px; max-width: 820px; opacity: 0.92; }
    .fbo-card { background: white; border: 1px solid var(--fbo-border); border-radius: 20px; padding: 18px 20px; box-shadow: 0 10px 28px rgba(16, 37, 27, 0.06); margin-bottom: 14px; }
    .fbo-section-title { color: var(--fbo-green); font-size: 1.15rem; font-weight: 850; margin: 4px 0 10px 0; }
    div[data-testid="stMetric"] { background: white; border: 1px solid var(--fbo-border); border-radius: 18px; padding: 16px 18px; box-shadow: 0 10px 26px rgba(16, 37, 27, 0.06); }
    div[data-testid="stMetricLabel"] p { color: #547565; font-weight: 700; }
    div[data-testid="stMetricValue"] { color: var(--fbo-green); font-weight: 900; }
    .stTabs [data-baseweb="tab-list"] { gap: 8px; background: var(--fbo-soft); border: 1px solid var(--fbo-border); padding: 7px; border-radius: 18px; }
    .stTabs [data-baseweb="tab"] { border-radius: 14px; padding: 10px 16px; color: #1c4632; font-weight: 750; }
    .stTabs [aria-selected="true"] { background: var(--fbo-green) !important; color: green !important; }
    .stDownloadButton button, .stButton button { background: var(--fbo-green); color: white; border: 0; border-radius: 14px; padding: 0.75rem 1rem; font-weight: 800; box-shadow: 0 10px 22px rgba(11, 93, 59, 0.16); }
    .stDownloadButton button:hover, .stButton button:hover { background: var(--fbo-green-2); color: white; border: 0; }
</style>
<div class="fbo-hero">
    <div class="fbo-kicker">Planning compliance dashboard</div>
    <h1 class="fbo-title">FBO Hours Control tool</h1>
    <div class="fbo-subtitle">Controleer Strobbo-weekroosters op contracturen, pauzes, rusttijd en minderjarigenregels. Fouten worden rechtstreeks gemarkeerd in de Excel-export.</div>
</div>
""", unsafe_allow_html=True)

MIN_DAGUREN = 2
MAX_DAGUREN_VOLWASSEN = 11
MAX_DAGUREN_MINDERJARIG = 8
MAX_WEEKUREN_VOLWASSEN = 50
MAX_WEEKUREN_MINDERJARIG = 40
MIN_RUSTUREN_TUSSEN_SHIFTS = 10
FUZZY_MATCH_SCORE = 65
MERGE_GAP_MINUTEN = 5

FILL_FOUT = PatternFill("solid", fgColor="FFFF00")
FONT_ROOD = Font(color="FF0000", bold=True)


def normaliseer_tekst(x):
    if pd.isna(x):
        return ""
    return str(x).strip()


def strip_strobbo_tags(naam):
    naam = str(naam)
    naam = naam.replace("#", "")
    naam = re.sub(r"<\s*\d+", "", naam)
    naam = re.sub(r"\b(MGR|FLX|FLEXI|STUDENT)\b", "", naam, flags=re.IGNORECASE)
    naam = naam.replace(".", "")
    naam = naam.replace("-", " ")
    naam = re.sub(r"\s+", " ", naam)
    return naam.strip()


def normaliseer_naam(naam):
    naam = strip_strobbo_tags(naam)
    naam = naam.lower().strip()
    naam = re.sub(r"\s+", " ", naam)
    return naam


def veilige_float(x):
    try:
        if pd.isna(x):
            return 0.0
        return float(str(x).replace(",", "."))
    except Exception:
        return 0.0


def veilige_int(x):
    try:
        if pd.isna(x):
            return None
        leeftijd = int(float(str(x).replace(",", ".")))
        if leeftijd <= 0 or leeftijd > 100:
            return None
        return leeftijd
    except Exception:
        return None


def maak_volledige_naam(row):
    achternaam = normaliseer_tekst(row.get("NAAM", ""))
    voornaam = normaliseer_tekst(row.get("VOORNAAM", ""))
    if achternaam and voornaam:
        return f"{voornaam} {achternaam}"
    if voornaam:
        return voornaam
    if achternaam:
        return achternaam
    return ""


def bepaal_type(row):
    volledige_rij = " ".join([str(v).upper() for v in row.values])
    if "STUDENT" in volledige_rij:
        return "student"
    if "FLEXI" in volledige_rij or "FLX" in volledige_rij:
        return "flexi"
    contract = row.get("CONTRACT. UREN", row.get("CONTRACT UREN", ""))
    if pd.isna(contract) or str(contract).strip() == "":
        return "student"
    return "vast"


def zoek_beste_match(strobbo_naam, crew_df):
    naam = normaliseer_naam(strobbo_naam)
    if not naam:
        return None, 0

    exacte_matches = []
    for _, row in crew_df.iterrows():
        voornaam = normaliseer_naam(row["VOORNAAM"])
        if naam == voornaam:
            exacte_matches.append(row["VOLLEDIGE_NAAM"])

    if len(exacte_matches) == 1:
        return exacte_matches[0], 100

    match_initiaal = re.match(r"^([a-zA-ZÀ-ÿ]+)\s+([a-zA-Z])$", naam)
    if match_initiaal:
        voornaam_gezocht = match_initiaal.group(1)
        initiaal_gezocht = match_initiaal.group(2)
        mogelijke = []
        for _, row in crew_df.iterrows():
            voornaam = normaliseer_naam(row["VOORNAAM"])
            achternaam = normaliseer_naam(row["NAAM"])
            if voornaam == voornaam_gezocht and achternaam.startswith(initiaal_gezocht):
                mogelijke.append(row["VOLLEDIGE_NAAM"])
        if len(mogelijke) == 1:
            return mogelijke[0], 100

    voornamen = crew_df["VOORNAAM"].astype(str).tolist()
    voornamen_norm = [normaliseer_naam(v) for v in voornamen]
    match = process.extractOne(naam, voornamen_norm, scorer=fuzz.ratio)
    if match:
        _, score, index = match
        if score >= 90:
            return crew_df.iloc[index]["VOLLEDIGE_NAAM"], round(score, 2)

    crew_namen = crew_df["VOLLEDIGE_NAAM"].tolist()
    crew_norms = [normaliseer_naam(n) for n in crew_namen]
    match = process.extractOne(naam, crew_norms, scorer=fuzz.token_sort_ratio)
    if not match:
        return None, 0
    _, score, index = match
    if score >= FUZZY_MATCH_SCORE:
        return crew_namen[index], round(score, 2)
    return None, round(score, 2)


def parse_datum(waarde):
    if isinstance(waarde, datetime):
        return waarde.date()
    tekst = str(waarde).strip().lower()
    maand_map = {
        "jan": 1, "feb": 2, "mrt": 3, "mar": 3, "apr": 4,
        "mei": 5, "jun": 6, "jul": 7, "aug": 8,
        "sep": 9, "okt": 10, "oct": 10, "nov": 11, "dec": 12,
    }
    match = re.search(r"(\d{1,2})[-/\s]([a-zA-ZÀ-ÿ]+)", tekst)
    if match:
        dag = int(match.group(1))
        maand_txt = match.group(2)[:3]
        maand = maand_map.get(maand_txt)
        if maand:
            return datetime(2026, maand, dag).date()
    return None


def vind_dag_kolommen(raw):
    dag_kolommen = {}
    for rij in range(min(10, len(raw))):
        for col in range(raw.shape[1]):
            datum = parse_datum(raw.iloc[rij, col])
            if datum:
                dag_kolommen[col] = datum
        if len(dag_kolommen) >= 5:
            return dag_kolommen
    return dag_kolommen


def vind_totaal_kolom(raw):
    for r in range(min(10, len(raw))):
        for c in range(raw.shape[1]):
            if str(raw.iloc[r, c]).strip().lower() == "totaal":
                return c
    return None


def parse_uren_uit_totaal(waarde):
    tekst = str(waarde)
    match = re.search(r"(\d{1,3}):(\d{2})", tekst)
    if not match:
        return None
    uren = int(match.group(1))
    minuten = int(match.group(2))
    return uren + minuten / 60


def parse_pauze_minuten(pauze_txt):
    if not pauze_txt:
        return 0
    match = re.search(r"(\d{1,2}):(\d{2})", str(pauze_txt))
    if not match:
        return 0
    return int(match.group(1)) * 60 + int(match.group(2))


def parse_shiftblokken(cell_text, datum):
    if pd.isna(cell_text) or not datum:
        return []
    tekst = str(cell_text)
    patroon = re.compile(
        r"(\d{1,2}:\d{2})\s*-\s*(\d{1,2}:\d{2})\s*\*?\s*(?:\n|\r|\s)*\(?\s*(\d{1,2}:\d{2})?\s*\)?",
        re.MULTILINE,
    )
    blokken = []
    for match in patroon.finditer(tekst):
        start_txt = match.group(1)
        einde_txt = match.group(2)
        pauze_txt = match.group(3)
        start_dt = datetime.combine(datum, datetime.strptime(start_txt, "%H:%M").time())
        einde_dt = datetime.combine(datum, datetime.strptime(einde_txt, "%H:%M").time())
        if einde_dt <= start_dt:
            einde_dt += timedelta(days=1)
        pauze_minuten = parse_pauze_minuten(pauze_txt)
        blokken.append({
            "datum": datum,
            "start": start_dt,
            "einde": einde_dt,
            "pauze_minuten": pauze_minuten,
            "origineel": tekst,
        })
    if not blokken:
        return []
    blokken = sorted(blokken, key=lambda x: x["start"])
    samengevoegd = []
    huidige = blokken[0].copy()
    for blok in blokken[1:]:
        gap_minuten = (blok["start"] - huidige["einde"]).total_seconds() / 60
        if gap_minuten <= MERGE_GAP_MINUTEN:
            huidige["einde"] = max(huidige["einde"], blok["einde"])
            huidige["pauze_minuten"] += blok["pauze_minuten"]
            huidige["origineel"] = str(huidige.get("origineel", "")) + "\n---\n" + str(blok.get("origineel", ""))
        else:
            samengevoegd.append(huidige)
            huidige = blok.copy()
    samengevoegd.append(huidige)
    for shift in samengevoegd:
        bruto_uren = (shift["einde"] - shift["start"]).total_seconds() / 3600
        shift["bruto_uren"] = bruto_uren
        shift["netto_uren"] = bruto_uren - (shift["pauze_minuten"] / 60)
    return samengevoegd


def voeg_fout(fouten, naam, datum, fouttype, detail, ernst="Fout", bron_cellen=None):
    fouten.append({
        "Medewerker": naam,
        "Datum": datum,
        "Ernst": ernst,
        "Fout": fouttype,
        "Detail": detail,
        "bron_cellen": bron_cellen or [],
    })


def is_mogelijke_naamcel(eerste_cel):
    if not eerste_cel:
        return False
    bevat_tijd = bool(re.search(r"\d{1,2}:\d{2}", eerste_cel))
    is_header = eerste_cel.lower() in ["wk21", "wk22", "wk23", "ma", "di", "wo", "do", "vr", "za", "zo", "totaal"]
    return not bevat_tijd and not is_header


st.markdown('<div class="fbo-card"><div class="fbo-section-title">Bestanden uploaden</div>', unsafe_allow_html=True)
up_col1, up_col2 = st.columns(2)
with up_col1:
    crew_file = st.file_uploader("👥 Crew-database Excel", type=["xlsx"], help="Database met naam, voornaam, leeftijd en contracturen.")
with up_col2:
    rooster_file = st.file_uploader("📅 Strobbo weekrooster Excel", type=["xlsx"], help="Export van Strobbo voor één week.")
st.markdown('</div>', unsafe_allow_html=True)

if not crew_file or not rooster_file:
    st.info("Upload beide Excel-bestanden om de controle te starten.")
    st.stop()

crew_bytes = crew_file.getvalue()
rooster_bytes = rooster_file.getvalue()

try:
    crew = pd.read_excel(BytesIO(crew_bytes))
except Exception as e:
    st.error(f"Kon crew-database niet lezen: {e}")
    st.stop()

crew.columns = [str(c).strip().upper() for c in crew.columns]
if "CONTRACT. UREN" not in crew.columns and "CONTRACT UREN" in crew.columns:
    crew["CONTRACT. UREN"] = crew["CONTRACT UREN"]

vereiste_kolommen = ["NAAM", "VOORNAAM", "LFTD"]
ontbrekend = [c for c in vereiste_kolommen if c not in crew.columns]
if ontbrekend:
    st.error(f"Deze kolommen ontbreken in je crew-database: {ontbrekend}")
    st.stop()
if "CONTRACT. UREN" not in crew.columns:
    crew["CONTRACT. UREN"] = ""

crew["VOLLEDIGE_NAAM"] = crew.apply(maak_volledige_naam, axis=1)
crew["TYPE"] = crew.apply(bepaal_type, axis=1)
crew["LEEFTIJD"] = crew["LFTD"].apply(veilige_int)
crew["CONTRACTUREN"] = crew["CONTRACT. UREN"].apply(veilige_float)
crew = crew[crew["VOLLEDIGE_NAAM"].str.strip() != ""].copy()

try:
    raw = pd.read_excel(BytesIO(rooster_bytes), sheet_name=0, header=None)
    wb = load_workbook(BytesIO(rooster_bytes))
    ws = wb.active
except Exception as e:
    st.error(f"Kon Strobbo-rooster niet lezen: {e}")
    st.stop()

dag_kolommen = vind_dag_kolommen(raw)
if not dag_kolommen:
    st.error("Ik kon geen datumkolommen vinden in het Strobbo-bestand.")
    st.stop()

totaal_col = vind_totaal_kolom(raw)
if totaal_col is None:
    totaal_col = max(dag_kolommen.keys()) + 1

# =========================
# SHIFTS EXTRACTEN
# =========================

shifts = []
huidige_medewerker = None

for rij_index, row in raw.iterrows():
    eerste_cel = normaliseer_tekst(row.iloc[0])
    if is_mogelijke_naamcel(eerste_cel):
        huidige_medewerker = eerste_cel.replace("#", "").strip()
    if not huidige_medewerker:
        continue
    for col_index, datum in dag_kolommen.items():
        if col_index >= len(row):
            continue
        cel = row.iloc[col_index]
        gevonden_shifts = parse_shiftblokken(cel, datum)
        for shift in gevonden_shifts:
            shift["strobbo_naam"] = huidige_medewerker
            shift["bron_rij"] = rij_index + 1
            shift["bron_kolom"] = col_index + 1
            shifts.append(shift)

shifts_df = pd.DataFrame(shifts)
if shifts_df.empty:
    st.warning("Geen shifts gevonden in het Strobbo-bestand.")
    st.stop()

# =========================
# SHIFTS SAMENVOEGEN OVER RIJEN
# =========================

shifts_df = shifts_df.sort_values(["strobbo_naam", "datum", "start", "einde"]).reset_index(drop=True)
samengevoegde_shifts = []
for (naam, datum), groep in shifts_df.groupby(["strobbo_naam", "datum"]):
    groep = groep.sort_values("start").reset_index(drop=True)
    huidige = groep.iloc[0].to_dict()
    for i in range(1, len(groep)):
        blok = groep.iloc[i].to_dict()
        gap_minuten = (blok["start"] - huidige["einde"]).total_seconds() / 60
        if gap_minuten <= MERGE_GAP_MINUTEN:
            huidige["einde"] = max(huidige["einde"], blok["einde"])
            huidige["pauze_minuten"] += blok["pauze_minuten"]
            huidige["origineel"] = str(huidige.get("origineel", "")) + "\n---\n" + str(blok.get("origineel", ""))
            huidige["bron_rij"] = f"{huidige.get('bron_rij', '')}, {blok.get('bron_rij', '')}"
            huidige["bron_kolom"] = str(huidige.get("bron_kolom", ""))
        else:
            bruto_uren = (huidige["einde"] - huidige["start"]).total_seconds() / 3600
            huidige["bruto_uren"] = bruto_uren
            huidige["netto_uren"] = bruto_uren - (huidige["pauze_minuten"] / 60)
            samengevoegde_shifts.append(huidige)
            huidige = blok
    bruto_uren = (huidige["einde"] - huidige["start"]).total_seconds() / 3600
    huidige["bruto_uren"] = bruto_uren
    huidige["netto_uren"] = bruto_uren - (huidige["pauze_minuten"] / 60)
    samengevoegde_shifts.append(huidige)
shifts_df = pd.DataFrame(samengevoegde_shifts)

# =========================
# STROBBO TOTALEN LEZEN
# =========================

strobbo_totalen = {}
huidige_medewerker = None
for rij_index, row in raw.iterrows():
    eerste_cel = normaliseer_tekst(row.iloc[0])
    if is_mogelijke_naamcel(eerste_cel):
        huidige_medewerker = eerste_cel.replace("#", "").strip()
    if not huidige_medewerker:
        continue
    totaal_uren = parse_uren_uit_totaal(row.iloc[totaal_col])
    if totaal_uren is not None:
        strobbo_totalen[normaliseer_naam(huidige_medewerker)] = totaal_uren

# =========================
# MATCHEN MET DATABASE
# =========================

match_resultaten = []
for naam in shifts_df["strobbo_naam"].unique():
    beste_naam, score = zoek_beste_match(naam, crew)
    match_resultaten.append({
        "Strobbo naam": naam,
        "Database naam": beste_naam if beste_naam else "NIET GEVONDEN",
        "Match score": round(score, 2),
    })
match_df = pd.DataFrame(match_resultaten)
match_map = {row["Strobbo naam"]: row["Database naam"] for _, row in match_df.iterrows() if row["Database naam"] != "NIET GEVONDEN"}
shifts_df["database_naam"] = shifts_df["strobbo_naam"].map(match_map)

# =========================
# FOUTEN CONTROLEREN
# =========================

fouten = []
for _, row in match_df.iterrows():
    if row["Database naam"] == "NIET GEVONDEN":
        voeg_fout(fouten, row["Strobbo naam"], "", "Naam niet gevonden", f"Geen goede match in crew-database. Match score: {row['Match score']}", "Waarschuwing")

for _, shift in shifts_df.dropna(subset=["database_naam"]).iterrows():
    naam = shift["database_naam"]
    datum = shift["datum"]
    persoon = crew[crew["VOLLEDIGE_NAAM"] == naam].iloc[0]
    leeftijd = persoon["LEEFTIJD"]
    netto_uren = shift["netto_uren"]
    bruto_uren = shift["bruto_uren"]
    pauze = shift["pauze_minuten"]
    eerste_rij = str(shift["bron_rij"]).split(",")[0].strip()
    bron_cellen = []
    if eerste_rij.isdigit() and str(shift["bron_kolom"]).isdigit():
        bron_cellen = [(int(eerste_rij), int(shift["bron_kolom"]))]

    if netto_uren < MIN_DAGUREN:
        voeg_fout(fouten, naam, datum, "Shift te kort", f"{netto_uren:.2f}u gewerkt. Minimum is {MIN_DAGUREN}u.", bron_cellen=bron_cellen)

    if leeftijd is not None and leeftijd < 18:
        if netto_uren > MAX_DAGUREN_MINDERJARIG:
            voeg_fout(fouten, naam, datum, "Minderjarige werkt te lang", f"{netto_uren:.2f}u gewerkt. Maximum voor <18 is {MAX_DAGUREN_MINDERJARIG}u.", bron_cellen=bron_cellen)
    else:
        if netto_uren > MAX_DAGUREN_VOLWASSEN:
            voeg_fout(fouten, naam, datum, "Shift te lang", f"{netto_uren:.2f}u gewerkt. Maximum is {MAX_DAGUREN_VOLWASSEN}u.", bron_cellen=bron_cellen)

    if leeftijd is None or leeftijd >= 18:
        if bruto_uren <= 5 and pauze > 0:
            voeg_fout(fouten, naam, datum, "Onnodige pauze", f"{bruto_uren:.2f}u shift heeft {pauze} min pauze, maar tot 5u is geen pauze nodig.", "Waarschuwing", bron_cellen)
        if bruto_uren > 5 and bruto_uren <= 8 and pauze < 20:
            voeg_fout(fouten, naam, datum, "Pauze ontbreekt", f"{bruto_uren:.2f}u shift. Minstens 20 min pauze nodig. Geplande pauze: {pauze} min.", bron_cellen=bron_cellen)
        if bruto_uren > 8 and pauze < 30:
            voeg_fout(fouten, naam, datum, "Pauze ontbreekt", f"{bruto_uren:.2f}u shift. Minstens 30 min pauze nodig. Geplande pauze: {pauze} min.", bron_cellen=bron_cellen)

    if leeftijd is not None and leeftijd < 18:
        if bruto_uren > 4.5 and bruto_uren <= 6 and pauze < 30:
            voeg_fout(fouten, naam, datum, "Pauze minderjarige ontbreekt", f"{bruto_uren:.2f}u shift. Minderjarige heeft minstens 30 min pauze nodig. Geplande pauze: {pauze} min.", bron_cellen=bron_cellen)
        if bruto_uren > 6 and pauze < 60:
            voeg_fout(fouten, naam, datum, "Pauze minderjarige ontbreekt", f"{bruto_uren:.2f}u shift. Minderjarige heeft minstens 60 min pauze nodig. Geplande pauze: {pauze} min.", bron_cellen=bron_cellen)

    if leeftijd is not None:
        einduur = shift["einde"].hour + shift["einde"].minute / 60
        if leeftijd <= 15 and einduur > 20:
            voeg_fout(fouten, naam, datum, "Nachtwerk minderjarige", f"{leeftijd} jaar en werkt tot {shift['einde'].strftime('%H:%M')}. Max tot 20:00.", bron_cellen=bron_cellen)
        if 16 <= leeftijd < 18 and einduur > 23:
            voeg_fout(fouten, naam, datum, "Nachtwerk minderjarige", f"{leeftijd} jaar en werkt tot {shift['einde'].strftime('%H:%M')}. Max tot 23:00.", bron_cellen=bron_cellen)

for naam, groep in shifts_df.dropna(subset=["database_naam"]).groupby("database_naam"):
    persoon = crew[crew["VOLLEDIGE_NAAM"] == naam].iloc[0]
    leeftijd = persoon["LEEFTIJD"]
    medewerker_type = persoon["TYPE"]
    contracturen = persoon["CONTRACTUREN"]
    strobbo_naam = groep.iloc[0]["strobbo_naam"]
    totaal_weekuren = strobbo_totalen.get(normaliseer_naam(strobbo_naam), None)
    if totaal_weekuren is None:
        totaal_weekuren = groep["netto_uren"].sum()

    if leeftijd is not None and leeftijd < 18:
        if totaal_weekuren > MAX_WEEKUREN_MINDERJARIG:
            voeg_fout(fouten, naam, "", "Weekuren overschreden", f"{totaal_weekuren:.2f}u gewerkt volgens Strobbo. Maximum voor <18 is {MAX_WEEKUREN_MINDERJARIG}u.")
    else:
        if totaal_weekuren > MAX_WEEKUREN_VOLWASSEN:
            voeg_fout(fouten, naam, "", "Weekuren overschreden", f"{totaal_weekuren:.2f}u gewerkt volgens Strobbo. Maximum is {MAX_WEEKUREN_VOLWASSEN}u.")

    if medewerker_type == "vast" and contracturen > 0:
        if totaal_weekuren < contracturen:
            tekort = contracturen - totaal_weekuren
            voeg_fout(fouten, naam, "", "Contracturen niet gehaald", f"{totaal_weekuren:.2f}u gepland volgens Strobbo, contract is {contracturen:.2f}u. Tekort: {tekort:.2f}u.")

for naam, groep in shifts_df.dropna(subset=["database_naam"]).groupby("database_naam"):
    groep = groep.sort_values("start").reset_index(drop=True)
    for i in range(1, len(groep)):
        vorige = groep.iloc[i - 1]
        huidige = groep.iloc[i]
        rusturen = (huidige["start"] - vorige["einde"]).total_seconds() / 3600
        if rusturen < MIN_RUSTUREN_TUSSEN_SHIFTS:
            voeg_fout(fouten, naam, huidige["datum"], "Te weinig rust", f"Slechts {rusturen:.2f}u rust tussen {vorige['einde'].strftime('%d/%m %H:%M')} en {huidige['start'].strftime('%d/%m %H:%M')}.")

fouten_df = pd.DataFrame(fouten)

if not fouten_df.empty and "bron_cellen" in fouten_df.columns:
    for _, fout in fouten_df.iterrows():
        for rr, cc in fout.get("bron_cellen", []):
            try:
                cel = ws.cell(row=rr, column=cc)
                cel.fill = FILL_FOUT
                cel.font = FONT_ROOD
                cel.comment = Comment(str(fout["Detail"]), "Strobbo Checker")
            except Exception:
                pass

annotated_buffer = BytesIO()
wb.save(annotated_buffer)


# =========================
# OUTPUT DESIGN
# =========================

# KPI's
fout_count = 0 if fouten_df.empty else int((fouten_df["Ernst"] == "Fout").sum())
waarschuwing_count = 0 if fouten_df.empty else int((fouten_df["Ernst"] == "Waarschuwing").sum())
shift_count = len(shifts_df)
medewerker_count = shifts_df["database_naam"].dropna().nunique()
match_fail_count = int((match_df["Database naam"] == "NIET GEVONDEN").sum()) if not match_df.empty else 0

k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("Fouten", fout_count)
k2.metric("Waarschuwingen", waarschuwing_count)
k3.metric("Shifts", shift_count)
k4.metric("Medewerkers", medewerker_count)
k5.metric("Naam-mismatch", match_fail_count)

# Tabellen voorbereiden
toon_shifts = shifts_df.copy()
toon_shifts["Start"] = toon_shifts["start"].dt.strftime("%d/%m/%Y %H:%M")
toon_shifts["Einde"] = toon_shifts["einde"].dt.strftime("%d/%m/%Y %H:%M")
toon_shifts["Netto uren"] = toon_shifts["netto_uren"].round(2)
toon_shifts["Bruto uren"] = toon_shifts["bruto_uren"].round(2)

weekuren_lijst = []
for naam, groep in shifts_df.dropna(subset=["database_naam"]).groupby("database_naam"):
    persoon = crew[crew["VOLLEDIGE_NAAM"] == naam].iloc[0]
    strobbo_naam = groep.iloc[0]["strobbo_naam"]
    totaal_strobbo = strobbo_totalen.get(normaliseer_naam(strobbo_naam), None)
    weekuren_lijst.append({
        "Medewerker": naam,
        "Weekuren volgens Strobbo": round(totaal_strobbo, 2) if totaal_strobbo is not None else None,
        "Berekende weekuren app": round(groep["netto_uren"].sum(), 2),
        "Type": persoon["TYPE"],
        "Leeftijd": persoon["LEEFTIJD"],
        "Contracturen": persoon["CONTRACTUREN"],
    })
weekuren = pd.DataFrame(weekuren_lijst)

# Los rapport voorbereiden
rapport_buffer = BytesIO()
with pd.ExcelWriter(rapport_buffer, engine="openpyxl") as writer:
    if fouten_df.empty:
        pd.DataFrame(columns=["Medewerker", "Datum", "Ernst", "Fout", "Detail"]).to_excel(writer, index=False, sheet_name="Foutenrapport")
    else:
        fouten_df.drop(columns=["bron_cellen"], errors="ignore").to_excel(writer, index=False, sheet_name="Foutenrapport")
    match_df.to_excel(writer, index=False, sheet_name="Naamkoppeling")
    toon_shifts.to_excel(writer, index=False, sheet_name="Gevonden shifts")
    weekuren.to_excel(writer, index=False, sheet_name="Weekuren")

st.markdown("<br>", unsafe_allow_html=True)
tab_overzicht, tab_fouten, tab_match, tab_shifts, tab_weekuren, tab_export = st.tabs([
    "📊 Overzicht", "🚨 Fouten", "🔗 Naamkoppeling", "🕒 Shifts", "📈 Weekuren", "📥 Export"
])

with tab_overzicht:
    st.markdown('<div class="fbo-card">', unsafe_allow_html=True)
    st.markdown('<div class="fbo-section-title">Controle-overzicht</div>', unsafe_allow_html=True)
    if fouten_df.empty:
        st.success("Geen fouten gevonden volgens de ingestelde regels.")
    else:
        belangrijkste = fouten_df.drop(columns=["bron_cellen"], errors="ignore").head(10)
        st.warning(f"Er zijn {len(fouten_df)} fout(en) of waarschuwing(en) gevonden. Bekijk de tab Fouten voor details.")
        st.dataframe(belangrijkste, use_container_width=True, hide_index=True)
    st.markdown('</div>', unsafe_allow_html=True)

with tab_fouten:
    st.markdown('<div class="fbo-card">', unsafe_allow_html=True)
    st.markdown('<div class="fbo-section-title">Foutenrapport</div>', unsafe_allow_html=True)
    if fouten_df.empty:
        st.success("Geen fouten gevonden.")
    else:
        zichtbare_fouten = fouten_df.drop(columns=["bron_cellen"], errors="ignore").copy()
        st.dataframe(zichtbare_fouten, use_container_width=True, hide_index=True)
    st.markdown('</div>', unsafe_allow_html=True)

with tab_match:
    st.markdown('<div class="fbo-card">', unsafe_allow_html=True)
    st.markdown('<div class="fbo-section-title">Naamkoppeling Strobbo ↔ Database</div>', unsafe_allow_html=True)
    st.dataframe(match_df, use_container_width=True, hide_index=True)
    st.markdown('</div>', unsafe_allow_html=True)

with tab_shifts:
    st.markdown('<div class="fbo-card">', unsafe_allow_html=True)
    st.markdown('<div class="fbo-section-title">Gevonden shifts na samenvoegen</div>', unsafe_allow_html=True)
    st.dataframe(
        toon_shifts[["strobbo_naam", "database_naam", "datum", "Start", "Einde", "pauze_minuten", "Bruto uren", "Netto uren", "bron_rij", "bron_kolom"]],
        use_container_width=True,
        hide_index=True,
    )
    st.markdown('</div>', unsafe_allow_html=True)

with tab_weekuren:
    st.markdown('<div class="fbo-card">', unsafe_allow_html=True)
    st.markdown('<div class="fbo-section-title">Weekuren per medewerker</div>', unsafe_allow_html=True)
    st.dataframe(weekuren, use_container_width=True, hide_index=True)
    st.markdown('</div>', unsafe_allow_html=True)

with tab_export:
    st.markdown('<div class="fbo-card">', unsafe_allow_html=True)
    st.markdown('<div class="fbo-section-title">Downloads</div>', unsafe_allow_html=True)
    dl1, dl2 = st.columns(2)
    with dl1:
        st.download_button(
            "📥 Download gemarkeerde Strobbo Excel",
            data=annotated_buffer.getvalue(),
            file_name="strobbo_rooster_gecontroleerd.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with dl2:
        st.download_button(
            "📥 Download los foutenrapport",
            data=rapport_buffer.getvalue(),
            file_name="foutenrapport_strobbo.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    st.caption("De gemarkeerde Strobbo Excel bevat gele cellen met opmerkingen op de foutieve vakken.")
    st.markdown('</div>', unsafe_allow_html=True)
