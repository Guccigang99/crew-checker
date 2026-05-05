import re
from datetime import datetime, timedelta
from io import BytesIO

import pandas as pd
import streamlit as st
from rapidfuzz import process, fuzz
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.comments import Comment

st.set_page_config(page_title="Strobbo Rooster Checker", layout="wide")
st.title("📊 Strobbo Rooster Checker")
st.write("Upload je crew-database en je Strobbo weekrooster. De app controleert fouten en markeert ze in je Strobbo Excel.")

# =========================
# REGELS
# =========================
MIN_DAGUREN = 2
MAX_DAGUREN_VOLWASSEN = 11
MAX_DAGUREN_MINDERJARIG = 8
MAX_WEEKUREN_VOLWASSEN = 50
MAX_WEEKUREN_MINDERJARIG = 40
MIN_RUSTUREN_TUSSEN_SHIFTS = 10
FUZZY_MATCH_SCORE = 65
MERGE_GAP_MINUTEN = 5

FILL_FOUT = PatternFill("solid", fgColor="FFFF00")
FONT_FOUT = Font(color="FF0000", bold=True)

# =========================
# HELPERS
# =========================
def tekst(x):
    if pd.isna(x):
        return ""
    return str(x).strip()


def normaliseer_naam(x):
    x = str(x).lower().strip()
    x = x.replace("#", "")
    x = re.sub(r"<\s*\d+", "", x)
    x = x.replace(".", "")
    x = x.replace("-", " ")
    x = re.sub(r"\s+", " ", x)
    return x.strip()


def veilige_int(x):
    try:
        if pd.isna(x):
            return None
        waarde = int(float(str(x).replace(",", ".")))
        if waarde <= 0 or waarde > 100:
            return None
        return waarde
    except Exception:
        return None


def veilige_float(x):
    try:
        if pd.isna(x):
            return 0.0
        return float(str(x).replace(",", "."))
    except Exception:
        return 0.0


def maak_volledige_naam(row):
    voornaam = tekst(row.get("VOORNAAM", ""))
    naam = tekst(row.get("NAAM", ""))
    if voornaam and naam:
        return f"{voornaam} {naam}"
    return voornaam or naam


def bepaal_type(row):
    volledige_rij = " ".join(str(v).upper() for v in row.values)
    if "STUDENT" in volledige_rij:
        return "student"
    if "FLEXI" in volledige_rij:
        return "flexi"
    contract = row.get("CONTRACT. UREN", row.get("CONTRACT UREN", ""))
    if pd.isna(contract) or str(contract).strip() == "":
        return "student"
    return "vast"


def zoek_beste_match(strobbo_naam, crew_df):
    naam = normaliseer_naam(strobbo_naam)

    exacte = crew_df[crew_df["VOORNAAM"].apply(normaliseer_naam) == naam]
    if len(exacte) == 1:
        return exacte.iloc[0]["VOLLEDIGE_NAAM"], 100

    m = re.match(r"^([a-zA-ZÀ-ÿ]+)\s+([a-zA-Z])$", naam)
    if m:
        voornaam_gezocht = m.group(1)
        initiaal = m.group(2)
        for _, row in crew_df.iterrows():
            if normaliseer_naam(row["VOORNAAM"]) == voornaam_gezocht and normaliseer_naam(row["NAAM"]).startswith(initiaal):
                return row["VOLLEDIGE_NAAM"], 100

    voornamen = crew_df["VOORNAAM"].astype(str).tolist()
    voornamen_norm = [normaliseer_naam(v) for v in voornamen]
    match = process.extractOne(naam, voornamen_norm, scorer=fuzz.ratio)
    if match:
        _, score, index = match
        if score >= 90:
            return crew_df.iloc[index]["VOLLEDIGE_NAAM"], round(score, 2)

    namen = crew_df["VOLLEDIGE_NAAM"].tolist()
    namen_norm = [normaliseer_naam(n) for n in namen]
    match = process.extractOne(naam, namen_norm, scorer=fuzz.token_sort_ratio)
    if match:
        _, score, index = match
        if score >= FUZZY_MATCH_SCORE:
            return namen[index], round(score, 2)
        return None, round(score, 2)

    return None, 0


def parse_datum(x):
    if isinstance(x, datetime):
        return x.date()
    s = str(x).lower().strip()
    maanden = {
        "jan": 1, "feb": 2, "mrt": 3, "mar": 3, "apr": 4, "mei": 5,
        "jun": 6, "jul": 7, "aug": 8, "sep": 9, "okt": 10, "oct": 10,
        "nov": 11, "dec": 12,
    }
    m = re.search(r"(\d{1,2})[-/\s]([a-zA-ZÀ-ÿ]+)", s)
    if not m:
        return None
    dag = int(m.group(1))
    maand = maanden.get(m.group(2)[:3])
    if not maand:
        return None
    return datetime(2026, maand, dag).date()


def vind_datumkolommen(raw):
    dag_kolommen = {}
    for r in range(min(10, len(raw))):
        for c in range(raw.shape[1]):
            d = parse_datum(raw.iloc[r, c])
            if d:
                dag_kolommen[c] = d
        if len(dag_kolommen) >= 5:
            return dag_kolommen
    return dag_kolommen


def vind_totaal_kolom(raw):
    # Zoek kolom met header Totaal, anders neem de eerste kolom rechts van zondag/datumkolommen.
    for r in range(min(10, len(raw))):
        for c in range(raw.shape[1]):
            if str(raw.iloc[r, c]).strip().lower() == "totaal":
                return c
    datums = vind_datumkolommen(raw)
    if datums:
        return max(datums.keys()) + 1
    return raw.shape[1] - 1


def parse_totaal_uren(x):
    # Leest bv. 34:45 / 35:00 => 34.75
    s = str(x)
    m = re.search(r"(\d{1,3}):(\d{2})", s)
    if not m:
        return None
    return int(m.group(1)) + int(m.group(2)) / 60


def parse_shiftblok(celtekst, datum):
    if pd.isna(celtekst) or not datum:
        return []
    s = str(celtekst)
    patroon = re.compile(
        r"(\d{1,2}:\d{2})\s*-\s*(\d{1,2}:\d{2})\s*\*?\s*(?:\r?\n|\s)*\(?\s*(\d{1,2}:\d{2})?\s*\)?",
        re.MULTILINE,
    )
    blokken = []
    for m in patroon.finditer(s):
        start_txt = m.group(1)
        einde_txt = m.group(2)
        pauze_txt = m.group(3) or "00:00"
        start = datetime.combine(datum, datetime.strptime(start_txt, "%H:%M").time())
        einde = datetime.combine(datum, datetime.strptime(einde_txt, "%H:%M").time())
        if einde <= start:
            einde += timedelta(days=1)
        ph, pm = map(int, pauze_txt.split(":"))
        blokken.append({
            "start": start,
            "einde": einde,
            "pauze_minuten": ph * 60 + pm,
        })
    return blokken


def merge_blokken_naar_shifts(blokken):
    if not blokken:
        return []
    blokken = sorted(blokken, key=lambda x: x["start"])
    shifts = []
    huidig = blokken[0].copy()
    for blok in blokken[1:]:
        gap = (blok["start"] - huidig["einde"]).total_seconds() / 60
        if gap <= MERGE_GAP_MINUTEN:
            huidig["einde"] = max(huidig["einde"], blok["einde"])
            huidig["pauze_minuten"] += blok["pauze_minuten"]
            huidig["bron_cellen"].extend(blok["bron_cellen"])
        else:
            shifts.append(huidig)
            huidig = blok.copy()
    shifts.append(huidig)

    for sh in shifts:
        bruto = (sh["einde"] - sh["start"]).total_seconds() / 3600
        sh["bruto_uren"] = bruto
        sh["netto_uren"] = bruto - (sh["pauze_minuten"] / 60)
    return shifts


def fout_toevoegen(fouten, medewerker, datum, fout, detail, cellen=None, ernst="Fout"):
    fouten.append({
        "Medewerker": medewerker,
        "Datum": datum,
        "Ernst": ernst,
        "Fout": fout,
        "Detail": detail,
        "Cellen": cellen or [],
    })


def is_medewerker_rij(eerste_cel, datum_kolommen, row):
    # Herkent zowel #Naam als naam zonder #. Vermijdt lege subrijen.
    eerste = tekst(eerste_cel)
    if not eerste:
        return False
    if eerste.startswith("#"):
        return True
    lage = eerste.lower()
    verboden = {"wk21", "ma", "di", "wo", "do", "vr", "za", "zo", "totaal"}
    if lage in verboden:
        return False
    if parse_datum(eerste):
        return False
    # Een medewerker-rij heeft vaak minstens 1 shift of totaal op dezelfde rij.
    for c in list(datum_kolommen.keys()):
        if c < len(row) and re.search(r"\d{1,2}:\d{2}\s*-\s*\d{1,2}:\d{2}", str(row.iloc[c])):
            return True
    return True

# =========================
# UPLOADS
# =========================
crew_file = st.file_uploader("👥 Upload crew-database Excel", type=["xlsx"])
rooster_file = st.file_uploader("📅 Upload Strobbo weekrooster Excel", type=["xlsx"])

if not crew_file or not rooster_file:
    st.info("Upload beide bestanden om te starten.")
    st.stop()

crew_bytes = crew_file.getvalue()
rooster_bytes = rooster_file.getvalue()

# =========================
# CREW INLEZEN
# =========================
try:
    crew = pd.read_excel(BytesIO(crew_bytes))
except Exception as e:
    st.error(f"Kon crew-database niet lezen: {e}")
    st.stop()

crew.columns = [str(c).strip().upper() for c in crew.columns]
if "CONTRACT. UREN" not in crew.columns and "CONTRACT UREN" in crew.columns:
    crew["CONTRACT. UREN"] = crew["CONTRACT UREN"]
if "CONTRACT. UREN" not in crew.columns:
    crew["CONTRACT. UREN"] = ""

vereist = ["NAAM", "VOORNAAM", "LFTD"]
ontbrekend = [c for c in vereist if c not in crew.columns]
if ontbrekend:
    st.error(f"Deze kolommen ontbreken in je crew-database: {ontbrekend}")
    st.stop()

crew["VOLLEDIGE_NAAM"] = crew.apply(maak_volledige_naam, axis=1)
crew["TYPE"] = crew.apply(bepaal_type, axis=1)
crew["LEEFTIJD"] = crew["LFTD"].apply(veilige_int)
crew["CONTRACTUREN"] = crew["CONTRACT. UREN"].apply(veilige_float)
crew = crew[crew["VOLLEDIGE_NAAM"].str.strip() != ""].copy()

# =========================
# STROBBO INLEZEN
# =========================
try:
    raw = pd.read_excel(BytesIO(rooster_bytes), sheet_name=0, header=None)
    wb = load_workbook(BytesIO(rooster_bytes))
    ws = wb.active
except Exception as e:
    st.error(f"Kon Strobbo-rooster niet lezen: {e}")
    st.stop()

dag_kolommen = vind_datumkolommen(raw)
if not dag_kolommen:
    st.error("Ik kon geen datumkolommen vinden in het Strobbo-bestand.")
    st.stop()

totaal_col = vind_totaal_kolom(raw)

# =========================
# BLOKKEN EXTRACTEN + STROBBO TOTALEN
# =========================
blokken = []
strobbo_totalen = {}
strobbo_totaal_cellen = {}
huidige_medewerker = None

for r_idx, row in raw.iterrows():
    eerste_cel = tekst(row.iloc[0])

    if is_medewerker_rij(eerste_cel, dag_kolommen, row):
        huidige_medewerker = eerste_cel.replace("#", "").strip()
        totaal_uren = parse_totaal_uren(row.iloc[totaal_col] if totaal_col < len(row) else "")
        if huidige_medewerker and totaal_uren is not None:
            key = normaliseer_naam(huidige_medewerker)
            strobbo_totalen[key] = totaal_uren
            strobbo_totaal_cellen[key] = [(r_idx + 1, totaal_col + 1)]

    if not huidige_medewerker:
        continue

    for c_idx, datum in dag_kolommen.items():
        if c_idx >= len(row):
            continue
        gevonden = parse_shiftblok(row.iloc[c_idx], datum)
        for b in gevonden:
            b["strobbo_naam"] = huidige_medewerker
            b["datum"] = datum
            b["bron_cellen"] = [(r_idx + 1, c_idx + 1)]
            blokken.append(b)

if not blokken:
    st.warning("Geen shifts gevonden in het Strobbo-bestand.")
    st.stop()

blokken_df = pd.DataFrame(blokken)

# =========================
# BLOKKEN SAMENVOEGEN TOT ECHTE SHIFTS
# =========================
samengevoegd = []
for (naam, datum), groep in blokken_df.groupby(["strobbo_naam", "datum"]):
    lijst = groep.sort_values("start").to_dict("records")
    samengevoegd.extend(merge_blokken_naar_shifts(lijst))

shifts_df = pd.DataFrame(samengevoegd)

# =========================
# NAMEN MATCHEN
# =========================
match_resultaten = []
for naam in shifts_df["strobbo_naam"].unique():
    beste, score = zoek_beste_match(naam, crew)
    match_resultaten.append({
        "Strobbo naam": naam,
        "Database naam": beste if beste else "NIET GEVONDEN",
        "Match score": score,
    })
match_df = pd.DataFrame(match_resultaten)
match_map = {r["Strobbo naam"]: r["Database naam"] for _, r in match_df.iterrows() if r["Database naam"] != "NIET GEVONDEN"}
shifts_df["database_naam"] = shifts_df["strobbo_naam"].map(match_map)

# =========================
# FOUTEN CONTROLEREN
# =========================
fouten = []

for _, r in match_df.iterrows():
    if r["Database naam"] == "NIET GEVONDEN":
        fout_toevoegen(
            fouten,
            r["Strobbo naam"],
            "",
            "Naam niet gevonden",
            f"Geen goede match in crew-database. Match score: {r['Match score']}",
            [],
            "Waarschuwing",
        )

# Per shift controles
for _, sh in shifts_df.dropna(subset=["database_naam"]).iterrows():
    naam = sh["database_naam"]
    persoon = crew[crew["VOLLEDIGE_NAAM"] == naam].iloc[0]
    leeftijd = persoon["LEEFTIJD"]
    bruto = sh["bruto_uren"]
    netto = sh["netto_uren"]
    pauze = sh["pauze_minuten"]
    cellen = sh["bron_cellen"]

    if netto < MIN_DAGUREN:
        fout_toevoegen(fouten, naam, sh["datum"], "Shift te kort", f"{netto:.2f}u gewerkt. Minimum is {MIN_DAGUREN}u.", cellen)

    if leeftijd is not None and leeftijd < 18:
        if netto > MAX_DAGUREN_MINDERJARIG:
            fout_toevoegen(fouten, naam, sh["datum"], "Minderjarige werkt te lang", f"{netto:.2f}u gewerkt. Maximum voor <18 is {MAX_DAGUREN_MINDERJARIG}u.", cellen)
    else:
        if netto > MAX_DAGUREN_VOLWASSEN:
            fout_toevoegen(fouten, naam, sh["datum"], "Shift te lang", f"{netto:.2f}u gewerkt. Maximum is {MAX_DAGUREN_VOLWASSEN}u.", cellen)

    if leeftijd is None or leeftijd >= 18:
        if bruto <= 5 and pauze > 0:
            fout_toevoegen(fouten, naam, sh["datum"], "Onnodige pauze", f"{bruto:.2f}u shift heeft {pauze} min pauze, maar tot 5u is geen pauze nodig.", cellen, "Waarschuwing")
        elif 5 < bruto <= 8 and pauze < 20:
            fout_toevoegen(fouten, naam, sh["datum"], "Pauze ontbreekt", f"{bruto:.2f}u shift. Minstens 20 min pauze nodig. Gepland: {pauze} min.", cellen)
        elif bruto > 8 and pauze < 30:
            fout_toevoegen(fouten, naam, sh["datum"], "Pauze ontbreekt", f"{bruto:.2f}u shift. Minstens 30 min pauze nodig. Gepland: {pauze} min.", cellen)

    if leeftijd is not None and leeftijd < 18:
        if 4.5 < bruto <= 6 and pauze < 30:
            fout_toevoegen(fouten, naam, sh["datum"], "Pauze minderjarige ontbreekt", f"{bruto:.2f}u shift. Minderjarige heeft minstens 30 min pauze nodig. Gepland: {pauze} min.", cellen)
        elif bruto > 6 and pauze < 60:
            fout_toevoegen(fouten, naam, sh["datum"], "Pauze minderjarige ontbreekt", f"{bruto:.2f}u shift. Minderjarige heeft minstens 60 min pauze nodig. Gepland: {pauze} min.", cellen)

    if leeftijd is not None:
        einduur = sh["einde"].hour + sh["einde"].minute / 60
        if leeftijd <= 15 and einduur > 20:
            fout_toevoegen(fouten, naam, sh["datum"], "Nachtwerk minderjarige", f"{leeftijd} jaar en werkt tot {sh['einde'].strftime('%H:%M')}. Max tot 20:00.", cellen)
        elif 16 <= leeftijd < 18 and einduur > 23:
            fout_toevoegen(fouten, naam, sh["datum"], "Nachtwerk minderjarige", f"{leeftijd} jaar en werkt tot {sh['einde'].strftime('%H:%M')}. Max tot 23:00.", cellen)

# Weekuren + contracturen: op basis van Strobbo totaalvak, niet opgetelde app-shifts.
for naam, groep in shifts_df.dropna(subset=["database_naam"]).groupby("database_naam"):
    persoon = crew[crew["VOLLEDIGE_NAAM"] == naam].iloc[0]
    leeftijd = persoon["LEEFTIJD"]
    medewerker_type = persoon["TYPE"]
    contracturen = persoon["CONTRACTUREN"]
    strobbo_naam = groep.iloc[0]["strobbo_naam"]
    key = normaliseer_naam(strobbo_naam)
    totaal_weekuren = strobbo_totalen.get(key)
    totaal_cellen = strobbo_totaal_cellen.get(key, [])

    if totaal_weekuren is None:
        totaal_weekuren = groep["netto_uren"].sum()

    if leeftijd is not None and leeftijd < 18:
        if totaal_weekuren > MAX_WEEKUREN_MINDERJARIG:
            fout_toevoegen(fouten, naam, "", "Weekuren overschreden", f"{totaal_weekuren:.2f}u gewerkt volgens Strobbo. Maximum voor <18 is {MAX_WEEKUREN_MINDERJARIG}u.", totaal_cellen)
    else:
        if totaal_weekuren > MAX_WEEKUREN_VOLWASSEN:
            fout_toevoegen(fouten, naam, "", "Weekuren overschreden", f"{totaal_weekuren:.2f}u gewerkt volgens Strobbo. Maximum is {MAX_WEEKUREN_VOLWASSEN}u.", totaal_cellen)

    if medewerker_type == "vast" and contracturen > 0 and totaal_weekuren < contracturen:
        tekort = contracturen - totaal_weekuren
        fout_toevoegen(fouten, naam, "", "Contracturen niet gehaald", f"{totaal_weekuren:.2f}u gepland volgens Strobbo, contract is {contracturen:.2f}u. Tekort: {tekort:.2f}u.", totaal_cellen)

# Rust tussen shifts
for naam, groep in shifts_df.dropna(subset=["database_naam"]).groupby("database_naam"):
    groep = groep.sort_values("start").reset_index(drop=True)
    for i in range(1, len(groep)):
        vorige = groep.iloc[i - 1]
        huidige = groep.iloc[i]
        rust = (huidige["start"] - vorige["einde"]).total_seconds() / 3600
        if rust < MIN_RUSTUREN_TUSSEN_SHIFTS:
            fout_toevoegen(fouten, naam, huidige["datum"], "Te weinig rust", f"Slechts {rust:.2f}u rust tussen {vorige['einde'].strftime('%d/%m %H:%M')} en {huidige['start'].strftime('%d/%m %H:%M')}.", huidige["bron_cellen"])

fouten_df = pd.DataFrame(fouten)

# =========================
# EXCEL MARKEREN MET COMMENTS
# =========================
if not fouten_df.empty:
    opmerkingen_per_cel = {}
    for _, f in fouten_df.iterrows():
        for cel in f.get("Cellen", []):
            opmerkingen_per_cel.setdefault(tuple(cel), []).append(f"{f['Fout']}: {f['Detail']}")

    for (r, c), opmerkingen in opmerkingen_per_cel.items():
        cell = ws.cell(row=r, column=c)
        cell.fill = FILL_FOUT
        cell.font = FONT_FOUT
        unieke_opmerkingen = []
        for opm in opmerkingen:
            if opm not in unieke_opmerkingen:
                unieke_opmerkingen.append(opm)
        cell.comment = Comment("\n\n".join(unieke_opmerkingen), "Strobbo Checker")

marked_buffer = BytesIO()
wb.save(marked_buffer)

# =========================
# OUTPUT
# =========================
st.subheader("✅ Naamkoppeling Strobbo ↔ Database")
st.dataframe(match_df, use_container_width=True)

st.subheader("📋 Gevonden shifts na samenvoegen")
toon_shifts = shifts_df.copy()
toon_shifts["Start"] = toon_shifts["start"].dt.strftime("%d/%m/%Y %H:%M")
toon_shifts["Einde"] = toon_shifts["einde"].dt.strftime("%d/%m/%Y %H:%M")
toon_shifts["Bruto uren"] = toon_shifts["bruto_uren"].round(2)
toon_shifts["Netto uren"] = toon_shifts["netto_uren"].round(2)
st.dataframe(
    toon_shifts[["strobbo_naam", "database_naam", "datum", "Start", "Einde", "pauze_minuten", "Bruto uren", "Netto uren"]],
    use_container_width=True,
)

st.subheader("🚨 Foutenrapport")
if fouten_df.empty:
    st.success("Geen fouten gevonden volgens de ingestelde regels.")
else:
    st.error(f"{len(fouten_df)} fout(en) of waarschuwing(en) gevonden.")
    st.dataframe(fouten_df.drop(columns=["Cellen"], errors="ignore"), use_container_width=True)

st.download_button(
    label="📥 Download gemarkeerde Strobbo Excel",
    data=marked_buffer.getvalue(),
    file_name="strobbo_rooster_gecontroleerd.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

rapport_buffer = BytesIO()
with pd.ExcelWriter(rapport_buffer, engine="openpyxl") as writer:
    match_df.to_excel(writer, index=False, sheet_name="Naamkoppeling")
    toon_shifts.to_excel(writer, index=False, sheet_name="Gevonden shifts")
    if fouten_df.empty:
        pd.DataFrame(columns=["Medewerker", "Datum", "Ernst", "Fout", "Detail"]).to_excel(writer, index=False, sheet_name="Foutenrapport")
    else:
        fouten_df.drop(columns=["Cellen"], errors="ignore").to_excel(writer, index=False, sheet_name="Foutenrapport")

st.download_button(
    label="📥 Download los foutenrapport",
    data=rapport_buffer.getvalue(),
    file_name="foutenrapport_strobbo.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

st.subheader("📊 Weekuren per medewerker")
weekuren_rows = []
for strobbo_naam, db_naam in match_map.items():
    key = normaliseer_naam(strobbo_naam)
    weekuren_rows.append({
        "Medewerker": db_naam,
        "Strobbo naam": strobbo_naam,
        "Weekuren volgens Strobbo": round(strobbo_totalen.get(key, 0), 2),
    })
weekuren = pd.DataFrame(weekuren_rows)
if not weekuren.empty:
    weekuren = weekuren.merge(
        crew[["VOLLEDIGE_NAAM", "TYPE", "LEEFTIJD", "CONTRACTUREN"]],
        left_on="Medewerker",
        right_on="VOLLEDIGE_NAAM",
        how="left",
    ).drop(columns=["VOLLEDIGE_NAAM"])
st.dataframe(weekuren, use_container_width=True)
