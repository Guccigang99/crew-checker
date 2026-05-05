import re
from datetime import datetime, timedelta
from io import BytesIO

import pandas as pd
import streamlit as st
from rapidfuzz import process, fuzz
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.comments import Comment

st.set_page_config(page_title="Strobbo Rooster Checker", layout="wide")
st.title("📊 Strobbo Rooster Checker")
st.write("Upload je crew-database en je Strobbo weekrooster. De app markeert fouten geel in de Strobbo Excel en zet uitleg als opmerking in het vak.")

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

FILL_FOUT = PatternFill("solid", fgColor="FFFF00")       # geel
FILL_OPMERKING = PatternFill("solid", fgColor="92D050")  # groen
FONT_ROOD = Font(color="FF0000", bold=True)
FONT_BOLD = Font(bold=True)

# =========================
# HELPERS
# =========================
def txt(x):
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


def is_header_of_geen_medewerker(s):
    s_norm = normaliseer_naam(s)
    if not s_norm:
        return True
    verboden = ["wk", "totaal", "ma", "di", "wo", "do", "vr", "za", "zo"]
    if s_norm in verboden:
        return True
    if re.search(r"\d{1,2}[-/]\w+", s_norm):
        return True
    if re.search(r"\d{1,2}:\d{2}", s_norm):
        return True
    return False


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
    voornaam = txt(row.get("VOORNAAM", ""))
    naam = txt(row.get("NAAM", ""))
    if voornaam and naam:
        return f"{voornaam} {naam}"
    return voornaam or naam


def bepaal_type(row):
    hele_rij = " ".join(str(v).upper() for v in row.values)
    if "STUDENT" in hele_rij:
        return "student"
    if "FLEXI" in hele_rij:
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

    # Voornaam + initiaal, bv. Luka B.
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
    for r in range(min(10, len(raw))):
        for c in range(raw.shape[1]):
            if txt(raw.iloc[r, c]).lower() == "totaal":
                return c
    return max(vind_datumkolommen(raw).keys()) + 1


def parse_shiftblok(celtekst, datum):
    if pd.isna(celtekst) or not datum:
        return []
    s = str(celtekst)
    pattern = re.compile(
        r"(\d{1,2}:\d{2})\s*-\s*(\d{1,2}:\d{2})\s*\*?\s*(?:\r?\n|\s)*\(?\s*(\d{1,2}:\d{2})?\s*\)?",
        re.MULTILINE,
    )
    blokken = []
    for m in pattern.finditer(s):
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


def fout_toevoegen(fouten, medewerker, datum, fout, detail, cellen=None, actie=None, ernst="Fout"):
    fouten.append({
        "Medewerker": medewerker,
        "Datum": datum,
        "Ernst": ernst,
        "Fout": fout,
        "Detail": detail,
        "Actie voor Excel": actie or detail,
        "Cellen": cellen or [],
    })

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
# CREW
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

ontbrekend = [c for c in ["NAAM", "VOORNAAM", "LFTD"] if c not in crew.columns]
if ontbrekend:
    st.error(f"Deze kolommen ontbreken in je crew-database: {ontbrekend}")
    st.stop()

crew["VOLLEDIGE_NAAM"] = crew.apply(maak_volledige_naam, axis=1)
crew["TYPE"] = crew.apply(bepaal_type, axis=1)
crew["LEEFTIJD"] = crew["LFTD"].apply(veilige_int)
crew["CONTRACTUREN"] = crew["CONTRACT. UREN"].apply(veilige_float)
crew = crew[crew["VOLLEDIGE_NAAM"].str.strip() != ""].copy()

# =========================
# STROBBO
# =========================
try:
    raw = pd.read_excel(BytesIO(rooster_bytes), header=None)
    wb = load_workbook(BytesIO(rooster_bytes))
    ws = wb.active
except Exception as e:
    st.error(f"Kon Strobbo-bestand niet lezen: {e}")
    st.stop()

dag_kolommen = vind_datumkolommen(raw)
if not dag_kolommen:
    st.error("Ik kon de datumkolommen niet vinden in je Strobbo-export.")
    st.stop()

totaal_col = vind_totaal_kolom(raw)

# =========================
# BLOKKEN UIT STROBBO HALEN
# =========================
blokken = []
huidige_medewerker = None
huidige_medewerker_rij = None

for r_idx, row in raw.iterrows():
    eerste_cel = txt(row.iloc[0])

    # BELANGRIJK: sommige exports hebben #Naam, andere enkel Naam.
    # Daarom accepteren we elke niet-lege naamcel als nieuwe medewerker, behalve headers.
    if eerste_cel and not is_header_of_geen_medewerker(eerste_cel):
        huidige_medewerker = eerste_cel.replace("#", "").strip()
        huidige_medewerker_rij = r_idx + 1

    if not huidige_medewerker:
        continue

    for c_idx, datum in dag_kolommen.items():
        celwaarde = row.iloc[c_idx] if c_idx < len(row) else ""
        gevonden = parse_shiftblok(celwaarde, datum)
        for b in gevonden:
            b["strobbo_naam"] = huidige_medewerker
            b["datum"] = datum
            b["bron_cellen"] = [(r_idx + 1, c_idx + 1)]
            b["medewerker_rij"] = huidige_medewerker_rij
            blokken.append(b)

if not blokken:
    st.warning("Geen shifts gevonden in het Strobbo-bestand.")
    st.stop()

blokken_df = pd.DataFrame(blokken)

# =========================
# SAMENVOEGEN TOT ECHTE SHIFTS
# =========================
samengevoegde = []
for (naam, datum), groep in blokken_df.groupby(["strobbo_naam", "datum"]):
    samengevoegde.extend(merge_blokken_naar_shifts(groep.sort_values("start").to_dict("records")))

shifts_df = pd.DataFrame(samengevoegde)

# =========================
# MATCHEN
# =========================
matches = []
for naam in shifts_df["strobbo_naam"].unique():
    beste, score = zoek_beste_match(naam, crew)
    matches.append({"Strobbo naam": naam, "Database naam": beste if beste else "NIET GEVONDEN", "Match score": score})
match_df = pd.DataFrame(matches)
match_map = {r["Strobbo naam"]: r["Database naam"] for _, r in match_df.iterrows() if r["Database naam"] != "NIET GEVONDEN"}
shifts_df["database_naam"] = shifts_df["strobbo_naam"].map(match_map)

# =========================
# FOUTEN
# =========================
fouten = []

for _, r in match_df.iterrows():
    if r["Database naam"] == "NIET GEVONDEN":
        fout_toevoegen(fouten, r["Strobbo naam"], "", "Naam niet gevonden", f"Geen match in crew-database. Score: {r['Match score']}", [], "Naam controleren", "Waarschuwing")

for _, sh in shifts_df.dropna(subset=["database_naam"]).iterrows():
    naam = sh["database_naam"]
    persoon = crew[crew["VOLLEDIGE_NAAM"] == naam].iloc[0]
    leeftijd = persoon["LEEFTIJD"]
    bruto = sh["bruto_uren"]
    netto = sh["netto_uren"]
    pauze = sh["pauze_minuten"]
    datum = sh["datum"]
    cellen = sh["bron_cellen"]

    if netto < MIN_DAGUREN:
        fout_toevoegen(fouten, naam, datum, "Shift te kort", f"{netto:.2f}u gewerkt. Minimum is {MIN_DAGUREN}u.", cellen, "Shift verlengen of verwijderen")

    if leeftijd is not None and leeftijd < 18:
        if netto > MAX_DAGUREN_MINDERJARIG:
            fout_toevoegen(fouten, naam, datum, "Minderjarige werkt te lang", f"{netto:.2f}u gewerkt. Max voor <18 is {MAX_DAGUREN_MINDERJARIG}u.", cellen, "Shift inkorten")
    else:
        if netto > MAX_DAGUREN_VOLWASSEN:
            fout_toevoegen(fouten, naam, datum, "Shift te lang", f"{netto:.2f}u gewerkt. Max is {MAX_DAGUREN_VOLWASSEN}u.", cellen, "Shift inkorten")

    if leeftijd is None or leeftijd >= 18:
        if bruto <= 5 and pauze > 0:
            fout_toevoegen(fouten, naam, datum, "Onnodige pauze", f"{bruto:.2f}u shift heeft {pauze} min pauze, maar tot 5u is geen pauze nodig.", cellen, "Hier pauze verwijderen", "Waarschuwing")
        elif 5 < bruto <= 8 and pauze < 20:
            fout_toevoegen(fouten, naam, datum, "Pauze ontbreekt", f"{bruto:.2f}u shift. Minstens 20 min pauze nodig. Gepland: {pauze} min.", cellen, "HIER 20 min pauze toedienen")
        elif bruto > 8 and pauze < 30:
            fout_toevoegen(fouten, naam, datum, "Pauze ontbreekt", f"{bruto:.2f}u shift. Minstens 30 min pauze nodig. Gepland: {pauze} min.", cellen, "HIER 30 min pauze toedienen")

    if leeftijd is not None and leeftijd < 18:
        if 4.5 < bruto <= 6 and pauze < 30:
            fout_toevoegen(fouten, naam, datum, "Pauze minderjarige ontbreekt", f"{bruto:.2f}u shift. Minderjarige heeft minstens 30 min pauze nodig. Gepland: {pauze} min.", cellen, "HIER 30 min pauze toedienen")
        elif bruto > 6 and pauze < 60:
            fout_toevoegen(fouten, naam, datum, "Pauze minderjarige ontbreekt", f"{bruto:.2f}u shift. Minderjarige heeft minstens 60 min pauze nodig. Gepland: {pauze} min.", cellen, "HIER 60 min pauze toedienen")

    if leeftijd is not None:
        einduur = sh["einde"].hour + sh["einde"].minute / 60
        if leeftijd <= 15 and einduur > 20:
            fout_toevoegen(fouten, naam, datum, "Nachtwerk minderjarige", f"{leeftijd} jaar en werkt tot {sh['einde'].strftime('%H:%M')}. Max tot 20:00.", cellen, "Shift vroeger laten eindigen")
        elif 16 <= leeftijd < 18 and einduur > 23:
            fout_toevoegen(fouten, naam, datum, "Nachtwerk minderjarige", f"{leeftijd} jaar en werkt tot {sh['einde'].strftime('%H:%M')}. Max tot 23:00.", cellen, "Shift vroeger laten eindigen")

for naam, groep in shifts_df.dropna(subset=["database_naam"]).groupby("database_naam"):
    persoon = crew[crew["VOLLEDIGE_NAAM"] == naam].iloc[0]
    leeftijd = persoon["LEEFTIJD"]
    medewerker_type = persoon["TYPE"]
    contracturen = persoon["CONTRACTUREN"]
    totaal = groep["netto_uren"].sum()
    alle_cellen = []
    for cellijst in groep["bron_cellen"]:
        alle_cellen.extend(cellijst)

    if leeftijd is not None and leeftijd < 18 and totaal > MAX_WEEKUREN_MINDERJARIG:
        fout_toevoegen(fouten, naam, "", "Weekuren overschreden", f"{totaal:.2f}u gewerkt. Max voor <18 is {MAX_WEEKUREN_MINDERJARIG}u.", alle_cellen, "Weekplanning verminderen")
    elif (leeftijd is None or leeftijd >= 18) and totaal > MAX_WEEKUREN_VOLWASSEN:
        fout_toevoegen(fouten, naam, "", "Weekuren overschreden", f"{totaal:.2f}u gewerkt. Max is {MAX_WEEKUREN_VOLWASSEN}u.", alle_cellen, "Weekplanning verminderen")

    if medewerker_type == "vast" and contracturen > 0 and totaal < contracturen:
        fout_toevoegen(fouten, naam, "", "Contracturen niet gehaald", f"{totaal:.2f}u gepland, contract is {contracturen:.2f}u. Tekort: {contracturen - totaal:.2f}u.", alle_cellen, "Extra uren plannen")

for naam, groep in shifts_df.dropna(subset=["database_naam"]).groupby("database_naam"):
    groep = groep.sort_values("start").reset_index(drop=True)
    for i in range(1, len(groep)):
        vorige = groep.iloc[i - 1]
        huidige = groep.iloc[i]
        rust = (huidige["start"] - vorige["einde"]).total_seconds() / 3600
        if rust < MIN_RUSTUREN_TUSSEN_SHIFTS:
            fout_toevoegen(fouten, naam, huidige["datum"], "Te weinig rust", f"Slechts {rust:.2f}u rust tussen {vorige['einde'].strftime('%d/%m %H:%M')} en {huidige['start'].strftime('%d/%m %H:%M')}.", huidige["bron_cellen"], "Rusttijd corrigeren")

fouten_df = pd.DataFrame(fouten)

# =========================
# EXCEL MARKEREN
# =========================
# Vanaf nu worden fouten enkel aangeduid in het originele Strobbo-vak:
# - geel gemarkeerd
# - rode tekst
# - uitleg als Excel-opmerking/comment wanneer je met je cursor op het vak staat

if not fouten_df.empty:
    opmerkingen_per_cel = {}

    for _, f in fouten_df.iterrows():
        cellen = f.get("Cellen", [])
        detail = str(f.get("Detail", ""))
        actie = str(f.get("Actie voor Excel", ""))
        tekst = f"{f['Medewerker']} - {f['Fout']}\n{detail}\nActie: {actie}"

        for rr, cc in cellen:
            opmerkingen_per_cel.setdefault((rr, cc), []).append(tekst)

    for (rr, cc), meldingen in opmerkingen_per_cel.items():
        uniek = []
        for melding in meldingen:
            if melding not in uniek:
                uniek.append(melding)

        cel = ws.cell(row=rr, column=cc)
        cel.fill = FILL_FOUT
        cel.font = FONT_ROOD
        cel.comment = Comment("\n\n".join(uniek), "Strobbo Checker")

annotated_buffer = BytesIO()
wb.save(annotated_buffer)

# =========================
# OUTPUT
# =========================
st.subheader("✅ Naamkoppeling Strobbo ↔ Database")
st.dataframe(match_df, use_container_width=True)

st.subheader("📋 Gevonden shifts na samenvoegen")
toon = shifts_df.copy()
toon["Start"] = toon["start"].dt.strftime("%d/%m/%Y %H:%M")
toon["Einde"] = toon["einde"].dt.strftime("%d/%m/%Y %H:%M")
toon["Bruto uren"] = toon["bruto_uren"].round(2)
toon["Netto uren"] = toon["netto_uren"].round(2)
st.dataframe(toon[["strobbo_naam", "database_naam", "datum", "Start", "Einde", "pauze_minuten", "Bruto uren", "Netto uren"]], use_container_width=True)

st.subheader("🚨 Foutenrapport")
if fouten_df.empty:
    st.success("Geen fouten gevonden volgens de ingestelde regels.")
else:
    st.error(f"{len(fouten_df)} fout(en) of waarschuwing(en) gevonden.")
    st.dataframe(fouten_df.drop(columns=["Cellen"], errors="ignore"), use_container_width=True)

st.download_button(
    "📥 Download gemarkeerde Strobbo Excel",
    data=annotated_buffer.getvalue(),
    file_name="strobbo_rooster_gecontroleerd.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

rapport_buffer = BytesIO()
with pd.ExcelWriter(rapport_buffer, engine="openpyxl") as writer:
    match_df.to_excel(writer, index=False, sheet_name="Naamkoppeling")
    toon.to_excel(writer, index=False, sheet_name="Gevonden shifts")
    fouten_df.drop(columns=["Cellen"], errors="ignore").to_excel(writer, index=False, sheet_name="Foutenrapport")

st.download_button(
    "📥 Download los foutenrapport",
    data=rapport_buffer.getvalue(),
    file_name="foutenrapport.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

st.subheader("📊 Weekuren per medewerker")
weekuren = shifts_df.dropna(subset=["database_naam"]).groupby("database_naam")["netto_uren"].sum().reset_index()
weekuren = weekuren.rename(columns={"database_naam": "Medewerker", "netto_uren": "Weekuren"})
weekuren["Weekuren"] = weekuren["Weekuren"].round(2)
weekuren = weekuren.merge(
    crew[["VOLLEDIGE_NAAM", "TYPE", "LEEFTIJD", "CONTRACTUREN"]],
    left_on="Medewerker",
    right_on="VOLLEDIGE_NAAM",
    how="left",
).drop(columns=["VOLLEDIGE_NAAM"])
st.dataframe(weekuren, use_container_width=True)
