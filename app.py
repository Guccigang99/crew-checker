import re
from datetime import datetime, timedelta
from io import BytesIO

import pandas as pd
import streamlit as st
from rapidfuzz import process, fuzz

st.set_page_config(page_title="Strobbo Rooster Checker", layout="wide")

st.title("📊 Strobbo Rooster Checker")
st.write("Upload je crew-database en je Strobbo weekrooster. De app controleert automatisch de fouten.")

# =========================
# INSTELLINGEN
# =========================

MIN_DAGUREN = 2
MAX_DAGUREN_VOLWASSEN = 11
MAX_DAGUREN_MINDERJARIG = 8

MAX_WEEKUREN_VOLWASSEN = 50
MAX_WEEKUREN_MINDERJARIG = 40

MIN_RUSTUREN_TUSSEN_SHIFTS = 10
FUZZY_MATCH_SCORE = 65

# Taken die direct op elkaar aansluiten worden als 1 shift gezien.
# Voorbeeld: 09:30-10:00 + 10:00-14:45 = 1 shift van 09:30 tot 14:45.
MERGE_GAP_MINUTEN = 5


# =========================
# HELPERS
# =========================

def normaliseer_tekst(x):
    if pd.isna(x):
        return ""
    return str(x).strip()


def normaliseer_naam(naam):
    naam = str(naam).lower().strip()
    naam = naam.replace("#", "")
    naam = re.sub(r"<\s*\d+", "", naam)
    naam = naam.replace(".", "")
    naam = naam.replace("-", " ")
    naam = re.sub(r"\s+", " ", naam)
    return naam.strip()


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
    if "FLEXI" in volledige_rij:
        return "flexi"

    contract = row.get("CONTRACT. UREN", row.get("CONTRACT UREN", ""))

    if pd.isna(contract) or str(contract).strip() == "":
        return "student"

    return "vast"


def veilige_float(x):
    try:
        if pd.isna(x):
            return 0
        return float(str(x).replace(",", "."))
    except Exception:
        return 0


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


def zoek_beste_match(strobbo_naam, crew_df):
    naam = normaliseer_naam(strobbo_naam)

    # 1. Exacte voornaam match
    exacte_matches = []
    for _, row in crew_df.iterrows():
        voornaam = normaliseer_naam(row["VOORNAAM"])
        if naam == voornaam:
            exacte_matches.append(row["VOLLEDIGE_NAAM"])

    if len(exacte_matches) == 1:
        return exacte_matches[0], 100

    # 2. Match op voornaam + initiaal achternaam, bv. Luka B.
    match_initiaal = re.match(r"^([a-zA-ZÀ-ÿ]+)\s+([a-zA-Z])$", naam)
    if match_initiaal:
        voornaam_gezocht = match_initiaal.group(1)
        initiaal_gezocht = match_initiaal.group(2)

        for _, row in crew_df.iterrows():
            voornaam = normaliseer_naam(row["VOORNAAM"])
            achternaam = normaliseer_naam(row["NAAM"])

            if voornaam == voornaam_gezocht and achternaam.startswith(initiaal_gezocht):
                return row["VOLLEDIGE_NAAM"], 100

    # 3. Fuzzy match op voornaam
    voornamen = crew_df["VOORNAAM"].astype(str).tolist()
    voornamen_norm = [normaliseer_naam(v) for v in voornamen]

    match = process.extractOne(naam, voornamen_norm, scorer=fuzz.ratio)
    if match:
        _, score, index = match
        if score >= 90:
            return crew_df.iloc[index]["VOLLEDIGE_NAAM"], score

    # 4. Fuzzy match op volledige naam
    crew_namen = crew_df["VOLLEDIGE_NAAM"].tolist()
    crew_norms = [normaliseer_naam(n) for n in crew_namen]

    match = process.extractOne(naam, crew_norms, scorer=fuzz.token_sort_ratio)
    if not match:
        return None, 0

    _, score, index = match

    if score >= FUZZY_MATCH_SCORE:
        return crew_namen[index], score

    return None, score


def parse_datum(waarde):
    if isinstance(waarde, datetime):
        return waarde.date()

    tekst = str(waarde).strip().lower()

    maand_map = {
        "jan": 1, "feb": 2, "mrt": 3, "mar": 3, "apr": 4,
        "mei": 5, "jun": 6, "jul": 7, "aug": 8,
        "sep": 9, "okt": 10, "oct": 10, "nov": 11, "dec": 12
    }

    match = re.search(r"(\d{1,2})[-/\s]([a-zA-ZÀ-ÿ]+)", tekst)

    if match:
        dag = int(match.group(1))
        maand_txt = match.group(2)[:3]
        maand = maand_map.get(maand_txt)

        if maand:
            # Jaar uit exportnaam kan moeilijk zijn; voor weekcontrole maakt dit niet uit.
            return datetime(2026, maand, dag).date()

    return None


def vind_dag_kolommen(raw):
    dag_kolommen = {}

    # Zoek in de eerste 8 rijen naar datums zoals 18-mei, 19-mei, ...
    for rij in range(min(8, len(raw))):
        for col in range(raw.shape[1]):
            datum = parse_datum(raw.iloc[rij, col])
            if datum:
                dag_kolommen[col] = datum

        if len(dag_kolommen) >= 5:
            return dag_kolommen

    return dag_kolommen


def parse_pauze_minuten(pauze_txt):
    if not pauze_txt:
        return 0

    match = re.search(r"(\d{1,2}):(\d{2})", pauze_txt)
    if not match:
        return 0

    uren = int(match.group(1))
    minuten = int(match.group(2))
    return uren * 60 + minuten


def parse_shiftblokken(cell_text, datum):
    """
    Leest 1 cel van Strobbo.

    Voorbeeld cel:
    09:30-10:00
    (00:00)
    Roeselare - Aanv. C.
    10:00-14:45
    (00:00)
    Roeselare - Service

    Dit zijn eerst losse taakblokken.
    Daarna worden aansluitende taakblokken samengevoegd tot 1 echte shift.
    """
    if pd.isna(cell_text) or not datum:
        return []

    tekst = str(cell_text)

    # Zoek elk tijdsblok en optionele pauze er vlak na.
    patroon = re.compile(
        r"(\d{1,2}:\d{2})\s*-\s*(\d{1,2}:\d{2})\s*\*?\s*(?:\n|\r|\s)*\(?\s*(\d{1,2}:\d{2})?\s*\)?",
        re.MULTILINE
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
            "origineel": tekst
        })

    if not blokken:
        return []

    blokken = sorted(blokken, key=lambda x: x["start"])

    # Aansluitende blokken samenvoegen tot 1 shift.
    samengevoegd = []
    huidige = blokken[0].copy()

    for blok in blokken[1:]:
        gap_minuten = (blok["start"] - huidige["einde"]).total_seconds() / 60

        if gap_minuten <= MERGE_GAP_MINUTEN:
            huidige["einde"] = max(huidige["einde"], blok["einde"])
            huidige["pauze_minuten"] += blok["pauze_minuten"]
            huidige["origineel"] += "\n---\n" + blok["origineel"]
        else:
            samengevoegd.append(huidige)
            huidige = blok.copy()

    samengevoegd.append(huidige)

    for shift in samengevoegd:
        bruto_uren = (shift["einde"] - shift["start"]).total_seconds() / 3600
        netto_uren = bruto_uren - (shift["pauze_minuten"] / 60)
        shift["bruto_uren"] = bruto_uren
        shift["netto_uren"] = netto_uren

    return samengevoegd


def voeg_fout(fouten, naam, datum, fouttype, detail, ernst="Fout"):
    fouten.append({
        "Medewerker": naam,
        "Datum": datum,
        "Ernst": ernst,
        "Fout": fouttype,
        "Detail": detail
    })


# =========================
# UPLOADS
# =========================

crew_file = st.file_uploader("👥 Upload crew-database Excel", type=["xlsx"])
rooster_file = st.file_uploader("📅 Upload Strobbo weekrooster Excel", type=["xlsx"])

if not crew_file or not rooster_file:
    st.info("Upload beide Excel-bestanden om te starten.")
    st.stop()


# =========================
# CREW DATABASE INLEZEN
# =========================

try:
    crew = pd.read_excel(crew_file)
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


# =========================
# STROBBO ROOSTER INLEZEN
# =========================

try:
    raw = pd.read_excel(rooster_file, sheet_name=0, header=None)
except Exception as e:
    st.error(f"Kon Strobbo-rooster niet lezen: {e}")
    st.stop()

dag_kolommen = vind_dag_kolommen(raw)

if not dag_kolommen:
    st.error("Ik kon geen datumkolommen vinden in het Strobbo-bestand.")
    st.stop()

# Meestal is kolom 0 de naamkolom. Totaalkolom negeren we automatisch omdat daar geen datum boven staat.


# =========================
# SHIFTS EXTRACTEN
# =========================

shifts = []
huidige_medewerker = None

for rij_index, row in raw.iterrows():
    eerste_cel = normaliseer_tekst(row.iloc[0])

    # Nieuwe medewerker begint met #
    if eerste_cel.startswith("#"):
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

# Strobbo zet soms 1 echte shift in meerdere vakken/rijen omdat de taak wijzigt.
# Voorbeeld:
# rij 1: 09:30-10:00 Aanv. C.
# rij 2: 10:00-14:45 Service
# Dit moet samen 1 shift worden.

if not shifts_df.empty:
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
               huidige["origineel"] = str(huidige.get("origineel", "")) + "\\n---\\n" + str(blok.get("origineel", ""))
                huidige["bron_rij"] = str(huidige.get("bron_rij", "")) + ", " + str(blok.get("bron_rij", ""))
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
# MATCHEN MET DATABASE
# =========================

match_resultaten = []

for naam in shifts_df["strobbo_naam"].unique():
    beste_naam, score = zoek_beste_match(naam, crew)
    match_resultaten.append({
        "Strobbo naam": naam,
        "Database naam": beste_naam if beste_naam else "NIET GEVONDEN",
        "Match score": round(score, 2)
    })

match_df = pd.DataFrame(match_resultaten)

match_map = {
    row["Strobbo naam"]: row["Database naam"]
    for _, row in match_df.iterrows()
    if row["Database naam"] != "NIET GEVONDEN"
}

shifts_df["database_naam"] = shifts_df["strobbo_naam"].map(match_map)


# =========================
# FOUTEN CONTROLEREN
# =========================

fouten = []

# Naam niet gevonden
for _, row in match_df.iterrows():
    if row["Database naam"] == "NIET GEVONDEN":
        voeg_fout(
            fouten,
            row["Strobbo naam"],
            "",
            "Naam niet gevonden",
            f"Geen goede match in database. Match score: {row['Match score']}",
            "Waarschuwing"
        )

# Per shift controles
for _, shift in shifts_df.dropna(subset=["database_naam"]).iterrows():
    naam = shift["database_naam"]
    datum = shift["datum"]

    persoon = crew[crew["VOLLEDIGE_NAAM"] == naam].iloc[0]

    leeftijd = persoon["LEEFTIJD"]
    netto_uren = shift["netto_uren"]
    bruto_uren = shift["bruto_uren"]
    pauze = shift["pauze_minuten"]

    # Minimum uren
    if netto_uren < MIN_DAGUREN:
        voeg_fout(
            fouten,
            naam,
            datum,
            "Shift te kort",
            f"{netto_uren:.2f}u gewerkt. Minimum is {MIN_DAGUREN}u."
        )

    # Maximum uren per dag
    if leeftijd is not None and leeftijd < 18:
        if netto_uren > MAX_DAGUREN_MINDERJARIG:
            voeg_fout(
                fouten,
                naam,
                datum,
                "Minderjarige werkt te lang",
                f"{netto_uren:.2f}u gewerkt. Maximum voor <18 is {MAX_DAGUREN_MINDERJARIG}u."
            )
    else:
        if netto_uren > MAX_DAGUREN_VOLWASSEN:
            voeg_fout(
                fouten,
                naam,
                datum,
                "Shift te lang",
                f"{netto_uren:.2f}u gewerkt. Maximum is {MAX_DAGUREN_VOLWASSEN}u."
            )

    # Pauzes volwassen / algemeen
    if leeftijd is None or leeftijd >= 18:
        if bruto_uren > 5 and bruto_uren <= 8 and pauze < 20:
            voeg_fout(
                fouten,
                naam,
                datum,
                "Pauze ontbreekt",
                f"{bruto_uren:.2f}u shift. Minstens 20 min pauze nodig. Geplande pauze: {pauze} min."
            )

        if bruto_uren > 8 and pauze < 30:
            voeg_fout(
                fouten,
                naam,
                datum,
                "Pauze ontbreekt",
                f"{bruto_uren:.2f}u shift. Minstens 30 min pauze nodig. Geplande pauze: {pauze} min."
            )

    # Pauzes minderjarigen
    if leeftijd is not None and leeftijd < 18:
        if bruto_uren > 4.5 and bruto_uren <= 6 and pauze < 30:
            voeg_fout(
                fouten,
                naam,
                datum,
                "Pauze minderjarige ontbreekt",
                f"{bruto_uren:.2f}u shift. Minderjarige heeft minstens 30 min pauze nodig. Geplande pauze: {pauze} min."
            )

        if bruto_uren > 6 and pauze < 60:
            voeg_fout(
                fouten,
                naam,
                datum,
                "Pauze minderjarige ontbreekt",
                f"{bruto_uren:.2f}u shift. Minderjarige heeft minstens 60 min pauze nodig. Geplande pauze: {pauze} min."
            )

    # Nachtwerk minderjarigen
    if leeftijd is not None:
        einduur = shift["einde"].hour + shift["einde"].minute / 60

        if leeftijd <= 15 and einduur > 20:
            voeg_fout(
                fouten,
                naam,
                datum,
                "Nachtwerk minderjarige",
                f"{leeftijd} jaar en werkt tot {shift['einde'].strftime('%H:%M')}. Max tot 20:00."
            )

        if 16 <= leeftijd < 18 and einduur > 23:
            voeg_fout(
                fouten,
                naam,
                datum,
                "Nachtwerk minderjarige",
                f"{leeftijd} jaar en werkt tot {shift['einde'].strftime('%H:%M')}. Max tot 23:00."
            )

# Weekuren + contracturen
for naam, groep in shifts_df.dropna(subset=["database_naam"]).groupby("database_naam"):
    persoon = crew[crew["VOLLEDIGE_NAAM"] == naam].iloc[0]

    leeftijd = persoon["LEEFTIJD"]
    medewerker_type = persoon["TYPE"]
    contracturen = persoon["CONTRACTUREN"]

    totaal_weekuren = groep["netto_uren"].sum()

    if leeftijd is not None and leeftijd < 18:
        if totaal_weekuren > MAX_WEEKUREN_MINDERJARIG:
            voeg_fout(
                fouten,
                naam,
                "",
                "Weekuren overschreden",
                f"{totaal_weekuren:.2f}u gewerkt. Maximum voor <18 is {MAX_WEEKUREN_MINDERJARIG}u."
            )
    else:
        if totaal_weekuren > MAX_WEEKUREN_VOLWASSEN:
            voeg_fout(
                fouten,
                naam,
                "",
                "Weekuren overschreden",
                f"{totaal_weekuren:.2f}u gewerkt. Maximum is {MAX_WEEKUREN_VOLWASSEN}u."
            )

    if medewerker_type == "vast" and contracturen > 0:
        if totaal_weekuren < contracturen:
            tekort = contracturen - totaal_weekuren
            voeg_fout(
                fouten,
                naam,
                "",
                "Contracturen niet gehaald",
                f"{totaal_weekuren:.2f}u gepland, contract is {contracturen:.2f}u. Tekort: {tekort:.2f}u."
            )

# Rust tussen shifts
for naam, groep in shifts_df.dropna(subset=["database_naam"]).groupby("database_naam"):
    groep = groep.sort_values("start").reset_index(drop=True)

    for i in range(1, len(groep)):
        vorige = groep.iloc[i - 1]
        huidige = groep.iloc[i]

        rusturen = (huidige["start"] - vorige["einde"]).total_seconds() / 3600

        if rusturen < MIN_RUSTUREN_TUSSEN_SHIFTS:
            voeg_fout(
                fouten,
                naam,
                huidige["datum"],
                "Te weinig rust",
                f"Slechts {rusturen:.2f}u rust tussen {vorige['einde'].strftime('%d/%m %H:%M')} en {huidige['start'].strftime('%d/%m %H:%M')}."
            )

fouten_df = pd.DataFrame(fouten)


# =========================
# OUTPUT
# =========================

st.subheader("✅ Naamkoppeling Strobbo ↔ Database")
st.dataframe(match_df, use_container_width=True)

st.subheader("📋 Gevonden shifts")
toon_shifts = shifts_df.copy()
toon_shifts["Start"] = toon_shifts["start"].dt.strftime("%d/%m/%Y %H:%M")
toon_shifts["Einde"] = toon_shifts["einde"].dt.strftime("%d/%m/%Y %H:%M")
toon_shifts["Netto uren"] = toon_shifts["netto_uren"].round(2)
toon_shifts["Bruto uren"] = toon_shifts["bruto_uren"].round(2)

st.dataframe(
    toon_shifts[[
        "strobbo_naam",
        "database_naam",
        "datum",
        "Start",
        "Einde",
        "pauze_minuten",
        "Bruto uren",
        "Netto uren",
        "bron_rij",
        "bron_kolom"
    ]],
    use_container_width=True
)

st.subheader("🚨 Foutenrapport")

if fouten_df.empty:
    st.success("Geen fouten gevonden volgens de ingestelde regels.")
else:
    st.error(f"{len(fouten_df)} fout(en) of waarschuwing(en) gevonden.")
    st.dataframe(fouten_df, use_container_width=True)

    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        fouten_df.to_excel(writer, index=False, sheet_name="Foutenrapport")
        match_df.to_excel(writer, index=False, sheet_name="Naamkoppeling")
        toon_shifts.to_excel(writer, index=False, sheet_name="Gevonden shifts")

    st.download_button(
        label="📥 Download foutenrapport als Excel",
        data=buffer.getvalue(),
        file_name="foutenrapport_strobbo.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

st.subheader("📊 Weekuren per medewerker")

weekuren = (
    shifts_df
    .dropna(subset=["database_naam"])
    .groupby("database_naam")["netto_uren"]
    .sum()
    .reset_index()
    .rename(columns={"database_naam": "Medewerker", "netto_uren": "Weekuren"})
)

weekuren["Weekuren"] = weekuren["Weekuren"].round(2)

weekuren = weekuren.merge(
    crew[["VOLLEDIGE_NAAM", "TYPE", "LEEFTIJD", "CONTRACTUREN"]],
    left_on="Medewerker",
    right_on="VOLLEDIGE_NAAM",
    how="left"
)

weekuren = weekuren.drop(columns=["VOLLEDIGE_NAAM"])

st.dataframe(weekuren, use_container_width=True)
