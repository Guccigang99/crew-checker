import re
from datetime import datetime, timedelta

import pandas as pd
import streamlit as st
from rapidfuzz import process, fuzz


st.set_page_config(page_title="Strobbo Rooster Checker", layout="wide")

st.title("📊 Strobbo Rooster Checker")
st.write("Upload je crew-database en je Strobbo weekrooster. De app toont automatisch de fouten.")


# =========================
# INSTELLINGEN
# =========================

MIN_DAGUREN = 2
MAX_DAGUREN_VOLWASSEN = 11
MAX_DAGUREN_MINDERJARIG = 8

MAX_WEEKUREN_VOLWASSEN = 50
MAX_WEEKUREN_MINDERJARIG = 40

MIN_RUSTUREN_TUSSEN_SHIFTS = 10
FUZZY_MATCH_SCORE = 75


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
    naam = naam.replace("<18", "")
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
    except:
        return 0


def veilige_int(x):
    try:
        if pd.isna(x):
            return None
        return int(float(str(x).replace(",", ".")))
    except:
        return None


def zoek_beste_match(strobbo_naam, crew_namen):
    strobbo_norm = normaliseer_naam(strobbo_naam)
    crew_norms = [normaliseer_naam(n) for n in crew_namen]

    match = process.extractOne(
        strobbo_norm,
        crew_norms,
        scorer=fuzz.token_sort_ratio
    )

    if not match:
        return None, 0

    gematchte_norm, score, index = match

    if score >= FUZZY_MATCH_SCORE:
        return crew_namen[index], score

    return None, score


def parse_datum(datum_waarde):
    if isinstance(datum_waarde, datetime):
        return datum_waarde.date()

    tekst = str(datum_waarde).strip().lower()

    maand_map = {
        "jan": 1, "feb": 2, "mrt": 3, "mar": 3, "apr": 4,
        "mei": 5, "jun": 6, "jul": 7, "aug": 8,
        "sep": 9, "okt": 10, "oct": 10, "nov": 11, "dec": 12
    }

    match = re.search(r"(\d{1,2})[-/\s]([a-zA-Z]+)", tekst)

    if match:
        dag = int(match.group(1))
        maand_txt = match.group(2)[:3]
        maand = maand_map.get(maand_txt, None)

        if maand:
            return datetime(2026, maand, dag).date()

    return None


def parse_shift(cell_text, datum):
    """
    Verwacht cellen zoals:
    14:00-22:00 (00:20)
    Roeselare - Keuken
    """

    if pd.isna(cell_text) or not datum:
        return []

    tekst = str(cell_text)
    shifts = []

    patronen = re.findall(
        r"(\d{1,2}:\d{2})\s*-\s*(\d{1,2}:\d{2})\s*\((\d{1,2}:\d{2})\)",
        tekst
    )

    for start_txt, einde_txt, pauze_txt in patronen:
        start_dt = datetime.combine(datum, datetime.strptime(start_txt, "%H:%M").time())
        einde_dt = datetime.combine(datum, datetime.strptime(einde_txt, "%H:%M").time())

        if einde_dt <= start_dt:
            einde_dt += timedelta(days=1)

        pauze_uren, pauze_min = map(int, pauze_txt.split(":"))
        pauze_minuten = pauze_uren * 60 + pauze_min

        bruto_uren = (einde_dt - start_dt).total_seconds() / 3600
        netto_uren = bruto_uren - (pauze_minuten / 60)

        shifts.append({
            "datum": datum,
            "start": start_dt,
            "einde": einde_dt,
            "pauze_minuten": pauze_minuten,
            "bruto_uren": bruto_uren,
            "netto_uren": netto_uren,
            "origineel": tekst
        })

    return shifts


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

crew["VOLLEDIGE_NAAM"] = crew.apply(maak_volledige_naam, axis=1)
crew["TYPE"] = crew.apply(bepaal_type, axis=1)
crew["LEEFTIJD"] = crew["LFTD"].apply(veilige_int)
crew["CONTRACTUREN"] = crew["CONTRACT. UREN"].apply(veilige_float)

crew = crew[crew["VOLLEDIGE_NAAM"].str.strip() != ""]

crew_namen = crew["VOLLEDIGE_NAAM"].tolist()


# =========================
# STROBBO ROOSTER INLEZEN
# =========================

try:
    raw = pd.read_excel(rooster_file, sheet_name=0, header=None)
except Exception as e:
    st.error(f"Kon Strobbo-rooster niet lezen: {e}")
    st.stop()

# Datums zoeken op rij 2 zoals in je export
datum_rij = raw.iloc[2].tolist()

dag_kolommen = {}

for col_index, waarde in enumerate(datum_rij):
    datum = parse_datum(waarde)
    if datum:
        dag_kolommen[col_index] = datum


if not dag_kolommen:
    st.error("Ik kon geen datums vinden in het Strobbo-bestand.")
    st.stop()


# =========================
# SHIFTS EXTRACTEN
# =========================

shifts = []
huidige_medewerker = None

for rij_index, row in raw.iterrows():
    eerste_cel = normaliseer_tekst(row.iloc[0])

    if eerste_cel.startswith("#"):
        huidige_medewerker = eerste_cel.replace("#", "").replace("<18", "").strip()

    if not huidige_medewerker:
        continue

    for col_index, datum in dag_kolommen.items():
        cel = row.iloc[col_index]
        gevonden_shifts = parse_shift(cel, datum)

        for shift in gevonden_shifts:
            shift["strobbo_naam"] = huidige_medewerker
            shifts.append(shift)

shifts_df = pd.DataFrame(shifts)

if shifts_df.empty:
    st.warning("Geen shifts gevonden in het Strobbo-bestand.")
    st.stop()


# =========================
# MATCHEN MET DATABASE
# =========================

match_resultaten = []

for naam in shifts_df["strobbo_naam"].unique():
    beste_naam, score = zoek_beste_match(naam, crew_namen)
    match_resultaten.append({
        "Strobbo naam": naam,
        "Database naam": beste_naam if beste_naam else "NIET GEVONDEN",
        "Match score": score
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

def voeg_fout(naam, datum, fouttype, detail, ernst="Fout"):
    fouten.append({
        "Medewerker": naam,
        "Datum": datum,
        "Ernst": ernst,
        "Fout": fouttype,
        "Detail": detail
    })


# Naam niet gevonden
for _, row in match_df.iterrows():
    if row["Database naam"] == "NIET GEVONDEN":
        voeg_fout(
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
    medewerker_type = persoon["TYPE"]

    netto_uren = shift["netto_uren"]
    bruto_uren = shift["bruto_uren"]
    pauze = shift["pauze_minuten"]

    # Minimum uren
    if netto_uren < MIN_DAGUREN:
        voeg_fout(
            naam,
            datum,
            "Shift te kort",
            f"{netto_uren:.2f}u gewerkt. Minimum is {MIN_DAGUREN}u."
        )

    # Maximum uren per dag
    if leeftijd is not None and leeftijd < 18:
        if netto_uren > MAX_DAGUREN_MINDERJARIG:
            voeg_fout(
                naam,
                datum,
                "Minderjarige werkt te lang",
                f"{netto_uren:.2f}u gewerkt. Maximum voor <18 is {MAX_DAGUREN_MINDERJARIG}u."
            )
    else:
        if netto_uren > MAX_DAGUREN_VOLWASSEN:
            voeg_fout(
                naam,
                datum,
                "Shift te lang",
                f"{netto_uren:.2f}u gewerkt. Maximum is {MAX_DAGUREN_VOLWASSEN}u."
            )

    # Pauzes volwassen / algemeen
    if leeftijd is None or leeftijd >= 18:
        if bruto_uren > 5 and bruto_uren <= 8 and pauze < 20:
            voeg_fout(
                naam,
                datum,
                "Pauze ontbreekt",
                f"{bruto_uren:.2f}u shift. Minstens 20 min pauze nodig."
            )

        if bruto_uren > 8 and pauze < 30:
            voeg_fout(
                naam,
                datum,
                "Pauze ontbreekt",
                f"{bruto_uren:.2f}u shift. Minstens 30 min pauze nodig."
            )

    # Pauzes minderjarigen
    if leeftijd is not None and leeftijd < 18:
        if bruto_uren > 4.5 and bruto_uren <= 6 and pauze < 30:
            voeg_fout(
                naam,
                datum,
                "Pauze minderjarige ontbreekt",
                f"{bruto_uren:.2f}u shift. Minderjarige heeft minstens 30 min pauze nodig."
            )

        if bruto_uren > 6 and pauze < 60:
            voeg_fout(
                naam,
                datum,
                "Pauze minderjarige ontbreekt",
                f"{bruto_uren:.2f}u shift. Minderjarige heeft minstens 60 min pauze nodig."
            )

    # Nachtwerk minderjarigen
    if leeftijd is not None:
        einduur = shift["einde"].hour + shift["einde"].minute / 60

        if leeftijd <= 15 and einduur > 20:
            voeg_fout(
                naam,
                datum,
                "Nachtwerk minderjarige",
                f"{leeftijd} jaar en werkt tot {shift['einde'].strftime('%H:%M')}. Max tot 20:00."
            )

        if 16 <= leeftijd < 18 and einduur > 23:
            voeg_fout(
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
                naam,
                "",
                "Weekuren overschreden",
                f"{totaal_weekuren:.2f}u gewerkt. Maximum voor <18 is {MAX_WEEKUREN_MINDERJARIG}u."
            )
    else:
        if totaal_weekuren > MAX_WEEKUREN_VOLWASSEN:
            voeg_fout(
                naam,
                "",
                "Weekuren overschreden",
                f"{totaal_weekuren:.2f}u gewerkt. Maximum is {MAX_WEEKUREN_VOLWASSEN}u."
            )

    if medewerker_type == "vast" and contracturen > 0:
        if totaal_weekuren < contracturen:
            tekort = contracturen - totaal_weekuren
            voeg_fout(
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

st.dataframe(
    toon_shifts[[
        "strobbo_naam",
        "database_naam",
        "datum",
        "Start",
        "Einde",
        "pauze_minuten",
        "Netto uren"
    ]],
    use_container_width=True
)

st.subheader("🚨 Foutenrapport")

if fouten_df.empty:
    st.success("Geen fouten gevonden volgens de ingestelde regels.")
else:
    st.error(f"{len(fouten_df)} fout(en) of waarschuwing(en) gevonden.")
    st.dataframe(fouten_df, use_container_width=True)

    excel_export = fouten_df.to_excel(index=False)

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
