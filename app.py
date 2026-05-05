import re
from datetime import datetime, timedelta
from io import BytesIO

import pandas as pd
import streamlit as st
from rapidfuzz import process, fuzz

st.set_page_config(page_title="Strobbo Checker", layout="wide")
st.title("📊 Strobbo Rooster Checker")

# ======================
# INSTELLINGEN
# ======================
MERGE_GAP_MINUTEN = 5

# ======================
# HELPERS
# ======================

def normaliseer_naam(x):
    x = str(x).lower().strip()
    x = x.replace("#", "")
    x = re.sub(r"<\d+", "", x)
    x = x.replace(".", "")
    return x.strip()

def parse_datum(x):
    maanden = {
        "jan":1,"feb":2,"mrt":3,"apr":4,"mei":5,
        "jun":6,"jul":7,"aug":8,"sep":9,"okt":10,"nov":11,"dec":12
    }

    x = str(x).lower()
    match = re.search(r"(\d{1,2})[- ]([a-z]+)", x)

    if match:
        dag = int(match.group(1))
        maand = maanden.get(match.group(2)[:3])
        if maand:
            return datetime(2026, maand, dag).date()

    return None

def parse_shift_blokken(text, datum):
    if pd.isna(text):
        return []

    pattern = r"(\d{1,2}:\d{2})-(\d{1,2}:\d{2}).*?\((\d{2}:\d{2})\)"
    matches = re.findall(pattern, str(text), re.S)

    blokken = []

    for start, einde, pauze in matches:
        start_dt = datetime.combine(datum, datetime.strptime(start, "%H:%M").time())
        end_dt = datetime.combine(datum, datetime.strptime(einde, "%H:%M").time())

        if end_dt <= start_dt:
            end_dt += timedelta(days=1)

        p_h, p_m = map(int, pauze.split(":"))
        pauze_min = p_h*60 + p_m

        blokken.append({
            "start": start_dt,
            "einde": end_dt,
            "pauze": pauze_min
        })

    return blokken

def merge_blokken(blokken):
    if not blokken:
        return []

    blokken = sorted(blokken, key=lambda x: x["start"])
    merged = [blokken[0]]

    for b in blokken[1:]:
        last = merged[-1]
        gap = (b["start"] - last["einde"]).total_seconds()/60

        if gap <= MERGE_GAP_MINUTEN:
            last["einde"] = max(last["einde"], b["einde"])
            last["pauze"] += b["pauze"]
        else:
            merged.append(b)

    return merged

def fuzzy_match(naam, crew_df):
    naam = normaliseer_naam(naam)

    # exact voornaam
    for _, r in crew_df.iterrows():
        if naam == normaliseer_naam(r["VOORNAAM"]):
            return r["VOLLEDIGE_NAAM"]

    # fuzzy
    lijst = crew_df["VOLLEDIGE_NAAM"].tolist()
    match = process.extractOne(naam, lijst, scorer=fuzz.token_sort_ratio)

    if match and match[1] > 65:
        return match[0]

    return None

# ======================
# UPLOAD
# ======================

crew_file = st.file_uploader("Crew database", type=["xlsx"])
rooster_file = st.file_uploader("Strobbo export", type=["xlsx"])

if not crew_file or not rooster_file:
    st.stop()

# ======================
# CREW
# ======================

crew = pd.read_excel(crew_file)
crew.columns = [c.upper() for c in crew.columns]

crew["VOLLEDIGE_NAAM"] = crew["VOORNAAM"] + " " + crew["NAAM"]

# ======================
# STROBBO
# ======================

raw = pd.read_excel(rooster_file, header=None)

dag_kol = {}
for r in range(5):
    for c in range(raw.shape[1]):
        d = parse_datum(raw.iloc[r, c])
        if d:
            dag_kol[c] = d

shifts = []
huidige = None

for i, row in raw.iterrows():
    cel = str(row.iloc[0])

    if cel.startswith("#"):
        huidige = cel.replace("#", "")

    if not huidige:
        continue

    for col, datum in dag_kol.items():
        text = row.iloc[col]

        blokken = parse_shift_blokken(text, datum)
        blokken = merge_blokken(blokken)

        for b in blokken:
            shifts.append({
                "naam": huidige,
                "datum": datum,
                "start": b["start"],
                "einde": b["einde"],
                "pauze": b["pauze"]
            })

df = pd.DataFrame(shifts)

# ======================
# MATCH
# ======================

df["match"] = df["naam"].apply(lambda x: fuzzy_match(x, crew))

# ======================
# OUTPUT
# ======================

df["uren"] = (df["einde"] - df["start"]).dt.total_seconds()/3600 - (df["pauze"]/60)

st.subheader("Shifts")
st.dataframe(df)

fouten = df[df["uren"] < 2]

st.subheader("Fouten")
st.dataframe(fouten)

# download
buffer = BytesIO()
df.to_excel(buffer, index=False)

st.download_button("Download Excel", buffer.getvalue(), "output.xlsx")
