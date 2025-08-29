# (paste the full helper module here – exactly the version you tested in Colab)
# ↓  Paste this *entire* cell into Colab and run ↓
from google.colab import files
import pandas as pd, re, os, datetime as dt
from datetime import timedelta
from zipfile import ZipFile

# ───────────────────────── helpers ─────────────────────────
def escape_ics(txt):
    """ICS-safe text."""
    if pd.isna(txt): return ''
    return (str(txt)
            .replace("\\", "\\\\")
            .replace(",",  "\\,")
            .replace(";",  "\\;")
            .replace("\n", "\\n"))

def normalise_notes_header(cols):
    """Force exactly one 'Notes' column header."""
    for i, c in enumerate(cols):
        if isinstance(c, str) and re.fullmatch(r"\s*notes?\s*", c, re.I):
            cols[i] = 'Notes'
            return cols
    cols[-1] = 'Notes'          # fallback = last column
    return cols

def coerce_date(val, fallback_year="2025"):
    """Accept Excel datetimes, or text like 12-Jan, return Timestamp."""
    if isinstance(val, (pd.Timestamp, dt.datetime)):
        return pd.Timestamp(val.date())
    if pd.isna(val) or str(val).strip() == '':
        return pd.NaT
    # try pandas' own parser first
    try:
        return pd.to_datetime(val, dayfirst=True, errors='raise')
    except Exception:
        pass
    # fallback regex
    m = re.search(r"(\d{1,2})\s*[-/ ]\s*([A-Za-z]{3,9})", str(val))
    if m:
        txt = f"{m.group(1)}-{m.group(2).title()} {fallback_year}"
        return pd.to_datetime(txt, dayfirst=True, errors='coerce')
    return pd.NaT

def detect_location(event_txt, note_txt=""):
    blob = f"{event_txt} {note_txt}".lower()
    if "fed chamber" in blob:
        return "Federation Chamber"
    if "chamber" in blob:
        return "House of Representatives Chamber"
    return "Unknown"

# ──────────────────────── main routine ─────────────────────
def build_monique_calendar(path, *, default_year="2025", out_dir="/content"):
    print("📂  Processing", path)
    ext  = os.path.splitext(path)[-1].lower()
    df   = pd.read_excel(path, header=None) if ext == ".xlsx" else pd.read_csv(path, header=None)

    # impose headers (row-0 weekdays/dates, row-1 event labels)
    weekday_col, date_col = "Weekday", "Date"
    event_headers   = normalise_notes_header(df.iloc[1].fillna('').tolist())
    data            = df.iloc[2:].copy()
    data.columns    = [weekday_col, date_col] + event_headers[2:]

    # coerce dates, drop blanks
    data[date_col]  = data[date_col].map(lambda v: coerce_date(v, fallback_year=default_year))
    data            = data.dropna(subset=[date_col])

    # melt to long & keep Ryan slots
    id_vars         = [weekday_col, date_col, "Notes"]
    melted          = data.melt(id_vars=id_vars, var_name="Event Type", value_name="MP")
    events          = melted[melted["MP"].str.contains(r"\bRyan\b", case=False, na=False)].copy()
    if events.empty:
        raise RuntimeError("❌ No Ryan entries found – check the file.")
    events["Location"] = events.apply(lambda r: detect_location(r["Event Type"], r["Notes"]), axis=1)
    events.sort_values(date_col, inplace=True)

    # export CSV
    csv_path = os.path.join(out_dir, "monique_ryan_calendar.csv")
    events.to_csv(csv_path, index=False)

    # generate .ics files
    ics_dir = os.path.join(out_dir, "monique_ryan_calendar_items_clean")
    os.makedirs(ics_dir, exist_ok=True)
    ics_paths = []

    for _, row in events.iterrows():
        start = row[date_col].replace(hour=9, minute=0)  # default 09:00 AEST
        end   = start + timedelta(hours=1)
        title = row["Event Type"].splitlines()[0].strip()
        summary = f"{title} – {start.strftime('%d %b')}"
        description = "\n".join(filter(None, [
            row["Event Type"].strip(),
            "",
            "-----------------------------",
            f"Speaker: {row['MP']}",
            f"Notes: {row['Notes']}" if pd.notna(row["Notes"]) else ""
        ]))

        ics_text = "\n".join([
            "BEGIN:VCALENDAR",
            "VERSION:2.0",
            "PRODID:-//Monique Ryan Calendar//EN",
            "BEGIN:VEVENT",
            f"DTSTART;TZID=Australia/Sydney:{start.strftime('%Y%m%dT%H%M%S')}",
            f"DTEND;TZID=Australia/Sydney:{end.strftime('%Y%m%dT%H%M%S')}",
            f"SUMMARY:{escape_ics(summary)}",
            f"DESCRIPTION:{escape_ics(description)}",
            f"LOCATION:{escape_ics(row['Location'])}",
            ("ATTENDEE;CN=Rosie Leonthomas;ROLE=REQ-PARTICIPANT;"
             "PARTSTAT=NEEDS-ACTION;RSVP=TRUE:mailto:Rosie.Leonthomas@aph.gov.au"),
            "END:VEVENT",
            "END:VCALENDAR"
        ])

        fname  = re.sub(r'[^\w\- ]', '_', f"{start.strftime('%Y-%m-%d')} – {title}")
        ics_fp = os.path.join(ics_dir, f"{fname}.ics")
        with open(ics_fp, "w") as f: f.write(ics_text)
        ics_paths.append(ics_fp)

    zip_path = os.path.join(out_dir, "monique_ryan_calendar_items_clean.zip")
    with ZipFile(zip_path, "w") as zf:
        for p in ics_paths:
            zf.write(p, arcname=os.path.basename(p))

    return csv_path, zip_path

# ──────────────────────── run interactively ─────────────────
uploaded = files.upload()           # choose your .xlsx or .csv roster
src      = next(iter(uploaded))
csv_out, zip_out = build_monique_calendar(src)
files.download(csv_out)
files.download(zip_out)
print("✅  Done – CSV & ZIP downloaded.")
