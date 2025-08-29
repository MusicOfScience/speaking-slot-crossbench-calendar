"""
calendar_utils.py – pure-Python helper for the Streamlit app.
Converts the Crossbench roster into:
  • pandas DataFrame
  • CSV bytes
  • ZIP bytes containing .ics files
"""
import pandas as pd, re, io, datetime as dt, zipfile
from datetime import timedelta

# ── tiny helpers ───────────────────────────────────────────────
def _escape_ics(txt):
    if pd.isna(txt): return ''
    return (str(txt)
            .replace("\\", "\\\\").replace(",", "\\,")
            .replace(";",  "\\;").replace("\n", "\\n"))

def _normalise_notes(cols):
    for i, c in enumerate(cols):
        if isinstance(c, str) and re.fullmatch(r"\s*notes?\s*", c, re.I):
            cols[i] = "Notes"; return cols
    cols[-1] = "Notes"
    return cols

def _coerce_date(val, fallback_year="2025"):
    if isinstance(val, (pd.Timestamp, dt.datetime)):
        return pd.Timestamp(val.date())
    if pd.isna(val) or str(val).strip() == '':
        return pd.NaT
    try:
        return pd.to_datetime(val, dayfirst=True, errors='raise')
    except Exception:
        pass
    m = re.search(r"(\d{1,2})\s*[-/ ]\s*([A-Za-z]{3,9})", str(val))
    if m:
        txt = f"{m.group(1)}-{m.group(2).title()} {fallback_year}"
        return pd.to_datetime(txt, dayfirst=True, errors='coerce')
    return pd.NaT

def _detect_location(event_txt, note_txt=""):
    blob = f"{event_txt} {note_txt}".lower()
    if "fed chamber" in blob: return "Federation Chamber"
    if "chamber" in blob:     return "House of Representatives Chamber"
    return "Unknown"

# ── public function ────────────────────────────────────────────
def parse_roster(file_like, default_year="2025"):
    """
    Parameters
    ----------
    file_like : UploadedFile (Streamlit) – .xlsx/.csv roster
    Returns
    -------
    events_df  : pandas.DataFrame
    csv_bytes  : bytes
    ics_zip    : bytes
    """
    # read Excel or CSV
    df = (pd.read_excel(file_like, header=None)
          if file_like.name.endswith(("xlsx", "xlsm"))
          else pd.read_csv(file_like, header=None))

    weekday_col, date_col = "Weekday", "Date"
    headers = _normalise_notes(df.iloc[1].fillna('').tolist())
    data    = df.iloc[2:].copy()
    data.columns = [weekday_col, date_col] + headers[2:]

    data[date_col] = data[date_col].map(lambda v: _coerce_date(v, default_year))
    data = data.dropna(subset=[date_col])

    id_vars = [weekday_col, date_col, "Notes"]
    melted  = data.melt(id_vars=id_vars, var_name="Event Type", value_name="MP")
    events  = melted[melted["MP"].str.contains(r"\bRyan\b", case=False, na=False)].copy()
    events["Location"] = events.apply(
        lambda r: _detect_location(r["Event Type"], r["Notes"]), axis=1
    )
    events.sort_values(date_col, inplace=True)

    # CSV bytes
    csv_buf = io.StringIO()
    events.to_csv(csv_buf, index=False)
    csv_bytes = csv_buf.getvalue().encode()

    # ZIP of .ics files
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        for _, row in events.iterrows():
            start = row[date_col].replace(hour=9, minute=0)
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
                f"SUMMARY:{_escape_ics(summary)}",
                f"DESCRIPTION:{_escape_ics(description)}",
                f"LOCATION:{_escape_ics(row['Location'])}",
                ("ATTENDEE;CN=Rosie Leonthomas;ROLE=REQ-PARTICIPANT;"
                 "PARTSTAT=NEEDS-ACTION;RSVP=TRUE:mailto:Rosie.Leonthomas@aph.gov.au"),
                "END:VEVENT",
                "END:VCALENDAR"
            ])
            safe_name = re.sub(r"[^\w\- ]", "_", f"{start.strftime('%Y-%m-%d')} – {title}")
            zf.writestr(f"{safe_name}.ics", ics_text)
    zip_buf.seek(0)

    return events, csv_bytes, zip_buf.read()
