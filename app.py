import streamlit as st
from calendar_utils import parse_roster

st.set_page_config(page_title="Monique Ryan Calendar Builder", layout="centered")
st.title("📅 Monique Ryan Crossbench Roster → CSV + ICS")

uploaded = st.file_uploader(
    "Upload the Crossbench Calendar roster (Excel or CSV)",
    type=["xlsx", "xls", "csv"]
)

if uploaded:
    try:
        events, csv_bytes, zip_bytes = parse_roster(uploaded)
        st.success(f"Parsed {len(events)} Ryan speaking slots.")
        st.dataframe(events.style.format({"Date": lambda d: d.strftime("%d %b %Y")}), height=420)
        st.download_button("⬇️ CSV", csv_bytes, "monique_ryan_calendar.csv", "text/csv")
        st.download_button("⬇️ ICS ZIP", zip_bytes,
                           "monique_ryan_calendar_items_clean.zip", "application/zip")
    except Exception as e:
        st.error(f"Something went wrong: {e}")
else:
    st.info("Drag & drop the roster file above to begin.")
