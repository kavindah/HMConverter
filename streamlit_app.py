"""
Tech-Pack Size Chart Extractor - web app (Streamlit)
Upload a tech-pack PDF, download an Excel size chart.

Privacy: the PDF is processed entirely in memory (never written to the
server's disk), the uploaded bytes are released as soon as the Excel is
built, and a "Clear" button wipes the result and the uploaded file.
"""
import io
import gc
import streamlit as st

from sizechart_app import convert_to_bytes  # in-memory engine

st.set_page_config(page_title="Size Chart Extractor", page_icon="📏", layout="centered")

# Hide Streamlit chrome: hamburger menu, "Made with Streamlit" footer, header bar
st.markdown(
    """
    <style>
      /* remove the entire top header (GitHub/fork badge, deploy, menu, status) */
      [data-testid="stHeader"] {display: none !important;}
      [data-testid="stToolbar"] {display: none !important;}
      [data-testid="stToolbarActions"] {display: none !important;}
      [data-testid="stAppDeployButton"] {display: none !important;}
      [data-testid="stMainMenu"] {display: none !important;}
      #MainMenu {visibility: hidden !important;}
      header {visibility: hidden !important;}
      footer {visibility: hidden !important;}
    </style>
    """,
    unsafe_allow_html=True,
)
st.title("📏 Tech-Pack Size Chart Extractor")
st.write(
    "Upload a tech-pack PDF and download an Excel size chart: the standard size "
    "run (2XS–5XL) with a grade-difference (Δ) column after each size, plus a "
    "sheet for the logo measurement set."
)
st.info(
    "Your PDF is processed in memory and is **never saved to the server**. "
    "It is cleared from memory as soon as the Excel is built.",
    icon="🔒",
)

# a counter we put in the uploader key so we can fully reset (drop) the upload
if "uploader_id" not in st.session_state:
    st.session_state.uploader_id = 0

pdf = st.file_uploader(
    "Choose a tech-pack PDF",
    type=["pdf"],
    key=f"pdf_{st.session_state.uploader_id}",
    help=(
        "🔒 Your document is secure:\n\n"
        "• Processed entirely in RAM — never written to disk\n"
        "• Raw PDF bytes are deleted from memory the moment the Excel is built\n"
        "• No copy is kept on the server between sessions\n"
        "• Use 'Clear file from this session' to wipe the result immediately"
    ),
)

if pdf is not None and st.button("Convert to Excel", type="primary"):
    status = st.empty()
    buf = io.BytesIO(pdf.getvalue())  # copy bytes into memory only
    try:
        with st.spinner("Reading the PDF and building the chart..."):
            data, sheets, n = convert_to_bytes(buf, progress=lambda m: status.write(m))
        st.session_state.xlsx = data
        st.session_state.fname = pdf.name.rsplit(".", 1)[0] + "_size_chart.xlsx"
        st.session_state.summary = f"{n} measurement points across {len(sheets)} sheet(s)."
    except Exception as e:
        st.error(f"Could not convert this PDF: {e}")
    finally:
        # release the in-memory PDF immediately
        buf.close()
        del buf
        gc.collect()
    status.empty()

if "xlsx" in st.session_state:
    st.success("Done — " + st.session_state.summary)
    st.download_button(
        "⬇️ Download Excel",
        data=st.session_state.xlsx,
        file_name=st.session_state.fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
    st.caption("Open in Excel or Google Sheets — the Δ columns calculate on open.")

    if st.button("Clear file from this session"):
        # wipe the result and drop the uploaded file from the widget
        for k in ("xlsx", "fname", "summary"):
            st.session_state.pop(k, None)
        st.session_state.uploader_id += 1   # new key => uploader forgets the PDF
        gc.collect()
        st.rerun()

st.divider()
st.caption("Works on PDFs with real text tables. Scanned image-only PDFs need OCR and won't extract.")
