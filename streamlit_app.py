"""
Size Chart Extractor - web app (Streamlit)
Upload a tech-pack PDF, download an Excel size chart.
The extraction engine lives in sizechart_app.py (imported below).
"""
import io
import os
import tempfile
import streamlit as st

from sizechart_app import convert  # reuses the same engine as the desktop app

st.set_page_config(page_title="Size Chart Extractor", page_icon="📏", layout="centered")

st.title("📏 Size Chart Extractor")
st.write(
    "Upload a tech-pack PDF and download an Excel size chart: the standard size "
    "run (2XS–4XL) with a grade-difference (Δ) column after each size, plus a "
    "sheet for the logo measurement set."
)
st.info("Your file is processed in memory to build the Excel and is not stored.", icon="🔒")

pdf = st.file_uploader("Choose a tech-pack PDF", type=["pdf"])

if pdf is not None:
    if st.button("Convert to Excel", type="primary"):
        status = st.empty()
        with st.spinner("Reading the PDF and building the chart..."):
            with tempfile.TemporaryDirectory() as tmp:
                in_path = os.path.join(tmp, "input.pdf")
                out_path = os.path.join(tmp, "size_chart.xlsx")
                with open(in_path, "wb") as f:
                    f.write(pdf.getbuffer())
                try:
                    sheets, n = convert(in_path, out_path,
                                        progress=lambda m: status.write(m))
                    with open(out_path, "rb") as f:
                        data = f.read()
                except Exception as e:
                    st.error(f"Could not convert this PDF: {e}")
                    st.stop()
        status.empty()
        st.success(f"Done — {n} measurement points across {len(sheets)} sheet(s).")
        base = os.path.splitext(pdf.name)[0]
        st.download_button(
            "⬇️ Download Excel",
            data=data,
            file_name=f"{base}_size_chart.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
        st.caption("Open in Excel or Google Sheets — the Δ columns calculate on open.")

st.divider()
st.caption("Works on PDFs with real text tables. Scanned image-only PDFs need OCR and won't extract.")
