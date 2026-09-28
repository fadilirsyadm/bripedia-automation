from __future__ import annotations

import base64
import hashlib
import html
from datetime import date, timedelta

import streamlit as st
import streamlit.components.v1 as components

from core import (
    generate_outputs,
    get_regional_offices,
    prepare_master,
    read_input_excels,
)


st.set_page_config(
    page_title="BRIPEDIA Report Generator",
    page_icon="📘",
    layout="wide",
)

st.markdown(
    """
    <style>
      .block-container {max-width: 1180px; padding-top: 2rem; padding-bottom: 4rem;}
      .step-title {font-size: 0.92rem; font-weight: 700; margin-bottom: .3rem;}
      div[data-testid="stFileUploader"] {padding-bottom: .2rem;}
      .result-box {padding: 1rem 1.1rem; border: 1px solid rgba(128,128,128,.25); border-radius: .75rem;}
    </style>
    """,
    unsafe_allow_html=True,
)


def previous_month_end(today: date) -> date:
    return today.replace(day=1) - timedelta(days=1)


def bytes_signature(*items: bytes) -> str:
    h = hashlib.sha256()
    for item in items:
        h.update(item)
    return h.hexdigest()


def auto_download_two_files(zip_bytes: bytes, zip_name: str, hist_bytes: bytes, hist_name: str):
    """
    Best-effort automatic download. Browsers may block multiple automatic downloads;
    the normal Streamlit download buttons remain as a fallback.
    """
    # Avoid embedding extremely large payloads in an iframe.
    max_total = 20 * 1024 * 1024
    if len(zip_bytes) + len(hist_bytes) > max_total:
        return

    zip_b64 = base64.b64encode(zip_bytes).decode("ascii")
    hist_b64 = base64.b64encode(hist_bytes).decode("ascii")
    zip_name = html.escape(zip_name, quote=True)
    hist_name = html.escape(hist_name, quote=True)

    components.html(
        f"""
        <html><body>
          <a id="zip" download="{zip_name}" href="data:application/zip;base64,{zip_b64}"></a>
          <a id="hist" download="{hist_name}" href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{hist_b64}"></a>
          <script>
            setTimeout(() => document.getElementById('zip').click(), 250);
            setTimeout(() => document.getElementById('hist').click(), 900);
          </script>
        </body></html>
        """,
        height=0,
        scrolling=False,
    )


st.title("📘 BRIPEDIA Report Generator")
st.caption(
    "Upload 4 file sumber, atur konfigurasi report, pilih Regional Office, lalu proses. "
    "Blok komentar `# # REQUEST ...` dari notebook tidak diterapkan."
)

st.subheader("1. Unggah data")
col1, col2 = st.columns(2)
with col1:
    data_uker_file = st.file_uploader(
        "data_uker (.xlsx)",
        type=["xlsx"],
        help="Dibaca dari sheet 'Data Uker' dengan header pada baris ke-2 (header=1), sesuai notebook.",
    )
    data_bripedia_file = st.file_uploader("data_bripedia (.xlsx)", type=["xlsx"])
with col2:
    data_hc_file = st.file_uploader("data_hc (.xlsx)", type=["xlsx"])
    data_historis_file = st.file_uploader("data_historis (.xlsx)", type=["xlsx"])

all_uploaded = all([data_uker_file, data_hc_file, data_bripedia_file, data_historis_file])

today = date.today()
st.subheader("2–4. Data Configuration")
cfg1, cfg2, cfg3 = st.columns(3)
with cfg1:
    current_year = st.number_input(
        "current_year",
        min_value=2000,
        max_value=2100,
        value=today.year,
        step=1,
    )
with cfg2:
    tanggal_data_hc = st.date_input(
        "tanggal_data_hc",
        value=previous_month_end(today),
        format="DD/MM/YYYY",
    )
with cfg3:
    tanggal_report = st.date_input(
        "tanggal report",
        value=today,
        format="DD/MM/YYYY",
    )

if tanggal_data_hc > tanggal_report:
    st.warning("tanggal_data_hc lebih besar daripada tanggal report. Periksa kembali konfigurasi tanggal.")

# Read and prepare as soon as all files exist so Regional Office can be a dropdown.
data_master = None
data_historis_raw = None
regional_options = []
input_signature = None

if all_uploaded:
    try:
        uker_bytes = data_uker_file.getvalue()
        hc_bytes = data_hc_file.getvalue()
        bripedia_bytes = data_bripedia_file.getvalue()
        historis_bytes = data_historis_file.getvalue()
        input_signature = bytes_signature(uker_bytes, hc_bytes, bripedia_bytes, historis_bytes)

        with st.spinner("Membaca dan memvalidasi 4 file..."):
            data_uker_raw, data_hc_raw, data_bripedia_raw, data_historis_raw = read_input_excels(
                uker_bytes, hc_bytes, bripedia_bytes, historis_bytes
            )
            data_master = prepare_master(data_uker_raw, data_hc_raw, data_bripedia_raw)
            regional_options = get_regional_offices(data_master)

        if not regional_options:
            st.error("Tidak ada Regional Office yang berhasil terbentuk dari data upload.")
    except Exception as exc:
        st.error(f"Gagal membaca / menyiapkan data: {exc}")
        st.stop()
else:
    st.info("Lengkapi keempat file Excel terlebih dahulu agar dropdown Regional Office dapat dibentuk.")

st.subheader("5. Regional Office")
regional_office = st.selectbox(
    "regional_office",
    options=regional_options,
    index=0 if regional_options else None,
    placeholder="Upload 4 file terlebih dahulu",
    disabled=not bool(regional_options),
)

st.subheader("6. Process")
process_clicked = st.button(
    "⚙️ Process",
    type="primary",
    use_container_width=True,
    disabled=not (all_uploaded and regional_office),
)

if process_clicked:
    try:
        with st.spinner("Memproses report, historical file, chart, dan ZIP..."):
            outputs = generate_outputs(
                data_master=data_master,
                data_historis_raw=data_historis_raw,
                regional_office=regional_office,
                current_year=int(current_year),
                tanggal_data_hc_value=tanggal_data_hc,
                tanggal_report_value=tanggal_report,
            )
        st.session_state["outputs"] = outputs
        st.session_state["outputs_context"] = (
            input_signature,
            regional_office,
            int(current_year),
            str(tanggal_data_hc),
            str(tanggal_report),
        )
        st.session_state["auto_download_pending"] = True
    except Exception as exc:
        st.error(f"Proses gagal: {exc}")

outputs = st.session_state.get("outputs")
current_context = (
    input_signature,
    regional_office,
    int(current_year),
    str(tanggal_data_hc),
    str(tanggal_report),
)

# Don't show stale outputs after any input/configuration changes.
if outputs is not None and st.session_state.get("outputs_context") != current_context:
    outputs = None

if outputs is not None:
    st.subheader("7. Output")
    st.success("Proses selesai. ZIP report + chart dan file historis sudah siap.")

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Regional Office", outputs.summary["regional_office"])
    m2.metric("Total Pekerja", f"{outputs.summary['total_pekerja']:,}".replace(",", "."))
    m3.metric("Pekerja Baru", f"{outputs.summary['pekerja_baru']:,}".replace(",", "."))
    m4.metric("Sudah Akses", outputs.summary["persentase"])

    dl1, dl2 = st.columns(2)
    with dl1:
        st.download_button(
            "⬇️ Download ZIP (Report + Chart)",
            data=outputs.zip_bytes,
            file_name=outputs.zip_name,
            mime="application/zip",
            use_container_width=True,
        )
        st.caption(f"Isi ZIP: `{outputs.report_name}` + `{outputs.chart_name}`")
    with dl2:
        st.download_button(
            "⬇️ Download Output Historis",
            data=outputs.historis_bytes,
            file_name=outputs.historis_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        st.caption(f"File: `{outputs.historis_name}`")

    if st.session_state.get("auto_download_pending", False):
        auto_download_two_files(
            outputs.zip_bytes,
            outputs.zip_name,
            outputs.historis_bytes,
            outputs.historis_name,
        )
        st.session_state["auto_download_pending"] = False
        st.caption(
            "Download otomatis dicoba setelah proses selesai. Jika browser memblokir multiple downloads, "
            "gunakan dua tombol download di atas."
        )
