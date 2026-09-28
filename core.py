from __future__ import annotations

import io
import os
import re
import zipfile
from dataclasses import dataclass
from datetime import date, datetime
from typing import Dict, Iterable, Tuple

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd


PERSEN_COL = "PERSENTASE PEKERJA SUDAH AKSES BRIPEDIA"
BULAN_ID = {
    1: "Januari", 2: "Februari", 3: "Maret", 4: "April",
    5: "Mei", 6: "Juni", 7: "Juli", 8: "Agustus",
    9: "September", 10: "Oktober", 11: "November", 12: "Desember",
}


@dataclass
class GeneratedOutputs:
    zip_bytes: bytes
    zip_name: str
    historis_bytes: bytes
    historis_name: str
    report_name: str
    chart_name: str
    summary: Dict[str, object]


def format_tanggal_id(value: date | datetime | pd.Timestamp) -> str:
    ts = pd.Timestamp(value)
    return f"{ts.day} {BULAN_ID[ts.month]} {ts.year}"


def format_periode_report(value: date | datetime | pd.Timestamp) -> str:
    return f"1 s.d. {format_tanggal_id(value)}"


def _safe_filename(text: str) -> str:
    return re.sub(r'[\\/:*?"<>|]+', "_", str(text)).strip()


def _regional_filename_label(regional_office: str) -> str:
    # Notebook: "Region 3 - Regional Office Padang" -> "Region 3 (Regional Office Padang)"
    if " - " in regional_office:
        kiri, kanan = regional_office.split(" - ", 1)
        return f"{kiri} ({kanan})"
    return regional_office


def _normalize_code(series: pd.Series, width: int | None = None) -> pd.Series:
    """Normalize Excel identifiers without scientific notation / trailing .0."""
    s = series.astype("string").str.strip()
    s = s.str.replace(r"\.0$", "", regex=True)
    if width:
        s = s.str.zfill(width)
    return s


def _require_columns(df: pd.DataFrame, columns: Iterable[str], dataset_name: str) -> None:
    missing = [c for c in columns if c not in df.columns]
    if missing:
        raise ValueError(
            f"{dataset_name}: kolom wajib tidak ditemukan: {', '.join(missing)}"
        )


def read_input_excels(
    data_uker_bytes: bytes,
    data_hc_bytes: bytes,
    data_bripedia_bytes: bytes,
    data_historis_bytes: bytes,
) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Read the four uploads with the same configuration as the notebook."""
    data_uker = pd.read_excel(
        io.BytesIO(data_uker_bytes), sheet_name="Data Uker", header=1
    )
    data_hc = pd.read_excel(io.BytesIO(data_hc_bytes))
    data_bripedia = pd.read_excel(io.BytesIO(data_bripedia_bytes))
    data_historis = pd.read_excel(io.BytesIO(data_historis_bytes))
    return data_uker, data_hc, data_bripedia, data_historis


def prepare_master(
    data_uker_raw: pd.DataFrame,
    data_hc_raw: pd.DataFrame,
    data_bripedia_raw: pd.DataFrame,
) -> pd.DataFrame:
    """Port of notebook Data Preparation; REQUEST comment blocks are intentionally omitted."""

    # ==============================
    # DATA UKER
    # ==============================
    data_uker = data_uker_raw.copy()
    data_uker.columns = data_uker.columns.astype(str).str.strip()

    if "Status Kantor / Jenis Uker" in data_uker.columns:
        kolom_uker = "Status Kantor / Jenis Uker"
    elif "Nama Internal" in data_uker.columns:
        kolom_uker = "Nama Internal"
    else:
        raise ValueError(
            "Data Uker: kolom 'Status Kantor / Jenis Uker' atau 'Nama Internal' tidak ditemukan."
        )

    _require_columns(
        data_uker,
        ["Branch Code Uker", "Regional Office", "Nama KC Induk", "Nama Kantor Unit Kerja", kolom_uker],
        "Data Uker",
    )

    data_uker = data_uker[data_uker[kolom_uker].notna()].copy()

    for col in ["Regional Office", "Nama KC Induk", "Nama Kantor Unit Kerja", kolom_uker]:
        data_uker[col] = data_uker[col].astype("string").str.strip()

    uker_map = {
        "KANTOR PUSAT": ["Kanpus"],
        "KANTOR WILAYAH & AIW": ["Regional Office", "AIW"],
        "KANTOR CABANG": ["KC"],
        "KANTOR CABANG PEMBANTU": ["KCP"],
        "KANTOR KAS": ["KK"],
        "BRI UNIT": ["BRI UNIT"],
        "LAYANAN SIM STNK BPKB": ["SSB"],
        "TEMPORARY OUTLET": ["TO"],
        "PAYMENT POINT": ["PP"],
        "TERAS KANTOR": ["TERAS"],
        "TERAS KELILING & KAPAL": ["TERAS KELILING", "TERAS KAPAL"],
        "E-BUZZ": ["E-BUZZ"],
        "UKLN & KF NON AIW": ["KF", "KCLN", "KCPLN"],
    }

    dict_uker = {}
    for uker, status in uker_map.items():
        dict_uker[uker] = (
            data_uker.loc[
                data_uker[kolom_uker].isin(status),
                ["Branch Code Uker", "Regional Office", "Nama KC Induk", "Nama Kantor Unit Kerja"],
            ]
            .reset_index(drop=True)
        )

    data_uker = pd.concat(list(dict_uker.values()), axis=0, ignore_index=True)
    data_uker["Branch Code Uker"] = pd.to_numeric(
        data_uker["Branch Code Uker"], errors="coerce"
    ).astype("Int64")
    data_uker = data_uker.dropna(subset=["Branch Code Uker"])
    data_uker = data_uker.drop_duplicates().reset_index(drop=True)

    data_uker_merge = data_uker.drop(["Nama Kantor Unit Kerja"], axis=1)
    data_uker_merge = data_uker_merge.drop_duplicates().reset_index(drop=True)

    data_uker_merge = data_uker_merge[
        ~(
            data_uker_merge["Branch Code Uker"].duplicated(keep=False)
            & data_uker_merge["Nama KC Induk"].isin(
                ["KC Bandung Ah Nasution", "Regional Office Medan"]
            )
        )
    ].reset_index(drop=True)

    for df in [data_uker, data_uker_merge]:
        df.rename(columns={"Branch Code Uker": "Branch Code"}, inplace=True)
        df["Branch Code"] = df["Branch Code"].astype("Int64").astype(str)
        df.rename(columns={"Nama KC Induk": "Cabang"}, inplace=True)

    data_uker = data_uker.rename(columns={"Nama Kantor Unit Kerja": "Unit Kerja"})

    # ==============================
    # DATA HC
    # ==============================
    data_hc = data_hc_raw.copy()
    _require_columns(
        data_hc,
        [
            "PERNR", "COMPLETENAME", "KODE BRANCH", "PSADESC", "ORGDESC",
            "TMT PEKERJA BRI", "JOBDESC", "ESGDESC", "STATUS PEKERJA", "KELOMPOK",
        ],
        "Data HC",
    )

    data_hc = data_hc.rename(columns={"PERNR": "Personal Number"})
    if "PERNR.1" in data_hc.columns:
        data_hc.drop(["PERNR.1"], axis=1, inplace=True)

    data_hc["Personal Number"] = _normalize_code(data_hc["Personal Number"], width=8)

    data_hc_worker = data_hc.pop("COMPLETENAME")
    data_hc.insert(1, "Nama Pekerja", data_hc_worker)

    data_hc_branch_code = data_hc.pop("KODE BRANCH")
    data_hc.insert(2, "Branch Code", data_hc_branch_code)
    data_hc["Branch Code"] = pd.to_numeric(
        data_hc["Branch Code"], errors="coerce"
    ).astype("Int64").astype(str)

    data_hc = data_hc.rename(columns={"PSADESC": "Cabang", "ORGDESC": "Unit Kerja"})
    data_hc["TMT PEKERJA BRI"] = pd.to_datetime(data_hc["TMT PEKERJA BRI"], errors="coerce")
    data_hc["Tahun Mulai Pekerja"] = data_hc["TMT PEKERJA BRI"].dt.year.astype("Int64")

    data_hc_tmt = data_hc.pop("Tahun Mulai Pekerja")
    data_hc.insert(5, "Tahun Mulai Pekerja", data_hc_tmt)
    data_hc.drop(["TMT PEKERJA BRI"], axis=1, inplace=True)

    data_hc = data_hc.rename(
        columns={
            "JOBDESC": "Jabatan",
            "ESGDESC": "Jenis Jabatan",
            "STATUS PEKERJA": "Status Pekerja",
            "KELOMPOK": "Kelompok",
        }
    )

    # REQUEST blocks from the notebook are deliberately not applied.

    # ==============================
    # DATA BRIPEDIA
    # ==============================
    data_bripedia = data_bripedia_raw.copy()
    _require_columns(
        data_bripedia,
        ["Personal Number", "Jumlah Akses", "Jumlah Ketentuan"],
        "Data BRIPEDIA",
    )

    pn_raw = data_bripedia["Personal Number"].astype("string").str.strip()
    data_bripedia = data_bripedia[pn_raw.str.lower() != "superadmin"].copy()
    data_bripedia["Personal Number"] = _normalize_code(
        data_bripedia["Personal Number"], width=8
    )

    data_bripedia = data_bripedia[
        data_bripedia["Personal Number"].isin(data_hc["Personal Number"])
    ][["Personal Number", "Jumlah Akses", "Jumlah Ketentuan"]].reset_index(drop=True)

    data_bripedia = data_bripedia.rename(
        columns={
            "Jumlah Akses": "Jumlah Akses BRIPEDIA",
            "Jumlah Ketentuan": "Jumlah Akses Ketentuan",
        }
    )
    for col in ["Jumlah Akses BRIPEDIA", "Jumlah Akses Ketentuan"]:
        data_bripedia[col] = pd.to_numeric(data_bripedia[col], errors="coerce").fillna(0)

    # ==============================
    # DATA MASTER
    # ==============================
    data_master = data_hc.drop(["Cabang"], axis=1)
    data_master = data_master.merge(data_uker_merge, on="Branch Code", how="left")
    data_master = data_master.merge(data_bripedia, on="Personal Number", how="left")

    data_master = data_master[
        ~data_master["Personal Number"].astype(str).str.startswith("9")
    ].copy()

    data_master_ro = data_master.pop("Regional Office")
    data_master.insert(2, "Regional Office", data_master_ro)
    data_master = data_master[data_master["Regional Office"] != "Kantor Pusat"].copy()

    data_master_branch = data_master.pop("Cabang")
    data_master.insert(4, "Cabang", data_master_branch)

    data_master["Jabatan"] = data_master["Jabatan"].astype("string")
    data_master["Jenis Jabatan"] = data_master["Jenis Jabatan"].astype("string")
    data_master = data_master[
        ~data_master["Jabatan"].str.contains("SAKIT", case=False, na=False)
    ]
    data_master = data_master[
        ~data_master["Jenis Jabatan"].isin(
            ["Sakit Lebih 1 Tahun", "Ijin - Tanpa Upah", "Non Aktif"]
        )
    ].copy()

    data_master[["Jumlah Akses BRIPEDIA", "Jumlah Akses Ketentuan"]] = (
        data_master[["Jumlah Akses BRIPEDIA", "Jumlah Akses Ketentuan"]]
        .apply(pd.to_numeric, errors="coerce")
        .fillna(0)
    )

    # Match notebook's final fillna(0), but only after text filters.
    data_master = data_master.fillna(0).reset_index(drop=True)

    data_master["Keterangan Inklusi"] = np.where(
        data_master["Jumlah Akses BRIPEDIA"] > 0,
        "Sudah Akses BRIPEDIA",
        "Belum Akses BRIPEDIA",
    )
    data_master["Keterangan Literasi"] = np.where(
        data_master["Jumlah Akses Ketentuan"] > 0,
        "Sudah Akses Dokumen Ketentuan di BRIPEDIA",
        "Belum Akses Dokumen Ketentuan di BRIPEDIA",
    )

    return data_master


def get_regional_offices(data_master: pd.DataFrame) -> list[str]:
    values = (
        data_master["Regional Office"]
        .astype(str)
        .str.strip()
        .replace({"0": np.nan, "nan": np.nan, "<NA>": np.nan})
        .dropna()
        .unique()
        .tolist()
    )

    def region_key(text: str):
        m = re.match(r"\s*Region\s+(\d+)", text, flags=re.I)
        return (0, int(m.group(1))) if m else (1, text)

    return sorted(values, key=region_key)


def _build_summary(df: pd.DataFrame, group_col: str) -> pd.DataFrame:
    out = (
        df.groupby(group_col)
        .agg(
            **{
                "BELUM AKSES BRIPEDIA": (
                    "Jumlah Akses BRIPEDIA", lambda akses: (akses == 0).sum()
                ),
                "SUDAH AKSES BRIPEDIA": (
                    "Jumlah Akses BRIPEDIA", lambda akses: (akses > 0).sum()
                ),
                "VOLUME AKSES KETENTUAN": ("Jumlah Akses Ketentuan", "sum"),
            }
        )
        .reset_index()
        .rename(columns={group_col: "UNIT KERJA"})
    )
    out["TOTAL PEKERJA"] = out["BELUM AKSES BRIPEDIA"] + out["SUDAH AKSES BRIPEDIA"]
    out[PERSEN_COL] = np.where(
        out["TOTAL PEKERJA"] > 0,
        out["SUDAH AKSES BRIPEDIA"] / out["TOTAL PEKERJA"],
        0,
    )
    out = out.sort_values(PERSEN_COL, ascending=False).reset_index(drop=True)
    out.insert(0, "RANK", range(1, len(out) + 1))
    out = out[
        [
            "RANK", "UNIT KERJA", "BELUM AKSES BRIPEDIA", "SUDAH AKSES BRIPEDIA",
            "TOTAL PEKERJA", "VOLUME AKSES KETENTUAN", PERSEN_COL,
        ]
    ]
    out[PERSEN_COL] = out[PERSEN_COL].map(lambda x: f"{x:.2%}")
    return out


def build_report_data(data_master: pd.DataFrame, regional_office: str, current_year: int):
    bripedia_selindo = _build_summary(data_master, "Regional Office")

    data_region = data_master[data_master["Regional Office"] == regional_office].copy()
    data_region = data_region.sort_values(["Cabang", "Unit Kerja"]).reset_index(drop=True)
    data_region.insert(0, "No.", range(1, len(data_region) + 1))
    bripedia_branch = _build_summary(data_region, "Cabang")

    data_region_pekerja_baru = data_region[
        pd.to_numeric(data_region["Tahun Mulai Pekerja"], errors="coerce") == current_year
    ].copy()
    data_region_pekerja_baru = data_region_pekerja_baru.sort_values(
        ["Cabang", "Unit Kerja"]
    ).reset_index(drop=True)
    if "No." in data_region_pekerja_baru.columns:
        data_region_pekerja_baru.drop(["No."], axis=1, inplace=True)
    data_region_pekerja_baru.insert(
        0, "No.", range(1, len(data_region_pekerja_baru) + 1)
    )
    bripedia_branch_pekerja_baru = _build_summary(data_region_pekerja_baru, "Cabang")

    data_pinca = data_region[data_region["Jabatan"] == "PEMIMPIN CABANG"].copy()
    bripedia_pinca = data_pinca[
        ["Cabang", "Jumlah Akses BRIPEDIA", "Jumlah Akses Ketentuan"]
    ].copy()
    bripedia_pinca[["Jumlah Akses BRIPEDIA", "Jumlah Akses Ketentuan"]] = (
        bripedia_pinca[["Jumlah Akses BRIPEDIA", "Jumlah Akses Ketentuan"]]
        .apply(pd.to_numeric, errors="coerce")
        .fillna(0)
        .astype(int)
    )
    bripedia_pinca.rename(
        columns={
            "Jumlah Akses BRIPEDIA": "Volume Akses",
            "Jumlah Akses Ketentuan": "Volume Ketentuan",
        },
        inplace=True,
    )

    list_kanca = pd.DataFrame(
        {
            "Cabang": sorted(
                data_master[data_master["Regional Office"] == regional_office]["Cabang"]
                .astype(str)
                .unique()
                .tolist()
            )
        }
    )
    list_kanca["_key"] = (
        list_kanca["Cabang"].astype(str).str.strip().str.replace(r"\s+", " ", regex=True).str.upper()
    )
    bripedia_pinca["_key"] = (
        bripedia_pinca["Cabang"].astype(str).str.strip().str.replace(r"\s+", " ", regex=True).str.upper()
    )
    bripedia_pinca = list_kanca.merge(
        bripedia_pinca[["_key", "Volume Akses", "Volume Ketentuan"]],
        on="_key",
        how="left",
        indicator=True,
    )
    mask = bripedia_pinca["_merge"].eq("left_only")
    bripedia_pinca[["Volume Akses", "Volume Ketentuan"]] = (
        bripedia_pinca[["Volume Akses", "Volume Ketentuan"]].fillna(0).astype(int)
    )
    bripedia_pinca["Keterangan"] = np.where(
        bripedia_pinca["Volume Akses"] > 0,
        "Sudah Pernah Mengakses Ketentuan di BRIPEDIA",
        "Belum Pernah Mengakses Ketentuan di BRIPEDIA",
    )
    bripedia_pinca.loc[mask, "Keterangan"] = "Data Belum Ditemukan"
    urutan_keterangan = {
        "Sudah Pernah Mengakses Ketentuan di BRIPEDIA": 1,
        "Belum Pernah Mengakses Ketentuan di BRIPEDIA": 2,
        "Data Belum Ditemukan": 3,
    }
    bripedia_pinca["_urutan"] = bripedia_pinca["Keterangan"].map(urutan_keterangan)
    bripedia_pinca = bripedia_pinca[
        ~bripedia_pinca["Cabang"].str.contains("Region", case=False, na=False)
    ]
    bripedia_pinca = (
        bripedia_pinca.sort_values(
            ["_urutan", "Volume Akses", "Cabang"], ascending=[True, False, True]
        )
        .drop(columns=["_urutan", "_key", "_merge"])
        .reset_index(drop=True)
    )

    return {
        "bripedia_selindo": bripedia_selindo,
        "data_region": data_region,
        "bripedia_branch": bripedia_branch,
        "data_region_pekerja_baru": data_region_pekerja_baru,
        "bripedia_branch_pekerja_baru": bripedia_branch_pekerja_baru,
        "bripedia_pinca": bripedia_pinca,
    }


def update_historical(
    data_historis_raw: pd.DataFrame,
    bripedia_selindo: pd.DataFrame,
    regional_office: str,
    tanggal_input: date | datetime | pd.Timestamp,
):
    data_historis = data_historis_raw.copy()
    _require_columns(data_historis, ["date"], "Data Historis")
    if regional_office not in data_historis.columns:
        raise ValueError(
            f"Data Historis belum memiliki kolom untuk '{regional_office}'. "
            "Tambahkan kolom Regional Office tersebut pada file historis agar tren dapat dibuat."
        )

    data_historis["date"] = pd.to_datetime(data_historis["date"], errors="coerce")
    if data_historis["date"].isna().any():
        raise ValueError("Data Historis: terdapat nilai 'date' yang tidak dapat dibaca sebagai tanggal.")

    tanggal_input = pd.Timestamp(tanggal_input).normalize()
    persentase_map = bripedia_selindo.set_index("UNIT KERJA")[PERSEN_COL].to_dict()

    row_baru = {"date": tanggal_input}
    for col in data_historis.columns:
        if col != "date":
            row_baru[col] = persentase_map.get(col)

    data_historis = data_historis[data_historis["date"] != tanggal_input]
    data_historis = pd.concat(
        [data_historis, pd.DataFrame([row_baru])], ignore_index=True
    ).sort_values("date").reset_index(drop=True)

    df_chart = data_historis[["date", regional_office]].copy()
    values_raw = df_chart[regional_office]
    had_percent = values_raw.astype(str).str.contains("%", regex=False)
    numeric = pd.to_numeric(
        values_raw.astype(str).str.replace("%", "", regex=False), errors="coerce"
    )
    # Support either intentional strings like "42.70%" or decimal ratios like 0.427.
    ratio_mask = (~had_percent) & numeric.notna() & (numeric.abs() <= 1)
    numeric.loc[ratio_mask] = numeric.loc[ratio_mask] * 100
    df_chart[regional_office] = numeric
    df_chart = df_chart.dropna(subset=[regional_office]).reset_index(drop=True)
    if df_chart.empty:
        raise ValueError(f"Data Historis untuk '{regional_office}' tidak memiliki nilai yang dapat diplot.")

    return data_historis, df_chart


def _persen_to_float(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    if df.empty:
        if PERSEN_COL not in df.columns:
            df[PERSEN_COL] = pd.Series(dtype=float)
        return df
    if df[PERSEN_COL].dtype == "object":
        df[PERSEN_COL] = pd.to_numeric(
            df[PERSEN_COL].astype(str).str.replace("%", "", regex=False),
            errors="coerce",
        ) / 100
    elif pd.to_numeric(df[PERSEN_COL], errors="coerce").max() > 1:
        df[PERSEN_COL] = df[PERSEN_COL] / 100
    return df


def _prepare_excel_tables(report_data: dict):
    bripedia_selindo_report = _persen_to_float(report_data["bripedia_selindo"])
    bripedia_branch_report = _persen_to_float(report_data["bripedia_branch"])
    bripedia_branch_pekerja_baru_report = _persen_to_float(
        report_data["bripedia_branch_pekerja_baru"]
    )

    data_region = report_data["data_region"]
    detail_pekerja = data_region[
        [
            "No.", "Personal Number", "Nama Pekerja", "Branch Code", "Cabang", "Unit Kerja",
            "Jabatan", "Status Pekerja", "Jumlah Akses BRIPEDIA", "Jumlah Akses Ketentuan",
            "Keterangan Inklusi", "Keterangan Literasi",
        ]
    ].copy().rename(columns={"Cabang": "KC Induk/RO/RAO"})

    data_region_pekerja_baru = report_data["data_region_pekerja_baru"]
    detail_pekerja_baru = data_region_pekerja_baru[
        [
            "No.", "Personal Number", "Nama Pekerja", "Branch Code", "Cabang", "Unit Kerja",
            "Tahun Mulai Pekerja", "Jabatan", "Status Pekerja", "Jumlah Akses BRIPEDIA",
            "Jumlah Akses Ketentuan", "Keterangan Inklusi", "Keterangan Literasi",
        ]
    ].copy().rename(columns={"Cabang": "KC Induk/RO/RAO"})

    pinca_report = report_data["bripedia_pinca"].copy().rename(
        columns={"Cabang": "UNIT KERJA", "Volume Ketentuan": "JUMLAH KETENTUAN"}
    )
    pinca_report.insert(0, "RANK", range(1, len(pinca_report) + 1))
    pinca_report = pinca_report[
        ["RANK", "UNIT KERJA", "Volume Akses", "JUMLAH KETENTUAN", "Keterangan"]
    ]
    pinca_report.columns = [
        "RANK", "UNIT KERJA", "VOLUME AKSES", "JUMLAH KETENTUAN", "KETERANGAN"
    ]

    return (
        bripedia_selindo_report,
        bripedia_branch_report,
        bripedia_branch_pekerja_baru_report,
        detail_pekerja,
        detail_pekerja_baru,
        pinca_report,
    )


def generate_main_excel(
    report_data: dict,
    regional_office: str,
    tanggal_report: str,
    tanggal_data_hc: str,
    current_year: int,
) -> bytes:
    (
        bripedia_selindo_report,
        bripedia_branch_report,
        bripedia_branch_pekerja_baru_report,
        detail_pekerja,
        detail_pekerja_baru,
        pinca_report,
    ) = _prepare_excel_tables(report_data)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book

        title_format = workbook.add_format({"bold": True, "font_size": 14, "align": "center", "valign": "vcenter"})
        subtitle_format = workbook.add_format({"bold": True, "font_size": 12, "align": "center", "valign": "vcenter"})
        header_format = workbook.add_format({"bold": True, "font_color": "white", "bg_color": "#2F75B5", "align": "center", "valign": "vcenter", "text_wrap": True, "border": 1, "border_color": "#D9E2F3"})
        header_red_format = workbook.add_format({"bold": True, "font_color": "white", "bg_color": "#C00000", "align": "center", "valign": "vcenter", "text_wrap": True, "border": 1, "border_color": "#D9E2F3"})
        header_green_format = workbook.add_format({"bold": True, "font_color": "white", "bg_color": "#2E7D32", "align": "center", "valign": "vcenter", "text_wrap": True, "border": 1, "border_color": "#D9E2F3"})
        rank_format = workbook.add_format({"align": "center", "valign": "vcenter", "border": 1, "border_color": "#D9E2F3"})
        text_format = workbook.add_format({"valign": "vcenter", "border": 1, "border_color": "#D9E2F3"})
        number_format = workbook.add_format({"num_format": "#,##0", "align": "right", "valign": "vcenter", "border": 1, "border_color": "#D9E2F3"})
        percent_format = workbook.add_format({"num_format": "0.00%", "align": "right", "valign": "vcenter", "border": 1, "border_color": "#D9E2F3"})
        grand_text_format = workbook.add_format({"bold": True, "font_color": "white", "bg_color": "#1F4E78", "align": "center", "valign": "vcenter", "border": 1, "border_color": "white"})
        grand_number_format = workbook.add_format({"bold": True, "font_color": "white", "bg_color": "#1F4E78", "num_format": "#,##0", "align": "right", "valign": "vcenter", "border": 1, "border_color": "white"})
        grand_percent_format = workbook.add_format({"bold": True, "font_color": "white", "bg_color": "#1F4E78", "num_format": "0.00%", "align": "right", "valign": "vcenter", "border": 1, "border_color": "white"})

        highlight_rank_format = workbook.add_format({"bold": True, "font_size": 14, "align": "center", "valign": "vcenter", "left": 6, "top": 6, "bottom": 6, "left_color": "#ED7D31", "top_color": "#ED7D31", "bottom_color": "#ED7D31", "right": 1, "right_color": "#D9E2F3"})
        highlight_text_format = workbook.add_format({"bold": True, "font_size": 14, "valign": "vcenter", "left": 1, "right": 1, "top": 6, "bottom": 6, "left_color": "#D9E2F3", "right_color": "#D9E2F3", "top_color": "#ED7D31", "bottom_color": "#ED7D31"})
        highlight_number_format = workbook.add_format({"bold": True, "font_size": 14, "num_format": "#,##0", "align": "right", "valign": "vcenter", "left": 1, "right": 1, "top": 6, "bottom": 6, "left_color": "#D9E2F3", "right_color": "#D9E2F3", "top_color": "#ED7D31", "bottom_color": "#ED7D31"})
        highlight_percent_format = workbook.add_format({"bold": True, "font_size": 14, "num_format": "0.00%", "align": "right", "valign": "vcenter", "left": 1, "right": 6, "top": 6, "bottom": 6, "left_color": "#D9E2F3", "right_color": "#ED7D31", "top_color": "#ED7D31", "bottom_color": "#ED7D31"})

        def write_summary_sheet(sheet_name, df, title, regional_name=None, highlight_regional=False):
            worksheet = workbook.add_worksheet(sheet_name)
            worksheet.hide_gridlines(2)
            worksheet.set_zoom(80)
            worksheet.set_column("A:A", 3)
            worksheet.set_column("B:B", 8)
            worksheet.set_column("C:C", 50)
            worksheet.set_column("D:F", 15)
            worksheet.set_column("G:G", 14)
            worksheet.set_column("H:H", 25)

            if regional_name is None:
                worksheet.merge_range("B1:H1", title, title_format)
                worksheet.merge_range("B2:H2", tanggal_report, subtitle_format)
                worksheet.set_row(0, 25); worksheet.set_row(1, 20); worksheet.set_row(2, 28); worksheet.set_row(3, 45)
                header_row, subheader_row, data_start_row = 2, 3, 4
                worksheet.freeze_panes(4, 3)
            else:
                worksheet.merge_range("B1:H1", title, title_format)
                worksheet.merge_range("B2:H2", regional_name.upper(), subtitle_format)
                worksheet.merge_range("B3:H3", tanggal_report, subtitle_format)
                worksheet.set_row(0, 25); worksheet.set_row(1, 20); worksheet.set_row(2, 20); worksheet.set_row(3, 28); worksheet.set_row(4, 45)
                header_row, subheader_row, data_start_row = 3, 4, 5
                worksheet.freeze_panes(5, 3)

            worksheet.merge_range(header_row, 1, subheader_row, 1, "RANK", header_format)
            worksheet.merge_range(header_row, 2, subheader_row, 2, "UNIT KERJA", header_format)
            worksheet.merge_range(header_row, 3, header_row, 4, "JUMLAH PEKERJA", header_format)
            worksheet.write(subheader_row, 3, "BELUM AKSES\nBRIPEDIA", header_red_format)
            worksheet.write(subheader_row, 4, "SUDAH AKSES\nBRIPEDIA", header_green_format)
            worksheet.merge_range(header_row, 5, subheader_row, 5, "TOTAL\nPEKERJA", header_format)
            worksheet.merge_range(header_row, 6, subheader_row, 6, "VOLUME AKSES\nKETENTUAN", header_format)
            worksheet.merge_range(header_row, 7, subheader_row, 7, "PERSENTASE PEKERJA\nSUDAH AKSES BRIPEDIA", header_format)

            for i, (_, row) in enumerate(df.iterrows()):
                excel_row = data_start_row + i
                is_highlight = highlight_regional and str(row["UNIT KERJA"]).strip() == regional_office
                if is_highlight:
                    worksheet.set_row(excel_row, 19.5)
                    formats = [highlight_rank_format, highlight_text_format, highlight_number_format, highlight_number_format, highlight_number_format, highlight_number_format, highlight_percent_format]
                else:
                    formats = [rank_format, text_format, number_format, number_format, number_format, number_format, percent_format]
                values = [row["RANK"], row["UNIT KERJA"], row["BELUM AKSES BRIPEDIA"], row["SUDAH AKSES BRIPEDIA"], row["TOTAL PEKERJA"], row["VOLUME AKSES KETENTUAN"], row[PERSEN_COL]]
                for col_idx, (value, fmt) in enumerate(zip(values, formats), start=1):
                    worksheet.write(excel_row, col_idx, value, fmt)

            data_end_row = data_start_row + len(df) - 1
            if len(df) > 0:
                worksheet.conditional_format(data_start_row, 7, data_end_row, 7, {"type": "data_bar", "bar_color": "#5B9BD5", "min_type": "num", "min_value": 0, "max_type": "num", "max_value": 1})

            total_row = data_start_row + len(df)
            total_belum = df["BELUM AKSES BRIPEDIA"].sum()
            total_sudah = df["SUDAH AKSES BRIPEDIA"].sum()
            total_pekerja = df["TOTAL PEKERJA"].sum()
            total_volume = df["VOLUME AKSES KETENTUAN"].sum()
            total_persen = total_sudah / total_pekerja if total_pekerja != 0 else 0
            worksheet.merge_range(total_row, 1, total_row, 2, "GRAND TOTAL", grand_text_format)
            worksheet.write(total_row, 3, total_belum, grand_number_format)
            worksheet.write(total_row, 4, total_sudah, grand_number_format)
            worksheet.write(total_row, 5, total_pekerja, grand_number_format)
            worksheet.write(total_row, 6, total_volume, grand_number_format)
            worksheet.write(total_row, 7, total_persen, grand_percent_format)

            footnote_format = workbook.add_format({"italic": True, "align": "right", "valign": "vcenter"})
            if sheet_name in ["RO Selindo", "Unit Kerja"]:
                worksheet.merge_range(total_row + 1, 1, total_row + 1, 7, f"*Jumlah Pekerja berdasarkan Data HC posisi tanggal {tanggal_data_hc}", footnote_format)
                worksheet.merge_range(total_row + 2, 1, total_row + 2, 7, "**Akses dimaksud merupakan akses ke dokumen ketentuan, tidak sebatas membuka aplikasi BRIPEDIA", footnote_format)
            elif sheet_name == "Pekerja Baru":
                worksheet.merge_range(total_row + 1, 1, total_row + 1, 7, f"*Jumlah Pekerja TMT Tahun {current_year} berdasarkan Data HC posisi tanggal {tanggal_data_hc}", footnote_format)
                worksheet.merge_range(total_row + 2, 1, total_row + 2, 7, "**Akses dimaksud merupakan akses ke dokumen ketentuan, tidak sebatas membuka aplikasi BRIPEDIA", footnote_format)

        write_summary_sheet("RO Selindo", bripedia_selindo_report, "DATA AKSES BRIPEDIA RAO & RO SELINDO BESERTA SUPERVISI", None, True)
        write_summary_sheet("Unit Kerja", bripedia_branch_report, "DATA AKSES BRIPEDIA BERDASARKAN UNIT KERJA BESERTA SUPERVISI", regional_office)
        write_summary_sheet("Pekerja Baru", bripedia_branch_pekerja_baru_report, "DATA AKSES BRIPEDIA PEKERJA BARU BESERTA SUPERVISI", regional_office)

        worksheet_pinca = workbook.add_worksheet("Pemimpin Cabang")
        worksheet_pinca.hide_gridlines(2); worksheet_pinca.set_zoom(80); worksheet_pinca.freeze_panes(4, 3)
        worksheet_pinca.set_column("A:A", 3); worksheet_pinca.set_column("B:B", 8); worksheet_pinca.set_column("C:C", 40); worksheet_pinca.set_column("D:E", 20); worksheet_pinca.set_column("F:F", 55)
        worksheet_pinca.set_row(0, 25); worksheet_pinca.set_row(1, 20); worksheet_pinca.set_row(2, 20); worksheet_pinca.set_row(3, 28)
        worksheet_pinca.merge_range("B1:F1", "DATA AKSES BRIPEDIA PEMIMPIN CABANG", title_format)
        worksheet_pinca.merge_range("B2:F2", regional_office.upper(), subtitle_format)
        worksheet_pinca.merge_range("B3:F3", tanggal_report, subtitle_format)
        headers_pinca = ["RANK", "UNIT KERJA", "VOLUME AKSES", "JUMLAH KETENTUAN", "KETERANGAN"]
        for col_num, header in enumerate(headers_pinca, start=1):
            worksheet_pinca.write(3, col_num, header, header_format)

        pinca_sudah_format = workbook.add_format({"bold": True, "font_color": "#008000", "align": "center", "valign": "vcenter", "border": 1, "border_color": "#D9E2F3"})
        pinca_belum_format = workbook.add_format({"bold": True, "font_color": "#FF0000", "align": "center", "valign": "vcenter", "border": 1, "border_color": "#D9E2F3"})
        pinca_missing_format = workbook.add_format({"italic": True, "font_color": "#808080", "align": "center", "valign": "vcenter", "border": 1, "border_color": "#D9E2F3"})

        for i, (_, row) in enumerate(pinca_report.iterrows()):
            excel_row = 4 + i
            worksheet_pinca.write(excel_row, 1, row["RANK"], rank_format)
            worksheet_pinca.write(excel_row, 2, row["UNIT KERJA"], text_format)
            worksheet_pinca.write(excel_row, 3, row["VOLUME AKSES"], number_format)
            worksheet_pinca.write(excel_row, 4, row["JUMLAH KETENTUAN"], number_format)
            if row["KETERANGAN"] == "Sudah Pernah Mengakses Ketentuan di BRIPEDIA":
                ket_format = pinca_sudah_format
            elif row["KETERANGAN"] == "Belum Pernah Mengakses Ketentuan di BRIPEDIA":
                ket_format = pinca_belum_format
            else:
                ket_format = pinca_missing_format
            worksheet_pinca.write(excel_row, 5, row["KETERANGAN"], ket_format)

        pinca_total_row = 4 + len(pinca_report)
        worksheet_pinca.merge_range(pinca_total_row, 1, pinca_total_row, 2, "GRAND TOTAL", grand_text_format)
        worksheet_pinca.write(pinca_total_row, 3, pinca_report["VOLUME AKSES"].sum(), grand_number_format)
        worksheet_pinca.write(pinca_total_row, 4, pinca_report["JUMLAH KETENTUAN"].sum(), grand_number_format)
        worksheet_pinca.write_blank(pinca_total_row, 5, None, grand_text_format)

        detail_pekerja.to_excel(writer, sheet_name="Detail Pekerja", index=False)
        worksheet_detail = writer.sheets["Detail Pekerja"]
        worksheet_detail.hide_gridlines(2); worksheet_detail.set_zoom(80); worksheet_detail.freeze_panes(1, 0)
        detail_header_format = workbook.add_format({"bold": True, "font_color": "white", "bg_color": "#2F75B5", "align": "center", "valign": "vcenter"})
        detail_center_bold_format = workbook.add_format({"bold": True, "align": "center", "valign": "vcenter"})
        for col_num, column_name in enumerate(detail_pekerja.columns):
            worksheet_detail.write(0, col_num, column_name, detail_header_format)
        worksheet_detail.set_column("A:A", 10); worksheet_detail.set_column("B:B", 22); worksheet_detail.set_column("C:C", 50); worksheet_detail.set_column("D:D", 18); worksheet_detail.set_column("E:E", 40); worksheet_detail.set_column("F:G", 80); worksheet_detail.set_column("H:H", 19); worksheet_detail.set_column("I:I", 27); worksheet_detail.set_column("J:J", 28); worksheet_detail.set_column("K:K", 23, detail_center_bold_format); worksheet_detail.set_column("L:L", 41, detail_center_bold_format)
        if len(detail_pekerja.columns) > 0:
            worksheet_detail.autofilter(0, 0, max(len(detail_pekerja), 1), len(detail_pekerja.columns) - 1)

        detail_pekerja_baru.to_excel(writer, sheet_name="Detail Pekerja Baru", index=False)
        worksheet_detail_baru = writer.sheets["Detail Pekerja Baru"]
        worksheet_detail_baru.hide_gridlines(2); worksheet_detail_baru.set_zoom(80); worksheet_detail_baru.freeze_panes(1, 0)
        for col_num, column_name in enumerate(detail_pekerja_baru.columns):
            worksheet_detail_baru.write(0, col_num, column_name, detail_header_format)
        worksheet_detail_baru.set_column("A:A", 10); worksheet_detail_baru.set_column("B:B", 22); worksheet_detail_baru.set_column("C:C", 50); worksheet_detail_baru.set_column("D:D", 18); worksheet_detail_baru.set_column("E:E", 40); worksheet_detail_baru.set_column("F:F", 80); worksheet_detail_baru.set_column("G:G", 24); worksheet_detail_baru.set_column("H:H", 80); worksheet_detail_baru.set_column("I:I", 19); worksheet_detail_baru.set_column("J:J", 27); worksheet_detail_baru.set_column("K:K", 28); worksheet_detail_baru.set_column("L:L", 23, detail_center_bold_format); worksheet_detail_baru.set_column("M:M", 41, detail_center_bold_format)
        if len(detail_pekerja_baru.columns) > 0:
            worksheet_detail_baru.autofilter(0, 0, max(len(detail_pekerja_baru), 1), len(detail_pekerja_baru.columns) - 1)

        green_text_format = workbook.add_format({"bold": True, "font_color": "#00B050"})
        red_text_format = workbook.add_format({"bold": True, "font_color": "#C00000"})
        if len(detail_pekerja) > 0:
            for col in [10, 11]:
                worksheet_detail.conditional_format(1, col, len(detail_pekerja), col, {"type": "text", "criteria": "containing", "value": "Sudah", "format": green_text_format})
                worksheet_detail.conditional_format(1, col, len(detail_pekerja), col, {"type": "text", "criteria": "containing", "value": "Belum", "format": red_text_format})
        if len(detail_pekerja_baru) > 0:
            worksheet_detail_baru.conditional_format(1, 11, len(detail_pekerja_baru), 12, {"type": "text", "criteria": "containing", "value": "Sudah", "format": green_text_format})
            worksheet_detail_baru.conditional_format(1, 11, len(detail_pekerja_baru), 12, {"type": "text", "criteria": "containing", "value": "Belum", "format": red_text_format})

    return output.getvalue()


def generate_historical_excel(data_historis: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter", datetime_format="yyyy-mm-dd") as writer:
        data_historis.to_excel(writer, sheet_name="Time Series", index=False)
        workbook = writer.book
        worksheet = writer.sheets["Time Series"]
        header_format = workbook.add_format({"font_name": "Calibri", "font_size": 11, "bold": True, "font_color": "#FFFFFF", "bg_color": "#2F75B5", "align": "center", "valign": "vcenter", "text_wrap": True, "border": 1, "border_color": "#FFFFFF"})
        date_format = workbook.add_format({"font_name": "Calibri", "font_size": 11, "num_format": "yyyy-mm-dd"})
        body_format = workbook.add_format({"font_name": "Calibri", "font_size": 11, "text_wrap": True})
        for col_num, col_name in enumerate(data_historis.columns):
            worksheet.write(0, col_num, col_name, header_format)
        worksheet.set_row(0, 43.5)
        worksheet.set_column("A:A", 10.09, date_format)
        if len(data_historis.columns) > 1:
            worksheet.set_column(1, len(data_historis.columns) - 1, 15.63, body_format)
        worksheet.set_default_row(14.5)
    return output.getvalue()


def generate_chart(df_chart: pd.DataFrame, regional_office: str) -> bytes:
    fig, ax = plt.subplots(figsize=(13, 6.5))
    x = pd.to_datetime(df_chart["date"]).reset_index(drop=True)
    y = pd.to_numeric(df_chart[regional_office], errors="coerce").reset_index(drop=True)
    valid = y.notna()
    x, y = x[valid].reset_index(drop=True), y[valid].reset_index(drop=True)
    if len(y) == 0:
        plt.close(fig)
        raise ValueError("Tidak ada data historis numerik yang dapat dibuat menjadi chart.")

    ax.plot(x, y, marker="o", markersize=6.5, linewidth=2.4, zorder=3)

    for i, (dt, value) in enumerate(zip(x, y)):
        offset_x, offset_y, va = 0, 10, "bottom"
        if len(y) > 1:
            if i == 0:
                if value > y.iloc[i + 1]:
                    offset_y, va = -13, "top"
            elif i == len(y) - 1:
                if value < y.iloc[i - 1]:
                    offset_y, va = 10, "bottom"
                else:
                    offset_y, va = -13, "top"
            else:
                prev_value, next_value = y.iloc[i - 1], y.iloc[i + 1]
                if value > prev_value and value > next_value:
                    offset_y, va = -13, "top"
                elif value < prev_value and value < next_value:
                    offset_y, va = 10, "bottom"
        fontweight = "bold" if i == len(y) - 1 else "semibold"
        ax.annotate(
            f"{value:.2f}%", xy=(dt, value), xytext=(offset_x, offset_y),
            textcoords="offset points", ha="center", va=va, fontsize=9.5,
            fontweight=fontweight, zorder=5,
            bbox=dict(boxstyle="round,pad=0.15", facecolor="white", edgecolor="none", alpha=0.90),
        )

    latest_date, latest_value = x.iloc[-1], y.iloc[-1]
    ax.scatter(latest_date, latest_value, s=90, zorder=4)
    fig.suptitle("Tren Persentase Pekerja Sudah Akses BRIPEDIA", fontsize=16, fontweight="bold", y=0.965)
    fig.text(0.5, 0.915, regional_office, ha="center", va="center", fontsize=12, fontweight="semibold")
    ax.set_xlabel("Periode", fontsize=11, labelpad=14)
    ax.set_ylabel("Persentase Pekerja Sudah Akses BRIPEDIA (%)", fontsize=11, labelpad=14)
    ax.set_xticks(x)
    ax.set_xticklabels(x.dt.strftime("%b %Y"), fontsize=10)
    ax.margins(x=0.045)

    y_data_min, y_data_max = y.min(), y.max()
    y_min = max(0, np.floor((y_data_min - 3) / 5) * 5)
    y_max = min(105, np.ceil((y_data_max + 4) / 5) * 5)
    if y_max - y_min < 15:
        y_min = max(0, y_min - 5)
        y_max = min(105, y_max + 5)
    if y_max <= y_min:
        y_max = min(105, y_min + 15)
    ax.set_ylim(y_min, y_max)
    tick_max = min(y_max, 100)
    ax.set_yticks(np.arange(y_min, tick_max + 0.1, 5))
    ax.tick_params(axis="y", labelsize=10)
    ax.grid(axis="y", linestyle="--", linewidth=0.8, alpha=0.25, zorder=0)
    ax.grid(axis="x", visible=False)
    ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False)
    ax.spines["left"].set_alpha(0.35); ax.spines["bottom"].set_alpha(0.35)
    fig.text(0.985, 0.025, f"Update terakhir: {latest_date.strftime('%d %b %Y')}", ha="right", va="bottom", fontsize=9, style="italic")
    plt.subplots_adjust(top=0.84, bottom=0.16, left=0.08, right=0.975)

    output = io.BytesIO()
    fig.savefig(output, format="png", dpi=300, bbox_inches="tight")
    plt.close(fig)
    return output.getvalue()


def generate_outputs(
    data_master: pd.DataFrame,
    data_historis_raw: pd.DataFrame,
    regional_office: str,
    current_year: int,
    tanggal_data_hc_value: date | datetime | pd.Timestamp,
    tanggal_report_value: date | datetime | pd.Timestamp,
) -> GeneratedOutputs:
    tanggal_data_hc = format_tanggal_id(tanggal_data_hc_value)
    tanggal_report_date = pd.Timestamp(tanggal_report_value)
    tanggal_report = format_periode_report(tanggal_report_date)

    report_data = build_report_data(data_master, regional_office, current_year)
    if report_data["data_region"].empty:
        raise ValueError(f"Tidak ada data pekerja untuk '{regional_office}'.")

    data_historis, df_chart = update_historical(
        data_historis_raw,
        report_data["bripedia_selindo"],
        regional_office,
        tanggal_report_date,
    )

    report_label = _regional_filename_label(regional_office)
    report_date_label = format_tanggal_id(tanggal_report_date)
    report_name = _safe_filename(f"Data Akses BRIPEDIA {report_label} - {report_date_label}.xlsx")
    chart_name = _safe_filename(f"Tren Akses BRIPEDIA - {regional_office}.png")

    hist_start = pd.to_datetime(data_historis["date"]).min()
    historis_name = _safe_filename(
        f"Data Historis Akses BRIPEDIA ({format_tanggal_id(hist_start)} - {report_date_label}).xlsx"
    )
    zip_name = _safe_filename(f"Data Akses BRIPEDIA - {regional_office}.zip")

    report_bytes = generate_main_excel(
        report_data,
        regional_office,
        tanggal_report,
        tanggal_data_hc,
        current_year,
    )
    historis_bytes = generate_historical_excel(data_historis)
    chart_bytes = generate_chart(df_chart, regional_office)

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(report_name, report_bytes)
        zf.writestr(chart_name, chart_bytes)

    summary_row = report_data["bripedia_selindo"]
    summary_row = summary_row[summary_row["UNIT KERJA"] == regional_office]
    summary = {
        "regional_office": regional_office,
        "total_pekerja": int(report_data["data_region"].shape[0]),
        "pekerja_baru": int(report_data["data_region_pekerja_baru"].shape[0]),
        "jumlah_unit_kerja": int(report_data["bripedia_branch"].shape[0]),
        "persentase": summary_row[PERSEN_COL].iloc[0] if not summary_row.empty else "-",
    }

    return GeneratedOutputs(
        zip_bytes=zip_buffer.getvalue(),
        zip_name=zip_name,
        historis_bytes=historis_bytes,
        historis_name=historis_name,
        report_name=report_name,
        chart_name=chart_name,
        summary=summary,
    )
