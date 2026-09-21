# -*- coding: utf-8 -*-
"""
Dashboard LKPP - Perencanaan (RUP) dan Realisasi Pengadaan
Jalankan: streamlit run app.py

Dependensi:
pip install streamlit pandas openpyxl plotly kaleido
"""

from __future__ import annotations

import io
import re
import hashlib
import unicodedata
from typing import Optional

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(
    page_title="Dashboard LKPP - RUP dan Realisasi",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# -------------------------------------------------------------------
# Konfigurasi kategori dan utilitas
# -------------------------------------------------------------------
METHOD_ORDER = [
    "Pengadaan Langsung",
    "Penunjukan Langsung",
    "Tender/Seleksi",
    "E-Purchasing",
    "Lain-lain",
    "Swakelola",
]
TYPE_ORDER = [
    "Barang",
    "Jasa Konsultansi",
    "Jasa Lainnya",
    "Pekerjaan Konstruksi",
    "Pekerjaan Terintegrasi",
]
MONTHS_ID = {
    1: "Januari", 2: "Februari", 3: "Maret", 4: "April",
    5: "Mei", 6: "Juni", 7: "Juli", 8: "Agustus",
    9: "September", 10: "Oktober", 11: "November", 12: "Desember",
}
METHOD_COLORS = {
    "Pengadaan Langsung": "#4472C4",
    "Penunjukan Langsung": "#ED7D31",
    "Tender/Seleksi": "#A5A5A5",
    "E-Purchasing": "#70AD47",
    "Lain-lain": "#8064A2",
    "Swakelola": "#5B9BD5",
}
TYPE_COLORS = {
    "Barang": "#4472C4",
    "Jasa Konsultansi": "#ED7D31",
    "Jasa Lainnya": "#A5A5A5",
    "Pekerjaan Konstruksi": "#70AD47",
    "Pekerjaan Terintegrasi": "#8064A2",
}

def norm(value) -> str:
    if value is None:
        return ""
    s = unicodedata.normalize("NFKD", str(value)).encode("ascii", "ignore").decode("ascii")
    s = s.lower().strip()
    s = re.sub(r"[^a-z0-9]+", "_", s)
    return re.sub(r"_+", "_", s).strip("_")

def find_col(columns, candidates, contains=False) -> Optional[str]:
    """Cari kolom dengan exact match lebih dahulu, lalu kandidat alias."""
    lookup = {norm(c): c for c in columns}
    for cand in candidates:
        if norm(cand) in lookup:
            return lookup[norm(cand)]
    if contains:
        normalized = [(norm(c), c) for c in columns]
        for cand in candidates:
            needle = norm(cand)
            for nc, original in normalized:
                if needle and needle in nc:
                    return original
    return None

def amount(series: pd.Series) -> pd.Series:
    """Konversi angka Excel maupun teks berformat Indonesia ke numerik."""
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce").fillna(0.0)
    s = series.astype("string").str.strip()
    s = s.replace({"": pd.NA, "-": pd.NA, "nan": pd.NA, "None": pd.NA})
    # Hilangkan simbol mata uang/spasi; titik ribuan Indonesia dihapus,
    # koma desimal Indonesia diubah menjadi titik.
    s = s.str.replace(r"(?i)rp\s*", "", regex=True)
    s = s.str.replace(r"\s+", "", regex=True)
    both = s.str.contains(",", na=False) & s.str.contains(r"\.", na=False)
    s.loc[both] = s.loc[both].str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
    comma_only = s.str.contains(",", na=False) & ~s.str.contains(r"\.", na=False)
    s.loc[comma_only] = s.loc[comma_only].str.replace(",", ".", regex=False)
    # Titik tanpa koma dianggap pemisah ribuan jika pola ribuan.
    thousands = s.str.match(r"^-?\d{1,3}(\.\d{3})+$", na=False)
    s.loc[thousands] = s.loc[thousands].str.replace(".", "", regex=False)
    return pd.to_numeric(s, errors="coerce").fillna(0.0)

def normalize_code(series: pd.Series) -> pd.Series:
    s = series.astype("string").str.strip()
    s = s.str.replace(r"\.0$", "", regex=True)
    s = s.replace({"": pd.NA, "nan": pd.NA, "None": pd.NA})
    return s

def read_workbook(uploaded_file, label: str):
    """Baca workbook dengan progress per sheet dan ringkasan baris."""
    raw = uploaded_file.getvalue()
    digest = hashlib.md5(raw).hexdigest()
    xls = pd.ExcelFile(io.BytesIO(raw), engine="openpyxl")
    frames, summaries = [], []
    progress = st.progress(0, text=f"Membaca file {label}…")
    status = st.empty()
    total = max(len(xls.sheet_names), 1)
    for i, sheet in enumerate(xls.sheet_names, start=1):
        status.info(f"📖 {label}: membaca sheet **{sheet}** ({i}/{total})")
        try:
            df = pd.read_excel(xls, sheet_name=sheet, dtype=object)
            df.columns = [str(c).strip() for c in df.columns]
            # Buang kolom tanpa nama/Unnamed sepenuhnya kosong.
            df = df.loc[:, ~pd.Index(df.columns).str.match(r"^Unnamed", case=False)]
            df = df.dropna(how="all")
            summaries.append({"Sheet": sheet, "Jumlah baris": len(df), "Status": "Terbaca"})
            if not df.empty:
                df["_sheet_source"] = sheet
                frames.append(df)
        except Exception as exc:
            summaries.append({"Sheet": sheet, "Jumlah baris": 0, "Status": f"Gagal: {exc}"})
        progress.progress(i / total, text=f"Membaca file {label}: {i}/{total} sheet")
    progress.empty()
    status.empty()
    if not frames:
        raise ValueError(f"Tidak ada data yang berhasil dibaca dari file {label}.")
    combined = pd.concat(frames, ignore_index=True, sort=False)
    return combined, pd.DataFrame(summaries), digest

def is_swakelola(sheet_value) -> bool:
    s = norm(sheet_value)
    return s in {"sw_kl", "sw_pemda"} or s.startswith("realisasi_swakelola") or "swakelola" in s

def classify_method(value, sheet_value="") -> str:
    if is_swakelola(sheet_value):
        return "Swakelola"
    s = norm(value)
    if not s or s in {"nan", "none", "tidak tersedia"}:
        return "Lain-lain"
    if "pengadaan_langsung" in s or s == "pl":
        return "Pengadaan Langsung"
    if "penunjukan_langsung" in s or "penunjukkan_langsung" in s or s in {"pilih_langsung", "pj"}:
        return "Penunjukan Langsung"
    if "e_purchasing" in s or "epurchasing" in s or "e_katalog" in s or "katalog_elektronik" in s:
        return "E-Purchasing"
    if "tender" in s or "seleksi" in s:
        return "Tender/Seleksi"
    return "Lain-lain"

def classify_type(value, sheet_value="") -> Optional[str]:
    # Swakelola tidak dimasukkan ke kategori "Lain-lain" pada jenis pengadaan.
    if is_swakelola(sheet_value):
        return None
    s = norm(value)
    if not s or s in {"nan", "none", "tidak tersedia"}:
        return None
    if "terintegrasi" in s or "terintegritas" in s:
        return "Pekerjaan Terintegrasi"
    if "konsult" in s:
        return "Jasa Konsultansi"
    if "konstruksi" in s:
        return "Pekerjaan Konstruksi"
    if "barang" in s:
        return "Barang"
    if "jasa_lain" in s or s == "jasa":
        return "Jasa Lainnya"
    return None

def identify_scope_column(df: pd.DataFrame) -> Optional[str]:
    return find_col(df.columns, [
        "nama_kementerian_lembaga", "nama_k_l_pd", "nama_klpd", "nama_instansi",
        "kementerian_lembaga", "klpd", "nama_pemda", "nama_satker", "nama_satuan_kerja",
        "jenis_instansi", "tipe_instansi", "kelompok_instansi",
    ], contains=False)

def scope_mask(df: pd.DataFrame, scope: str, scope_col: Optional[str]) -> pd.Series:
    if scope == "Nasional":
        return pd.Series(True, index=df.index)
    if not scope_col:
        return pd.Series(False, index=df.index)
    vals = df[scope_col].astype("string").fillna("").str.lower()
    # Gunakan klasifikasi label yang tersedia. Untuk Kementerian, kecualikan
    # label yang jelas menunjukkan pemerintah daerah.
    pemda = vals.str.contains(r"pemda|pemerintah daerah|kabupaten|kota provinsi|provinsi|dinas daerah", regex=True)
    if scope == "Pemda":
        return pemda
    return ~pemda & vals.ne("")

def first_nonempty(df: pd.DataFrame, candidates, default=None):
    col = find_col(df.columns, candidates, contains=False)
    if not col:
        return pd.Series(default, index=df.index), None
    return df[col], col

def prepare_rup(df: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    d = df.copy()
    sheet_col = "_sheet_source"
    sheet = d[sheet_col] if sheet_col in d else pd.Series("", index=d.index)
    val = pd.Series(0.0, index=d.index)
    normal_col = find_col(d.columns, ["pagu_realisasi"])
    sw_col = find_col(d.columns, ["pag_ta"])
    sw_mask = sheet.map(is_swakelola)
    if normal_col:
        val.loc[~sw_mask] = amount(d.loc[~sw_mask, normal_col])
    if sw_col:
        val.loc[sw_mask] = amount(d.loc[sw_mask, sw_col])
    # Jika kolom pagu_realisasi tidak tersedia pada sheet tertentu, coba pag_ta
    # untuk baris swakelola; nilai selain swakelola tetap mengikuti spesifikasi.
    d["_nilai_rup"] = val
    method_col = find_col(d.columns, ["metode_pengadaan"])
    type_col = find_col(d.columns, ["jenis_pengadaan"])
    unit_col = find_col(d.columns, ["nama_satuan_kerja"])
    scope_col = identify_scope_column(d)
    d["_metode"] = [
        classify_method(d.at[i, method_col] if method_col else None, sheet.iloc[pos])
        for pos, i in enumerate(d.index)
    ]
    d["_jenis"] = [
        classify_type(d.at[i, type_col] if type_col else None, sheet.iloc[pos])
        for pos, i in enumerate(d.index)
    ]
    d["_unit"] = d[unit_col].astype("string").fillna("Tidak diketahui") if unit_col else "Tidak diketahui"
    d["_scope_label"] = d[scope_col].astype("string").fillna("") if scope_col else ""
    id_col = find_col(d.columns, ["kode_rup", "kd_rup", "kode_rup_paket"])
    d["_id"] = normalize_code(d[id_col]) if id_col else pd.Series(pd.NA, index=d.index, dtype="string")
    # Bulan jadwal/awal pelaksanaan, bila tersedia.
    date_col = find_col(d.columns, [
        "awal_pemilihan", "tanggal_awal_pemilihan", "awal_pekerjaan",
        "tanggal_awal_pekerjaan", "awal_pelaksanaan", "tanggal_awal_pelaksanaan",
        "jadwal_pemilihan_awal", "bulan_rencana", "bulan",
    ])
    d["_bulan"] = pd.to_datetime(d[date_col], errors="coerce") if date_col else pd.NaT
    return d, {"scope_col": scope_col, "date_col": date_col, "id_col": id_col}

def prepare_real(df: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    d = df.copy()
    sheet = d["_sheet_source"] if "_sheet_source" in d else pd.Series("", index=d.index)
    val_col = find_col(d.columns, ["nilai_realisasi"])
    d["_nilai_real"] = amount(d[val_col]) if val_col else 0.0
    method_col = find_col(d.columns, ["metode_pengadaan", "mtd_pemilihan"])
    type_col = find_col(d.columns, ["jenis_pengadaan"])
    unit_col = find_col(d.columns, ["nama_satker", "nama_satuan_kerja"])
    scope_col = identify_scope_column(d)
    d["_metode"] = [
        classify_method(d.at[i, method_col] if method_col else None, sheet.iloc[pos])
        for pos, i in enumerate(d.index)
    ]
    d["_jenis"] = [
        classify_type(d.at[i, type_col] if type_col else None, sheet.iloc[pos])
        for pos, i in enumerate(d.index)
    ]
    d["_unit"] = d[unit_col].astype("string").fillna("Tidak diketahui") if unit_col else "Tidak diketahui"
    d["_scope_label"] = d[scope_col].astype("string").fillna("") if scope_col else ""
    id_col = find_col(d.columns, ["kd_rup_paket", "kd_rup", "kode_rup", "kode_rup_paket"])
    d["_id"] = normalize_code(d[id_col]) if id_col else pd.Series(pd.NA, index=d.index, dtype="string")
    date_col = find_col(d.columns, [
        "tanggal_realisasi", "tgl_realisasi", "bulan_realisasi", "tanggal_bayar",
        "tgl_bayar", "tanggal_transaksi", "bulan", "tanggal",
    ])
    d["_bulan"] = pd.to_datetime(d[date_col], errors="coerce") if date_col else pd.NaT
    return d, {"scope_col": scope_col, "date_col": date_col, "id_col": id_col}

def fmt_rp(x):
    try:
        return "Rp " + f"{float(x):,.0f}".replace(",", ".")
    except Exception:
        return "Rp 0"

def fmt_num(x):
    return f"{int(x):,}".replace(",", ".")

def download_png(fig, key: str, label="Unduh grafik PNG"):
    """PNG memakai Kaleido; jika belum terpasang, tampilkan instruksi dan HTML fallback."""
    try:
        image_bytes = fig.to_image(format="png", scale=2)
        st.download_button(label, data=image_bytes, file_name=f"{key}.png",
                           mime="image/png", key=f"png_{key}")
    except Exception:
        st.caption("Ekspor PNG memerlukan paket Kaleido. Jalankan: `pip install -U kaleido`")
        try:
            html = fig.to_html(include_plotlyjs="cdn", full_html=True).encode("utf-8")
            st.download_button(f"Unduh grafik HTML ({key})", data=html,
                               file_name=f"{key}.html", mime="text/html", key=f"html_{key}")
        except Exception:
            pass

def pie_chart(data, names, values, title, key, color_map=None):
    d = data.copy()
    d = d[d[values].fillna(0) > 0]
    if d.empty:
        st.info("Belum ada nilai untuk ditampilkan pada grafik ini.")
        return
    fig = px.pie(d, names=names, values=values, title=title, hole=0.28,
                 color=names, color_discrete_map=color_map)
    fig.update_traces(textposition="inside", textinfo="percent", insidetextorientation="auto")
    fig.update_layout(legend_title_text="", margin=dict(l=20, r=20, t=65, b=20),
                      legend=dict(orientation="v"))
    st.plotly_chart(fig, use_container_width=True)
    download_png(fig, key)

def bar_chart(data, x, y, title, key, horizontal=False, color=None):
    if data.empty:
        st.info("Belum ada data untuk ditampilkan pada grafik ini.")
        return
    if horizontal:
        fig = px.bar(data, x=y, y=x, orientation="h", title=title, color=color,
                     text_auto=".3s")
        fig.update_layout(yaxis={"categoryorder": "total ascending"})
    else:
        fig = px.bar(data, x=x, y=y, title=title, color=color, text_auto=".3s")
    fig.update_layout(margin=dict(l=20, r=20, t=65, b=25), legend_title_text="")
    fig.update_yaxes(title="Nilai (Rp)", tickformat="~s")
    st.plotly_chart(fig, use_container_width=True)
    download_png(fig, key)

def grouped_totals(df, category, value, ordered_categories=None):
    out = df.groupby(category, dropna=False)[value].sum().reset_index()
    out[category] = out[category].fillna("Tidak dikategorikan")
    if ordered_categories:
        out[category] = pd.Categorical(out[category], categories=ordered_categories, ordered=True)
        out = out.sort_values(category)
        out[category] = out[category].astype(str)
    return out

def stacked_plan_real(rup_df, real_df, category, categories, title, key, colors=None):
    """
    Grafik 100% stacked: setiap kategori memiliki tinggi 100%.
    Persentase terealisasi = min(realisasi, RUP) / RUP; sisanya belum terealisasi.
    """
    plan = rup_df.groupby(category, dropna=False)["_nilai_rup"].sum()
    real = real_df.groupby(category, dropna=False)["_nilai_real"].sum()
    rows = []
    for cat in categories:
        p = float(plan.get(cat, 0) or 0)
        r = float(real.get(cat, 0) or 0)
        executed = min(max(r, 0), max(p, 0))
        remaining = max(p - executed, 0)
        pct_done = (executed / p * 100) if p > 0 else 0.0
        pct_left = (remaining / p * 100) if p > 0 else 0.0
        rows.append({
            category: cat,
            "Terealisasi (%)": pct_done,
            "Belum terealisasi (%)": pct_left,
            "RUP (Rp)": p,
            "Realisasi aktual (Rp)": r,
            "Nilai terealisasi dalam batas RUP (Rp)": executed,
            "Sisa RUP (Rp)": remaining,
        })
    out = pd.DataFrame(rows)

    fig = go.Figure()
    fig.add_bar(
        x=out[category], y=out["Terealisasi (%)"], name="Terealisasi",
        marker_color="#4472C4",
        customdata=out[["RUP (Rp)", "Realisasi aktual (Rp)"]].to_numpy(),
        hovertemplate=(
            "%{x}<br>Terealisasi: %{y:.1f}%"
            "<br>RUP: Rp %{customdata[0]:,.0f}"
            "<br>Realisasi aktual: Rp %{customdata[1]:,.0f}<extra></extra>"
        ),
    )
    fig.add_bar(
        x=out[category], y=out["Belum terealisasi (%)"], name="Belum terealisasi",
        marker_color="#ED7D31",
        customdata=out[["RUP (Rp)"]].to_numpy(),
        hovertemplate="%{x}<br>Belum terealisasi: %{y:.1f}%<br>RUP: Rp %{customdata[0]:,.0f}<extra></extra>",
    )
    fig.update_layout(
        barmode="stack", title=title, yaxis_title="Persentase terhadap RUP",
        yaxis=dict(range=[0, 100], ticksuffix="%", dtick=20),
        legend_title_text="", margin=dict(l=20, r=20, t=65, b=25),
        hovermode="x unified",
    )
    st.plotly_chart(fig, use_container_width=True)
    download_png(fig, key)

    over = out[out["Realisasi aktual (Rp)"] > out["RUP (Rp)"]]
    if not over.empty:
        st.caption(
            "Catatan: realisasi aktual melebihi RUP pada sebagian kategori. "
            "Grafik membatasi persentase terealisasi hingga 100%; nilai aktual tetap ditampilkan pada tabel."
        )
    st.dataframe(
        out, use_container_width=True, hide_index=True,
        column_config={
            "Terealisasi (%)": st.column_config.NumberColumn(format="%.1f%%"),
            "Belum terealisasi (%)": st.column_config.NumberColumn(format="%.1f%%"),
            "RUP (Rp)": st.column_config.NumberColumn(format="Rp %.0f"),
            "Realisasi aktual (Rp)": st.column_config.NumberColumn(format="Rp %.0f"),
            "Nilai terealisasi dalam batas RUP (Rp)": st.column_config.NumberColumn(format="Rp %.0f"),
            "Sisa RUP (Rp)": st.column_config.NumberColumn(format="Rp %.0f"),
        },
    )

def monthly_comparison(rup_df, real_df):
    r = rup_df.dropna(subset=["_bulan"]).copy()
    a = real_df.dropna(subset=["_bulan"]).copy()
    if r.empty or a.empty:
        st.info("Grafik bulanan belum dapat dibuat karena kolom tanggal/bulan rencana atau realisasi tidak ditemukan atau tidak terbaca. Periksa kolom tanggal pada workbook.")
        return
    r["Bulan"] = r["_bulan"].dt.to_period("M").dt.to_timestamp()
    a["Bulan"] = a["_bulan"].dt.to_period("M").dt.to_timestamp()
    plan = r.groupby("Bulan")["_nilai_rup"].sum().rename("Rancangan")
    actual = a.groupby("Bulan")["_nilai_real"].sum().rename("Realisasi")
    monthly = pd.concat([plan, actual], axis=1).fillna(0).sort_index().reset_index()
    monthly["Bulan"] = monthly["Bulan"].dt.strftime("%Y-%m")
    long = monthly.melt(id_vars="Bulan", value_vars=["Rancangan", "Realisasi"],
                        var_name="Seri", value_name="Nilai")
    fig = px.bar(long, x="Bulan", y="Nilai", color="Seri", barmode="group",
                 title="Perbandingan Rancangan dan Realisasi per Bulan",
                 color_discrete_map={"Rancangan": "#4472C4", "Realisasi": "#ED7D31"})
    fig.update_layout(yaxis_title="Nilai (Rp)", yaxis_tickformat="~s", legend_title_text="",
                      margin=dict(l=20, r=20, t=65, b=25))
    st.plotly_chart(fig, use_container_width=True)
    download_png(fig, "perbandingan_rup_realisasi_bulanan")
    st.dataframe(monthly, use_container_width=True, hide_index=True,
                 column_config={
                     "Rancangan": st.column_config.NumberColumn(format="Rp %.0f"),
                     "Realisasi": st.column_config.NumberColumn(format="Rp %.0f"),
                 })

# -------------------------------------------------------------------
# Unggah sumber data RUP dan Realisasi
# -------------------------------------------------------------------
st.title("📊 Dashboard Pengadaan LKPP")
st.caption("Dashboard Perencanaan Pengadaan (RUP) dan Realisasi Pengadaan")
st.sidebar.header("📂 Sumber Data")
st.sidebar.caption("Unggah file Excel RUP dan Realisasi untuk menampilkan dashboard.")

rup_file = st.sidebar.file_uploader(
    "Unggah File RUP (.xlsx)", type=["xlsx"], key="rup_upload"
)
real_file = st.sidebar.file_uploader(
    "Unggah File Realisasi (.xlsx)", type=["xlsx"], key="realisasi_upload"
)

if rup_file is None or real_file is None:
    st.info("Silakan unggah kedua file Excel (RUP dan Realisasi) melalui panel sebelah kiri.")
    st.stop()

rup_bytes = rup_file.getvalue()
real_bytes = real_file.getvalue()
rup_hash = hashlib.md5(rup_bytes).hexdigest()
real_hash = hashlib.md5(real_bytes).hexdigest()

# Cache data berdasarkan isi file; Streamlit akan menghindari pembacaan ulang saat filter/halaman berubah.
@st.cache_data(show_spinner=False, max_entries=4)
def cached_read_rup(raw: bytes, digest: str):
    xls = pd.ExcelFile(io.BytesIO(raw), engine="openpyxl")
    frames, summary = [], []
    for sheet_no, sheet in enumerate(xls.sheet_names, start=1):
        st.write(f"📖 RUP: membaca sheet **{sheet}** ({sheet_no}/{len(xls.sheet_names)})")
        df = pd.read_excel(xls, sheet_name=sheet, dtype=object)
        df.columns = [str(c).strip() for c in df.columns]
        df = df.loc[:, ~pd.Index(df.columns).str.match(r"^Unnamed", case=False)].dropna(how="all")
        summary.append({"Sheet": sheet, "Jumlah baris": len(df), "Status": "Terbaca"})
        if not df.empty:
            df["_sheet_source"] = sheet
            frames.append(df)
    return (pd.concat(frames, ignore_index=True, sort=False) if frames else pd.DataFrame(),
            pd.DataFrame(summary))

@st.cache_data(show_spinner=False, max_entries=4)
def cached_read_real(raw: bytes, digest: str):
    xls = pd.ExcelFile(io.BytesIO(raw), engine="openpyxl")
    frames, summary = [], []
    for sheet_no, sheet in enumerate(xls.sheet_names, start=1):
        st.write(f"📖 Realisasi: membaca sheet **{sheet}** ({sheet_no}/{len(xls.sheet_names)})")
        df = pd.read_excel(xls, sheet_name=sheet, dtype=object)
        df.columns = [str(c).strip() for c in df.columns]
        df = df.loc[:, ~pd.Index(df.columns).str.match(r"^Unnamed", case=False)].dropna(how="all")
        summary.append({"Sheet": sheet, "Jumlah baris": len(df), "Status": "Terbaca"})
        if not df.empty:
            df["_sheet_source"] = sheet
            frames.append(df)
    return (pd.concat(frames, ignore_index=True, sort=False) if frames else pd.DataFrame(),
            pd.DataFrame(summary))

try:
    with st.status("Membaca dan menyiapkan data dashboard…", expanded=True) as status:
        st.write("⏳ Membaca seluruh sheet file RUP…")
        rup_raw, rup_summary = cached_read_rup(rup_bytes, rup_hash)
        st.write(f"✅ RUP terbaca: {len(rup_summary)} sheet, {len(rup_raw):,} baris.")
        st.write("⏳ Membaca seluruh sheet file Realisasi…")
        real_raw, real_summary = cached_read_real(real_bytes, real_hash)
        st.write(f"✅ Realisasi terbaca: {len(real_summary)} sheet, {len(real_raw):,} baris.")
        st.write("⏳ Membersihkan dan mengelompokkan data…")
        rup, rup_meta = prepare_rup(rup_raw)
        real, real_meta = prepare_real(real_raw)
        status.update(label="Pembacaan dan pengolahan data selesai", state="complete", expanded=False)
except Exception as exc:
    st.error("Terjadi kendala saat membaca atau mengolah file. Pastikan file berformat Excel (.xlsx) dan memiliki struktur kolom yang sesuai.")
    st.exception(exc)
    st.stop()

if rup.empty or real.empty:
    st.error("Salah satu workbook tidak memiliki baris data yang dapat dibaca.")
    st.stop()

with st.expander("Ringkasan sheet yang berhasil dibaca", expanded=False):
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**File RUP**")
        st.dataframe(rup_summary, use_container_width=True, hide_index=True)
    with c2:
        st.markdown("**File Realisasi**")
        st.dataframe(real_summary, use_container_width=True, hide_index=True)

# -------------------------------------------------------------------
# Filter cakupan bersama untuk kedua halaman
# -------------------------------------------------------------------
st.sidebar.divider()
page = st.sidebar.radio("Pilih Halaman", ["Perencanaan Pengadaan (RUP)", "Realisasi Pengadaan"])
scope = st.sidebar.selectbox("Cakupan Data", ["Nasional", "Kementerian", "Pemda"])

scope_col_rup = rup_meta.get("scope_col")
scope_col_real = real_meta.get("scope_col")
if scope != "Nasional" and (not scope_col_rup or not scope_col_real):
    st.warning("Kolom identitas K/L/Pemda tidak terdeteksi pada salah satu file. Untuk cakupan Kementerian/Pemda, dashboard membutuhkan kolom nama K/L/PD atau identitas instansi yang sesuai.")

rup_f = rup[scope_mask(rup, scope, scope_col_rup)].copy()
real_f = real[scope_mask(real, scope, scope_col_real)].copy()

# Pilihan tahun berdasarkan tahun anggaran bila tersedia.
year_col_rup = find_col(rup_f.columns, ["tahun_anggaran_rup", "tahun_anggaran"])
year_col_real = find_col(real_f.columns, ["tahun_anggaran", "tahun_anggaran_realisasi"])
year_values = set()
for d, col in [(rup_f, year_col_rup), (real_f, year_col_real)]:
    if col:
        year_values.update(d[col].dropna().astype(str).str.extract(r"(\d{4})", expand=False).dropna().tolist())
if year_values:
    years = ["Semua Tahun"] + sorted(year_values)
    year_choice = st.sidebar.selectbox("Tahun Anggaran", years)
    if year_choice != "Semua Tahun":
        if year_col_rup:
            rup_f = rup_f[rup_f[year_col_rup].astype(str).str.contains(year_choice, na=False)]
        if year_col_real:
            real_f = real_f[real_f[year_col_real].astype(str).str.contains(year_choice, na=False)]

# -------------------------------------------------------------------
# Dashboard: Perencanaan RUP
# -------------------------------------------------------------------
if page == "Perencanaan Pengadaan (RUP)":
    st.header("Perencanaan Pengadaan (RUP)")
    total_rup = rup_f["_nilai_rup"].sum()
    # Jumlah paket dihitung unik bila kode tersedia; fallback ke jumlah baris.
    n_paket_rup = rup_f["_id"].dropna().nunique() if rup_f["_id"].notna().any() else len(rup_f)
    k1, k2 = st.columns(2)
    k1.metric("Total Nilai RUP", fmt_rp(total_rup))
    k2.metric("Total Paket RUP", fmt_num(n_paket_rup))

    st.subheader("Komposisi Nilai RUP")
    c1, c2 = st.columns(2)
    with c1:
        method_data = grouped_totals(rup_f, "_metode", "_nilai_rup", METHOD_ORDER)
        pie_chart(method_data, "_metode", "_nilai_rup", "RUP berdasarkan Metode Pemilihan",
                  "rup_metode_pemilihan", METHOD_COLORS)
    with c2:
        type_df = rup_f[rup_f["_jenis"].notna()].copy()
        type_data = grouped_totals(type_df, "_jenis", "_nilai_rup", TYPE_ORDER)
        pie_chart(type_data, "_jenis", "_nilai_rup", "RUP berdasarkan Jenis Barang/Jasa",
                  "rup_jenis_pengadaan", TYPE_COLORS)
        sw_value = rup_f.loc[rup_f["_metode"].eq("Swakelola"), "_nilai_rup"].sum()
        st.caption(f"Nilai Swakelola yang tidak dimasukkan ke pie jenis barang/jasa: {fmt_rp(sw_value)}")

    st.subheader("Nilai RUP berdasarkan Unit Kerja")
    unit_data = grouped_totals(rup_f, "_unit", "_nilai_rup").sort_values("_nilai_rup", ascending=False).head(30)
    bar_chart(unit_data, "_unit", "_nilai_rup", "30 Unit Kerja dengan Nilai RUP Terbesar",
              "rup_nilai_per_unit", horizontal=True)
    with st.expander("Lihat rekap nilai RUP per unit kerja"):
        st.dataframe(unit_data.rename(columns={"_unit": "Unit Kerja", "_nilai_rup": "Nilai RUP"}),
                     use_container_width=True, hide_index=True)

# -------------------------------------------------------------------
# Dashboard: Realisasi
# -------------------------------------------------------------------
else:
    st.header("Realisasi Pengadaan")
    total_real = real_f["_nilai_real"].sum()
    n_paket_real = real_f["_id"].dropna().nunique() if real_f["_id"].notna().any() else len(real_f)
    k1, k2 = st.columns(2)
    k1.metric("Total Nilai Realisasi", fmt_rp(total_real))
    k2.metric("Jumlah Paket Terealisasi", fmt_num(n_paket_real))

    st.subheader("Realisasi terhadap RUP")
    st.caption("Bagian terealisasi dibatasi maksimal sebesar nilai RUP pada kategori yang sama. Sisa nilai RUP ditampilkan sebagai belum terealisasi.")
    c1, c2 = st.columns(2)
    with c1:
        stacked_plan_real(rup_f, real_f, "_metode", METHOD_ORDER,
                          "RUP: Terealisasi vs Belum Terealisasi per Metode Pemilihan",
                          "realisasi_stacked_metode")
    with c2:
        # Swakelola tidak dikategorikan sebagai "Lain-lain" dalam jenis pengadaan.
        rup_types = rup_f[rup_f["_jenis"].notna()]
        real_types = real_f[real_f["_jenis"].notna()]
        stacked_plan_real(rup_types, real_types, "_jenis", TYPE_ORDER,
                          "RUP: Terealisasi vs Belum Terealisasi per Jenis Barang/Jasa",
                          "realisasi_stacked_jenis")
        sw_rup = rup_f.loc[rup_f["_metode"].eq("Swakelola"), "_nilai_rup"].sum()
        sw_real = real_f.loc[real_f["_metode"].eq("Swakelola"), "_nilai_real"].sum()
        st.caption(f"Swakelola tidak dimasukkan ke grafik jenis barang/jasa. Nilai RUP Swakelola: {fmt_rp(sw_rup)}; realisasi Swakelola: {fmt_rp(sw_real)}.")

    st.subheader("Perbandingan Rancangan dan Realisasi Setiap Bulan")
    monthly_comparison(rup_f, real_f)

    with st.expander("Ringkasan total realisasi per metode dan jenis"):
        method_summary = grouped_totals(real_f, "_metode", "_nilai_real", METHOD_ORDER)
        method_summary = method_summary.rename(columns={"_metode": "Metode Pemilihan", "_nilai_real": "Nilai Realisasi"})
        st.markdown("**Metode Pemilihan**")
        st.dataframe(method_summary, use_container_width=True, hide_index=True)
        type_summary = grouped_totals(real_f[real_f["_jenis"].notna()], "_jenis", "_nilai_real", TYPE_ORDER)
        type_summary = type_summary.rename(columns={"_jenis": "Jenis Pengadaan", "_nilai_real": "Nilai Realisasi"})
        st.markdown("**Jenis Barang/Jasa (Swakelola dikecualikan)**")
        st.dataframe(type_summary, use_container_width=True, hide_index=True)

st.divider()
st.caption("Catatan: hasil cakupan Kementerian/Pemda mengikuti kolom identitas instansi yang terdeteksi pada file. Perbandingan bulanan memerlukan kolom tanggal/bulan yang dapat dibaca sebagai tanggal. Untuk data swakelola, metode diperlakukan sebagai kategori tersendiri dan jenis pengadaan tidak dialihkan ke kategori Lain-lain.")
