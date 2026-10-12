"""
app_cuti.py - Streamlit UI untuk manajemen cuti SMK Negeri 1 Koba.

Menu:
  1. Data Cuti     -> tabel data cuti tahunan (dari data.xlsx)
  2. Rekap Cuti    -> rekap saldo cuti per pegawai per tahun
  3. Kartu Kendali -> preview & download PDF kartu kendali
  4. Surat Cuti    -> form generate surat permohonan cuti
                      + tombol "Isi Otomatis dari Master"

Jalankan:
    streamlit run app_cuti.py
"""

import os
from datetime import datetime, date

import pandas as pd
import streamlit as st

import ambil_data_cuti as adc


# ---------------------------------------------------------------------------
# Konfigurasi path & konstanta
# ---------------------------------------------------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, "data.xlsx")
LIBUR_PATH = os.path.join(BASE_DIR, "libur.xlsx")
PEGAWAI_PATH = os.path.join(BASE_DIR, "data_pegawai.xlsx")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
KARTU_DIR = os.path.join(BASE_DIR, "kartu_kendali")

os.makedirs(OUTPUT_DIR, exist_ok=True)

# Data kepala sekolah (untuk blok tanda tangan surat)
KEPALA_NAMA = "Syahryanto, S.T.,M.Pd"
KEPALA_PANGKAT = "Pembina Tk.I, IV/b"
KEPALA_NIP = "197708262006041005"

# Format tanggal default untuk semua tabel
DATE_FORMAT = "DD-MM-YYYY"

st.set_page_config(
    page_title="Manajemen Cuti - SMK Negeri 1 Koba",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ---------------------------------------------------------------------------
# Cache loader
# ---------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def _load_libur():
    return adc.load_libur(LIBUR_PATH)


@st.cache_data(show_spinner=False)
def _load_pegawai():
    return adc.load_pegawai(PEGAWAI_PATH)


@st.cache_data(show_spinner=False)
def _load_cuti():
    libur = _load_libur()
    return adc.load_cuti_dataframe(EXCEL_PATH, libur)


@st.cache_data(show_spinner=False)
def _load_rekap():
    output = _load_cuti()
    return adc.hitung_rekap(output)


# ---------------------------------------------------------------------------
# Helper: config kolom tanggal & angka untuk st.dataframe
# ---------------------------------------------------------------------------
def _date_col(label):
    """Buat DateColumn dengan format dd-mm-yyyy."""
    return st.column_config.DateColumn(label, format=DATE_FORMAT)


def _int_col(label):
    """Buat NumberColumn integer tanpa pemisah ribuan."""
    return st.column_config.NumberColumn(label, format="%d")


# ---------------------------------------------------------------------------
# Sidebar
# ---------------------------------------------------------------------------
st.sidebar.title("📋 Manajemen Cuti")
st.sidebar.caption("SMK Negeri 1 Koba")

menu = st.sidebar.radio(
    "Menu",
    ["📄 Data Cuti", "📊 Rekap Cuti", "🗂️ Kartu Kendali", "✉️ Surat Cuti"],
)

st.sidebar.markdown("---")
if st.sidebar.button("🔄 Refresh Data", width="stretch"):
    st.cache_data.clear()
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.caption(f"Data: `{os.path.basename(EXCEL_PATH)}`")


# ===========================================================================
# MENU 1: Data Cuti
# ===========================================================================
if menu == "📄 Data Cuti":
    st.header("📄 Data Cuti Tahunan")
    st.caption("Data hasil ekstraksi dari `data.xlsx` (Sheet1).")

    try:
        output = _load_cuti()
    except Exception as e:
        st.error(f"Gagal memuat data cuti: {e}")
        st.stop()

    if output.empty:
        st.warning("Tidak ada data cuti tahunan.")
        st.stop()

    col_f1, col_f2, col_f3 = st.columns([2, 2, 1])
    with col_f1:
        tahun_list = ["Semua"] + sorted(
            output["Tahun_cuti"].dropna().astype(int).unique().tolist(),
            reverse=True,
        )
        tahun_pilih = st.selectbox("Filter Tahun", tahun_list, key="dc_tahun")
    with col_f2:
        nama_list = ["Semua"] + sorted(output["Nama"].dropna().unique().tolist())
        nama_pilih = st.selectbox("Filter Nama", nama_list, key="dc_nama")
    with col_f3:
        st.metric("Total Record", len(output))

    df_view = output.copy()
    if tahun_pilih != "Semua":
        df_view = df_view[df_view["Tahun_cuti"] == tahun_pilih]
    if nama_pilih != "Semua":
        df_view = df_view[df_view["Nama"] == nama_pilih]

    st.write(f"Menampilkan **{len(df_view)}** record")
    st.dataframe(
        df_view,
        width="stretch",
        hide_index=True,
        column_config={
            "No": _int_col("No"),
            "Nama": st.column_config.TextColumn("Nama"),
            "Keterangan": st.column_config.TextColumn("Keterangan"),
            "Tgl_mulai_cuti": _date_col("Tgl Mulai"),
            "Tgl_selesai_cuti": _date_col("Tgl Selesai"),
            "Tahun_cuti": _int_col("Tahun"),
            "Jumlah_hari_cuti": _int_col("Jumlah Hari"),
        },
    )

    csv = df_view.to_csv(index=False).encode("utf-8")
    st.download_button(
        "⬇️ Download CSV",
        csv,
        "data_cuti.csv",
        "text/csv",
    )


# ===========================================================================
# MENU 2: Rekap Cuti
# ===========================================================================
elif menu == "📊 Rekap Cuti":
    st.header("📊 Rekap Saldo Cuti")
    st.caption("Perhitungan kuota, akumulasi, dan sisa cuti per pegawai per tahun.")

    try:
        rekap = _load_rekap()
    except Exception as e:
        st.error(f"Gagal memuat rekap: {e}")
        st.stop()

    if rekap.empty:
        st.warning("Tidak ada data rekap.")
        st.stop()

    col_f1, col_f2, col_f3 = st.columns([2, 2, 1])
    with col_f1:
        tahun_list = ["Semua"] + sorted(
            rekap["tahun"].dropna().astype(int).unique().tolist(), reverse=True
        )
        tahun_pilih = st.selectbox("Filter Tahun", tahun_list, key="rc_tahun")
    with col_f2:
        nama_list = ["Semua"] + sorted(rekap["nama"].dropna().unique().tolist())
        nama_pilih = st.selectbox("Filter Nama", nama_list, key="rc_nama")
    with col_f3:
        st.metric("Total Record", len(rekap))

    df_view = rekap.copy()
    if tahun_pilih != "Semua":
        df_view = df_view[df_view["tahun"] == tahun_pilih]
    if nama_pilih != "Semua":
        df_view = df_view[df_view["nama"] == nama_pilih]

    st.write(f"Menampilkan **{len(df_view)}** record")
    st.dataframe(
        df_view,
        width="stretch",
        hide_index=True,
        column_config={
            "No": _int_col("No"),
            "nama": st.column_config.TextColumn("Nama"),
            "tahun": _int_col("Tahun"),
            "cuti_current_year": _int_col("Cuti Tahun Ini"),
            "sisa_tahun_n_minus_1": _int_col("Sisa N-1"),
            "kuota_current_year": _int_col("Kuota"),
            "sisa_di_tahun_ini": _int_col("Sisa Tahun Ini"),
            "sisa_cuti_current_year": _int_col("Sisa Cuti"),
            "total_hari": _int_col("Total Hari"),
        },
    )

    csv = df_view.to_csv(index=False).encode("utf-8")
    st.download_button(
        "⬇️ Download CSV",
        csv,
        "rekap_cuti.csv",
        "text/csv",
    )


# ===========================================================================
# MENU 3: Kartu Kendali
# ===========================================================================
elif menu == "🗂️ Kartu Kendali":
    st.header("🗂️ Kartu Kendali Cuti")

    try:
        output = _load_cuti()
        rekap = _load_rekap()
        pegawai_df = _load_pegawai()
    except Exception as e:
        st.error(f"Gagal memuat data: {e}")
        st.stop()

    if output.empty:
        st.warning("Tidak ada data cuti.")
        st.stop()

    tahun_list = sorted(
        output["Tahun_cuti"].dropna().astype(int).unique().tolist(), reverse=True
    )

    col1, col2 = st.columns(2)
    with col1:
        tahun = st.selectbox("Tahun", tahun_list, key="kk_tahun")
    with col2:
        daftar_nama = sorted(
            output[output["Tahun_cuti"] == tahun]["Nama"].dropna().unique().tolist()
        )
        nama = st.selectbox("Pegawai", daftar_nama, key="kk_nama")

    if nama:
        kartu = adc.kartu_kendali(output, rekap, nama=nama, tahun=int(tahun))
        if kartu.empty:
            st.warning("Tidak ada data cuti untuk pegawai ini di tahun tersebut.")
        else:
            info = adc.get_pegawai_info(pegawai_df, nama)

            st.markdown("---")
            c1, c2, c3 = st.columns(3)
            c1.write(f"**Nama:** {info['nama']}")
            c1.write(f"**NIP:** {info['nip']}")
            c2.write(f"**Pangkat/Gol:** {adc.format_pangkat_gol(info['pangkat'], info['gol'])}")
            c2.write(f"**Jabatan:** {info['jabatan']}")
            c3.write(f"**Unit Kerja:** {adc.UNIT_KERJA_DEFAULT}")
            c3.write(f"**Tahun:** {tahun}")

            st.markdown("### Tabel Kartu Kendali")
            st.dataframe(
                kartu,
                width="stretch",
                hide_index=True,
                column_config={
                    "No": _int_col("No"),
                    "Surat Izin Nomor": st.column_config.TextColumn("Surat Izin Nomor"),
                    "Surat Izin Tanggal": st.column_config.TextColumn("Surat Izin Tanggal"),
                    "Dari Tgl": _date_col("Dari Tgl"),
                    "Sampai Tgl": _date_col("Sampai Tgl"),
                    "Jenis Cuti": st.column_config.TextColumn("Jenis Cuti"),
                    "Paraf": st.column_config.TextColumn("Paraf"),
                    "ATB": _int_col("ATB"),
                    "LHC": _int_col("LHC"),
                    "STB": _int_col("STB"),
                },
            )

            safe = "".join(c if c.isalnum() else "_" for c in str(info["nama"])).strip("_")
            path_pdf = os.path.join(OUTPUT_DIR, f"kartu_kendali_{tahun}_{safe}.pdf")

            if st.button("📄 Generate PDF Kartu Kendali", width="stretch"):
                with st.spinner("Membuat PDF..."):
                    adc.generate_pdf(
                        kartu,
                        nama=info["nama"],
                        nip=info["nip"],
                        pangkat=info["pangkat"],
                        jabatan=info["jabatan"],
                        unit_kerja=adc.UNIT_KERJA_DEFAULT,
                        tahun=int(tahun),
                        path_out=path_pdf,
                    )
                st.success(f"PDF berhasil dibuat: `{path_pdf}`")
                with open(path_pdf, "rb") as f:
                    st.download_button(
                        "⬇️ Download PDF Kartu Kendali",
                        f,
                        file_name=os.path.basename(path_pdf),
                        mime="application/pdf",
                        width="stretch",
                    )

    st.markdown("---")
    st.subheader("🗂️ Generate Semua Kartu Kendali")
    col_a, col_b = st.columns(2)
    with col_a:
        tahun_semua = st.selectbox(
            "Tahun (generate semua)", tahun_list, key="kk_all_tahun"
        )
    with col_b:
        bersihkan = st.checkbox("Hapus folder lama sebelum generate", value=False)

    if st.button("📦 Generate Semua Kartu Kendali", width="stretch"):
        with st.spinner("Membuat semua kartu kendali..."):
            pdfs = adc.generate_semua_kartu(
                tahun=int(tahun_semua),
                out_dir=KARTU_DIR,
                bersihkan=bersihkan,
            )
        st.success(f"Berhasil membuat **{len(pdfs)}** file PDF.")
        for p in pdfs:
            st.write(f"- `{os.path.basename(p)}`")


# ===========================================================================
# MENU 4: Surat Cuti
# ===========================================================================
elif menu == "✉️ Surat Cuti":
    st.header("✉️ Surat Permohonan Cuti Tahunan")
    st.caption(
        "Pilih pegawai dari master, klik **Isi Otomatis** untuk mengisi form "
        "dari `data_pegawai.xlsx` + data cuti terakhir dari `data.xlsx`."
    )

    try:
        pegawai_df = _load_pegawai()
        libur_dates = _load_libur()
    except Exception as e:
        st.error(f"Gagal memuat master pegawai: {e}")
        st.stop()

    # --- Inisialisasi session state form ---
    if "form_surat" not in st.session_state:
        st.session_state.form_surat = {
            "nama": "",
            "nip": "-",
            "pangkat_gol": "-",
            "jabatan": "-",
            "unit_kerja": adc.UNIT_KERJA_DEFAULT,
            "tanggal_surat": date.today(),
            "lama_hari": 1,
            "tgl_mulai": date.today(),
            "tgl_selesai": date.today(),
            "alasan": "",
            "alamat": "",
            "kontak": "",
        }

    if "selected_pegawai" not in st.session_state:
        st.session_state.selected_pegawai = ""

    # --- Daftar nama dari master ---
    daftar_nama = [""] + (
        list(pegawai_df["NAMA"].astype(str)) if not pegawai_df.empty else []
    )

    # --- Pilih pegawai ---
    current = st.session_state.selected_pegawai
    idx_default = daftar_nama.index(current) if current in daftar_nama else 0

    nama_pilihan = st.selectbox(
        "Pilih Pegawai (dari master data_pegawai.xlsx)",
        daftar_nama,
        index=idx_default,
        key="selectbox_pegawai_surat",
    )

    # --- Tombol Isi Otomatis ---
    col_btn1, col_btn2 = st.columns([1, 3])
    with col_btn1:
        klik_isi = st.button(
            "🔄 Isi Otomatis dari Master",
            width="stretch",
            type="primary",
        )

    if klik_isi:
        if not nama_pilihan:
            st.warning("⚠️ Pilih pegawai terlebih dahulu.")
        else:
            # 1) Ambil info master pegawai
            info = adc.get_pegawai_info(pegawai_df, nama_pilihan)
            master_row = pegawai_df[
                pegawai_df["NAMA"].astype(str).str.casefold()
                == nama_pilihan.casefold()
            ]
            master_data = master_row.iloc[0] if not master_row.empty else {}
            alamat_master = info.get("alamat", master_data.get("ALAMAT", ""))
            kontak_master = info.get("contact", master_data.get("CONTACT", ""))
            alamat_master = "" if pd.isna(alamat_master) else str(alamat_master)
            kontak_master = "" if pd.isna(kontak_master) else str(kontak_master)

            # 2) Ambil data cuti terakhir dari data.xlsx
            last = adc.get_last_cuti_untuk_nama(
                EXCEL_PATH, nama_pilihan, libur_dates
            )

            # 3) Isi session state form
            st.session_state.form_surat["nama"] = info["nama"]
            st.session_state.form_surat["nip"] = info["nip"]
            st.session_state.form_surat["pangkat_gol"] = adc.format_pangkat_gol(
                info["pangkat"], info["gol"]
            )
            st.session_state.form_surat["jabatan"] = info["jabatan"]
            st.session_state.form_surat["unit_kerja"] = adc.UNIT_KERJA_DEFAULT
            st.session_state.form_surat["alamat"] = alamat_master
            st.session_state.form_surat["kontak"] = kontak_master
            st.session_state.selected_pegawai = nama_pilihan

            if last:
                st.session_state.form_surat["tgl_mulai"] = last["tgl_mulai"].date()
                st.session_state.form_surat["tgl_selesai"] = last["tgl_selesai"].date()
                st.session_state.form_surat["lama_hari"] = int(last["jumlah_hari"])
                st.session_state.form_surat["alasan"] = (
                    last["alasan"] or "Keperluan pribadi"
                )
                st.success(
                    f"✅ Form terisi otomatis!\n\n"
                    f"- **Master**: {info['nama']} / NIP {info['nip']}\n"
                    f"- **Cuti terakhir**: "
                    f"{last['tgl_mulai'].strftime('%d-%m-%Y')} s/d "
                    f"{last['tgl_selesai'].strftime('%d-%m-%Y')} "
                    f"({last['jumlah_hari']} hari kerja)"
                )
            else:
                st.info(
                    f"ℹ️ Data master **{info['nama']}** berhasil diisi.\n\n"
                    f"Namun tidak ada riwayat cuti tahunan di `data.xlsx`, "
                    f"sehingga field tanggal & alasan tidak diubah."
                )
            st.rerun()

    # --- Form Surat ---
    fs = st.session_state.form_surat

    st.markdown("---")
    with st.form("form_surat_cuti"):
        st.subheader("👤 Data Pemohon")
        c1, c2 = st.columns(2)
        with c1:
            nama = st.text_input("Nama", value=fs["nama"])
            nip = st.text_input("NIP", value=fs["nip"])
            pangkat_gol = st.text_input("Pangkat/Gol.Ruang", value=fs["pangkat_gol"])
        with c2:
            jabatan = st.text_input("Jabatan", value=fs["jabatan"])
            unit_kerja = st.text_input("Unit Kerja", value=fs["unit_kerja"])

        st.subheader("📅 Detail Cuti")
        c3, c4, c5 = st.columns(3)
        with c3:
            tanggal_surat = st.date_input(
                "Tanggal Surat",
                value=fs["tanggal_surat"],
                format="DD/MM/YYYY",
            )
        with c4:
            tgl_mulai = st.date_input(
                "Tanggal Mulai",
                value=fs["tgl_mulai"],
                format="DD/MM/YYYY",
            )
        with c5:
            tgl_selesai = st.date_input(
                "Tanggal Selesai",
                value=fs["tgl_selesai"],
                format="DD/MM/YYYY",
            )

        lama_hari = st.number_input(
            "Lama Hari (hari kerja)",
            min_value=1,
            value=int(fs["lama_hari"]),
            help="Jumlah hari kerja efektif (tanpa Sabtu/Minggu/libur nasional).",
        )

        st.subheader("📞 Kontak & Alasan")
        alasan = st.text_input("Alasan Cuti", value=fs["alasan"])
        c6, c7 = st.columns(2)
        with c6:
            alamat = st.text_input("Alamat selama cuti", value=fs["alamat"])
        with c7:
            kontak = st.text_input("Nomor Kontak", value=fs["kontak"])

        st.markdown("---")
        submit = st.form_submit_button(
            "📄 Buat Surat Cuti PDF",
            width="stretch",
            type="primary",
        )

    if submit:
        # Validasi
        if not nama.strip():
            st.error("Nama wajib diisi.")
            st.stop()
        if tgl_selesai < tgl_mulai:
            st.error("Tanggal selesai tidak boleh sebelum tanggal mulai.")
            st.stop()

        # Simpan ke session state
        st.session_state.form_surat.update({
            "nama": nama,
            "nip": nip,
            "pangkat_gol": pangkat_gol,
            "jabatan": jabatan,
            "unit_kerja": unit_kerja,
            "tanggal_surat": tanggal_surat,
            "lama_hari": lama_hari,
            "tgl_mulai": tgl_mulai,
            "tgl_selesai": tgl_selesai,
            "alasan": alasan,
            "alamat": alamat,
            "kontak": kontak,
        })

        safe = "".join(c if c.isalnum() else "_" for c in str(nama)).strip("_")
        path_pdf = os.path.join(OUTPUT_DIR, f"surat_cuti_{safe}.pdf")

        try:
            with st.spinner("Membuat surat cuti PDF..."):
                adc.generate_surat_cuti(
                    nama=nama,
                    nip=nip,
                    pangkat_gol=pangkat_gol,
                    jabatan=jabatan,
                    unit_kerja=unit_kerja,
                    tanggal_surat=tanggal_surat,
                    lama_hari=lama_hari,
                    tgl_mulai=tgl_mulai,
                    tgl_selesai=tgl_selesai,
                    alasan=alasan,
                    alamat=alamat,
                    kontak=kontak,
                    kepala_nama=KEPALA_NAMA,
                    kepala_pangkat=KEPALA_PANGKAT,
                    kepala_nip=KEPALA_NIP,
                    path_out=path_pdf,
                )
            st.success(f"✅ Surat cuti berhasil dibuat: `{path_pdf}`")
            with open(path_pdf, "rb") as f:
                st.download_button(
                    "⬇️ Download Surat Cuti PDF",
                    f,
                    file_name=os.path.basename(path_pdf),
                    mime="application/pdf",
                    width="stretch",
                )
        except Exception as e:
            st.error(f"Gagal membuat surat cuti: {e}")