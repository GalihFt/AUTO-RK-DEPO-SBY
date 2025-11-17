import streamlit as st
import pandas as pd
import numpy as np
import datetime
import random
import io
from collections import Counter

# ---------------------------------------------------------------------
# KONFIGURASI APLIKASI STREAMLIT
# ---------------------------------------------------------------------
st.set_page_config(layout="wide", page_title="Pencocokan Hutang/Piutang CABANG - SBY")

st.title("Pencocokan Hutang/Piutang Afiliasi CABANG - SBY")
# Diubah: Menghapus referensi ke file gantungan
st.markdown("Unggah 2 file CSV wajib (CABANG SBY & SBY CABANG) untuk memulai.")

# ---------------------------------------------------------------------
# FUNGSI-FUNGSI (DARI KODE ANDA - TIDAK DIUBAH)
# ---------------------------------------------------------------------

# Fungsi reconcile_accounts_table (Persis seperti kode Anda)
def reconcile_accounts_table(df):
    """
    Melakukan rekonsiliasi Debet dan Kredit, menghasilkan dua DataFrame:
    1. offset_table -> baris-baris yang berhasil di-offset
    2. gantung_table -> baris-baris yang belum ketemu pasangan
    """

    # --- Persiapan Data ---
    debits_list = df['Debet'].fillna(0).loc[df['Debet'] > 0].to_list()
    credits_list = df['Kredit'].fillna(0).loc[df['Kredit'] > 0].to_list()

    # --- Langkah 1: Pasangan nilai persis ---
    debit_counts = Counter(debits_list)
    credit_counts = Counter(credits_list)
    offset_persis = {}

    common_values = set(debit_counts.keys()) & set(credit_counts.keys())
    for value in common_values:
        matched_count = min(debit_counts[value], credit_counts[value])
        if matched_count > 0:
            offset_persis[value] = matched_count
            debit_counts[value] -= matched_count
            credit_counts[value] -= matched_count

    # --- Sisa nilai ---
    sisa_debit = []
    for val, count in debit_counts.items():
        sisa_debit.extend([val] * count)

    sisa_kredit = []
    for val, count in credit_counts.items():
        sisa_kredit.extend([val] * count)

    sisa_debit.sort(reverse=True)
    sisa_kredit.sort(reverse=True)

    # --- Langkah 2: Greedy kombinasi ---
    offset_penjumlahan = []
    used_debit = [False] * len(sisa_debit)
    used_credit = [False] * len(sisa_kredit)

    # (a) 1 debit = banyak kredit
    for i in range(len(sisa_debit)):
        target_debit = sisa_debit[i]
        current_sum = 0
        matched_credits_indices = []
        for j in range(len(sisa_kredit)):
            if not used_credit[j] and current_sum + sisa_kredit[j] <= target_debit:
                current_sum += sisa_kredit[j]
                matched_credits_indices.append(j)
        if abs(current_sum - target_debit) < 1e-6:
            used_debit[i] = True
            for k in matched_credits_indices:
                used_credit[k] = True
            offset_penjumlahan.append({
                "debit": target_debit,
                "kredit_group": [sisa_kredit[k] for k in matched_credits_indices]
            })

    # (b) 1 kredit = banyak debit
    for i in range(len(sisa_kredit)):
        if used_credit[i]:
            continue
        target_kredit = sisa_kredit[i]
        current_sum = 0
        matched_debits_indices = []
        for j in range(len(sisa_debit)):
            if not used_debit[j] and current_sum + sisa_debit[j] <= target_kredit:
                current_sum += sisa_debit[j]
                matched_debits_indices.append(j)
        if abs(current_sum - target_kredit) < 1e-6:
            used_credit[i] = True
            for k in matched_debits_indices:
                used_debit[k] = True
            offset_penjumlahan.append({
                "debit_group": [sisa_debit[k] for k in matched_debits_indices],
                "kredit": target_kredit
            })

    # --- Buat tabel hasil ---
    offset_values = set(offset_persis.keys())
    for g in offset_penjumlahan:
        if "debit" in g:
            offset_values.add(g["debit"])
            offset_values.update(g["kredit_group"])
        else:
            offset_values.add(g["kredit"])
            offset_values.update(g["debit_group"])

    df = df.copy()
    df["Debet"] = df["Debet"].fillna(0)
    df["Kredit"] = df["Kredit"].fillna(0)

    df["Posisi"] = df.apply(
        lambda x: "OFFSET" if (x["Debet"] in offset_values or x["Kredit"] in offset_values) else "GANTUNG",
        axis=1
    )

    offset_table = df[df["Posisi"] == "OFFSET"].sort_values(by=["Debet", "Kredit"], ascending=False)
    gantung_table = df[df["Posisi"] == "GANTUNG"].sort_values(by=["Debet", "Kredit"], ascending=False)

    return offset_table, gantung_table

# Fungsi combine_with_spacing (Persis seperti kode Anda)
def combine_with_spacing(list_df, all_columns):
    combined = []
    for name in list_df:
        if name in globals():
            df = globals()[name]
            df_copy = df.copy()
            df_copy = df_copy.reindex(columns=all_columns)
            combined.append(df_copy)
            empty = pd.DataFrame([[""] * len(all_columns)] * 2, columns=all_columns)
            combined.append(empty)
        else:
            st.warning(f"DataFrame {name} tidak ditemukan, dilewati.")
    
    if not combined:
        return pd.DataFrame(columns=all_columns)
        
    return pd.concat(combined, ignore_index=True)

# ---------------------------------------------------------------------
# TATA LETAK INPUT (UI) - DISEDERHANAKAN
# ---------------------------------------------------------------------

# --- Bagian CABANG - SBY ---
st.header("Bagian CABANG - SBY")
# Diubah: Menghapus st.columns dan file gantungan
cabang_sby_file = st.file_uploader("Input File CABANG SBY (Wajib)", type="csv")

# --- Bagian SBY - CABANG ---
st.header("Bagian SBY - CABANG")
# Diubah: Menghapus st.columns dan file gantungan
sby_cabang_file = st.file_uploader("Input file SBY CABANG (Wajib)", type="csv")

# --- Input Selisih ---
st.header("Input Selisih")
selisih_input = st.number_input("Input selisih periode sebelumnya", value=0, step=1, help="Masukkan nilai selisih dari periode sebelumnya.")

st.divider()

# Tombol untuk memulai proses
process_button = st.button("Mulai Proses Pencocokan", type="primary", use_container_width=True)

# ---------------------------------------------------------------------
# LOGIKA UTAMA (SAAT TOMBOL DITEKAN)
# ---------------------------------------------------------------------
if process_button:
    # Validasi input (Logika ini sudah benar, tidak perlu diubah)
    if not all([cabang_sby_file, sby_cabang_file]):
        st.error("Harap unggah file WAJIB (CABANG SBY & SBY CABANG) untuk melanjutkan.")
    else:
        try:
            with st.spinner("Sedang memproses... Harap tunggu..."):
                
                # == MULAI KODE ASLI ANDA ==
                
                selisih_sebelumnya = selisih_input

                columns = [
                'Tanggal Kasir', 'ID Dokumen', 'Nomor Dokumen', 'Dibayarkan (ke/dari)', 'Keperluan', 'Vessel Voyage',
                'Debet', 'Kredit', 'Tempat Pembayaran', 'Pembuat', 'Sumber Dokumen', 'Jenis Dokumen',
                'Tanggal Delivery', 'Nama Kode', 'Kode Accounting', 'User Pengakuan', 'Unit', 'Divisi',
                'Flag KBM/KDRT', 'Target_First', 'Target_Jenis', 'Target_Second'
                ]

                # --- PROSES CABANG SBY ---
                # Logika pemisah fleksibel Anda tetap dipertahankan
                cabang_sby = pd.read_csv(cabang_sby_file, sep=None, engine='python', encoding='utf-8-sig')
                available_cols_cabang_sby = [col for col in columns if col in cabang_sby.columns]
                cabang_sby = cabang_sby[available_cols_cabang_sby].copy()
                
                # (DIHAPUS) Blok 'if gantungan_cabang_sby_file is not None' dihapus

                # (Logika asli Anda berlanjut)
                cabang_sby["ID_1"] = np.nan
                cabang_sby = cabang_sby.reset_index()

                cabang_sby["Debet"] = cabang_sby["Debet"].replace("-", 0)
                cabang_sby["Kredit"] = cabang_sby["Kredit"].replace("-", 0)
                cabang_sby["Debet"] = cabang_sby["Debet"].astype(str).str.replace(",", "", regex=False)
                cabang_sby["Kredit"] = cabang_sby["Kredit"].astype(str).str.replace(",", "", regex=False)
                cabang_sby["Debet"] = pd.to_numeric(cabang_sby["Debet"]).fillna(0)
                cabang_sby["Kredit"] = pd.to_numeric(cabang_sby["Kredit"]).fillna(0)

                cabang_sby_PN = cabang_sby[cabang_sby['Keperluan'].str.contains("PEMBAYARAN ATAS NOTA", case=False, na=False)].copy()
                cabang_sby = cabang_sby[~cabang_sby['Keperluan'].str.contains("PEMBAYARAN ATAS NOTA", case=False, na=False)].copy()

                cabang_sby["ID_1"] = cabang_sby["Keperluan"].str.extract(r'(?:ID)?BKK\s*[:\-]?\s*(\d+/\d{4})', expand=False)
                cabang_sby_bkk = cabang_sby[cabang_sby["ID_1"].notna()].copy()
                cabang_sby = cabang_sby[cabang_sby["ID_1"].isna()].copy()

                cabang_sby["ID_1"] = cabang_sby["Keperluan"].str.extract(r'(?:ID)?BKM\s*[:\-]?\s*(\d+/\d{4})', expand=False)
                cabang_sby_bkm = cabang_sby[cabang_sby["ID_1"].notna()].copy()
                cabang_sby = cabang_sby[cabang_sby["ID_1"].isna()].copy()

                cabang_sby_va_ri = cabang_sby[
                    cabang_sby["Keperluan"].str.contains("PENERIMAAN GIRO DENGAN VA|KODE LAWAN RI", case=False, na=False)
                ].copy()
                cabang_sby = cabang_sby[
                    ~cabang_sby["Keperluan"].str.contains("PENERIMAAN GIRO DENGAN VA|KODE LAWAN RI", case=False, na=False)
                ].copy()
                
                # --- PROSES SBY CABANG ---
                # Logika pemisah fleksibel Anda tetap dipertahankan
                sby_cabang = pd.read_csv(sby_cabang_file, sep=None, engine='python', encoding='utf-8-sig')
                available_cols_sby_cabang = [col for col in columns if col in sby_cabang.columns]
                sby_cabang = sby_cabang[available_cols_sby_cabang].copy()

                # (DIHAPUS) Blok 'if gantungan_sby_cabang_file is not None' dihapus
                
                # (Logika asli Anda berlanjut)
                sby_cabang["ID_1"] = np.nan
                sby_cabang = sby_cabang.reset_index()

                sby_cabang["Debet"] = sby_cabang["Debet"].replace("-", 0)
                sby_cabang["Kredit"] = sby_cabang["Kredit"].replace("-", 0)
                sby_cabang["Debet"] = sby_cabang["Debet"].astype(str).str.replace(",", "", regex=False)
                sby_cabang["Kredit"] = sby_cabang["Kredit"].astype(str).str.replace(",", "", regex=False)
                sby_cabang["Debet"] = pd.to_numeric(sby_cabang["Debet"]).fillna(0)
                sby_cabang["Kredit"] = pd.to_numeric(sby_cabang["Kredit"]).fillna(0)

                sby_cabang_PN = sby_cabang[sby_cabang['Keperluan'].str.contains("PEMBAYARAN ATAS NOTA", case=False, na=False)].copy()
                sby_cabang = sby_cabang[~sby_cabang['Keperluan'].str.contains("PEMBAYARAN ATAS NOTA", case=False, na=False)].copy()

                sby_cabang_jmu = sby_cabang[
                    sby_cabang["Keperluan"].str.contains("JMU ASD|JMU ASK", case=False, na=False)
                ].copy()
                sby_cabang = sby_cabang[~sby_cabang["Keperluan"].str.contains("JMU ASD|JMU ASK", case=False, na=False)].copy()

                sby_cabang["ID_1"] = sby_cabang["Keperluan"].str.extract(r'(?:ID)?BKK\s*[:\-]?\s*(\d+/\d{4})', expand=False)
                sby_cabang_bkk = sby_cabang[sby_cabang["ID_1"].notna()].copy()
                sby_cabang = sby_cabang[sby_cabang["ID_1"].isna()].copy()

                sby_cabang["ID_1"] = sby_cabang["Keperluan"].str.extract(r'(?:ID)?BKM\s*[:\-]?\s*(\d+/\d{4})', expand=False)
                sby_cabang_bkm = sby_cabang[sby_cabang["ID_1"].notna()].copy()
                sby_cabang = sby_cabang[sby_cabang["ID_1"].isna()].copy()

                sby_cabang_va_ri = sby_cabang[
                    sby_cabang["Keperluan"].str.contains("PEMBAYARAN DPP GIRO|KODE LAWAN RO", case=False, na=False)
                ].copy()
                sby_cabang = sby_cabang[
                    ~sby_cabang["Keperluan"].str.contains("PEMBAYARAN DPP GIRO|KODE LAWAN RO", case=False, na=False)
                ].copy()

                # --- PROSES TOTAL (VA/RI) --- (TIDAK DIUBAH)
                df_sebelumnya = pd.DataFrame({"Debet": [selisih_sebelumnya]})
                cabang_sby_va_ri = pd.concat([df_sebelumnya, cabang_sby_va_ri], ignore_index=True)

                total_va_ri = pd.DataFrame({
                    "Debet": [cabang_sby_va_ri["Debet"].sum() + sby_cabang_va_ri["Debet"].sum()],
                    "Kredit": [cabang_sby_va_ri["Kredit"].sum() + sby_cabang_va_ri["Kredit"].sum()],
                    "Tempat Pembayaran": [(cabang_sby_va_ri["Debet"].sum() + sby_cabang_va_ri["Debet"].sum()) - (cabang_sby_va_ri["Kredit"].sum() + sby_cabang_va_ri["Kredit"].sum())]
                })

                sby_cabang_va_ri_total = pd.concat([sby_cabang_va_ri, total_va_ri], ignore_index=True)
                cabang_sby_va_ri_total = pd.concat([cabang_sby_va_ri, total_va_ri], ignore_index=True)
                sby_cabang_va_ri_total["Grup"] = "A1"
                cabang_sby_va_ri_total["Grup"] = "A1"


                # --- PROSES TOTAL (PN) --- (TIDAK DIUBAH)
                total_row_PN = pd.DataFrame({
                    "Debet": [cabang_sby_PN["Debet"].sum() + sby_cabang_PN["Debet"].sum()],
                    "Kredit": [cabang_sby_PN["Kredit"].sum() + sby_cabang_PN["Kredit"].sum()],
                    "Tempat Pembayaran": [(cabang_sby_PN["Debet"].sum() + sby_cabang_PN["Debet"].sum()) - (cabang_sby_PN["Kredit"].sum() + sby_cabang_PN["Kredit"].sum())]
                })
                cabang_sby_PN_total = pd.concat([cabang_sby_PN, total_row_PN], ignore_index=True)
                sby_cabang_PN_total = pd.concat([sby_cabang_PN, total_row_PN], ignore_index=True)
                cabang_sby_PN_total["Grup"] = "A2"
                sby_cabang_PN_total["Grup"] = "A2"

                # --- PROSES MATCHING (BKK/BKM) --- (TIDAK DIUBAH)
                list_bkk_cabang_sby = cabang_sby_bkk["ID_1"].unique().tolist()
                sby_cabang_bkk_matched = sby_cabang[sby_cabang["ID Dokumen"].isin(list_bkk_cabang_sby)].copy()
                sby_cabang = sby_cabang[~sby_cabang["ID Dokumen"].isin(list_bkk_cabang_sby)].copy()

                total_row_bkk = pd.DataFrame({
                    "Debet": [cabang_sby_bkk["Debet"].sum() + sby_cabang_bkk_matched["Debet"].sum()],
                    "Kredit": [cabang_sby_bkk["Kredit"].sum() + sby_cabang_bkk_matched["Kredit"].sum()],
                    "Tempat Pembayaran": [(cabang_sby_bkk["Debet"].sum() + sby_cabang_bkk_matched["Debet"].sum()) - (cabang_sby_bkk["Kredit"].sum() + sby_cabang_bkk_matched["Kredit"].sum())]
                })
                cabang_sby_bkk_total = pd.concat([cabang_sby_bkk, total_row_bkk], ignore_index=True)
                sby_cabang_bkk_matched_total = pd.concat([sby_cabang_bkk_matched, total_row_bkk], ignore_index=True)
                cabang_sby_bkk_total["Grup"] = "A3"
                sby_cabang_bkk_matched_total["Grup"] = "A3"


                list_bkm_cabang_sby = cabang_sby_bkm["ID_1"].unique().tolist()
                sby_cabang_bkm_matched = sby_cabang[sby_cabang["ID Dokumen"].isin(list_bkm_cabang_sby)].copy()
                sby_cabang = sby_cabang[~sby_cabang["ID Dokumen"].isin(list_bkm_cabang_sby)].copy()

                total_row_bkm = pd.DataFrame({
                    "Debet": [cabang_sby_bkm["Debet"].sum() + sby_cabang_bkm_matched["Debet"].sum()],
                    "Kredit": [cabang_sby_bkm["Kredit"].sum() + sby_cabang_bkm_matched["Kredit"].sum()],
                    "Tempat Pembayaran": [(cabang_sby_bkm["Debet"].sum() + sby_cabang_bkm_matched["Debet"].sum()) - (cabang_sby_bkm["Kredit"].sum() + sby_cabang_bkm_matched["Kredit"].sum())]
                })
                cabang_sby_bkm_total = pd.concat([cabang_sby_bkm, total_row_bkm], ignore_index=True)
                sby_cabang_bkm_matched_total = pd.concat([sby_cabang_bkm_matched, total_row_bkm], ignore_index=True)
                cabang_sby_bkm_total["Grup"] = "A4"
                sby_cabang_bkm_matched_total["Grup"] = "A4"

                list_bkk_sby_cabang = sby_cabang_bkk["ID_1"].unique().tolist()
                cabang_sby_bkk_matched = cabang_sby[cabang_sby["ID Dokumen"].isin(list_bkk_sby_cabang)].copy()
                cabang_sby = cabang_sby[~cabang_sby["ID Dokumen"].isin(list_bkk_sby_cabang)].copy()

                total_row_bkk_2 = pd.DataFrame({
                    "Debet": [sby_cabang_bkk["Debet"].sum() + cabang_sby_bkk_matched["Debet"].sum()],
                    "Kredit": [sby_cabang_bkk["Kredit"].sum() + cabang_sby_bkk_matched["Kredit"].sum()],
                    "Tempat Pembayaran": [(sby_cabang_bkk["Debet"].sum() + cabang_sby_bkk_matched["Debet"].sum()) - (sby_cabang_bkk["Kredit"].sum() + cabang_sby_bkk_matched["Kredit"].sum())]
                })
                sby_cabang_bkk_total = pd.concat([sby_cabang_bkk, total_row_bkk_2], ignore_index=True)
                cabang_sby_bkk_matched_total = pd.concat([cabang_sby_bkk_matched, total_row_bkk_2], ignore_index=True)
                sby_cabang_bkk_total["Grup"] = "A5"
                cabang_sby_bkk_matched_total["Grup"] = "A5"

                list_bkm_sby_cabang = sby_cabang_bkm["ID_1"].unique().tolist()
                cabang_sby_bkm_matched = cabang_sby[cabang_sby["ID Dokumen"].isin(list_bkm_sby_cabang)].copy()
                cabang_sby = cabang_sby[~cabang_sby["ID Dokumen"].isin(list_bkm_sby_cabang)].copy()

                total_row_bkm_2 = pd.DataFrame({
                    "Debet": [sby_cabang_bkm["Debet"].sum() + cabang_sby_bkm_matched["Debet"].sum()],
                    "Kredit": [sby_cabang_bkm["Kredit"].sum() + cabang_sby_bkm_matched["Kredit"].sum()],
                    "Tempat Pembayaran": [(sby_cabang_bkm["Debet"].sum() + cabang_sby_bkm_matched["Debet"].sum()) - (sby_cabang_bkm["Kredit"].sum() + cabang_sby_bkm_matched["Kredit"].sum())]
                })
                sby_cabang_bkm_total = pd.concat([sby_cabang_bkm, total_row_bkm_2], ignore_index=True)
                cabang_sby_bkm_matched_total = pd.concat([cabang_sby_bkm_matched, total_row_bkm_2], ignore_index=True)
                sby_cabang_bkm_total["Grup"] = "A6"
                cabang_sby_bkm_matched_total["Grup"] = "A6"

                # --- PROSES TOTAL (JMU) --- (TIDAK DIUBAH)
                total_jmu = pd.DataFrame({
                    "Debet": [sby_cabang_jmu["Debet"].sum()],
                    "Kredit": [sby_cabang_jmu["Kredit"].sum()],
                    "Tempat Pembayaran": [sby_cabang_jmu["Debet"].sum() - sby_cabang_jmu["Kredit"].sum()]
                })
                sby_cabang_jmu_total = pd.concat([sby_cabang_jmu, total_jmu], ignore_index=True)
                sby_cabang_jmu_total["Grup"] = "B1"

                # --- PROSES REKONSILIASI GANTUNGAN (OFFSET) --- (TIDAK DIUBAH)
                cabang_sby["Sumber"] = "cabang_sby"
                sby_cabang["Sumber"] = "sby_cabang"
                gabungan = pd.concat([cabang_sby, sby_cabang], ignore_index=True)

                offset_df, gantung_df = reconcile_accounts_table(gabungan)

                offset_df_cabang_sby = offset_df[offset_df["Sumber"] == "cabang_sby"].copy()
                offset_df_sby_cabang = offset_df[offset_df["Sumber"] == "sby_cabang"].copy()

                total_offset = pd.DataFrame({
                    "Debet": [offset_df_sby_cabang["Debet"].sum() + offset_df_cabang_sby["Debet"].sum()],
                    "Kredit": [offset_df_sby_cabang["Kredit"].sum() + offset_df_cabang_sby["Kredit"].sum()],
                    "Tempat Pembayaran": [(offset_df_sby_cabang["Debet"].sum() + offset_df_cabang_sby["Debet"].sum()) - (offset_df_sby_cabang["Kredit"].sum() + offset_df_cabang_sby["Kredit"].sum())]
                })
                offset_df_cabang_sby_total = pd.concat([offset_df_cabang_sby, total_offset], ignore_index=True)
                offset_df_sby_cabang_total = pd.concat([offset_df_sby_cabang, total_offset], ignore_index=True)
                offset_df_cabang_sby_total["Grup"] = "C1"
                offset_df_sby_cabang_total["Grup"] = "C1"

                gantung_df_cabang_sby = gantung_df[gantung_df["Sumber"] == "cabang_sby"].copy()
                total_gantung_cabang_sby = pd.DataFrame({
                    "Debet": [gantung_df_cabang_sby["Debet"].sum()],
                    "Kredit": [gantung_df_cabang_sby["Kredit"].sum()]
                })
                gantung_df_cabang_sby_total = pd.concat([gantung_df_cabang_sby, total_gantung_cabang_sby], ignore_index=True)
                gantung_df_cabang_sby_total["Grup"] = "D1"

                gantung_sby_cabang = gantung_df[gantung_df["Sumber"] == "sby_cabang"].copy()
                total_gantung_sby_cabang = pd.DataFrame({
                    "Debet": [gantung_sby_cabang["Debet"].sum()],
                    "Kredit": [gantung_sby_cabang["Kredit"].sum()]
                })
                gantung_df_sby_cabang_total = pd.concat([gantung_sby_cabang, total_gantung_sby_cabang], ignore_index=True)
                gantung_df_sby_cabang_total["Grup"] = "D2"

                # --- PROSES FINALISASI & EXPORT --- (TIDAK DIUBAH)
                columns_final = [
                    'Grup',
                    'index', 'Tanggal Kasir', 'ID Dokumen', 'Nomor Dokumen', 'Dibayarkan (ke/dari)', 'Keperluan',
                    'Vessel Voyage', 'Debet', 'Kredit', 'Tempat Pembayaran', 'Pembuat', 'Sumber Dokumen',
                    'Jenis Dokumen', 'Tanggal Delivery', 'Nama Kode', 'Kode Accounting', 'User Pengakuan',
                    'Unit', 'Divisi', 'Flag KBM/KDRT', 'Target_First', 'Target_Jenis', 'Target_Second'
                ]
                
                if 'Sumber' not in columns_final:
                    columns_final.append('Sumber')
                if 'Posisi' not in columns_final:
                    columns_final.append('Posisi')


                list_df_names = [
                    "cabang_sby_va_ri_total", "cabang_sby_PN_total", "cabang_sby_bkk_total", "cabang_sby_bkm_total",
                    "cabang_sby_bkk_matched_total", "cabang_sby_bkm_matched_total", "offset_df_cabang_sby_total",
                    "gantung_df_cabang_sby_total", "sby_cabang_va_ri_total", "sby_cabang_PN_total",
                    "sby_cabang_jmu_total", "sby_cabang_bkk_matched_total", "sby_cabang_bkm_matched_total",
                    "sby_cabang_bkk_total", "sby_cabang_bkm_total", "offset_df_sby_cabang_total",
                    "gantung_df_sby_cabang_total"
                ]
                
                for name in list_df_names:
                    if name in globals():
                        df = globals()[name]
                        available_cols = [col for col in columns_final if col in df.columns]
                        globals()[name] = df[available_cols]

                list_df_cabang_sby = [
                    "gantung_df_cabang_sby_total", "cabang_sby_va_ri_total", "cabang_sby_PN_total",
                    "cabang_sby_bkk_total", "cabang_sby_bkm_total", "cabang_sby_bkk_matched_total",
                    "cabang_sby_bkm_matched_total", "offset_df_cabang_sby_total"
                ]

                list_df_sby_cabang = [
                    "gantung_df_sby_cabang_total", "sby_cabang_va_ri_total", "sby_cabang_PN_total",
                    "sby_cabang_jmu_total", "sby_cabang_bkk_matched_total", "sby_cabang_bkm_matched_total",
                    "sby_cabang_bkk_total", "sby_cabang_bkm_total", "offset_df_sby_cabang_total"
                ]

                cabang_sby_all_combined = combine_with_spacing(list_df_cabang_sby, columns_final)
                sby_cabang_all_combined = combine_with_spacing(list_df_sby_cabang, columns_final)

                # == AKHIR KODE ASLI ANDA ==

                # --- BUAT FILE EXCEL DI MEMORY ---
                output_buffer = io.BytesIO()
                with pd.ExcelWriter(output_buffer, engine="xlsxwriter") as writer:
                    cabang_sby_all_combined.to_excel(writer, sheet_name="cabang_sby", index=False)
                    sby_cabang_all_combined.to_excel(writer, sheet_name="sby_cabang", index=False)
                
                output_buffer.seek(0) 

                st.success("✅ Proses Selesai! File Excel siap diunduh.")
                
                output_file_name = f"hasil_RK_{datetime.datetime.now():%Y%m%d_%H%M}.xlsx"

                # Tampilkan tombol download
                st.download_button(
                    label="📥 Download Hasil Excel",
                    data=output_buffer,
                    file_name=output_file_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

        except Exception as e:
            st.error(f"Terjadi error saat pemrosesan: {e}")
            st.exception(e)