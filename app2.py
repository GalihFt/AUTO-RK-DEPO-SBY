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
st.set_page_config(layout="wide", page_title="Pencocokan Hutang/Piutang DEPO - SBY")

st.title("Pencocokan Hutang/Piutang Afiliasi DEPO - SBY")
st.markdown("Unggah 2 file CSV wajib (dan 2 file gantungan opsional) untuk memulai.")

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
# TATA LETAK INPUT (UI) - DENGAN UPDATE OPSIONAL
# ---------------------------------------------------------------------

# --- Bagian DEPO - SBY ---
st.header("Bagian DEPO - SBY")
col1, col2 = st.columns(2)
with col1:
    depo_sby_file = st.file_uploader("Input File DEPO SBY (Wajib)", type="csv")
with col2:
    gantungan_depo_sby_file = st.file_uploader("Input file gantungan depo sby (Opsional)", type="csv")

# --- Bagian SBY - DEPO ---
st.header("Bagian SBY - DEPO")
col3, col4 = st.columns(2)
with col3:
    sby_depo_file = st.file_uploader("Input file SBY DEPO (Wajib)", type="csv")
with col4:
    gantungan_sby_depo_file = st.file_uploader("Input file gantungan sby depo (Opsional)", type="csv")

# --- Input Selisih ---
st.header("Input Selisih")
selisih_input = st.number_input("Input selisih periode sebelumnya", value=869412210, step=1, help="Masukkan nilai selisih dari periode sebelumnya.")

st.divider()

# Tombol untuk memulai proses
process_button = st.button("Mulai Proses Pencocokan", type="primary", use_container_width=True)

# ---------------------------------------------------------------------
# LOGIKA UTAMA (SAAT TOMBOL DITEKAN - HANYA VALIDASI & LOAD DIUBAH)
# ---------------------------------------------------------------------
if process_button:
    # Validasi input (DIUBAH)
    if not all([depo_sby_file, sby_depo_file]):
        st.error("Harap unggah file WAJIB (DEPO SBY & SBY DEPO) untuk melanjutkan.")
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

                # --- PROSES DEPO SBY (DENGAN KONDISI) ---
                depo_sby = pd.read_csv(depo_sby_file, sep=None, engine='python')
                available_cols_depo_sby = [col for col in columns if col in depo_sby.columns]
                depo_sby = depo_sby[available_cols_depo_sby].copy()
                
                # (DIUBAH) Hanya muat dan gabung jika file diunggah
                if gantungan_depo_sby_file is not None:
                    depo_sby_gantungan = pd.read_csv(gantungan_depo_sby_file, sep=None, engine='python')
                    available_cols_gantungan = [col for col in columns if col in depo_sby_gantungan.columns]
                    depo_sby_gantungan = depo_sby_gantungan[available_cols_gantungan].copy()
                    depo_sby = pd.concat([depo_sby, depo_sby_gantungan], ignore_index=True)

                # (Logika asli Anda berlanjut)
                depo_sby["ID_1"] = np.nan
                depo_sby = depo_sby.reset_index()

                depo_sby["Debet"] = depo_sby["Debet"].replace("-", 0)
                depo_sby["Kredit"] = depo_sby["Kredit"].replace("-", 0)
                depo_sby["Debet"] = depo_sby["Debet"].astype(str).str.replace(",", "", regex=False)
                depo_sby["Kredit"] = depo_sby["Kredit"].astype(str).str.replace(",", "", regex=False)
                depo_sby["Debet"] = pd.to_numeric(depo_sby["Debet"]).fillna(0)
                depo_sby["Kredit"] = pd.to_numeric(depo_sby["Kredit"]).fillna(0)

                depo_sby_PN = depo_sby[depo_sby['Keperluan'].str.contains("PEMBAYARAN ATAS NOTA", case=False, na=False)].copy()
                depo_sby = depo_sby[~depo_sby['Keperluan'].str.contains("PEMBAYARAN ATAS NOTA", case=False, na=False)].copy()

                depo_sby["ID_1"] = depo_sby["Keperluan"].str.extract(r'(?:ID)?BKK\s*[:\-]?\s*(\d+/\d{4})', expand=False)
                depo_sby_bkk = depo_sby[depo_sby["ID_1"].notna()].copy()
                depo_sby = depo_sby[depo_sby["ID_1"].isna()].copy()

                depo_sby["ID_1"] = depo_sby["Keperluan"].str.extract(r'(?:ID)?BKM\s*[:\-]?\s*(\d+/\d{4})', expand=False)
                depo_sby_bkm = depo_sby[depo_sby["ID_1"].notna()].copy()
                depo_sby = depo_sby[depo_sby["ID_1"].isna()].copy()

                depo_sby_va_ri = depo_sby[
                    depo_sby["Keperluan"].str.contains("PENERIMAAN GIRO DENGAN VA|KODE LAWAN RI", case=False, na=False)
                ].copy()
                depo_sby = depo_sby[
                    ~depo_sby["Keperluan"].str.contains("PENERIMAAN GIRO DENGAN VA|KODE LAWAN RI", case=False, na=False)
                ].copy()
                
                # --- PROSES SBY DEPO (DENGAN KONDISI) ---
                sby_depo = pd.read_csv(sby_depo_file, sep=None, engine='python')
                available_cols_sby_depo = [col for col in columns if col in sby_depo.columns]
                sby_depo = sby_depo[available_cols_sby_depo].copy()

                # (DIUBAH) Hanya muat dan gabung jika file diunggah
                if gantungan_sby_depo_file is not None:
                    sby_depo_gantungan = pd.read_csv(gantungan_sby_depo_file, sep=None, engine='python')
                    available_cols_gantungan_sby_depo = [col for col in columns if col in sby_depo_gantungan.columns]
                    sby_depo_gantungan = sby_depo_gantungan[available_cols_gantungan_sby_depo].copy()
                    sby_depo = pd.concat([sby_depo, sby_depo_gantungan], ignore_index=True)
                
                # (Logika asli Anda berlanjut)
                sby_depo["ID_1"] = np.nan
                sby_depo = sby_depo.reset_index()

                sby_depo["Debet"] = sby_depo["Debet"].replace("-", 0)
                sby_depo["Kredit"] = sby_depo["Kredit"].replace("-", 0)
                sby_depo["Debet"] = sby_depo["Debet"].astype(str).str.replace(",", "", regex=False)
                sby_depo["Kredit"] = sby_depo["Kredit"].astype(str).str.replace(",", "", regex=False)
                sby_depo["Debet"] = pd.to_numeric(sby_depo["Debet"]).fillna(0)
                sby_depo["Kredit"] = pd.to_numeric(sby_depo["Kredit"]).fillna(0)

                sby_depo_PN = sby_depo[sby_depo['Keperluan'].str.contains("PEMBAYARAN ATAS NOTA", case=False, na=False)].copy()
                sby_depo = sby_depo[~sby_depo['Keperluan'].str.contains("PEMBAYARAN ATAS NOTA", case=False, na=False)].copy()

                sby_depo_jmu = sby_depo[
                    sby_depo["Keperluan"].str.contains("JMU ASD|JMU ASK", case=False, na=False)
                ].copy()
                sby_depo = sby_depo[~sby_depo["Keperluan"].str.contains("JMU ASD|JMU ASK", case=False, na=False)].copy()

                sby_depo["ID_1"] = sby_depo["Keperluan"].str.extract(r'(?:ID)?BKK\s*[:\-]?\s*(\d+/\d{4})', expand=False)
                sby_depo_bkk = sby_depo[sby_depo["ID_1"].notna()].copy()
                sby_depo = sby_depo[sby_depo["ID_1"].isna()].copy()

                sby_depo["ID_1"] = sby_depo["Keperluan"].str.extract(r'(?:ID)?BKM\s*[:\-]?\s*(\d+/\d{4})', expand=False)
                sby_depo_bkm = sby_depo[sby_depo["ID_1"].notna()].copy()
                sby_depo = sby_depo[sby_depo["ID_1"].isna()].copy()

                sby_depo_va_ri = sby_depo[
                    sby_depo["Keperluan"].str.contains("PEMBAYARAN DPP GIRO|KODE LAWAN RO", case=False, na=False)
                ].copy()
                sby_depo = sby_depo[
                    ~sby_depo["Keperluan"].str.contains("PEMBAYARAN DPP GIRO|KODE LAWAN RO", case=False, na=False)
                ].copy()

                # --- PROSES TOTAL (VA/RI) ---
                df_sebelumnya = pd.DataFrame({"Debet": [selisih_sebelumnya]})
                depo_sby_va_ri = pd.concat([df_sebelumnya, depo_sby_va_ri], ignore_index=True)

                total_va_ri = pd.DataFrame({
                    "Debet": [depo_sby_va_ri["Debet"].sum() + sby_depo_va_ri["Debet"].sum()],
                    "Kredit": [depo_sby_va_ri["Kredit"].sum() + sby_depo_va_ri["Kredit"].sum()],
                    "Tempat Pembayaran": [(depo_sby_va_ri["Debet"].sum() + sby_depo_va_ri["Debet"].sum()) - (depo_sby_va_ri["Kredit"].sum() + sby_depo_va_ri["Kredit"].sum())]
                })

                sby_depo_va_ri_total = pd.concat([sby_depo_va_ri, total_va_ri], ignore_index=True)
                depo_sby_va_ri_total = pd.concat([depo_sby_va_ri, total_va_ri], ignore_index=True)

                # --- PROSES TOTAL (PN) ---
                total_row_PN = pd.DataFrame({
                    "Debet": [depo_sby_PN["Debet"].sum() + sby_depo_PN["Debet"].sum()],
                    "Kredit": [depo_sby_PN["Kredit"].sum() + sby_depo_PN["Kredit"].sum()],
                    "Tempat Pembayaran": [(depo_sby_PN["Debet"].sum() + sby_depo_PN["Debet"].sum()) - (depo_sby_PN["Kredit"].sum() + sby_depo_PN["Kredit"].sum())]
                })
                depo_sby_PN_total = pd.concat([depo_sby_PN, total_row_PN], ignore_index=True)
                sby_depo_PN_total = pd.concat([sby_depo_PN, total_row_PN], ignore_index=True)

                # --- PROSES MATCHING (BKK/BKM) ---
                list_bkk_depo_sby = depo_sby_bkk["ID_1"].unique().tolist()
                sby_depo_bkk_matched = sby_depo[sby_depo["ID Dokumen"].isin(list_bkk_depo_sby)].copy()
                sby_depo = sby_depo[~sby_depo["ID Dokumen"].isin(list_bkk_depo_sby)].copy()

                total_row_bkk = pd.DataFrame({
                    "Debet": [depo_sby_bkk["Debet"].sum() + sby_depo_bkk_matched["Debet"].sum()],
                    "Kredit": [depo_sby_bkk["Kredit"].sum() + sby_depo_bkk_matched["Kredit"].sum()],
                    "Tempat Pembayaran": [(depo_sby_bkk["Debet"].sum() + sby_depo_bkk_matched["Debet"].sum()) - (depo_sby_bkk["Kredit"].sum() + sby_depo_bkk_matched["Kredit"].sum())]
                })
                depo_sby_bkk_total = pd.concat([depo_sby_bkk, total_row_bkk], ignore_index=True)
                sby_depo_bkk_matched_total = pd.concat([sby_depo_bkk_matched, total_row_bkk], ignore_index=True)

                list_bkm_depo_sby = depo_sby_bkm["ID_1"].unique().tolist()
                sby_depo_bkm_matched = sby_depo[sby_depo["ID Dokumen"].isin(list_bkm_depo_sby)].copy()
                sby_depo = sby_depo[~sby_depo["ID Dokumen"].isin(list_bkm_depo_sby)].copy()

                total_row_bkm = pd.DataFrame({
                    "Debet": [depo_sby_bkm["Debet"].sum() + sby_depo_bkm_matched["Debet"].sum()],
                    "Kredit": [depo_sby_bkm["Kredit"].sum() + sby_depo_bkm_matched["Kredit"].sum()],
                    "Tempat Pembayaran": [(depo_sby_bkm["Debet"].sum() + sby_depo_bkm_matched["Debet"].sum()) - (depo_sby_bkm["Kredit"].sum() + sby_depo_bkm_matched["Kredit"].sum())]
                })
                depo_sby_bkm_total = pd.concat([depo_sby_bkm, total_row_bkm], ignore_index=True)
                sby_depo_bkm_matched_total = pd.concat([sby_depo_bkm_matched, total_row_bkm], ignore_index=True)

                list_bkk_sby_depo = sby_depo_bkk["ID_1"].unique().tolist()
                depo_sby_bkk_matched = depo_sby[depo_sby["ID Dokumen"].isin(list_bkk_sby_depo)].copy()
                depo_sby = depo_sby[~depo_sby["ID Dokumen"].isin(list_bkk_sby_depo)].copy()

                total_row_bkk_2 = pd.DataFrame({
                    "Debet": [sby_depo_bkk["Debet"].sum() + depo_sby_bkk_matched["Debet"].sum()],
                    "Kredit": [sby_depo_bkk["Kredit"].sum() + depo_sby_bkk_matched["Kredit"].sum()],
                    "Tempat Pembayaran": [(sby_depo_bkk["Debet"].sum() + depo_sby_bkk_matched["Debet"].sum()) - (sby_depo_bkk["Kredit"].sum() + depo_sby_bkk_matched["Kredit"].sum())]
                })
                sby_depo_bkk_total = pd.concat([sby_depo_bkk, total_row_bkk_2], ignore_index=True)
                depo_sby_bkk_matched_total = pd.concat([depo_sby_bkk_matched, total_row_bkk_2], ignore_index=True)

                list_bkm_sby_depo = sby_depo_bkm["ID_1"].unique().tolist()
                depo_sby_bkm_matched = depo_sby[depo_sby["ID Dokumen"].isin(list_bkm_sby_depo)].copy()
                depo_sby = depo_sby[~depo_sby["ID Dokumen"].isin(list_bkm_sby_depo)].copy()

                total_row_bkm_2 = pd.DataFrame({
                    "Debet": [sby_depo_bkm["Debet"].sum() + depo_sby_bkm_matched["Debet"].sum()],
                    "Kredit": [sby_depo_bkm["Kredit"].sum() + depo_sby_bkm_matched["Kredit"].sum()],
                    "Tempat Pembayaran": [(sby_depo_bkm["Debet"].sum() + depo_sby_bkm_matched["Debet"].sum()) - (sby_depo_bkm["Kredit"].sum() + depo_sby_bkm_matched["Kredit"].sum())]
                })
                sby_depo_bkm_total = pd.concat([sby_depo_bkm, total_row_bkm_2], ignore_index=True)
                depo_sby_bkm_matched_total = pd.concat([depo_sby_bkm_matched, total_row_bkm_2], ignore_index=True)

                # --- PROSES TOTAL (JMU) ---
                total_jmu = pd.DataFrame({
                    "Debet": [sby_depo_jmu["Debet"].sum()],
                    "Kredit": [sby_depo_jmu["Kredit"].sum()],
                    "Tempat Pembayaran": [sby_depo_jmu["Debet"].sum() - sby_depo_jmu["Kredit"].sum()]
                })
                sby_depo_jmu_total = pd.concat([sby_depo_jmu, total_jmu], ignore_index=True)

                # --- PROSES REKONSILIASI GANTUNGAN (OFFSET) ---
                depo_sby["Sumber"] = "depo_sby"
                sby_depo["Sumber"] = "sby_depo"
                gabungan = pd.concat([depo_sby, sby_depo], ignore_index=True)

                offset_df, gantung_df = reconcile_accounts_table(gabungan)

                offset_df_depo_sby = offset_df[offset_df["Sumber"] == "depo_sby"].copy()
                offset_df_sby_depo = offset_df[offset_df["Sumber"] == "sby_depo"].copy()

                total_offset = pd.DataFrame({
                    "Debet": [offset_df_sby_depo["Debet"].sum() + offset_df_depo_sby["Debet"].sum()],
                    "Kredit": [offset_df_sby_depo["Kredit"].sum() + offset_df_depo_sby["Kredit"].sum()],
                    "Tempat Pembayaran": [(offset_df_sby_depo["Debet"].sum() + offset_df_depo_sby["Debet"].sum()) - (offset_df_sby_depo["Kredit"].sum() + offset_df_depo_sby["Kredit"].sum())]
                })
                offset_df_depo_sby_total = pd.concat([offset_df_depo_sby, total_offset], ignore_index=True)
                offset_df_sby_depo_total = pd.concat([offset_df_sby_depo, total_offset], ignore_index=True)

                gantung_df_depo_sby = gantung_df[gantung_df["Sumber"] == "depo_sby"].copy()
                total_gantung_depo_sby = pd.DataFrame({
                    "Debet": [gantung_df_depo_sby["Debet"].sum()],
                    "Kredit": [gantung_df_depo_sby["Kredit"].sum()]
                })
                gantung_df_depo_sby_total = pd.concat([gantung_df_depo_sby, total_gantung_depo_sby], ignore_index=True)

                gantung_sby_depo = gantung_df[gantung_df["Sumber"] == "sby_depo"].copy()
                total_gantung_sby_depo = pd.DataFrame({
                    "Debet": [gantung_sby_depo["Debet"].sum()],
                    "Kredit": [gantung_sby_depo["Kredit"].sum()]
                })
                gantung_df_sby_depo_total = pd.concat([gantung_sby_depo, total_gantung_sby_depo], ignore_index=True)

                # --- PROSES FINALISASI & EXPORT ---
                columns_final = [
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
                    "depo_sby_va_ri_total", "depo_sby_PN_total", "depo_sby_bkk_total", "depo_sby_bkm_total",
                    "depo_sby_bkk_matched_total", "depo_sby_bkm_matched_total", "offset_df_depo_sby_total",
                    "gantung_df_depo_sby_total", "sby_depo_va_ri_total", "sby_depo_PN_total",
                    "sby_depo_jmu_total", "sby_depo_bkk_matched_total", "sby_depo_bkm_matched_total",
                    "sby_depo_bkk_total", "sby_depo_bkm_total", "offset_df_sby_depo_total",
                    "gantung_df_sby_depo_total"
                ]
                
                for name in list_df_names:
                    if name in globals():
                        df = globals()[name]
                        available_cols = [col for col in columns_final if col in df.columns]
                        globals()[name] = df[available_cols]

                list_df_depo_sby = [
                    "gantung_df_depo_sby_total", "depo_sby_va_ri_total", "depo_sby_PN_total",
                    "depo_sby_bkk_total", "depo_sby_bkm_total", "depo_sby_bkk_matched_total",
                    "depo_sby_bkm_matched_total", "offset_df_depo_sby_total"
                ]

                list_df_sby_depo = [
                    "gantung_df_sby_depo_total", "sby_depo_va_ri_total", "sby_depo_PN_total",
                    "sby_depo_jmu_total", "sby_depo_bkk_matched_total", "sby_depo_bkm_matched_total",
                    "sby_depo_bkk_total", "sby_depo_bkm_total", "offset_df_sby_depo_total"
                ]

                depo_sby_all_combined = combine_with_spacing(list_df_depo_sby, columns_final)
                sby_depo_all_combined = combine_with_spacing(list_df_sby_depo, columns_final)

                # == AKHIR KODE ASLI ANDA ==

                # --- BUAT FILE EXCEL DI MEMORY ---
                output_buffer = io.BytesIO()
                with pd.ExcelWriter(output_buffer, engine="xlsxwriter") as writer:
                    depo_sby_all_combined.to_excel(writer, sheet_name="depo_sby", index=False)
                    sby_depo_all_combined.to_excel(writer, sheet_name="sby_depo", index=False)
                
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