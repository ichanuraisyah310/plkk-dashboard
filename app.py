"""
Dashboard Utilization Review PLKK  —  versi Streamlit
=====================================================
Jalankan di VS Code (terminal):

    pip install -r requirements.txt
    streamlit run app.py

Dashboard langsung tampil memakai data bawaan (data/Data_Dummy.xlsx).
File Excel lain (format template BMIV-01 s.d. BMIV-04) bisa dipakai lewat
sidebar → "Ganti sumber data (opsional)".
"""
from __future__ import annotations

import html
import io
import re
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ─────────────────────────────────────────────────────────────────────────────
# 1. KONSTANTA & SPESIFIKASI TEMPLATE
# ─────────────────────────────────────────────────────────────────────────────
BS = ["RJTP", "RJTL", "RANAP", "Khusus"]
BN = ["BMIV-01 · RJTP", "BMIV-02 · RJTL", "BMIV-03 · RANAP", "BMIV-04 · Khusus"]
BC = ["#2a9d8f", "#3b82f6", "#e9a23b", "#a855f7"]
BD = [
    "Rawat Jalan Tingkat Pertama — dokter umum, gigi, obat, penunjang diagnostik, tindakan medis.",
    "Rawat Jalan Tingkat Lanjutan — dokter spesialis, laboratorium, radiologi, fisioterapi, ODC, emergensi.",
    "Rawat Inap — akomodasi, operasi, ICU, transfusi, implan ortopedi, penunjang.",
    "Pelayanan khusus & alat bantu — kaca mata, gigi palsu, protesis, ambulans, program RTW.",
]
MON = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"]
ACC = "#0f766e"
MON_FULL = ["JANUARI", "FEBRUARI", "MARET", "APRIL", "MEI", "JUNI", "JULI", "AGUSTUS", "SEPTEMBER",
            "OKTOBER", "NOVEMBER", "DESEMBER"]
# komponen yang hanya berisi volume informasi (bukan biaya) — tidak ikut tab analisis biasa
INFO_CODES = {"RJTP-03K", "RJTP-03L", "RJTP-RUJ", "RI-02K", "RI-02L"}

# Metadata kolom tiap sheet BMIV  (index = kode BMIV 0..3)
SHEETS = [
    dict(key="BMIV-01", date="TGL PELAYANAN", plkk="NAMA PLKK TK I",
         total="TOTAL BIAYA DISETUJUI", visit="KUNJUNGAN KE-"),
    dict(key="BMIV-02", date="TGL PELAYANAN", plkk="NAMA PLKK TK II / RS",
         total="TOTAL BIAYA DISETUJUI", visit="KUNJUNGAN KE-", poli="POLI / SPESIALISASI"),
    dict(key="BMIV-03", date="TGL MASUK", plkk="NAMA PLKK TK II / RS",
         total="TOTAL BIAYA DISETUJUI", los="JML HR RWT", kelas="KELAS RAWAT"),
    dict(key="BMIV-04", date="TGL PELAYANAN", plkk="TEMPAT PELAYANAN / PENYEDIA",
         total="JUMLAH JAMINAN (DISETUJUI)"),
]
COMMON_COLS = ["TGL KECELAKAAN", "NAMA TK / PASIEN", "NO KPJ", "NAMA PERUSAHAAN",
               "KODE CAB", "NO. KASUS KK", "JENIS KASUS", "DIAGNOSA", "KODE ICD-10", "KANWIL"]

# Komponen layanan: (kode, nama, satuan, [kolom volume], [kolom biaya])
def _pairs(rows, prefix, start=1):
    out = []
    for i, (nama, sat, vol, biaya) in enumerate(rows, start):
        out.append((f"{prefix}-{i:02d}", nama, sat, [vol], [biaya]))
    return out


SPEC = {
    0: [
        ("RJTP-01", "Dokter Umum", "kunjungan", ["JASA MEDIS DR UMUM | JML KUNJ"], ["BIAYA (RP)"]),
        ("RJTP-02", "Tindakan Gigi", "tindakan",
         ["JASA MEDIS DR GIGI (TRAUMA GIGI) | PX + KONSUL", "TS", "TP", "ExGT", "ExKOMP", "SPLINT / REPOSISI"],
         ["BIAYA (RP) [2]"]),
        ("RJTP-03", "Obat/Resep", "resep", ["RX | JML"],
         ["BIAYA (RP) [3]", "BIAYA (RP) [4]", "BIAYA (RP) [5]"]),
        ("RJTP-04", "Penunjang Diagnostik", "kunjungan",
         ["PENUNJANG DIAGNOSTIK SEDERHANA | DR", "UR", "RO EKSTREMITAS", "RO THORAX / LAIN", "RO GIGI"],
         ["BIAYA (RP) [6]"]),
        ("RJTP-05", "Tindakan Medis Dokter Umum", "tindakan",
         ["TIND MEDIS DR UMUM | JAHIT LUKA", "BERSIH LUKA", "GANTI BALUT", "INSISI", "EKSISI", "EKSPLORASI",
          "DEBRIDEMENT", "EKSTRAKSI BENDA ASING", "BIDAI / SPALK", "REPOSISI SEDERHANA", "LUKA BAKAR"],
         ["BIAYA (RP) [7]"]),
        ("RJTP-06", "Tindakan Medis Dokter Gigi", "tindakan",
         ["TIND MEDIS DR GIGI | ODONTEKTOMI", "ALVEOLEKTOMI", "FLAP OPERASI", "INSISI [2]", "EKSISI [2]",
          "REPLANTASI / REPOSISI"],
         ["BIAYA (RP) [8]"]),
        # baris informasi (volume saja, tidak dijumlah ke biaya) — dipakai LB-ST
        ("RJTP-03K", "Obat - kasus, KS (informasi)", "kasus", ["OBAT | KS | JML"], []),
        ("RJTP-03L", "Obat - lembar resep, LB (informasi)", "lembar", ["LB | JML"], []),
    ],
    1: [
        ("RJTL-01", "Jasa Dokter Spesialis", "kunjungan", ["JASA DOKTER | JML KUNJ"], ["BIAYA (RP)"]),
        ("RJTL-02", "Obat KS", "kasus", ["OBAT / RESEP | KS | JML"], ["BIAYA (RP) [2]"]),
        ("RJTL-03", "Obat LB", "kasus", ["LB | JML"], ["BIAYA (RP) [3]"]),
        ("RJTL-04", "Obat RX", "resep", ["RX | JML"], ["BIAYA (RP) [4]"]),
        ("RJTL-05", "Laboratorium", "kunjungan", ["LAB | JML KUNJ"], ["BIAYA (RP) [5]"]),
        ("RJTL-06", "Radiologi (Rontgen)", "kunjungan", ["RONTGEN | JML KUNJ"], ["BIAYA (RP) [6]"]),
        ("RJTL-07", "EKG", "kunjungan", ["ELEKTROMEDIK / PEMERIKSAAN KHUSUS | EKG | JML KUNJ"], ["BIAYA (RP) [7]"]),
        ("RJTL-08", "CT Scan", "kunjungan", ["CT SCAN | JML KUNJ"], ["BIAYA (RP) [8]"]),
        ("RJTL-09", "MRI", "kunjungan", ["MRI | JML KUNJ"], ["BIAYA (RP) [9]"]),
        ("RJTL-10", "USG", "kunjungan", ["USG | JML KUNJ"], ["BIAYA (RP) [10]"]),
        ("RJTL-11", "EMG/NCV", "kunjungan", ["EMG / NCV | JML KUNJ"], ["BIAYA (RP) [11]"]),
        ("RJTL-12", "Elektromedik Lain", "kunjungan", ["LAIN-LAIN | JML KUNJ"], ["BIAYA (RP) [12]"]),
        ("RJTL-13", "Tindakan Spesialis", "kunjungan",
         ["TINDAKAN MEDIS DOKTER SPESIALIS | TINDAKAN | JML KUNJ"], ["BIAYA (RP) [13]"]),
        ("RJTL-14", "ODC", "kunjungan", ["ODC (ONE DAY CARE) | JML KUNJ"], ["BIAYA (RP) [14]"]),
        ("RJTL-15", "Fisioterapi", "kasus", ["REHABILITASI MEDIK | FISIOTERAPI | KS"], ["BIAYA (RP) [15]"]),
        ("RJTL-16", "Rehab Medik Lain", "kasus", ["REHABILITASI MEDIK LAIN | KS"], ["BIAYA (RP) [16]"]),
        ("RJTL-17", "Emergensi", "kunjungan", ["EMERGENSI | JML KUNJ"], ["BIAYA (RP) [17]"]),
        ("RJTL-18", "Ambulans", "unit", ["AMBULANS | JML"], ["BIAYA (RP) [18]"]),
        ("RJTL-19", "Homecare", "kunjungan", ["HOMECARE | JML KUNJ."], ["BIAYA (RP) [19]"]),
    ],
    2: [
        ("RI-01", "Akomodasi Rawat Inap", "hari", ["JML HR RWT"], ["BIAYA DISETUJUI (RP)"]),
        ("RI-02", "Obat/Resep", "resep", ["RX"], ["BIAYA DISETUJUI (RP) [2]"]),
        ("RI-03", "Tindakan Non Operatif", "kunjungan", ["TINDAKAN NON OPERATIF | JML KUNJ"], ["BIAYA (RP)"]),
        ("RI-04", "Operasi Kecil", "kunjungan", ["TINDAKAN OPERATIF | KECIL | JML KUNJ"], ["BIAYA (RP) [2]"]),
        ("RI-05", "Operasi Sedang", "kunjungan", ["SEDANG | JML KUNJ"], ["BIAYA (RP) [3]"]),
        ("RI-06", "Operasi Besar", "kunjungan", ["BESAR | JML KUNJ"], ["BIAYA (RP) [4]"]),
        ("RI-07", "Operasi Khusus", "kunjungan", ["KHUSUS | JML KUNJ"], ["BIAYA (RP) [5]"]),
        ("RI-08", "ICU/ICCU/HCU/Burn Unit", "hari", ["JML HR RWT [2]"], ["BIAYA DISETUJUI (RP) [4]"]),
        ("RI-09", "Lab", "kunjungan", ["LAB | JML KUNJ"], ["BIAYA (RP) [6]"]),
        ("RI-10", "Rontgen", "kunjungan", ["RONTGEN | JML KUNJ"], ["BIAYA (RP) [7]"]),
        ("RI-11", "EKG", "kunjungan", ["ELEKTROMEDIK / PEMERIKSAAN KHUSUS | EKG | JML KUNJ"], ["BIAYA (RP) [8]"]),
        ("RI-12", "CT Scan", "kunjungan", ["CT SCAN | JML KUNJ"], ["BIAYA (RP) [9]"]),
        ("RI-13", "MRI", "kunjungan", ["MRI | JML KUNJ"], ["BIAYA (RP) [10]"]),
        ("RI-14", "USG", "kunjungan", ["USG | JML KUNJ"], ["BIAYA (RP) [11]"]),
        ("RI-15", "EMG/NCV", "kunjungan", ["EMG / NCV | JML KUNJ"], ["BIAYA (RP) [12]"]),
        ("RI-16", "Elektromedik Lain", "kunjungan", ["LAIN-LAIN | JML KUNJ"], ["BIAYA (RP) [13]"]),
        ("RI-17", "Pelayanan Darah", "kasus", ["PELAYANAN DARAH | KS"], ["BIAYA (RP) [14]"]),
        ("RI-18", "Transfusi Darah", "kasus", ["TRANSFUSI DARAH | KS"], ["BIAYA (RP) [15]"]),
        ("RI-19", "Implan Ortopedi", "kasus",
         ["ALAT KESEHATAN IMPLAN ORTOPEDI (PIN, PLATE, SCREW, DLL) | KS"], ["BIAYA (RP) [16]"]),
        ("RI-20", "Fisioterapi", "kasus", ["REHABILITASI MEDIK | FISIOTERAPI | KS"], ["BIAYA (RP) [17]"]),
        ("RI-21", "Rehab Medik Lain", "kasus", ["REHABILITASI MEDIK LAIN | KS"], ["BIAYA (RP) [18]"]),
        ("RI-22", "Ambulans", "unit", ["AMBULANS | JML"], ["BIAYA (RP) [19]"]),
        ("RI-23", "Homecare", "kunjungan", ["HOMECARE | JML KUNJ."], ["BIAYA (RP) [20]"]),
        ("RI-02L", "Obat - lembar resep, LB (informasi)", "lembar", ["OBAT / RESEP | LB"], []),
        ("RI-02K", "Obat - kasus, KS (informasi)", "kasus", ["KS"], []),
    ],
    3: [
        ("AB-01", "Kacamata", "kasus", ["KACAMATA | KS"], ["BIAYA (RP)"]),
        ("AB-02", "Gigi Palsu", "kasus", ["GIGI PALSU | KS"], ["BIAYA (RP) [2]"]),
        ("AB-03", "Protesis Anggota Gerak", "kasus", ["PROTESIS ANGGOTA GERAK | KS"], ["BIAYA (RP) [3]"]),
        ("AB-04", "Ortosis", "kasus", ["ORTOSIS | KS"], ["BIAYA (RP) [4]"]),
        ("AB-05", "Alat Bantu Jalan", "kasus", ["ALAT BANTU JALAN | KS"], ["BIAYA (RP) [5]"]),
        ("AB-06", "Alat Bantu Dengar", "kasus", ["ALAT BANTU DENGAR | KS"], ["BIAYA (RP) [6]"]),
        ("AB-07", "Mata Palsu", "kasus", ["MATA PALSU | KS"], ["BIAYA (RP) [7]"]),
        ("AB-08", "Ambulans/Transportasi", "kasus", ["AMBULANS / TRANSPORTASI | KS"], ["BIAYA (RP) [8]"]),
        ("AB-09", "Program Kembali Bekerja (RTW)", "kasus", ["PROGRAM KEMBALI BEKERJA (RTW) | KS"],
         ["BIAYA (RP) [9]"]),
    ],
}

# Nama cabang per KODE CAB (di template raw hanya ada kodenya)
CABANG = {
    "K00": "Jakarta Raya (cabang pelapor)", "K01": "Tangerang", "K02": "Bekasi-Karawang", "K03": "Bandung",
    "K04": "Cirebon", "K05": "Serang-Cilegon", "K06": "Semarang", "K07": "Yogyakarta", "K08": "Surakarta",
    "K09": "Surabaya", "K10": "Sidoarjo", "K11": "Malang", "K12": "Denpasar", "K13": "Mataram",
    "K14": "Kupang", "K15": "Medan", "K16": "Banda Aceh", "K17": "Batam", "K18": "Pekanbaru",
    "K19": "Padang", "K20": "Palembang", "K21": "Bandar Lampung", "K22": "Jambi", "K23": "Pontianak",
    "K24": "Banjarmasin", "K25": "Balikpapan-Samarinda", "K26": "Makassar", "K27": "Manado",
    "K28": "Kendari", "K29": "Jayapura-Timika",
}

# Sektor usaha diturunkan dari kata kunci pada NAMA PERUSAHAAN (urutan = prioritas)
SEKTOR_RULES = [
    ("Jasa Pendidikan", ["pendidikan"]),
    ("Jasa Penunjang", ["cleaning", "sekuriti"]),
    ("Perhotelan", ["hotel"]),
    ("Keuangan", ["finansial"]),
    ("Migas & Pertambangan", ["migas"]),
    ("Teknologi Informasi", ["teknologi"]),
    ("Pertanian & Perkebunan", ["sawit", "pertanian"]),
    ("Konstruksi", ["infrastruktur", "teknik sipil", "konstruksi", "pondasi", "struktur"]),
    ("Industri Manufaktur", ["kertas", "plastik", "kimia", "makanan", "logam"]),
    ("Perdagangan & Ritel", ["perdagangan", "niaga", "mart", "retail"]),
    ("Transportasi & Logistik", ["angkutan", "logistik", "terminal", "ekspedisi", "kargo", "pergudangan"]),
]


def sektor_of(nama: str) -> str:
    s = str(nama).lower()
    for sektor, kws in SEKTOR_RULES:
        if any(k in s for k in kws):
            return sektor
    return "Lainnya"


# ─────────────────────────────────────────────────────────────────────────────
# 2. LOADER: baca template Excel → tabel episode + tabel komponen
# ─────────────────────────────────────────────────────────────────────────────
def _norm(s) -> str:
    return re.sub(r"\s+", " ", str(s)).strip().upper()


class TemplateError(Exception):
    pass


def _find_sheet(xl: pd.ExcelFile, key: str) -> str | None:
    """Cari sheet 'BMIV-0x' (boleh ada akhiran seperti _Dummy)."""
    for name in xl.sheet_names:
        if _norm(name).replace("_", " ").startswith(key):
            return name
    return None


def required_columns(b: int) -> list[str]:
    m = SHEETS[b]
    cols = list(COMMON_COLS) + [m["date"], m["plkk"], m["total"]]
    for k in ("visit", "poli", "los", "kelas"):
        if k in m:
            cols.append(m[k])
    cols += ["JUMLAH KUNJUNGAN / UNIT"] if b == 3 else ["TOTAL BIAYA DIAJUKAN", "TOTAL SELISIH"]
    for _, _, _, vols, biaya in SPEC[b]:
        cols += vols + biaya
    seen, out = set(), []
    for c in cols:
        if c not in seen:
            seen.add(c)
            out.append(c)
    return out


def read_template(file) -> tuple[dict, list[str]]:
    """Baca & validasi 4 sheet BMIV. Return (dict b→DataFrame, daftar peringatan)."""
    try:
        xl = pd.ExcelFile(file)
    except Exception as e:  # noqa: BLE001
        raise TemplateError(f"File tidak bisa dibaca sebagai Excel (.xlsx): {e}")
    frames, warns, errors = {}, [], []
    for b, m in enumerate(SHEETS):
        sname = _find_sheet(xl, m["key"])
        if sname is None:
            errors.append(f"Sheet **{m['key']}** tidak ditemukan.")
            continue
        df = xl.parse(sname)
        df.columns = [_norm(c) for c in df.columns]
        miss = [c for c in required_columns(b) if _norm(c) not in df.columns]
        if miss:
            errors.append(f"Sheet **{sname}** kurang kolom: " + ", ".join(f"`{c}`" for c in miss[:12])
                          + (" …" if len(miss) > 12 else ""))
            continue
        frames[b] = df
    if errors:
        raise TemplateError("\n\n".join(errors))
    return frames, warns


def _num(df: pd.DataFrame, cols: list[str]) -> pd.DataFrame:
    return df[[_norm(c) for c in cols]].apply(pd.to_numeric, errors="coerce").fillna(0.0)


# Kolom penutup (total) & kolom subtotal per sheet — dipakai untuk deteksi pergeseran kolom
FOOTER_COL = {0: "RUJUKAN | R. JALAN", 1: "TOTAL BIAYA DIAJUKAN", 2: "TOTAL BIAYA DIAJUKAN",
              3: "JUMLAH KUNJUNGAN / UNIT"}
SUBTOTAL_COLS = {
    0: [],
    1: ["SUB TOTAL PENUNJANG"],
    2: ["BIAYA DISETUJUI (RP) [3]", "SUB TOTAL PENUNJANG", "SUB TOTAL LAIN-LAIN"],
    3: [],
}


def _plausible(vol: np.ndarray, cost: np.ndarray) -> np.ndarray:
    """Volume = bilangan bulat kecil, biaya = 0 atau nominal rupiah (>= 100)."""
    return (vol >= 0) & (vol <= 200) & (vol == np.floor(vol)) & ((cost == 0) | (cost >= 100))


def read_components(df: pd.DataFrame, b: int, ref: np.ndarray | None = None):
    """
    Baca volume & biaya tiap komponen layanan.
    Header template dan isi data kadang tidak sejajar (mis. file dummy: pasangan [volume, biaya]
    bergeser 1 kolom). Untuk tiap komponen dicoba offset 0 / +1 / -1 dan dipilih yang paling
    masuk akal (volume = bilangan bulat kecil, biaya = nominal rupiah).
    Return: (list komponen, daftar nama komponen yang dikoreksi).
    """
    cols = list(df.columns)
    pos = {c: i for i, c in enumerate(cols)}
    M = df.apply(pd.to_numeric, errors="coerce").fillna(0.0).to_numpy(dtype=float)
    limit = pos.get(_norm(FOOTER_COL[b]), len(cols))
    # kolom subtotal (header subtotal + 1 kolom sesudahnya yang berisi nilai subtotal) tidak dibaca sebagai komponen
    blocked = set()
    for c in SUBTOTAL_COLS[b]:
        if _norm(c) in pos:
            p = pos[_norm(c)]
            blocked.add(p)
            if _norm(c).startswith("SUB TOTAL"):
                blocked.add(p + 1)
    prev, out, shifted = 0, [], []
    for kode, nama, sat, vcols, bcols in SPEC[b]:
        vi = [pos[_norm(c)] for c in vcols]
        bi = [pos[_norm(c)] for c in bcols]
        best = None
        for o in dict.fromkeys([prev, 0, 1]):
            vv = [i + o for i in vi if 0 <= i + o < limit and (i + o) not in blocked]
            bb = [i + o for i in bi if 0 <= i + o < limit and (i + o) not in blocked]
            vol = M[:, vv].sum(axis=1) if vv else np.zeros(len(M))
            cost = M[:, bb].sum(axis=1) if bb else np.zeros(len(M))
            has = (vol != 0) | (cost != 0)
            ok = _plausible(vol, cost)
            net = int((has & ok).sum()) - int((has & ~ok).sum())
            if best is None or net > best[0]:
                best = (net, o, vol, cost)
        net, o, vol, cost = best
        if net > 0:
            prev = o
            if o != 0:
                shifted.append(nama)
        out.append([kode, nama, sat, vol, cost])
    # Komponen terakhir (Homecare / RTW): kolom biayanya bertabrakan dengan kolom total pada file yang
    # bergeser → biaya dihitung sebagai sisa (total - biaya komponen lain)
    if ref is not None and out and b in (1, 3):
        last = out[-1]
        others = sum(x[4] for x in out[:-1])
        need = (last[3] > 0) & (last[4] == 0)
        last[4] = np.where(need, np.maximum(ref - others, 0), last[4])
    # BMIV-03: biaya 'Operasi Khusus' bertabrakan dengan kolom subtotal operatif → hitung dari selisih subtotal
    if b == 2 and _norm("BIAYA DISETUJUI (RP) [3]") in pos:
        st = M[:, pos[_norm("BIAYA DISETUJUI (RP) [3]")]]
        names = [x[1] for x in out]
        if "Operasi Khusus" in names:
            k = names.index("Operasi Khusus")
            others = sum(x[4] for x in out if x[1] in ("Operasi Kecil", "Operasi Sedang", "Operasi Besar"))
            need = (out[k][3] > 0) & (out[k][4] == 0)
            out[k][4] = np.where(need, np.maximum(st - others, 0), out[k][4])
    return [tuple(x) for x in out], shifted


def read_total(df: pd.DataFrame, b: int) -> tuple[np.ndarray, np.ndarray, str | None]:
    """Total biaya disetujui per episode, dengan deteksi kolom total yang tertukar/bergeser."""
    cols = list(df.columns)
    pos = {c: i for i, c in enumerate(cols)}
    M = df.apply(pd.to_numeric, errors="coerce").fillna(0.0).to_numpy(dtype=float)
    m = SHEETS[b]
    if b == 3:  # BMIV-04: ambil kolom yang berisi nominal rupiah (bukan jumlah unit)
        c1, c2 = pos[_norm("JUMLAH KUNJUNGAN / UNIT")], pos[_norm(m["total"])]
        use = c1 if M[:, c1].sum() > M[:, c2].sum() else c2
        note = None if use == c2 else (
            "BMIV-04: kolom 'JUMLAH KUNJUNGAN / UNIT' dan 'JUMLAH JAMINAN (DISETUJUI)' tampak tertukar; "
            "kolom berisi nominal rupiah dipakai sebagai total disetujui.")
        return M[:, use], M[:, use], note
    d, a, sl = (pos[_norm(x)] for x in ("TOTAL BIAYA DIAJUKAN", m["total"], "TOTAL SELISIH"))
    best = None
    for o in (0, 1):
        if sl + o >= len(cols):
            continue
        ident = (M[:, d + o] > 0) & (np.abs(M[:, d + o] - M[:, a + o] - M[:, sl + o]) <= 1)
        score = int(ident.sum())
        if best is None or score > best[0]:
            best = (score, o)
    o = best[1]
    note = None if o == 0 else (
        f"{m['key']}: kolom TOTAL (diajukan/disetujui/selisih) tampak bergeser +1 kolom terhadap header; "
        "nilai disetujui dibaca dari posisi yang sesuai.")
    return M[:, a + o], M[:, d + o], note


# ─────────────────────────────────────────────────────────────────────────────
# 2b. PROFIL LAYOUT Data_Dummy.xlsx  (posisi kolom ISI data yang sebenarnya)
# ─────────────────────────────────────────────────────────────────────────────
# Pada Data_Dummy.xlsx isi data TIDAK selalu sejajar dengan nama header:
#   • BMIV-01 : sejajar penuh                                      (70 kolom)
#   • BMIV-02 : pasangan [volume, biaya] mulai dari kolom 50 (header 'BIAYA (RP) [13]') bergeser
#               +1 kolom; kolom 48-49 = subtotal penunjang (jumlah, biaya); Homecare tidak punya kolom
#               biaya sendiri (bertabrakan dgn TOTAL DIAJUKAN)                                (68 kolom)
#   • BMIV-03 : blok operatif (kecil→khusus) bergeser +1 kolom; setelah subtotal penunjang (kol 62-63)
#               semua pasangan bergeser +1; TOTAL diajukan/disetujui/selisih ada di kolom 79/80/81
#               (header-nya tertulis 78/79/80)                                                (84 kolom)
#   • BMIV-04 : 9 pasangan [volume, biaya] dimulai kolom 22; kolom 39 (header 'JUMLAH KUNJUNGAN / UNIT')
#               sebenarnya TOTAL RUPIAH disetujui, kolom 40 = jumlah unit                     (43 kolom)
# Indeks di bawah 0-based, urut seperti kolom pada sheet.
DUMMY_NCOLS = {0: 70, 1: 68, 2: 84, 3: 43}
DUMMY_TOTAL_POS = {0: (65, 66), 1: (63, 64), 2: (79, 80), 3: (39, 39)}  # (diajukan, disetujui)


def _dummy_layout(b: int, pos: dict) -> dict:
    """kode komponen → (kolom volume, kolom biaya, flag). flag: ''|'resid'|'out'|'out_cond'."""
    lay = {}
    if b == 0:  # sejajar → pakai nama header
        for kode, _n, _s, vcols, bcols in SPEC[0]:
            lay[kode] = ([pos[_norm(c)] for c in vcols], [pos[_norm(c)] for c in bcols], "")
    elif b == 1:
        for i, x in enumerate(SPEC[1][:12]):          # RJTL-01..12  : (24,25) … (46,47)
            lay[x[0]] = ([24 + 2 * i], [25 + 2 * i], "")
        for j, x in enumerate(SPEC[1][12:18]):        # RJTL-13..18  : (50,51) … (60,61)
            lay[x[0]] = ([50 + 2 * j], [51 + 2 * j], "")
        lay["RJTL-19"] = ([62], [], "resid")          # Homecare: biaya = sisa terhadap TOTAL DIAJUKAN
    elif b == 2:
        lay.update({"RI-01": ([25], [26], ""), "RI-02": ([29], [31], ""), "RI-02L": ([27], [], ""), "RI-02K": ([28], [], ""), "RI-03": ([32], [33], ""),
                    "RI-04": ([35], [36], "out"), "RI-05": ([37], [38], "out"), "RI-06": ([39], [40], "out"),
                    "RI-07": ([41], [42], "out_cond"), "RI-08": ([44], [45], "")})
        for i, k in enumerate(range(9, 17)):          # RI-09..16 : (46,47) … (60,61)
            lay[f"RI-{k:02d}"] = ([46 + 2 * i], [47 + 2 * i], "")
        lay.update({"RI-17": ([64], [65], ""), "RI-18": ([66], [67], ""), "RI-19": ([68], [69], ""),
                    "RI-20": ([71], [72], ""), "RI-21": ([73], [74], ""), "RI-22": ([75], [76], ""),
                    "RI-23": ([77], [78], "")})
    else:
        for i, x in enumerate(SPEC[3][:8]):           # AB-01..08 : (22,23) … (36,37)
            lay[x[0]] = ([22 + 2 * i], [23 + 2 * i], "")
        lay["AB-09"] = ([38], [], "")                 # RTW: kolom biaya = kolom total → tidak tersedia
    return lay


def read_dummy(df: pd.DataFrame, b: int):
    """
    Baca komponen & total memakai posisi kolom Data_Dummy. Return None bila struktur sheet tidak cocok
    (→ loader memakai deteksi otomatis). Return (total disetujui, total diajukan, komponen, kode di luar total).
    """
    M = df.apply(pd.to_numeric, errors="coerce").fillna(0.0).to_numpy(dtype=float)
    if M.shape[1] != DUMMY_NCOLS[b]:
        return None
    i_dj, i_ds = DUMMY_TOTAL_POS[b]
    dj, ds = M[:, i_dj], M[:, i_ds]
    if b < 3:   # uji identitas: diajukan − disetujui = selisih (kolom tepat di sebelah kanan disetujui)
        ok = dj > 0
        if ok.sum() == 0 or (np.abs(dj - ds - M[:, i_ds + 1])[ok] <= 1).mean() < 0.9:
            return None
    elif M[:, 39].sum() < M[:, 40].sum():   # kolom 39 harus yang berisi nominal rupiah
        return None
    pos = {c: i for i, c in enumerate(df.columns)}
    lay = _dummy_layout(b, pos)
    n, comps, outside = len(M), [], set()
    for kode, nama, sat, _v, _b in SPEC[b]:
        vi, bi_, flag = lay[kode]
        vol = M[:, vi].sum(axis=1) if vi else np.zeros(n)
        cost = M[:, bi_].sum(axis=1) if bi_ else np.zeros(n)
        if flag == "out_cond":
            cost = np.where(vol > 0, cost, 0.0)
        if flag.startswith("out"):
            outside.add(kode)
        comps.append([kode, nama, sat, vol, cost])
    if b == 1:   # Homecare: sisa = diajukan − komponen lain − subtotal penunjang (kol 49)
        others = sum(x[4] for x in comps[:-1])
        resid = np.maximum(dj - others - M[:, 49], 0.0)
        comps[-1][4] = np.where(comps[-1][3] > 0, resid, 0.0)
    return ds, dj, [tuple(x) for x in comps], outside


def _dummy_notes(b: int, comps: list, outside: set) -> list[str]:
    key, out = SHEETS[b]["key"], []
    if b >= 1:
        out.append(f"{key}: posisi kolom isi data disesuaikan dengan layout Data_Dummy (isi data bergeser dari "
                   "nama header; total dibaca dari kolom yang benar).")
    if b in (1, 2):
        out.append(f"{key}: TOTAL BIAYA pada file sudah memuat SUB TOTAL PENUNJANG sebagai tambahan (penunjang "
                   "terhitung ganda). Biaya komponen dialokasikan proporsional terhadap Total Disetujui, sehingga "
                   "dapat lebih besar dari angka mentah di kolom komponen.")
    if outside:
        ops = sum(float(x[4].sum()) for x in comps if x[0] in outside)
        out.append(f"{key}: biaya tindakan operatif (Rp {fN(ops)}) terisi di kolom komponen tetapi TIDAK termasuk "
                   "TOTAL BIAYA pada file → tidak dijumlahkan ke Total Biaya; hanya volumenya yang ditampilkan.")
    if b == 3:
        n = int((comps[-1][3] > 0).sum())
        out.append(f"{key}: kolom biaya RTW bertabrakan dengan kolom total, sehingga {fN(n)} episode RTW tercatat "
                   "tanpa nominal biaya. 'JUMLAH KUNJUNGAN / UNIT' berisi total Rp; 'JUMLAH JAMINAN (DISETUJUI)' "
                   "berisi jumlah unit.")
    return out


def _rujukan_flags(df: pd.DataFrame) -> np.ndarray:
    """Episode BMIV-01 yang dirujuk: ada isian pada kolom 'RUJUKAN | ...' (maks. 3 kolom sesudahnya)."""
    out = np.zeros(len(df), dtype=bool)
    cols = list(df.columns)
    start = next((i for i, c in enumerate(cols) if str(c).startswith("RUJUKAN")), None)
    if start is None:
        return out
    for c in cols[start:start + 3]:
        s = df[c]
        n = pd.to_numeric(s, errors="coerce")
        t = s.astype(str).str.strip().str.upper()
        txt = n.isna() & ~t.isin(["", "NAN", "NONE", "-", "0", "TIDAK", "T", "N", "NO"])
        out |= ((n.fillna(0) > 0) | txt).to_numpy()
    return out


def build_dataset(frames: dict) -> tuple[pd.DataFrame, pd.DataFrame, list[str]]:
    """
    Return:
      eps  — 1 baris = 1 episode pelayanan (1 baris sheet BMIV)
      comp — 1 baris = 1 komponen layanan pada 1 episode (biaya sudah dialokasikan
             proporsional terhadap TOTAL BIAYA DISETUJUI episode)
      warns — catatan kualitas data
    """
    ep_parts, comp_parts, warns = [], [], []
    uid = 0
    for b in range(4):
        df, m = frames[b], SHEETS[b]
        n = len(df)
        ids = np.arange(uid, uid + n)
        uid += n
        ep = pd.DataFrame({"ep": ids, "b": b})
        ep["tgl"] = pd.to_datetime(df[_norm(m["date"])], errors="coerce")
        ep["tgl_kec"] = pd.to_datetime(df["TGL KECELAKAAN"], errors="coerce")
        ep["kpj"] = df["NO KPJ"].astype(str).str.replace(r"\.0$", "", regex=True).str.strip()
        ep["peserta"] = df["NAMA TK / PASIEN"].astype(str)
        ep["kasus"] = df["NO. KASUS KK"].astype(str).str.strip()
        ep["kanwil"] = df["KANWIL"].astype(str).str.strip()
        ep["kode_cab"] = df["KODE CAB"].astype(str).str.strip()
        ep["cabang"] = ep["kode_cab"].map(CABANG).fillna(ep["kode_cab"])
        ep["plkk"] = df[_norm(m["plkk"])].astype(str).str.strip()
        ep["jenis_kasus"] = df["JENIS KASUS"].astype(str).str.strip()
        ep["perusahaan"] = df["NAMA PERUSAHAAN"].astype(str).str.strip()
        ep["sektor"] = ep["perusahaan"].map(sektor_of)
        ep["diagnosa"] = df["DIAGNOSA"].astype(str).str.strip()
        dm = read_dummy(df, b)   # layout Data_Dummy terdeteksi → posisi kolom eksplisit
        if dm is not None:
            tot_arr, diaj_arr, comps, outside = dm
            shifted = []
        else:                    # file lain → deteksi pergeseran kolom otomatis
            tot_arr, diaj_arr, tnote = read_total(df, b)
            if tnote:
                warns.append(tnote)
            comps, shifted = read_components(df, b, diaj_arr)
            outside = set()
        ep["total"] = tot_arr
        ep["diaj"] = diaj_arr if b < 3 else np.nan
        ep["icd"] = df["KODE ICD-10"].astype(str).str.strip()
        ep["visit"] = pd.to_numeric(df[_norm(m["visit"])], errors="coerce") if "visit" in m else np.nan
        ep["los"] = pd.to_numeric(df[_norm(m["los"])], errors="coerce").fillna(0) if "los" in m else np.nan
        ep["kelas"] = df[_norm(m["kelas"])].astype(str) if "kelas" in m else ""
        ep["poli"] = df[_norm(m["poli"])].astype(str) if "poli" in m else ""

        # ── komponen layanan ──
        if dm is not None:
            warns.extend(_dummy_notes(b, comps, outside))
        if shifted:
            warns.append(f"{m['key']}: komponen {', '.join(shifted)} terbaca dengan koreksi geser kolom "
                         "(isi data tidak sejajar dengan header template).")
        raw_sum = np.zeros(n)
        pieces = []
        for kode, nama, sat, vol, biaya in comps:
            if kode not in outside:
                raw_sum += biaya
            keep = (biaya > 0) | (vol > 0)
            if keep.any():
                pieces.append(pd.DataFrame({
                    "ep": ids[keep], "kode": kode, "komponen": nama, "satuan": sat, "b": b,
                    "volume": vol[keep], "biaya_raw": biaya[keep], "_i": np.flatnonzero(keep)}))
        if b == 0:   # rujukan (volume saja) — kolom 'RUJUKAN | ...' sampai akhir sheet
            rc = _rujukan_flags(df)
            if rc.any():
                pieces.append(pd.DataFrame({
                    "ep": ids[rc], "kode": "RJTP-RUJ", "komponen": "Rujukan", "satuan": "rujukan", "b": b,
                    "volume": 1.0, "biaya_raw": 0.0, "_i": np.flatnonzero(rc)}))
        if pieces:
            c = pd.concat(pieces, ignore_index=True)
            tot = ep["total"].to_numpy()
            scale = np.where(raw_sum > 0, tot / np.where(raw_sum > 0, raw_sum, 1), 0.0)
            c["biaya"] = c["biaya_raw"] * scale[c["_i"].to_numpy()]
            if outside:   # komponen di luar total file: volume tetap, biaya tidak dialokasikan
                c.loc[c["kode"].isin(outside), "biaya"] = 0.0
            comp_parts.append(c.drop(columns=["_i"]))
        ep_parts.append(ep)

    eps = pd.concat(ep_parts, ignore_index=True)
    comp = pd.concat(comp_parts, ignore_index=True) if comp_parts else pd.DataFrame(
        columns=["ep", "kode", "komponen", "satuan", "b", "volume", "biaya_raw", "biaya"])

    comp["info"] = comp["kode"].isin(INFO_CODES)
    bad_date = int(eps["tgl"].isna().sum())
    if bad_date:
        warns.append(f"{bad_date} baris dibuang karena tanggal pelayanan/masuk kosong atau tidak valid.")
        eps = eps[eps["tgl"].notna()]
        comp = comp[comp["ep"].isin(eps["ep"])]
    eps["periode"] = eps["tgl"].dt.strftime("%Y-%m")
    eps["tgl_s"] = eps["tgl"].dt.strftime("%Y-%m-%d")
    comp = comp.merge(eps[["ep", "periode", "kpj", "kasus", "kanwil", "cabang", "plkk",
                           "jenis_kasus", "sektor", "tgl"]], on="ep", how="left")
    return eps.reset_index(drop=True), comp.reset_index(drop=True), warns


CACHE_VER = 4   # naikkan bila struktur hasil build_dataset berubah (membuang cache lama)


@st.cache_data(show_spinner="Membaca & menormalisasi data…")
def load_data(content: bytes, ver: int = CACHE_VER):
    frames, w1 = read_template(io.BytesIO(content))
    eps, comp, w2 = build_dataset(frames)
    return eps, comp, w1 + w2


@st.cache_data(show_spinner=False)
def export_xlsx(content: bytes, ver: int = CACHE_VER) -> bytes:
    """File ternormalisasi dibuat sekali per file upload (bukan di setiap perubahan filter)."""
    eps, comp, _ = load_data(content, ver)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        eps.drop(columns=["tgl_s"]).to_excel(xw, sheet_name="Episode", index=False)
        comp.drop(columns=["tgl"]).to_excel(xw, sheet_name="Komponen", index=False)
    return buf.getvalue()


# ─────────────────────────────────────────────────────────────────────────────
# 3. FORMATTER (format Indonesia)
# ─────────────────────────────────────────────────────────────────────────────
def fN(n) -> str:
    return f"{round(float(n)):,}".replace(",", ".")


def fF(n) -> str:
    return "Rp " + fN(n)


def fR(n) -> str:
    n = float(n)
    a = abs(n)
    if a >= 1e9:
        return "Rp " + f"{n / 1e9:.2f}".replace(".", ",") + " M"
    if a >= 1e6:
        return "Rp " + f"{n / 1e6:.1f}".replace(".", ",") + " jt"
    return fF(n)


def f1(n) -> str:
    r = round(float(n), 1)
    if r == int(r):
        return fN(r)
    return f"{r:,.1f}".replace(",", "§").replace(".", ",").replace("§", ".")


def fU(n) -> str:
    """Utilisasi per 1.000 TK: 2 desimal bila kecil (kejadian JKK jarang)."""
    n = float(n)
    return f"{n:,.2f}".replace(",", "§").replace(".", ",").replace("§", ".") if abs(n) < 10 else f1(n)


def fP(n) -> str:
    return f1(float(n) * 100) + "%"


def ml(p: str) -> str:
    return f"{MON[int(p[5:7]) - 1]} {p[2:4]}"


# ─────────────────────────────────────────────────────────────────────────────
# 4. AGREGASI
# ─────────────────────────────────────────────────────────────────────────────
def tk_aktif() -> int:
    """Jumlah tenaga kerja aktif JKK (input sidebar) = penyebut per kapita pada template LB-ST."""
    try:
        return max(0, int(st.session_state.get("f_tk") or 0))
    except Exception:  # noqa: BLE001
        return 0


def den_of(pes) -> float:
    """Penyebut per kapita: TK aktif bila diisi; bila kosong, proxy = peserta unik yang melapor klaim."""
    tk = tk_aktif()
    return float(tk) if tk > 0 else float(pes)


def pm_caption(m: int) -> str:
    return (f"biaya / TK aktif / bulan ({m} bln)" if tk_aktif() > 0
            else f"proxy: biaya / peserta pelapor / bulan ({m} bln)")


def agg(L: pd.DataFrame, m: int) -> dict:
    """
    Unit cost   = biaya / volume                       (tidak tergantung populasi)
    PMPM        = biaya / (TK aktif x bulan)           (per kapita)
    Utilisasi   = volume / (TK aktif x bulan) x 1.000  (per 1.000 TK per bulan)
    Identitas   : PMPM = Unit Cost x Utilisasi / 1.000
    """
    k, c, n, t = L["kpj"].nunique(), L["kasus"].nunique(), len(L), float(L["total"].sum())
    d = den_of(k)
    return dict(pes=k, kas=c, kun=n, bi=t, uc=t / n if n else 0.0, pm=t / (d * m) if d else 0.0,
                cpc=t / c if c else 0.0, ut=n / (d * m) * 1000 if d else 0.0)


def cagg(C: pd.DataFrame) -> pd.DataFrame:
    if C.empty:
        return pd.DataFrame(columns=["kode", "komponen", "b", "vol", "bi"])
    return (C.groupby(["kode", "komponen", "b"], as_index=False)
            .agg(vol=("volume", "sum"), bi=("biaya", "sum"))
            .sort_values("bi", ascending=False).reset_index(drop=True))


# ─────────────────────────────────────────────────────────────────────────────
# 5. HELPER UI
# ─────────────────────────────────────────────────────────────────────────────
_KEY = [0]


def _nk() -> str:
    _KEY[0] += 1
    return f"plt{_KEY[0]}"


def plot(fig: go.Figure, height: int = 360):
    fig.update_layout(height=height, margin=dict(l=10, r=10, t=30, b=10), separators=",.",
                      legend=dict(orientation="h", yanchor="top", y=-0.15, x=0))
    try:
        st.plotly_chart(fig, width="stretch", key=_nk())
    except TypeError:  # Streamlit versi lama
        st.plotly_chart(fig, use_container_width=True, key=_nk())


def table(df: pd.DataFrame, fmts: dict | None = None, height: int | None = None, total: list | None = None):
    sty = df.style.format(fmts or {}, na_rep="–")
    kw = dict(hide_index=True)
    if height:
        kw["height"] = height
    try:
        st.dataframe(sty, width="stretch", **kw)
    except TypeError:
        st.dataframe(sty, use_container_width=True, **kw)
    if total is not None:
        t = pd.DataFrame([total], columns=df.columns)
        tf = {k: v for k, v in (fmts or {}).items()}
        try:
            st.dataframe(t.style.format(tf, na_rep=""), width="stretch", hide_index=True)
        except TypeError:
            st.dataframe(t.style.format(tf, na_rep=""), use_container_width=True, hide_index=True)


CSS = """
<style>
.kpi{border:1px solid rgba(128,128,128,.35);border-top:3px solid #0f766e;border-radius:12px;
     background:rgba(128,128,128,.06);height:116px;padding:8px 10px;box-sizing:border-box;
     display:flex;flex-direction:column;align-items:center;justify-content:center;text-align:center;
     overflow:hidden;margin-bottom:.6rem}
.kpi-t{font-size:.8rem;font-weight:600;opacity:.75;line-height:1.2;max-width:100%}
.kpi-v{font-size:1.5rem;font-weight:700;line-height:1.25;margin:4px 0 2px;max-width:100%;
       overflow-wrap:anywhere}
.kpi-v.sm{font-size:1.02rem;line-height:1.2}
.kpi-s{font-size:.72rem;opacity:.65;line-height:1.2;max-width:100%;white-space:nowrap;
       overflow:hidden;text-overflow:ellipsis}
</style>
"""


def inject_css():
    st.markdown(CSS, unsafe_allow_html=True)
    st.markdown(LB_CSS, unsafe_allow_html=True)


def kpis(items: list[tuple]):
    """Kotak KPI: tinggi tetap, lebar sama dalam satu baris, isi rata tengah (simetris)."""
    cols = st.columns(len(items))
    for col, it in zip(cols, items):
        title, val = html.escape(str(it[0])), str(it[1])
        sub = html.escape(str(it[2])) if len(it) > 2 and it[2] != "" else "&nbsp;"
        cls = "kpi-v sm" if len(val) > 16 else "kpi-v"
        col.markdown(f"<div class='kpi'><div class='kpi-t'>{title}</div>"
                     f"<div class='{cls}'>{html.escape(val)}</div><div class='kpi-s'>{sub}</div></div>",
                     unsafe_allow_html=True)


def insight_box(items: list[str], limit: int = 8):
    """Kotak insight otomatis (2 kolom). Teks boleh memakai **tebal** markdown."""
    items = [i for i in items if i][:limit]
    if not items:
        return
    with st.container(border=True):
        st.markdown("**💡 Insight**")
        if len(items) == 1:
            st.markdown(f"- {items[0]}")
            return
        half = (len(items) + 1) // 2
        c1, c2 = st.columns(2)
        c1.markdown("\n".join(f"- {i}" for i in items[:half]))
        c2.markdown("\n".join(f"- {i}" for i in items[half:]))


def money_axis(fig: go.Figure, vmax: float, axis: str = "x"):
    """Tick uang ringkas (jt / M) pada sumbu."""
    if vmax <= 0:
        return
    ticks = np.linspace(0, vmax, 6)
    txt = [fR(v).replace("Rp ", "") if v else "0" for v in ticks]
    (fig.update_xaxes if axis == "x" else fig.update_yaxes)(tickvals=ticks, ticktext=txt)


def hbar(labels, values, color, fmt=fF, title=None, height=None, axis_money=True):
    labels, values = list(labels), [float(v) for v in values]
    fig = go.Figure(go.Bar(x=values, y=labels, orientation="h", marker_color=color,
                           text=[fmt(v) for v in values], textposition="outside", cliponaxis=False,
                           hovertemplate="%{y}<br>%{text}<extra></extra>"))
    fig.update_yaxes(autorange="reversed", automargin=True)
    vmax = max(values) if values else 0
    # ruang kanan otomatis menyesuaikan panjang label angka agar tidak terpotong
    maxlen = max((len(fmt(v)) for v in values), default=0)
    pad = 1.28 + 0.03 * max(0, maxlen - 12)
    fig.update_xaxes(range=[0, vmax * pad if vmax else 1])
    if axis_money:
        money_axis(fig, vmax * pad)
    if title:
        fig.update_layout(title=title)
    plot(fig, height or max(260, 28 * len(labels) + 70))


def empty():
    st.info("Tidak ada data untuk kombinasi filter ini. Ubah atau reset filter.")


# ─────────────────────────────────────────────────────────────────────────────
# 5b. INSIGHT OTOMATIS (dihitung dari data terfilter)
# ─────────────────────────────────────────────────────────────────────────────
def _x(v: float) -> str:
    return f1(v) + "×"


def ins_exec(L, C, m) -> list[str]:
    out, tot = [], float(L["total"].sum())
    if tot <= 0:
        return out
    bb = L.groupby("b")["total"].sum()
    b = int(bb.idxmax())
    out.append(f"**{BN[b]}** menyerap biaya terbesar: {fR(bb[b])} (**{fP(bb[b] / tot)}** dari total biaya).")
    cs = L.groupby("kasus")["total"].sum().sort_values(ascending=False)
    n10 = max(1, int(round(len(cs) * 0.1)))
    out.append(f"**10% kasus teratas** ({fN(n10)} dari {fN(len(cs))} kasus) menyumbang **{fP(cs.head(n10).sum() / tot)}** "
               "biaya — kandidat utama untuk review kasus berbiaya tinggi.")
    if not C.empty:
        t = (C.assign(lbl=C["komponen"] + " (" + C["b"].map(dict(enumerate(BS))) + ")")
             .groupby("lbl")["biaya"].sum().sort_values(ascending=False))
        if t.sum() > 0:
            out.append(f"Komponen terbesar: **{t.index[0]}** — {fR(t.iloc[0])} ({fP(t.iloc[0] / t.sum())} dari biaya komponen).")
    dx = L.groupby("diagnosa")["total"].sum().sort_values(ascending=False)
    out.append(f"Diagnosa berbiaya tertinggi: **{dx.index[0]}** ({fP(dx.iloc[0] / tot)}); 3 diagnosa teratas "
               f"mencakup {fP(dx.head(3).sum() / tot)} biaya.")
    pm = L.groupby("periode")["total"].sum()
    if len(pm) >= 2:
        out.append(f"Bulan dengan biaya tertinggi: **{ml(pm.idxmax())}** ({fR(pm.max())}); rata-rata bulanan {fR(pm.mean())}.")
    nb = L.groupby("kasus")["b"].nunique()
    if len(nb):
        multi = float((nb > 1).mean())
        c_multi = L[L["kasus"].isin(nb[nb > 1].index)].groupby("kasus")["total"].sum().mean() if (nb > 1).any() else 0.0
        c_single = L[L["kasus"].isin(nb[nb == 1].index)].groupby("kasus")["total"].sum().mean() if (nb == 1).any() else 0.0
        if multi > 0 and c_single > 0:
            out.append(f"**{fP(multi)}** kasus melewati ≥2 jenis layanan; biaya per kasusnya {_x(c_multi / c_single)} "
                       "kasus yang hanya satu jenis layanan.")
    Lx = L[L["diaj"].notna()] if "diaj" in L.columns else L.iloc[0:0]
    dsum = float(Lx["diaj"].sum()) if len(Lx) else 0.0
    if dsum > 0:
        out.append(f"Rasio persetujuan (BMIV-01 s.d. 03) **{fP(Lx['total'].sum() / dsum)}**; {fR(dsum - Lx['total'].sum())} "
                   "dari biaya diajukan tidak disetujui.")
    jk = L.groupby("jenis_kasus").agg(bi=("total", "sum"), kas=("kasus", "nunique"))
    jk = jk[jk["kas"] > 0]
    if len(jk) >= 2:
        jk["cpc"] = jk["bi"] / jk["kas"]
        h, avg = jk["cpc"].idxmax(), L["total"].sum() / L["kasus"].nunique()
        out.append(f"Jenis kasus **{h}** punya biaya per kasus tertinggi ({fR(jk.loc[h, 'cpc'])}, "
                   f"{_x(jk.loc[h, 'cpc'] / avg)} rata-rata).")
    return out


def ins_trend(L, periods) -> list[str]:
    out = []
    mon = L.groupby("periode")["total"].sum().reindex(periods, fill_value=0.0)
    vis = L.groupby("periode").size().reindex(periods, fill_value=0)
    if len(mon) < 2 or mon.sum() <= 0:
        return out
    out.append(f"Puncak biaya di **{ml(mon.idxmax())}** ({fR(mon.max())}), terendah di **{ml(mon.idxmin())}** "
               f"({fR(mon.min())}); rata-rata {fR(mon.mean())} per bulan.")
    if mon.iloc[-2] > 0:
        ch = mon.iloc[-1] / mon.iloc[-2] - 1
        out.append(f"Biaya {ml(periods[-1])} **{'naik' if ch >= 0 else 'turun'} {f1(abs(ch) * 100)}%** dibanding {ml(periods[-2])}.")
    pc = mon.pct_change().replace([np.inf, -np.inf], np.nan).dropna()
    if len(pc) >= 2:
        out.append(f"Lonjakan MoM terbesar: **{ml(pc.idxmax())}** (+{f1(pc.max() * 100)}%); penurunan terdalam: "
                   f"**{ml(pc.idxmin())}** ({f1(pc.min() * 100)}%).")
    if len(mon) >= 4:
        h = len(mon) // 2
        a1, a2 = mon.iloc[:h].mean(), mon.iloc[h:].mean()
        if a1 > 0:
            out.append(f"Rata-rata bulanan paruh akhir periode **{'lebih tinggi' if a2 >= a1 else 'lebih rendah'} "
                       f"{f1(abs(a2 / a1 - 1) * 100)}%** dibanding paruh awal.")
    uc = (mon / vis.replace(0, np.nan)).dropna()
    if len(uc):
        out.append(f"Unit cost tertinggi pada **{ml(uc.idxmax())}** ({fF(uc.max())}/kunjungan) vs terendah "
                   f"{ml(uc.idxmin())} ({fF(uc.min())}).")
    bm = L.groupby(["periode", "b"])["total"].sum().unstack(fill_value=0.0)
    if not bm.empty:
        lead = bm.idxmax(axis=1).value_counts()
        out.append(f"**{BN[int(lead.index[0])]}** menjadi layanan berbiaya terbesar di {fN(lead.iloc[0])} dari {fN(len(bm))} bulan.")
    return out


def ins_bmiv(L, C, b, m, cg, a) -> list[str]:
    out, tot = [], a["bi"]
    if tot <= 0:
        return out
    s = float(cg["bi"].sum()) if not cg.empty else 0.0
    if s > 0:
        t = cg.iloc[0]
        out.append(f"**{t['komponen']}** komponen terbesar: {fR(t['bi'])} (**{fP(t['bi'] / s)}**); 3 komponen teratas "
                   f"= {fP(cg['bi'].head(3).sum() / s)} biaya {BS[b]}.")
    dx = L.groupby("diagnosa").agg(bi=("total", "sum"), kas=("kasus", "nunique")).sort_values("bi", ascending=False)
    if not dx.empty:
        out.append(f"Diagnosa **{dx.index[0]}** menyerap biaya terbesar ({fR(dx['bi'].iloc[0])}, "
                   f"{fP(dx['bi'].iloc[0] / tot)}) dari {fN(dx['kas'].iloc[0])} kasus.")
        r = dx[dx["kas"] >= 5].copy()
        if len(r) >= 2:
            r["cpc"] = r["bi"] / r["kas"]
            h = r["cpc"].idxmax()
            if h == dx.index[0]:
                out[-1] += f" Sekaligus biaya per kasus tertinggi ({fF(r.loc[h, 'cpc'])}, {_x(r.loc[h, 'cpc'] / a['cpc'])} rata-rata)."
            else:
                out.append(f"Biaya per kasus tertinggi: **{h}** ({fF(r.loc[h, 'cpc'])}), {_x(r.loc[h, 'cpc'] / a['cpc'])} "
                           f"rata-rata {BS[b]}.")
    # ── spesifik per kelompok layanan ──
    if b == 0:
        ks = L.groupby("kasus").size()
        out.append(f"Rata-rata **{f1(ks.mean())} kunjungan per kasus**; {fN((ks >= 3).sum())} kasus ({fP((ks >= 3).mean())}) "
                   "berkunjung ≥3 kali — pantau kunjungan berulang.")
        ob = cg[cg["kode"] == "RJTP-03"]
        if s > 0 and not ob.empty:
            out.append(f"Obat/resep = **{fP(ob['bi'].iloc[0] / s)}** biaya RJTP ({fN(ob['vol'].iloc[0])} resep).")
    elif b == 1:
        pl = L.groupby("poli").agg(kun=("ep", "size"), bi=("total", "sum"))
        if len(pl):
            p1, p2 = pl["kun"].idxmax(), pl["bi"].idxmax()
            out.append(f"Poli tersibuk **{p1}** ({fN(pl.loc[p1, 'kun'])} kunjungan, {fP(pl.loc[p1, 'kun'] / len(L))}); "
                       f"biaya terbesar ada di poli **{p2}** ({fR(pl.loc[p2, 'bi'])}).")
        hv = cg[cg["kode"].isin(["RJTL-08", "RJTL-09"])]
        if s > 0 and hv["vol"].sum() > 0:
            out.append(f"CT Scan & MRI hanya **{fP(hv['vol'].sum() / len(L))}** kunjungan namun menyerap "
                       f"**{fP(hv['bi'].sum() / s)}** biaya RJTL.")
        pen = cg[cg["kode"].isin([f"RJTL-{k:02d}" for k in range(5, 13)])]
        if s > 0:
            out.append(f"Penunjang (lab, radiologi, elektromedik, USG, dll) = **{fP(pen['bi'].sum() / s)}** biaya RJTL.")
    elif b == 2:
        los = L["los"]
        out.append(f"ALOS **{f1(los.mean())} hari**; {fP((los >= 8).mean())} episode dirawat ≥8 hari; biaya rata-rata "
                   f"{fF(tot / los.sum()) if los.sum() else '–'} per hari rawat.")
        kg = L.groupby("kelas").agg(n=("ep", "size"), bi=("total", "sum"))
        kg = kg[kg["n"] >= 3]
        if len(kg) >= 2:
            kg["cpc"] = kg["bi"] / kg["n"]
            h = kg["cpc"].idxmax()
            out.append(f"Kelas **{h}** punya biaya per kasus tertinggi ({fR(kg.loc[h, 'cpc'])}); terendah "
                       f"**{kg['cpc'].idxmin()}** ({fR(kg['cpc'].min())}).")
        icu = C[C["komponen"].str.startswith("ICU")]
        if len(icu):
            out.append(f"ICU/HCU dipakai pada **{fP(icu['ep'].nunique() / len(L))}** episode namun menyerap "
                       f"**{fP(icu['biaya'].sum() / tot)}** biaya RANAP.")
        op = C[C["kode"].isin(["RI-04", "RI-05", "RI-06", "RI-07"]) & (C["volume"] > 0)]
        if len(op):
            out.append(f"**{fP(op['ep'].nunique() / len(L))}** episode melibatkan tindakan operatif "
                       f"({fN(op['ep'].nunique())} episode); biaya operatif tidak termasuk Total Biaya pada file.")
    else:
        if s > 0 and len(cg) > 1:
            v = cg.sort_values("vol", ascending=False).iloc[0]
            out.append(f"Alat bantu/layanan terbanyak: **{v['komponen']}** ({fN(v['vol'])} unit).")
        u = cg[cg["vol"] >= 3].copy()
        if len(u) >= 2:
            u["ucc"] = u["bi"] / u["vol"]
            h = u.sort_values("ucc").iloc[-1]
            out.append(f"Unit cost tertinggi: **{h['komponen']}** ({fF(h['ucc'])}/unit), {_x(h['ucc'] / a['uc'])} "
                       "rata-rata per episode.")
        rtw = cg[cg["kode"] == "AB-09"]
        if len(rtw):
            out.append(f"Program RTW tercatat pada **{fN(rtw['vol'].iloc[0])}** episode "
                       f"({fP(rtw['vol'].iloc[0] / len(L))}), nominal biayanya tidak tersedia pada file.")
    pk = L.groupby("plkk")["total"].sum().sort_values(ascending=False)
    if len(pk):
        out.append(f"PLKK dengan biaya terbesar: **{pk.index[0]}** ({fP(pk.iloc[0] / tot)}); total {fN(len(pk))} PLKK aktif.")
    kw = L.groupby("kanwil")["total"].sum().sort_values(ascending=False)
    if len(kw) > 1:
        out.append(f"Kanwil **{kw.index[0]}** berkontribusi {fP(kw.iloc[0] / tot)} biaya {BS[b]}.")
    return out


def ins_plkk(g, mt) -> list[str]:
    out = []
    tot = float(g["bi"].sum())
    if g.empty or tot <= 0:
        return out
    top = g.sort_values("bi", ascending=False)
    out.append(f"**{top['plkk'].iloc[0]}** berbiaya terbesar: {fR(top['bi'].iloc[0])} (**{fP(top['bi'].iloc[0] / tot)}**); "
               f"3 PLKK teratas = {fP(top['bi'].head(3).sum() / tot)} dari {fN(len(g))} PLKK.")
    kv = g.sort_values("kun", ascending=False).iloc[0]
    out.append(f"Kunjungan terbanyak: **{kv['plkk']}** ({fN(kv['kun'])} kunjungan, {fP(kv['kun'] / g['kun'].sum())}).")
    r = g[g["kun"] >= 5]
    if len(r) >= 3:
        h = r.sort_values("uc").iloc[-1]
        out.append(f"Unit cost tertinggi: **{h['plkk']}** ({fF(h['uc'])}), {_x(h['uc'] / r['uc'].median())} median PLKK "
                   f"({fF(r['uc'].median())}).")
        c = r.sort_values("cpc").iloc[-1]
        out.append(f"Cost per case tertinggi: **{c['plkk']}** ({fF(c['cpc'])}) vs median {fF(r['cpc'].median())}.")
        u = r.sort_values("ut").iloc[-1]
        out.append(f"Utilisasi tertinggi: **{u['plkk']}** ({f1(u['ut'])} kunjungan / 1.000 peserta).")
    out.append(f"Rata-rata biaya per PLKK {fR(g['bi'].mean())}; median {fR(g['bi'].median())} — "
               f"{'sebaran condong ke beberapa PLKK besar' if g['bi'].mean() > 1.3 * g['bi'].median() else 'sebaran relatif merata'}.")
    return out


def ins_lbst(X, periods, val) -> list[str]:
    out = []
    if X.empty or float(X[val].sum()) <= 0:
        return out
    f = fR if val == "biaya" else fN
    tot = float(X[val].sum())
    cp = X.groupby(["b", "komponen"])[val].sum().sort_values(ascending=False)
    (bb, nm), v = cp.index[0], cp.iloc[0]
    out.append(f"**{nm}** ({BS[bb]}) komponen terbesar: {f(v)} (**{fP(v / tot)}**); 5 teratas = {fP(cp.head(5).sum() / tot)}.")
    lb = X.groupby("b")[val].sum()
    out.append(f"Layanan dominan: **{BN[int(lb.idxmax())]}** ({fP(lb.max() / tot)} dari total).")
    mon = X.groupby("periode")[val].sum().reindex(periods, fill_value=0.0)
    if len(mon) >= 2:
        out.append(f"Bulan puncak **{ml(mon.idxmax())}** ({f(mon.max())}); terendah {ml(mon.idxmin())} ({f(mon.min())}).")
    if len(periods) >= 6:
        pv = X.pivot_table(index=["b", "komponen"], columns="periode", values=val, aggfunc="sum", fill_value=0.0)
        pv = pv.reindex(columns=periods, fill_value=0.0)
        first, last = pv[periods[:3]].sum(axis=1), pv[periods[-3:]].sum(axis=1)
        d = (last - first).sort_values()
        if d.iloc[-1] > 0:
            (b1, n1) = d.index[-1]
            out.append(f"Kenaikan terbesar (3 bulan akhir vs 3 bulan awal): **{n1}** ({BS[b1]}), naik {f(d.iloc[-1])}.")
        if d.iloc[0] < 0:
            (b2, n2) = d.index[0]
            out.append(f"Penurunan terbesar: **{n2}** ({BS[b2]}), turun {f(abs(d.iloc[0]))}.")
    out.append(f"Tercatat **{fN(len(cp))}** komponen layanan aktif dari {fN(X['b'].nunique())} kelompok BMIV.")
    return out


def ins_case_all(eps) -> list[str]:
    out = []
    cs = eps.groupby("kasus").agg(bi=("total", "sum"), n=("ep", "size"), nb=("b", "nunique")).sort_values("bi", ascending=False)
    tot = float(cs["bi"].sum())
    if cs.empty or tot <= 0:
        return out
    out.append(f"12 kasus teratas menyumbang **{fP(cs['bi'].head(12).sum() / tot)}** biaya dari {fN(len(cs))} kasus.")
    out.append(f"Rata-rata biaya per kasus {fR(cs['bi'].mean())} vs median {fR(cs['bi'].median())} — "
               f"{'distribusi condong kanan: sedikit kasus sangat mahal' if cs['bi'].mean() > 1.3 * cs['bi'].median() else 'distribusi relatif merata'}.")
    mx = cs["n"].idxmax()
    out.append(f"Kasus dengan layanan terbanyak: **{mx}** ({fN(cs.loc[mx, 'n'])} episode, biaya {fR(cs.loc[mx, 'bi'])}).")
    ri = set(eps.loc[eps["b"] == 2, "kasus"])
    if ri:
        a_ri, a_no = cs.loc[cs.index.isin(ri), "bi"].mean(), cs.loc[~cs.index.isin(ri), "bi"].mean()
        if a_no > 0:
            out.append(f"Kasus yang berujung rawat inap ({fP(len(ri) / len(cs))}) berbiaya {_x(a_ri / a_no)} kasus tanpa rawat inap.")
    out.append(f"**{fP((cs['nb'] > 1).mean())}** kasus melibatkan ≥2 jenis layanan BMIV.")
    return out


def ins_case_one(E, cc, eps) -> list[str]:
    out = []
    t = float(E["total"].sum())
    avg = float(eps.groupby("kasus")["total"].sum().mean())
    if avg > 0:
        out.append(f"Biaya kasus ini **{_x(t / avg)}** rata-rata biaya per kasus ({fR(avg)}).")
    bb = E.groupby("b")["total"].sum()
    if t > 0:
        out.append(f"**{BN[int(bb.idxmax())]}** mendominasi biaya kasus ({fP(bb.max() / t)}).")
    seq = []
    for x in E["b"]:
        if not seq or seq[-1] != x:
            seq.append(x)
    out.append("Alur layanan: " + " → ".join(BS[int(x)] for x in seq) + f" ({len(E)} episode).")
    days = (E["tgl"].max() - E["tgl"].min()).days
    out.append(f"Melibatkan **{E['plkk'].nunique()} PLKK** dalam rentang **{days} hari**.")
    if not cc.empty and cc["biaya"].sum() > 0:
        k = cc.groupby("komponen")["biaya"].sum().sort_values(ascending=False)
        out.append(f"Komponen termahal: **{k.index[0]}** ({fR(k.iloc[0])}, {fP(k.iloc[0] / k.sum())} biaya komponen).")
    if (E["b"] == 2).any():
        r = E[E["b"] == 2]
        out.append(f"Rawat inap **{fN(r['los'].sum())} hari** dengan biaya {fR(r['total'].sum())}.")
    return out


# ─────────────────────────────────────────────────────────────────────────────
# 6. TAB-TAB DASHBOARD
# ─────────────────────────────────────────────────────────────────────────────
def tab_exec(L, C, m):
    a = agg(L, m)
    kpis([("Total Peserta", fN(a["pes"]), "NO KPJ"), ("Kasus", fN(a["kas"]), "No. Kasus KK"),
          ("Kunjungan", fN(a["kun"]), "episode layanan"), ("Total Biaya", fR(a["bi"]), fF(a["bi"])),
          ("Unit Cost", fF(a["uc"]), "biaya / kunjungan"),
          ("PMPM", fF(a["pm"]), pm_caption(m))])
    insight_box(ins_exec(L, C, m))
    c1, c2 = st.columns(2)
    with c1.container(border=True):
        st.markdown("**Komposisi Biaya per Kelompok Pelayanan (BMIV)**")
        st.caption("Porsi biaya disetujui")
        bb = [float(L.loc[L.b == b, "total"].sum()) for b in range(4)]
        fig = go.Figure(go.Pie(labels=BN, values=bb, hole=0.6, marker=dict(colors=BC), sort=False,
                               texttemplate="%{percent:.1%}", hovertemplate="%{label}<br>Rp %{value:,.0f}<extra></extra>"))
        plot(fig, 340)
    with c2.container(border=True):
        st.markdown("**Top 10 Komponen Layanan**")
        st.caption("Penyerapan biaya tertinggi")
        if not C.empty:
            t = C.assign(lbl=C["komponen"] + " · " + C["b"].map(dict(enumerate(BS)))) \
                 .groupby("lbl")["biaya"].sum().sort_values(ascending=False).head(10)
            hbar(t.index, t.values, ACC, height=340)
    st.caption("PMPM = biaya ÷ (TK aktif × bulan). Isi 'Jumlah TK aktif JKK' di sidebar; bila kosong dipakai proxy peserta "
               "unik yang melapor klaim. Biaya per komponen dialokasikan proporsional terhadap Total Biaya "
               "Disetujui tiap baris.")


def tab_trend(L, periods, m):
    sel = st.selectbox("Metrik", ["Biaya", "Volume (Kunjungan)", "Utilisasi (kunjungan per 1.000 peserta)"], key="tm")
    rows = []
    for p in periods:
        X = L[L.periode == p]
        a = agg(X, 1)
        rows.append(dict(p=p, X=X, pes=a["pes"], tot=a))
    labels = [ml(p) for p in periods]
    insight_box(ins_trend(L, periods))

    def val(o, b):
        X = o["X"] if b < 0 else o["X"][o["X"].b == b]
        if sel == "Biaya":
            return float(X["total"].sum())
        if sel.startswith("Volume"):
            return float(len(X))
        d = den_of(o["pes"])
        return len(X) / d * 1000 if d else 0.0

    with st.container(border=True):
        st.markdown("**Tren Bulanan**")
        st.caption("Seluruh bulan dalam rentang filter, urut kronologis")
        fig = go.Figure()
        if sel.startswith("Utilisasi"):
            fig.add_scatter(x=labels, y=[val(o, -1) for o in rows], name="Total", mode="lines+markers",
                            line=dict(color="#1c2430", width=3))
            for b in range(4):
                fig.add_scatter(x=labels, y=[val(o, b) for o in rows], name=BN[b], mode="lines+markers",
                                line=dict(color=BC[b]))
            fig.update_yaxes(title="kunjungan / 1.000 TK / bulan" if tk_aktif() else "kunjungan / 1.000 peserta pelapor")
        else:
            for b in range(4):
                fig.add_bar(x=labels, y=[val(o, b) for o in rows], name=BN[b], marker_color=BC[b])
            fig.update_layout(barmode="stack")
            if sel == "Biaya":
                vmax = max([sum(val(o, b) for b in range(4)) for o in rows] or [0])
                money_axis(fig, vmax * 1.05, "y")
        fig.update_xaxes(type="category")
        plot(fig, 380)
    df = pd.DataFrame(dict(
        Bulan=[ml(o["p"]) for o in rows], Peserta=[o["tot"]["pes"] for o in rows],
        Kasus=[o["tot"]["kas"] for o in rows], Kunjungan=[o["tot"]["kun"] for o in rows],
        Biaya=[o["tot"]["bi"] for o in rows], **{"Unit Cost": [o["tot"]["uc"] for o in rows]},
        **{"Util /1.000": [o["tot"]["ut"] for o in rows]}))
    bi = df["Biaya"].to_numpy()
    df["Δ Biaya MoM"] = [np.nan] + [(bi[j] / bi[j - 1] - 1) if bi[j - 1] else np.nan for j in range(1, len(bi))]
    with st.container(border=True):
        st.markdown("**Rekap Bulanan**")
        table(df, {"Peserta": fN, "Kasus": fN, "Kunjungan": fN, "Biaya": fF, "Unit Cost": fF, "Util /1.000": f1,
                   "Δ Biaya MoM": lambda v: ("+" if v >= 0 else "") + f1(v * 100) + "%"})


def _dx_table(L: pd.DataFrame) -> pd.DataFrame:
    if "icd" not in L.columns:
        L = L.assign(icd="")
    g = (L.groupby("diagnosa")
         .agg(icd=("icd", "first"), bi=("total", "sum"), kas=("kasus", "nunique"), kun=("ep", "size"))
         .reset_index())
    ok = ~g["icd"].isin(["", "nan", "None"])
    g["lbl"] = np.where(ok, g["diagnosa"] + " (" + g["icd"] + ")", g["diagnosa"])
    return g


def _wrap(t: str, w: int = 18, maxlines: int = 3) -> str:
    words, lines, cur = str(t).split(), [], ""
    for wd in words:
        if cur and len(cur) + 1 + len(wd) > w:
            lines.append(cur)
            cur = wd
        else:
            cur = f"{cur} {wd}".strip()
    lines.append(cur)
    if len(lines) > maxlines:
        lines = lines[:maxlines]
        lines[-1] = lines[-1][: max(0, w - 1)] + "…"
    return "<br>".join(lines)


def combo_dx(t: pd.DataFrame, color: str):
    """Batang jumlah kasus (angka di dasar batang) + garis biaya (label berlatar putih) per diagnosa."""
    t = t.reset_index(drop=True)
    labels = [_wrap(x) for x in t["lbl"]]
    kmax, bmax = float(t["kas"].max() or 1), float(t["bi"].max() or 1)
    fig = go.Figure()
    fig.add_bar(x=labels, y=t["kas"], name="Jumlah Kasus", marker_color=color, customdata=t["lbl"],
                text=[fN(v) for v in t["kas"]], textposition="inside", insidetextanchor="start",
                textfont=dict(color="#fff", size=13),
                hovertemplate="%{customdata}<br>Kasus: %{y:,.0f}<extra></extra>")
    fig.add_scatter(x=labels, y=t["bi"], name="Total Biaya (Rp)", mode="lines+markers", yaxis="y2",
                    line=dict(color="#ed7d31", width=3),
                    marker=dict(color="#ed7d31", size=9, line=dict(color="#fff", width=1.5)),
                    customdata=t["lbl"], hovertemplate="%{customdata}<br>Biaya: Rp %{y:,.0f}<extra></extra>")
    for x, v in zip(labels, t["bi"]):       # label biaya dengan latar putih agar tidak tumpang tindih
        fig.add_annotation(x=x, y=float(v), xref="x", yref="y2", text=f"<b>{fR(v)}</b>", showarrow=False,
                           yshift=20, bgcolor="rgba(255,255,255,0.95)", bordercolor="#ed7d31", borderwidth=1,
                           borderpad=3, font=dict(color="#7c2d12", size=12))
    # satu bidang grafik: batang & garis berbagi area yang sama (combo), garis berada di atas batang
    bt = np.linspace(0, bmax * 1.2, 6)
    fig.update_layout(
        yaxis=dict(title="Jumlah kasus", range=[0, kmax * 1.5], rangemode="tozero"),
        yaxis2=dict(title="Biaya", overlaying="y", side="right", showgrid=False,
                    range=[0, bmax * 1.2], tickvals=bt,
                    ticktext=[fR(v).replace("Rp ", "") if v else "0" for v in bt]))
    fig.update_xaxes(type="category", tickfont=dict(size=11))
    plot(fig, 480)


def tab_bmiv(L, C, b, m):
    L = L[L.b == b]
    if L.empty:
        return empty()
    C = C[C.b == b]
    a = agg(L, m)
    cg = cagg(C)
    items = [("Total Biaya", fR(a["bi"]), fF(a["bi"])), ("Kunjungan / Episode", fN(a["kun"]), "episode layanan"),
             ("Kasus", fN(a["kas"]), "No. Kasus KK"), ("Peserta", fN(a["pes"]), "NO KPJ"),
             ("Unit Cost", fF(a["uc"]), "biaya / kunjungan")]
    if b == 2:
        h = float(L["los"].sum())
        items.append(("ALOS", f1(h / len(L)) + " hari", f"total {fN(h)} hari rawat"))
    else:
        items.append(("PMPM", fF(a["pm"]), pm_caption(m)))
    kpis(items)
    insight_box(ins_bmiv(L, C, b, m, cg, a))
    def _komp(col):
        with col.container(border=True):
            st.markdown("**Biaya per Komponen**")
            top = cg.head(15)
            hbar(top["komponen"], top["bi"], BC[b])

    if b == 0:      # BMIV-01: grafik tren biaya bulanan dihapus → komponen tampil selebar layar
        _komp(st)
    else:
        c1, c2 = st.columns(2)
        _komp(c1)
        with c2.container(border=True):
            if b == 2:
                st.markdown("**Distribusi ALOS (hari rawat)**")
                bk = [("1–3", 1, 3), ("4–7", 4, 7), ("8–10", 8, 10), ("11–14", 11, 14), ("15+", 15, 999)]
                fig = go.Figure(go.Bar(x=[x[0] + " hr" for x in bk],
                                       y=[int(((L.los >= lo) & (L.los <= hi)).sum()) for _, lo, hi in bk],
                                       marker_color=BC[2], name="Kasus"))
                plot(fig, 340)
            elif b == 1:
                st.markdown("**Kunjungan per Poli/Spesialisasi**")
                t = L["poli"].value_counts().head(10)
                hbar(t.index, t.values, BC[1], fmt=fN, axis_money=False)
            else:
                st.markdown("**Tren Biaya Bulanan**")
                t = L.groupby("periode")["total"].sum().sort_index()
                fig = go.Figure(go.Bar(x=[ml(p) for p in t.index], y=t.values, marker_color=BC[b],
                                       hovertemplate="%{x}<br>Rp %{y:,.0f}<extra></extra>"))
                money_axis(fig, float(t.max()) * 1.05 if len(t) else 0, "y")
                fig.update_xaxes(type="category")
                plot(fig, 340)
    # ── Top Diagnosa (semua tab BMIV): batang = jumlah kasus, titik = biaya ──
    dx = _dx_table(L)
    with st.container(border=True):
        st.markdown("**Top 10 Diagnosa — Jumlah Kasus & Biaya**")
        combo_dx(dx.sort_values(["kas", "bi"], ascending=False).head(10), BC[b])
    with st.container(border=True):
        st.markdown("**Rincian Komponen**")
        df = pd.DataFrame(dict(Kode=cg["kode"], Komponen=cg["komponen"], Volume=cg["vol"], Biaya=cg["bi"],
                               **{"% Biaya": cg["bi"] / a["bi"] if a["bi"] else 0.0},
                               **{"Unit Cost": np.where(cg["vol"] > 0, cg["bi"] / cg["vol"].where(cg["vol"] > 0, 1), 0.0)}))
        table(df, {"Volume": fN, "Biaya": fF, "% Biaya": fP, "Unit Cost": fF},
              total=["Total", "", cg["vol"].sum(), a["bi"], 1.0, np.nan])
    if b == 2:
        with st.container(border=True):
            st.markdown("**ALOS per Kelas Rawat**")
            g = L.groupby("kelas").agg(Kasus=("ep", "size"), Hari=("los", "sum"), Biaya=("total", "sum")).reset_index()
            g["ALOS"] = g["Hari"] / g["Kasus"]
            g["Biaya / Kasus"] = g["Biaya"] / g["Kasus"]
            g = g.rename(columns={"kelas": "Kelas", "Hari": "Hari Rawat", "ALOS": "ALOS (hari)"})
            g = g[["Kelas", "Kasus", "Hari Rawat", "ALOS (hari)", "Biaya", "Biaya / Kasus"]].sort_values("Kelas")
            table(g, {"Kasus": fN, "Hari Rawat": fN, "ALOS (hari)": f1, "Biaya": fF, "Biaya / Kasus": fF})


def tab_plkk(L, m, comp=None):
    c1, c2 = st.columns(2)
    bs = c1.selectbox("Basis perbandingan", ["Semua layanan"] + BN, key="pb")
    mt = c2.selectbox("Metrik", ["Total Biaya", "Kunjungan (utilisasi)", "Unit Cost", "Cost per Case"], key="pm")
    X = L if bs == "Semua layanan" else L[L.b == BN.index(bs)]
    if X.empty:
        return empty()
    g = X.groupby("plkk").agg(pes=("kpj", "nunique"), kas=("kasus", "nunique"), kun=("ep", "size"),
                              bi=("total", "sum")).reset_index()
    g["uc"] = g["bi"] / g["kun"]
    g["cpc"] = g["bi"] / g["kas"]
    tk = tk_aktif()
    g["ut"] = g["kun"] / (tk * m) * 1000 if tk else g["kun"] / (g["pes"] * m) * 1000
    col, fmt = {"Total Biaya": ("bi", fF), "Kunjungan (utilisasi)": ("kun", fN), "Unit Cost": ("uc", fF),
                "Cost per Case": ("cpc", fF)}[mt]
    T = g[g["kun"] >= 5] if mt in ("Unit Cost", "Cost per Case") else g
    T = T.sort_values(col, ascending=False).head(15)
    insight_box(ins_plkk(g, mt))
    with st.container(border=True):
        st.markdown("**Peringkat PLKK (Top 15)**")
        st.caption("Bandingkan PLKK pada level layanan yang sama (apple-to-apple) dengan memilih basis BMIV. "
                   "Metrik rasio hanya untuk PLKK dengan ≥5 kunjungan.")
        hbar(T["plkk"], T[col], ACC, fmt=fmt, axis_money=col != "kun")
    with st.container(border=True):
        st.markdown(f"**Rekapitulasi PLKK** — {len(g)} PLKK")
        df = g.rename(columns=dict(plkk="PLKK", pes="Peserta", kas="Kasus", kun="Kunjungan", bi="Total Biaya",
                                   uc="Unit Cost", cpc="Cost per Case", ut="Util /1.000")) \
              .sort_values("Total Biaya", ascending=False)
        table(df, {"Peserta": fN, "Kasus": fN, "Kunjungan": fN, "Total Biaya": fF, "Unit Cost": fF,
                   "Cost per Case": fF, "Util /1.000": f1}, height=460)
    if comp is not None:
        plkk_overprice(L, comp, comp["plkk"].unique())


# ─────────────────────────────────────────────────────────────────────────────
# 6b. LB-ST & UNIT COST (mengikuti template LB_ST_DAN_UNIT_COST_PER_KAPITA.xlsx)
# ─────────────────────────────────────────────────────────────────────────────
def _R(kode, label, nm, sat, jenis, vol=(), bi=(), ket="", lvl=1, cnt=False):
    return dict(kode=kode, label=label, nm=nm, sat=sat, jenis=jenis, vol=list(vol), bi=list(bi), ket=ket,
                lvl=lvl, cnt=cnt)


def _LBL(label, lvl=1):
    return dict(kode="", label=label, nm="", sat="", jenis="label", vol=[], bi=[], ket="", lvl=lvl, cnt=False)


_INFO_KS = "Informasi. Tidak dijumlah ke biaya."
_RES = "Volume = RX. Biaya obat dibaca dari kolom BIAYA pada RX."

# vol/bi = kode komponen pada comp (SPEC); cnt=True → volume = jumlah episode (kasus) yang memakai komponen itu
LB_GROUPS = [
    dict(b=0, no="1", short="RJTP", title="RAWAT JALAN TK. I (PLKK TINGKAT I)", sub="Sub-Total Rawat Jalan Tk. I", rows=[
        _R("RJTP-01", "A. Dokter Umum", "Dokter Umum", "kunjungan", "biaya", ["RJTP-01"], ["RJTP-01"]),
        _R("RJTP-02", "B. Dokter Gigi (trauma gigi)", "Dokter Gigi (trauma gigi)", "tindakan", "biaya",
           ["RJTP-02"], ["RJTP-02"], "Volume = jumlah tindakan gigi, bukan kunjungan."),
        _R("RJTP-03", "C. Obat-obatan / Resep", "Obat-obatan / Resep", "resep", "biaya", ["RJTP-03"], ["RJTP-03"], _RES),
        _R("RJTP-04", "Obat - kasus, KS (informasi)", "", "kasus", "info", ["RJTP-03K"], [], _INFO_KS, 2),
        _R("RJTP-05", "Obat - lembar resep, LB (informasi)", "", "lembar", "info", ["RJTP-03L"], [], _INFO_KS, 2),
        _R("RJTP-06", "D. Penunjang Diagnostik Sederhana", "Penunjang Diagnostik Sederhana", "pemeriksaan", "biaya",
           ["RJTP-04"], ["RJTP-04"], "DR, UR, rontgen ekstremitas/thorax/gigi."),
        _R("RJTP-07", "E. Tindakan Medis Dokter Umum", "Tindakan Medis Dokter Umum", "tindakan", "biaya",
           ["RJTP-05"], ["RJTP-05"], "Termasuk bidai, reposisi sederhana, luka bakar."),
        _R("RJTP-08", "F. Tindakan Medis Dokter Gigi", "Tindakan Medis Dokter Gigi", "tindakan", "biaya",
           ["RJTP-06"], ["RJTP-06"]),
        _R("RJTP-09", "G. Rujukan", "", "rujukan", "volume", ["RJTP-RUJ"], [], "Volume saja, tanpa biaya."),
    ]),
    dict(b=1, no="2", short="RJTL", title="RAWAT JALAN TK. II (PLKK TINGKAT II)", sub="Sub-Total Rawat Jalan Tk. II", rows=[
        _R("RJTL-01", "A. Dokter Spesialis", "Dokter Spesialis", "kunjungan", "biaya", ["RJTL-01"], ["RJTL-01"]),
        _R("RJTL-02", "B. Obat-obatan / Resep", "Obat-obatan / Resep", "resep", "biaya",
           ["RJTL-04"], ["RJTL-02", "RJTL-03", "RJTL-04"], _RES),
        _R("RJTL-03", "Obat - kasus, KS (informasi)", "", "kasus", "info", ["RJTL-02"], [], _INFO_KS, 2),
        _R("RJTL-04", "Obat - lembar resep, LB (informasi)", "", "lembar", "info", ["RJTL-03"], [], _INFO_KS, 2),
        _LBL("C. Penunjang Diagnostik"),
        _R("RJTL-05", "- Laboratorium", "Penunjang Diagnostik - Laboratorium", "pemeriksaan", "biaya", ["RJTL-05"], ["RJTL-05"], "", 2),
        _R("RJTL-06", "- Radiologi / Rontgen", "Penunjang Diagnostik - Radiologi / Rontgen", "pemeriksaan", "biaya", ["RJTL-06"], ["RJTL-06"], "", 2),
        _R("RJTL-07", "- Elektromedik: EKG", "Penunjang Diagnostik - Elektromedik: EKG", "pemeriksaan", "biaya", ["RJTL-07"], ["RJTL-07"],
           "Elektromedik dipisah per modalitas sesuai BMIV-02.", 2),
        _R("RJTL-08", "- Elektromedik: CT Scan", "Penunjang Diagnostik - Elektromedik: CT Scan", "pemeriksaan", "biaya", ["RJTL-08"], ["RJTL-08"], "", 2),
        _R("RJTL-09", "- Elektromedik: MRI", "Penunjang Diagnostik - Elektromedik: MRI", "pemeriksaan", "biaya", ["RJTL-09"], ["RJTL-09"], "", 2),
        _R("RJTL-10", "- Elektromedik: USG", "Penunjang Diagnostik - Elektromedik: USG", "pemeriksaan", "biaya", ["RJTL-10"], ["RJTL-10"], "", 2),
        _R("RJTL-11", "- Elektromedik: EMG / NCV", "Penunjang Diagnostik - Elektromedik: EMG / NCV", "pemeriksaan", "biaya", ["RJTL-11"], ["RJTL-11"], "", 2),
        _R("RJTL-12", "- Elektromedik: Lain-lain", "Penunjang Diagnostik - Elektromedik: Lain-lain", "pemeriksaan", "biaya", ["RJTL-12"], ["RJTL-12"], "", 2),
        _R("RJTL-13", "D. Tindakan Medis Dokter Spesialis", "Tindakan Medis Dokter Spesialis", "tindakan", "biaya",
           ["RJTL-13"], ["RJTL-13"], "Tindakan rawat jalan, di luar ODC."),
        _R("RJTL-14", "- One Day Care (ODC)", "Tindakan One Day Care (ODC)", "tindakan", "biaya", ["RJTL-14"], ["RJTL-14"], "", 2),
        _R("RJTL-15", "E. Fisioterapi", "Fisioterapi", "kunjungan", "biaya", ["RJTL-15"], ["RJTL-15"]),
        _R("RJTL-16", "F. Rehabilitasi Medik Lain", "Rehabilitasi Medik Lain", "kunjungan", "biaya", ["RJTL-16"], ["RJTL-16"],
           "Terapi okupasi, psikologi, dll."),
        _R("RJTL-17", "G. Emergensi", "Emergensi", "kunjungan", "biaya", ["RJTL-17"], ["RJTL-17"]),
        _R("RJTL-18", "H. Ambulans", "Ambulans", "perjalanan", "biaya", ["RJTL-18"], ["RJTL-18"]),
        _R("RJTL-19", "I. Homecare", "Homecare", "kunjungan", "biaya", ["RJTL-19"], ["RJTL-19"]),
    ]),
    dict(b=2, no="3", short="RANAP", title="RAWAT INAP", sub="Sub-Total Rawat Inap", rows=[
        _R("RANAP-01", "A. Akomodasi Rawat Inap", "Akomodasi Rawat Inap", "kasus", "biaya", ["RI-01"], ["RI-01"], "Volume = kasus.", 1, True),
        _R("RANAP-02", "Hari rawat inap (informasi)", "", "hari", "info", ["RI-01"], [], "Untuk ALOS dan biaya per hari.", 2),
        _R("RANAP-03", "B. Obat-obatan / Resep", "Obat-obatan / Resep", "resep", "biaya", ["RI-02"], ["RI-02"], "Volume = RX."),
        _R("RANAP-04", "Obat - kasus, KS (informasi)", "", "kasus", "info", ["RI-02K"], [], _INFO_KS, 2),
        _R("RANAP-05", "Obat - lembar resep, LB (informasi)", "", "lembar", "info", ["RI-02L"], [], _INFO_KS, 2),
        _R("RANAP-06", "C. Tindakan Non Operatif", "Tindakan Non Operatif", "kunjungan", "biaya", ["RI-03"], ["RI-03"], "Tindakan non bedah."),
        _R("RANAP-07", "D. Tindakan Operatif", "Tindakan Operatif", "tindakan", "biaya", ["RI-04", "RI-05", "RI-06", "RI-07"],
           ["RI-04", "RI-05", "RI-06", "RI-07"], "Kecil + sedang + besar + khusus. Biaya = BIAYA DISETUJUI operatif."),
        _R("RANAP-08", "Operasi kecil (informasi)", "", "tindakan", "info", ["RI-04"], [], "Volume per kelas operasi.", 2),
        _R("RANAP-09", "Operasi sedang (informasi)", "", "tindakan", "info", ["RI-05"], [], _INFO_KS, 2),
        _R("RANAP-10", "Operasi besar (informasi)", "", "tindakan", "info", ["RI-06"], [], _INFO_KS, 2),
        _R("RANAP-11", "Operasi khusus (informasi)", "", "tindakan", "info", ["RI-07"], [], _INFO_KS, 2),
        _R("RANAP-12", "E. Perawatan Khusus (ICU/ICCU/HCU/Burn Unit)", "Perawatan Khusus (ICU/ICCU/HCU/Burn Unit)", "kasus",
           "biaya", ["RI-08"], ["RI-08"], "", 1, True),
        _R("RANAP-13", "Hari perawatan khusus (informasi)", "", "hari", "info", ["RI-08"], [], _INFO_KS, 2),
        _LBL("F. Penunjang Diagnostik"),
        _R("RANAP-14", "- Laboratorium", "Penunjang Diagnostik - Laboratorium", "pemeriksaan", "biaya", ["RI-09"], ["RI-09"], "", 2),
        _R("RANAP-15", "- Radiologi / Rontgen", "Penunjang Diagnostik - Radiologi / Rontgen", "pemeriksaan", "biaya", ["RI-10"], ["RI-10"], "", 2),
        _R("RANAP-16", "- Elektromedik: EKG", "Penunjang Diagnostik - Elektromedik: EKG", "pemeriksaan", "biaya", ["RI-11"], ["RI-11"],
           "Elektromedik dipisah per modalitas sesuai BMIV-03.", 2),
        _R("RANAP-17", "- Elektromedik: CT Scan", "Penunjang Diagnostik - Elektromedik: CT Scan", "pemeriksaan", "biaya", ["RI-12"], ["RI-12"], "", 2),
        _R("RANAP-18", "- Elektromedik: MRI", "Penunjang Diagnostik - Elektromedik: MRI", "pemeriksaan", "biaya", ["RI-13"], ["RI-13"], "", 2),
        _R("RANAP-19", "- Elektromedik: USG", "Penunjang Diagnostik - Elektromedik: USG", "pemeriksaan", "biaya", ["RI-14"], ["RI-14"], "", 2),
        _R("RANAP-20", "- Elektromedik: EMG / NCV", "Penunjang Diagnostik - Elektromedik: EMG / NCV", "pemeriksaan", "biaya", ["RI-15"], ["RI-15"], "", 2),
        _R("RANAP-21", "- Elektromedik: Lain-lain", "Penunjang Diagnostik - Elektromedik: Lain-lain", "pemeriksaan", "biaya", ["RI-16"], ["RI-16"], "", 2),
        _LBL("G. Lain-lain"),
        _R("RANAP-22", "- Pelayanan Darah", "Pelayanan Darah", "kasus", "biaya", ["RI-17"], ["RI-17"], "", 2),
        _R("RANAP-23", "- Transfusi Darah", "Transfusi Darah", "kasus", "biaya", ["RI-18"], ["RI-18"], "", 2),
        _R("RANAP-24", "- Implan Ortopedi (Pin, Plate, Screw, dll)", "Implan Ortopedi (Pin, Plate, Screw, dll)", "kasus", "biaya",
           ["RI-19"], ["RI-19"], "Biaya implan/alat saja. Biaya operasi ada di Tindakan Operatif.", 2),
        _LBL("H. Rehabilitasi Medik"),
        _R("RANAP-25", "- Fisioterapi", "Rehabilitasi Medik - Fisioterapi", "sesi", "biaya", ["RI-20"], ["RI-20"], "", 2),
        _R("RANAP-26", "- Rehabilitasi Medik Lain", "Rehabilitasi Medik Lain", "sesi", "biaya", ["RI-21"], ["RI-21"], "", 2),
        _R("RANAP-27", "I. Ambulans", "Ambulans", "perjalanan", "biaya", ["RI-22"], ["RI-22"]),
        _R("RANAP-28", "J. Homecare", "Homecare", "kunjungan", "biaya", ["RI-23"], ["RI-23"]),
    ]),
    dict(b=3, no="4", short="KHUSUS", title="PELAYANAN KHUSUS, REHABILITASI & PENDUKUNG",
         sub="Sub-Total Pelayanan Khusus, Rehabilitasi & Pendukung", rows=[
        _R("KHUSUS-01", "A. Kaca Mata", "Kaca Mata", "unit", "biaya", ["AB-01"], ["AB-01"], "Plafon per unit lebih relevan daripada rata-rata."),
        _R("KHUSUS-02", "B. Gigi Palsu", "Gigi Palsu", "unit", "biaya", ["AB-02"], ["AB-02"]),
        _R("KHUSUS-03", "C. Protesis Anggota Gerak", "Protesis Anggota Gerak", "unit", "biaya", ["AB-03"], ["AB-03"], "Kaki / tangan palsu."),
        _R("KHUSUS-04", "D. Ortosis", "Ortosis", "unit", "biaya", ["AB-04"], ["AB-04"], "Brace, korset, spalk fungsional."),
        _R("KHUSUS-05", "E. Alat Bantu Jalan", "Alat Bantu Jalan", "unit", "biaya", ["AB-05"], ["AB-05"], "Kursi roda, tongkat, walker."),
        _R("KHUSUS-06", "F. Alat Bantu Dengar", "Alat Bantu Dengar", "unit", "biaya", ["AB-06"], ["AB-06"]),
        _R("KHUSUS-07", "G. Mata Palsu", "Mata Palsu", "unit", "biaya", ["AB-07"], ["AB-07"]),
        _R("KHUSUS-08", "H. Ambulans / Transportasi", "Ambulans / Transportasi", "perjalanan", "biaya", ["AB-08"], ["AB-08"]),
        _R("KHUSUS-09", "I. Program Kembali Bekerja (RTW)", "Program Kembali Bekerja (RTW)", "kasus", "biaya", ["AB-09"], ["AB-09"]),
    ]),
]
LB_BY_B = {g["b"]: g for g in LB_GROUPS}
BASIS_OPTS = ["Alokasi (selaras dashboard)"]   # basis tunggal: biaya komponen dialokasikan ke Total Disetujui


def lb_vals(C: pd.DataFrame, basis: str, by: str | None = None) -> dict:
    """kode baris LB-ST → (Series volume, Series biaya) per kunci grup (by) atau satu kunci '' bila by=None."""
    if C is None or C.empty:
        return {}
    col = "biaya" if basis.startswith("Alokasi") else "biaya_raw"
    d = C.assign(_k=C[by] if by else "", _bi=C[col])
    g = d.groupby(["_k", "kode"]).agg(vol=("volume", "sum"), bi=("_bi", "sum"), n=("ep", "size"))
    V, B, N = (g[x].unstack(fill_value=0.0) for x in ("vol", "bi", "n"))

    def S(T, codes):
        cs = [c for c in codes if c in T.columns]
        return T[cs].sum(axis=1) if cs else pd.Series(0.0, index=T.index)

    out = {}
    for grp in LB_GROUPS:
        for r in grp["rows"]:
            if r["jenis"] == "label":
                continue
            out[r["kode"]] = (S(N if r["cnt"] else V, r["vol"]), S(B, r["bi"]))
    return out


def _vb(vals: dict, kode: str) -> tuple[float, float]:
    v = vals.get(kode)
    if v is None:
        return 0.0, 0.0
    return float(v[0].get("", 0.0)), float(v[1].get("", 0.0))


def _c(v) -> str:
    return "" if v is None or (isinstance(v, float) and np.isnan(v)) else fN(v)


def periode_label(sel: list[str]) -> str:
    if not sel:
        return "-"
    a, b = sel[0], sel[-1]
    ta, tb = f"{MON_FULL[int(a[5:7]) - 1]} {a[:4]}", f"{MON_FULL[int(b[5:7]) - 1]} {b[:4]}"
    return ta if a == b else f"{ta} s.d. {tb}"


def cabang_label(L: pd.DataFrame) -> tuple[str, str]:
    cb, kw = L["cabang"].dropna().unique(), L["kanwil"].dropna().unique()
    cab = f"{L['kode_cab'].iloc[0]} - {cb[0]}" if len(cb) == 1 and len(L) else f"{len(cb)} cabang (hasil filter)"
    kan = kw[0] if len(kw) == 1 else f"{len(kw)} kanwil (hasil filter)"
    return cab, kan


def kasus_baru(L: pd.DataFrame, sel: list[str]) -> int:
    kb = L.groupby("kasus")["tgl_kec"].min()
    if kb.isna().all():
        return int(len(kb))
    return int(kb.dropna().dt.strftime("%Y-%m").isin(sel).sum())


def lb_rows(C: pd.DataFrame, L: pd.DataFrame, basis: str, sel_b: list[int], hide_empty: bool = False):
    """Susun baris Bagian A (rincian) & Bagian B (kontrol) LB-ST persis mengikuti template."""
    vals = lb_vals(C, basis)
    A, tot_lb = [], 0.0
    for grp in LB_GROUPS:
        b = grp["b"]
        if b not in sel_b:
            continue
        A.append(dict(cls="grp", cells=[grp["no"], grp["title"]], colspan=True))
        sub, buf = 0.0, []
        for r in grp["rows"]:
            if r["jenis"] == "label":
                buf = [dict(cls="lab", lvl=r["lvl"], cells=["", r["label"], "", "", "", "", "", "", ""])]
                A.extend(buf)
                continue
            vol, bi = _vb(vals, r["kode"])
            empty_row = vol == 0 and bi == 0
            if hide_empty and empty_row:
                continue
            if r["jenis"] == "biaya":
                sub += bi
                uc = bi / vol if vol > 0 and bi > 0 else None
                cells = ["", r["label"], r["sat"], fN(vol), "", fN(bi), "", _c(uc), r["ket"]]
                cls = "row"
            else:
                cells = ["", r["label"], r["sat"], fN(vol), "", "", "", "", r["ket"]]
                cls = "inf"
            A.append(dict(cls=cls, lvl=r["lvl"], cells=cells))
        Lb = L[L["b"] == b]
        tb = float(Lb["total"].sum())
        if b < 3:
            dj = float(Lb["diaj"].sum())
            e, g_, k = fN(dj), fN(dj - tb), "Diajukan dan selisih dari baris JUMLAH BMIV (total per BMIV)."
        else:
            e, g_, k = "", "", "BMIV-04 tidak memuat kolom biaya diajukan."
        A.append(dict(cls="sub", cells=["", grp["sub"], "campuran", "", e, fN(sub), g_, "", k]))
        tot_lb += sub
    A.append(dict(cls="tot", cells=["", "TOTAL PELAYANAN KESEHATAN JKK", "", "", "", fN(tot_lb), "", "", ""]))

    # Bagian B — kontrol terhadap baris JUMLAH BMIV
    Bt, sE, sF = [], 0.0, 0.0
    nm = {0: "BMIV-01 Rawat Jalan Tk. I", 1: "BMIV-02 Rawat Jalan Tk. II", 2: "BMIV-03 Rawat Inap", 3: "BMIV-04 Pelayanan Khusus"}
    tot_bmiv, n = 0.0, 0
    for b in sel_b:
        Lb = L[L["b"] == b]
        tb = float(Lb["total"].sum())
        tot_bmiv += tb
        n += 1
        if b < 3:
            dj = float(Lb["diaj"].sum())
            sE, sF = sE + dj, sF + tb
            Bt.append(dict(cls="row", cells=[str(n), nm[b], "", "", fN(dj), fN(tb), fN(dj - tb),
                                              fP((dj - tb) / dj) if dj else "", ""]))
        else:
            Bt.append(dict(cls="row", cells=[str(n), nm[b], "", "", "", fN(tb), "", "",
                                              "Tidak ada kolom diajukan."]))
    if sE > 0:
        lbl = "Total BMIV-01 s.d. BMIV-03"
        Bt.append(dict(cls="sub", cells=["", lbl, "", "", fN(sE), fN(sF), fN(sE - sF), fP((sE - sF) / sE), ""]))
    diff = tot_lb - tot_bmiv
    ok = abs(diff) < 1
    Bt.append(dict(cls="tot", cells=["", "Kontrol: Total LB-ST dikurangi total disetujui BMIV (harus 0)", "", "", "",
                                      fN(diff), "", "",
                                      "OK — seluruh kolom biaya terbaca" if ok else
                                      "PERIKSA: ada kolom biaya BMIV yang tidak terbaca LB-ST / tidak sama dengan JUMLAH"]))
    return A, Bt, tot_lb, diff


LB_HEAD_A = ["No", "Jenis Pelayanan", "Satuan", "Volume\n(Kunjungan/Kasus)", "Biaya Diajukan\n(Rp)", "Biaya Disetujui\n(Rp)",
             "Selisih\n(Rp)", "Unit Cost\n(Rp)", "Keterangan"]
LB_HEAD_B = ["No", "Sumber", "", "", "Biaya Diajukan\n(Rp)", "Biaya Disetujui\n(Rp)", "Selisih\n(Rp)",
             "Tingkat Koreksi\n(Selisih / Diajukan)", "Keterangan"]
LB_CSS = """
<style>
.lbw{overflow:auto;max-height:640px;border:1px solid rgba(128,128,128,.35);border-radius:8px;margin-bottom:.6rem}
.lb{width:100%;border-collapse:collapse;font-size:.8rem}
.lb th{background:#0f766e;color:#fff;padding:6px 8px;text-align:center;border:1px solid rgba(128,128,128,.35);
       position:sticky;top:0;z-index:1;white-space:pre-line}
.lb td{padding:4px 8px;border:1px solid rgba(128,128,128,.22);vertical-align:top}
.lb td.n{text-align:right;white-space:nowrap}
.lb td.k{font-size:.72rem;opacity:.7;min-width:160px}
.lb tr.grp td{background:rgba(15,118,110,.20);font-weight:700}
.lb tr.lab td{font-weight:600;opacity:.9}
.lb tr.inf td{opacity:.62;font-style:italic}
.lb tr.sub td{background:rgba(128,128,128,.16);font-weight:700}
.lb tr.tot td{background:rgba(15,118,110,.30);font-weight:800}
.lbh{display:grid;grid-template-columns:230px 1fr;gap:2px 14px;font-size:.88rem;margin:.2rem 0 .8rem}
.lbh b{opacity:.75;font-weight:600}
</style>
"""


def html_table(heads: list[str], rows: list[dict], num_from: int = 3) -> str:
    th = "".join(f"<th>{html.escape(h)}</th>" for h in heads)
    body = []
    for r in rows:
        cls, cells, lvl = r.get("cls", "row"), r["cells"], r.get("lvl", 0)
        tds = []
        if r.get("colspan"):
            tds.append(f"<td>{html.escape(str(cells[0]))}</td><td colspan='{len(heads) - 1}'>{html.escape(str(cells[1]))}</td>")
        else:
            for i, v in enumerate(cells):
                v = html.escape(str(v))
                if i == 1:
                    pad = f" style='padding-left:{8 + max(lvl - 1, 0) * 18}px'" if lvl > 1 else ""
                    tds.append(f"<td{pad}>{v}</td>")
                elif num_from <= i < len(cells) - 1:
                    tds.append(f"<td class='n'>{v}</td>")
                elif i == len(cells) - 1 and len(cells) == len(heads) and heads[-1] == "Keterangan":
                    tds.append(f"<td class='k'>{v}</td>")
                else:
                    tds.append(f"<td>{v}</td>")
        body.append(f"<tr class='{cls}'>{''.join(tds)}</tr>")
    return f"<div class='lbw'><table class='lb'><thead><tr>{th}</tr></thead><tbody>{''.join(body)}</tbody></table></div>"


def rows_to_df(heads: list[str], rows: list[dict]) -> pd.DataFrame:
    out = []
    for r in rows:
        c = list(r["cells"]) + [""] * (len(heads) - len(r["cells"]))
        out.append(c[:len(heads)])
    return pd.DataFrame(out, columns=[h.replace("\n", " ") for h in heads])


def to_xlsx(sheets: dict[str, pd.DataFrame]) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        for name, df in sheets.items():
            df.to_excel(xw, sheet_name=name[:31], index=False)
    return buf.getvalue()


def month_pick(periods: list[str], key: str) -> list[str]:
    """Filter bulan (multi-select). Key memuat rentang periode agar state tidak basi saat filter tanggal berubah."""
    sel = st.multiselect("Filter bulan", periods, default=periods, format_func=ml,
                         key=f"{key}_{periods[0]}_{periods[-1]}_{len(periods)}" if periods else key,
                         placeholder="Pilih bulan")
    return [p for p in periods if p in sel]


def basis_now() -> str:
    return BASIS_OPTS[0]


def tab_lbst(L, C, periods):
    """LB-ST — Rekapitulasi Kunjungan dan Pembayaran Jaminan JKK per PLKK (format template)."""
    c1, c2 = st.columns([3, 2])
    bsel = c1.radio("Kelompok BMIV", ["Semua BMIV"] + BN, horizontal=True, key="lb_b")
    with c2:
        sel = month_pick(periods, "lb_m")
    if not sel:
        return st.info("Pilih minimal satu bulan.")
    sel_b = [0, 1, 2, 3] if bsel == "Semua BMIV" else [BN.index(bsel)]
    X = L[L["periode"].isin(sel) & L["b"].isin(sel_b)]
    CX = C[C["periode"].isin(sel) & C["b"].isin(sel_b)]
    if X.empty:
        return empty()
    basis = basis_now()
    hide = st.checkbox("Sembunyikan baris tanpa data", value=False, key="lb_h")
    cab, kan = cabang_label(X)
    tk, m = tk_aktif(), len(sel)
    hdr = [("Kantor Cabang", cab), ("Kantor Wilayah / Deputi", kan), ("Periode Laporan", periode_label(sel)),
           ("Jumlah Tenaga Kerja Aktif JKK (jiwa)", fN(tk) if tk else "(belum diisi — isi di sidebar)"),
           ("Jumlah Bulan dalam Periode Data", str(m)), ("Kasus Baru JKK Dilaporkan", fN(kasus_baru(X, sel)))]
    with st.container(border=True):
        st.markdown("**BPJS KETENAGAKERJAAN**  \n**LB-ST — REKAPITULASI KUNJUNGAN DAN PEMBAYARAN JAMINAN JKK - PLKK**")
        st.markdown("<div class='lbh'>" + "".join(f"<b>{html.escape(a)}</b><span>{html.escape(b)}</span>" for a, b in hdr)
                    + "</div>", unsafe_allow_html=True)
    A, B, tot_lb, diff = lb_rows(CX, X, basis, sel_b, hide)
    insight_box(ins_lbst(CX[~CX["info"]], sel, "biaya"))
    with st.container(border=True):
        st.markdown("**A. Rincian Pelayanan Kesehatan JKK — Volume dan Biaya Disetujui**")
        st.markdown(html_table(LB_HEAD_A, A), unsafe_allow_html=True)
        st.caption("Volume obat = jumlah RX (KS & LB hanya informasi). Volume rawat inap = kasus (hari rawat = informasi). "
                   "Unit cost = Biaya Disetujui ÷ Volume. Biaya Diajukan & Selisih di sub-total bersumber dari JUMLAH BMIV.")
    with st.container(border=True):
        st.markdown("**B. Kontrol Verifikasi dan Kesesuaian Total (dari baris JUMLAH BMIV)**")
        st.markdown(html_table(LB_HEAD_B, B), unsafe_allow_html=True)
        if abs(diff) >= 1:
            op = CX[CX["kode"].isin(["RI-04", "RI-05", "RI-06", "RI-07"])]["biaya_raw"].sum()
            msg = (f"Selisih {fF(diff)}. ")
            if op > 0:
                msg += (f"Biaya tindakan operatif ({fF(op)}) terisi di kolom BMIV-03 tetapi tidak termasuk TOTAL pada file, "\
                        "sehingga tidak dialokasikan.")
            else:
                msg += "Ada episode dengan total disetujui tetapi tanpa rincian komponen / kolom biaya yang tidak terbaca."
            st.warning(msg)
    d1, d2 = st.columns([1, 4])
    d1.download_button("Unduh LB-ST (.xlsx)",
                       to_xlsx({"LB-ST A": rows_to_df(LB_HEAD_A, A), "LB-ST B": rows_to_df(LB_HEAD_B, B)}),
                       f"LB-ST_{sel[0]}_{sel[-1]}.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with st.expander("Rekap komponen × bulan"):
        sv = st.selectbox("Tampilkan", ["Biaya (Rp)", "Volume"], key="ls")
        val = "biaya" if sv.startswith("Biaya") else "volume"
        Xc = CX[~CX["info"]]
        if Xc.empty:
            empty()
        else:
            pv = Xc.pivot_table(index=["b", "kode", "komponen"], columns="periode", values=val, aggfunc="sum",
                                fill_value=0.0).reindex(columns=sel, fill_value=0.0)
            pv["Total"] = pv.sum(axis=1)
            pv = pv.reset_index().sort_values(["b", "kode"])
            pv["Layanan"] = pv["b"].map(dict(enumerate(BS)))
            pv = pv.rename(columns={**{p: ml(p) for p in sel}, "komponen": "Komponen"})
            cols = ["Layanan", "Komponen"] + [ml(p) for p in sel] + ["Total"]
            pv = pv[cols]
            num = cols[2:]
            table(pv, {c: fN for c in num}, height=420, total=["Total", ""] + [pv[c].sum() for c in num])
    with st.container(border=True):
        st.markdown("**Pengesahan**")
        s1, s2 = st.columns(2)
        s1.markdown("[Kota, tanggal]  \n**Pejabat Pembuat Laporan**  \n\n\n[isi nama]")
        s2.markdown("&nbsp;  \n**Kepala Cabang**  \n\n\n[isi nama]")


def tab_uc(L, C, periods):
    """Unit Cost & Nilai Per Kapita — format template + unit cost per PLKK."""
    c1, c2 = st.columns([2, 2])
    bs = c1.selectbox("Kelompok", ["Semua BMIV"] + BN, key="ub")
    with c2:
        sel = month_pick(periods, "ub_m")
    if not sel:
        return st.info("Pilih minimal satu bulan.")
    sel_b = [0, 1, 2, 3] if bs == "Semua BMIV" else [BN.index(bs)]
    X = L[L["periode"].isin(sel) & L["b"].isin(sel_b)]
    CX = C[C["periode"].isin(sel) & C["b"].isin(sel_b) & ~C["info"]]
    if X.empty:
        return empty()
    basis, tk, m = basis_now(), tk_aktif(), len(sel)
    denom = tk * m
    vals = lb_vals(CX, basis)
    cab, _ = cabang_label(X)

    # ── Parameter (sama dengan blok Parameter pada template) ──
    with st.container(border=True):
        st.markdown("**UNIT COST DAN NILAI PER KAPITA — JKK PLKK**")
        st.caption(f"{cab} | Periode {periode_label(sel)} | Basis biaya: {basis}")
        p1, p2 = st.columns(2)
        p1.metric("Jumlah TK aktif terdaftar JKK (orang)", fN(tk) if tk else "belum diisi")
        p2.metric("Jumlah bulan dalam periode data", m)
        if not tk:
            st.warning("Isi **Jumlah TK aktif JKK** di sidebar. Tanpa penyebut ini, PMPM dan Utilisasi per 1.000 TK tidak "
                       "dihitung (sama seperti template). Unit Cost tetap tampil karena tidak bergantung populasi.")

    # ── Tabel utama ──
    rows, no, gtot = [], 0, 0.0
    flat = []
    for grp in LB_GROUPS:
        if grp["b"] not in sel_b:
            continue
        for r in grp["rows"]:
            if r["jenis"] == "biaya":
                vol, bi = _vb(vals, r["kode"])
                flat.append((grp, r, vol, bi))
                gtot += bi
    hide = st.checkbox("Sembunyikan komponen tanpa data", value=True, key="uc_h")
    cur = None
    sub_bi = {g["b"]: 0.0 for g in LB_GROUPS}
    for grp, r, vol, bi in flat:
        sub_bi[grp["b"]] += bi
    for grp in LB_GROUPS:
        if grp["b"] not in sel_b:
            continue
        for g2, r, vol, bi in [x for x in flat if x[0] is grp]:
            if hide and vol == 0 and bi == 0:
                continue
            no += 1
            uc = bi / vol if vol > 0 and bi > 0 else None
            ut = vol / denom * 1000 if denom else None
            pm = bi / denom if denom else None
            rows.append(dict(cls="row", cells=[str(no), grp["short"], r["nm"], r["sat"], fN(vol), fN(bi), _c(uc),
                                                fU(ut) if ut is not None else "", _c(pm) if pm is not None else "",
                                                fP(bi / gtot) if gtot else "0%"]))
        sb_ = sub_bi[grp["b"]]
        rows.append(dict(cls="sub", cells=["", "", f"Sub Total {grp['short']}", "campuran", "", fN(sb_), "", "",
                                            _c(sb_ / denom) if denom else "", fP(sb_ / gtot) if gtot else "0%"]))
    rows.append(dict(cls="tot", cells=["", "", "TOTAL SELURUH KOMPONEN", "", "", fN(gtot), "", "",
                                        _c(gtot / denom) if denom else "", "100%" if gtot else "0%"]))
    heads = ["No", "Kelompok", "Komponen Layanan", "Satuan", "Volume", "Biaya Disetujui (Rp)", "Unit Cost (Rp)",
             "Utilisasi per 1.000 TK per bulan", "PMPM (Rp/TK/bln)", "% Biaya"]
    ep_n = len(X)
    kpis([("Total Biaya", fR(gtot), fF(gtot)), ("Cost per Case", fF(gtot / X["kasus"].nunique()) if len(X) else "–", "biaya / kasus"),
          ("Unit Cost / Episode", fF(gtot / ep_n) if ep_n else "–", "biaya / kunjungan"),
          ("PMPM", fF(gtot / denom) if denom else "–", "biaya / TK / bulan"),
          ("Utilisasi /1.000 TK/bln", fU(ep_n / denom * 1000) if denom else "–", "kunjungan per 1.000 TK")])
    cg = pd.DataFrame([dict(lbl=f"{r['nm']} ({grp['short']})", vol=vol, bi=bi, ucc=bi / vol if vol > 0 else 0.0,
                            pmpm=bi / denom if denom else 0.0) for grp, r, vol, bi in flat if bi > 0 or vol > 0])
    insight_box(ins_uc(cg, gtot, denom, ep_n, X))
    with st.container(border=True):
        st.markdown("**Unit Cost, Utilisasi & PMPM per Komponen**")
        st.markdown(html_table(heads, rows, num_from=4), unsafe_allow_html=True)
        st.download_button("Unduh tabel (.xlsx)", to_xlsx({"Unit Cost & Per Kapita": rows_to_df(heads, rows)}),
                           f"UnitCost_PerKapita_{sel[0]}_{sel[-1]}.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    if not cg.empty:
        g1, g2 = st.columns(2)
        with g1.container(border=True):
            st.markdown("**Layanan dengan " + ("Biaya per Kapita (PMPM) Terbesar" if denom else "Total Biaya Terbesar") + "**")
            st.caption("PMPM = biaya ÷ (TK aktif × bulan). Dipengaruhi harga dan seberapa sering layanan dipakai."
                       if denom else "Biaya = harga per satuan × frekuensi pakai. Menunjukkan di mana uang paling banyak terserap.")
            t = cg.sort_values("pmpm" if denom else "bi", ascending=False).head(10)
            hbar(t["lbl"], t["pmpm"] if denom else t["bi"], ACC)
        with g2.container(border=True):
            st.markdown("**Layanan dengan Harga per Satuan Tertinggi**")
            st.caption("Unit cost = biaya ÷ volume (hanya komponen dengan volume ≥3). "
                       "Satuan tiap komponen berbeda, jadi peringkat bersifat indikatif.")
            t = cg[cg["vol"] >= 3].sort_values("ucc", ascending=False).head(10)
            hbar(t["lbl"], t["ucc"], "#a855f7")

    # ── Rumus ──
    with st.expander("Perbedaan rumus Unit Cost vs Per Kapita (agar selaras dengan Utilisasi)", expanded=False):
        st.markdown(
            "| Indikator | Rumus | Bergantung populasi? |\n|---|---|---|\n"
            "| **Unit Cost** | Biaya Disetujui ÷ Volume | Tidak |\n"
            "| **Utilisasi** per 1.000 TK per bulan | Volume ÷ (TK aktif × bulan) × 1.000 | Ya |\n"
            "| **PMPM** (Rp/TK/bulan) | Biaya Disetujui ÷ (TK aktif × bulan) | Ya |\n\n"
            "**Identitas:** `PMPM = Unit Cost × Utilisasi ÷ 1.000`. Biaya per kapita naik karena harga per unit naik "
            "(unit cost) **atau** karena frekuensi pemakaian naik (utilisasi). Kedua indikator memakai penyebut yang sama "\
            "(TK aktif × bulan), bukan jumlah peserta yang kebetulan mengklaim.")
        ex = [x for x in flat if x[2] > 0 and x[3] > 0]
        if denom and ex:
            _g, r, vol, bi = max(ex, key=lambda x: x[3])
            st.markdown(f"Contoh dari data ini — **{r['nm']}**: Unit Cost {fF(bi / vol)} × Utilisasi {fU(vol / denom * 1000)} ÷ 1.000 "
                        f"= **{fF(bi / vol * (vol / denom * 1000) / 1000)}** = PMPM {fF(bi / denom)}.")
        st.caption("Pembilang dan penyebut harus satu populasi (TK aktif harus sesuai cabang/kanwil yang difilter). "
                   "Kejadian JKK jarang dan sangat bervariasi; gunakan data kumulatif ≥12 bulan untuk PMPM yang stabil. "
                   "Biaya hanya pelayanan kesehatan PLKK & rehabilitasi (tanpa santunan uang/beasiswa). "
                   "Unit cost Tindakan Operatif adalah gabungan kelas kecil–khusus sehingga rentangnya lebar.")

    # ── Per PLKK ──
    st.markdown("### Unit Cost per PLKK")
    byp = lb_vals(CX, basis, by="plkk")
    opts = {"Semua komponen (biaya per episode)": None}
    for grp, r, vol, bi in flat:
        if bi > 0 or vol > 0:
            opts[f"{r['nm']} ({grp['short']})"] = r["kode"]
    pick = st.selectbox("Komponen yang dibandingkan antar-PLKK", list(opts), key="uc_pk")
    kd = opts[pick]
    if kd is None:
        vs = X.groupby("plkk").size().astype(float)
        bs_ = X.groupby("plkk")["total"].sum()
        unit_lbl = "episode"
    else:
        vs, bs_ = byp.get(kd, (pd.Series(dtype=float), pd.Series(dtype=float)))
        unit_lbl = next(r["sat"] for _g, r, _v, _b in flat if r["kode"] == kd)
    P = pd.DataFrame({"vol": vs, "bi": bs_}).fillna(0.0)
    P = P[(P["vol"] > 0) | (P["bi"] > 0)]
    if P.empty:
        st.info("Tidak ada data PLKK untuk komponen ini.")
        return
    P["uc"] = np.where(P["vol"] > 0, P["bi"] / P["vol"].where(P["vol"] > 0, 1), np.nan)
    med = float(P.loc[P["vol"] > 0, "uc"].median()) if (P["vol"] > 0).any() else np.nan
    P["rasio"] = P["uc"] / med if med and not np.isnan(med) else np.nan
    P["pct"] = P["bi"] / P["bi"].sum() if P["bi"].sum() else 0.0
    P["ut"] = P["vol"] / denom * 1000 if denom else np.nan
    P["pm"] = P["bi"] / denom if denom else np.nan
    P = P.reset_index().rename(columns={"index": "plkk", "_k": "plkk"})
    P = P.sort_values("uc", ascending=False)
    st.caption(f"Unit cost = Biaya ÷ Volume ({unit_lbl}). Median antar-PLKK = {fF(med) if not np.isnan(med) else '–'}; "
               "rasio >1,5× median layak ditelaah. PMPM & utilisasi PLKK = kontribusi PLKK terhadap seluruh TK aktif.")
    with st.container(border=True):
        st.markdown("**Ranking Unit Cost PLKK (Top 15)**")
        T = P[P["vol"] > 0].head(15)
        hbar(T["plkk"], T["uc"], "#a855f7")
    with st.container(border=True):
        st.markdown(f"**Rekap Unit Cost per PLKK** — {len(P)} PLKK")
        df = P.rename(columns=dict(plkk="PLKK", vol="Volume", bi="Biaya Disetujui", uc="Unit Cost", rasio="× Median PLKK",
                                   pct="% Biaya", ut="Util /1.000 TK/bln", pm="PMPM kontribusi"))
        df = df[["PLKK", "Volume", "Biaya Disetujui", "Unit Cost", "× Median PLKK", "% Biaya", "Util /1.000 TK/bln", "PMPM kontribusi"]]
        table(df, {"Volume": fN, "Biaya Disetujui": fF, "Unit Cost": fF, "× Median PLKK": lambda v: f1(v) + "×", "% Biaya": fP,
                   "Util /1.000 TK/bln": fU, "PMPM kontribusi": fF}, height=420)
    with st.expander("Matriks Unit Cost PLKK × Komponen"):
        top = [(grp, r) for grp, r, vol, bi in sorted(flat, key=lambda x: -x[3])[:10] if vol > 0 and bi > 0]
        mat = {}
        for grp, r in top:
            v, b_ = byp.get(r["kode"], (pd.Series(dtype=float), pd.Series(dtype=float)))
            mat[f"{r['nm']} ({grp['short']})"] = (b_ / v.where(v > 0)).dropna()
        if mat:
            M = pd.DataFrame(mat).reset_index().rename(columns={"index": "PLKK", "_k": "PLKK"})
            table(M, {c: fF for c in M.columns if c != "PLKK"}, height=420)
        else:
            st.caption("Belum ada data.")


def ins_uc(cg, gtot, denom, ep_n, X) -> list[str]:
    out = []
    if cg.empty or gtot <= 0:
        return out
    if denom:
        p = cg.sort_values("pmpm", ascending=False).iloc[0]
        out.append(f"**{p['lbl']}** penyumbang PMPM terbesar: {fF(p['pmpm'])} (**{fP(p['pmpm'] / (gtot / denom))}** dari PMPM {fF(gtot / denom)}).")
    u = cg[cg["vol"] >= 3]
    if len(u) >= 2:
        h = u.sort_values("ucc").iloc[-1]
        out.append(f"Unit cost tertinggi: **{h['lbl']}** ({fF(h['ucc'])} per satuan), sedangkan biaya rata-rata per episode {fF(gtot / ep_n) if ep_n else '–'}.")
    t3 = cg.sort_values("bi", ascending=False).head(3)
    out.append(f"3 komponen teratas ({', '.join(x.split(' (')[0] for x in t3['lbl'])}) = **{fP(t3['bi'].sum() / gtot)}** total biaya.")
    k = X["kasus"].nunique()
    out.append(f"Cost per case {fF(gtot / k) if k else '–'} dengan rata-rata **{f1(ep_n / k if k else 0)} episode per kasus**.")
    if denom:
        out.append(f"Utilisasi {fU(ep_n / denom * 1000)} kunjungan per 1.000 TK per bulan; PMPM = Unit Cost/episode × Utilisasi ÷ 1.000.")
    return out


# ─────────────────────────────────────────────────────────────────────────────
# 6c. CASE EXPLORER + DETEKSI OVERPRICE
# ─────────────────────────────────────────────────────────────────────────────
OV_MODES = ["Median seluruh data", "Median PLKK yang sama"]
OV_MIN_N = 5
OV_THR = 1.5   # rasio >= 1,5× normal = indikasi overprice; 1,2–1,5× = waspada


def _price_rows(comp: pd.DataFrame) -> pd.DataFrame:
    """Baris komponen berharga: harga jaminan = nilai kolom BIAYA (mentah) per volume."""
    P = comp[(~comp["info"]) & (comp["volume"] > 0)].copy()
    P["harga"] = np.where(P["biaya_raw"] > 0, P["biaya_raw"], P["biaya"])
    P = P[P["harga"] > 0]
    P["unit"] = P["harga"] / P["volume"]
    return P


def _eval_price(R: pd.DataFrame, ref: pd.DataFrame, mode: str, thr: float) -> pd.DataFrame:
    """Bandingkan harga/unit tiap baris R dengan harga normal (median) dari ref."""
    gl = ref.groupby("kode")["unit"].agg(norm="median", p90=lambda s: float(s.quantile(0.9)), n="size").reset_index()
    R = R.merge(gl, on="kode", how="left")
    R["acuan"] = "semua PLKK"
    if mode == OV_MODES[1]:
        gp = ref.groupby(["plkk", "kode"])["unit"].agg(norm_p="median", n_p="size").reset_index()
        R = R.merge(gp, on=["plkk", "kode"], how="left")
        use = R["n_p"].fillna(0) >= 3
        R["norm"] = np.where(use, R["norm_p"], R["norm"])
        R["acuan"] = np.where(use, "PLKK sama", "semua PLKK")
    R["ok"] = R["n"].fillna(0) >= OV_MIN_N
    R["ratio"] = np.where(R["ok"] & (R["norm"] > 0), R["unit"] / R["norm"], np.nan)
    R["lebih"] = np.where(R["ratio"] >= thr, (R["unit"] - R["norm"]) * R["volume"], 0.0)
    R["status"] = np.select([R["ratio"].isna(), R["ratio"] >= thr, R["ratio"] >= 1.2],
                            ["⚪ Pembanding kurang", "🔴 Indikasi overprice", "🟠 Waspada"], "🟢 Wajar")
    return R


def tab_case(eps, comp):
    st.caption("Case Explorer menelusuri seluruh data (tidak terpengaruh filter sidebar).")
    info = (eps.sort_values("tgl").groupby("kasus")
            .agg(nama=("peserta", "first"), dx=("diagnosa", "first"), bi=("total", "sum")))
    top = info["bi"].sort_values(ascending=False).head(12)
    cases = sorted(info.index)

    def pick(c):
        st.session_state["case_sel"] = c

    st.selectbox("Case ID / nama pasien (ketik untuk mencari)", [""] + cases, key="case_sel",
                 format_func=lambda x: "— pilih No. Kasus / nama pasien —" if x == "" else
                 f"{x} — {info.loc[x, 'nama']} · {info.loc[x, 'dx']}")
    cur = st.session_state.get("case_sel", "")
    if not cur:
        insight_box(ins_case_all(eps))
        with st.container(border=True):
            st.markdown("**Pilih Case ID**")
            st.caption("Pilih No. Kasus / nama di atas, atau klik salah satu kasus berbiaya tertinggi:")
            cols = st.columns(4)
            for i, (c, v) in enumerate(top.items()):
                stretch_button(cols[i % 4], f"{info.loc[c, 'nama']} · {fR(v)}", key=f"chip{i}", on_click=pick, args=(c,))
        return
    E = eps[eps.kasus == cur].sort_values(["tgl_s", "b"])
    cc = comp[comp.ep.isin(E.ep) & ~comp["info"]]
    names = {}
    for ep, g in cc.groupby("ep"):
        names[ep] = ", ".join(f"{r.komponen}{' ×' + fN(r.volume) if r.volume > 1 else ''}" for r in g.itertuples())
    t = float(E["total"].sum())
    days = (E["tgl"].max() - E["tgl"].min()).days
    first = E.iloc[0]
    kpis([("Nama Pasien / TK", first["peserta"], "KPJ " + first["kpj"]), ("Perusahaan", first["perusahaan"], first["sektor"]),
          ("Diagnosa", first["diagnosa"], first["icd"]), ("Jenis Kasus", first["jenis_kasus"], f"Kasus {cur}")])
    kpis([("Total Biaya", fF(t)), ("Layanan / Episode", str(len(E))), ("PLKK Terlibat", str(E["plkk"].nunique())),
          ("Rentang", f"{days} hari", f"{E.tgl_s.iloc[0]} → {E.tgl_s.iloc[-1]}")])
    insight_box(ins_case_one(E, cc, eps))
    with st.container(border=True):
        st.markdown("**Timeline / Journey Layanan**")
        st.caption("Urut tanggal pelayanan / masuk RS")
        parts = []
        for e in E.itertuples():
            extra = f" · {int(e.los)} hari rawat" if e.b == 2 else ""
            parts.append(
                f"<div style='border-left:4px solid {BC[e.b]};padding:2px 0 2px 12px;margin:0 0 12px 4px'>"
                f"<span style='background:{BC[e.b]};color:#fff;border-radius:99px;padding:1px 9px;font-size:12px'>{BS[e.b]}</span> "
                f"<b>{e.tgl_s}</b> · {html.escape(e.plkk)}<br><span style='opacity:.7'>{html.escape(names.get(e.ep, '–'))}{extra}</span>"
                f"<br>{fF(e.total)}</div>")
        st.markdown("".join(parts), unsafe_allow_html=True)
    with st.container(border=True):
        st.markdown("**Detail Baris Data**")
        df = pd.DataFrame(dict(Tanggal=E["tgl_s"], Layanan=E["b"].map(dict(enumerate(BN))), PLKK=E["plkk"],
                               **{"Kunjungan ke": E["visit"], "Komponen": E["ep"].map(names).fillna("–"),
                                  "Biaya Disetujui": E["total"]}))
        table(df, {"Kunjungan ke": lambda v: "–" if pd.isna(v) or not v else fN(v), "Biaya Disetujui": fF})


def plkk_overprice(L, comp, plkk_all):
    """Per PLKK: kasus / layanan mana yang harganya melebihi normal (median PLKK lain)."""
    st.markdown("### Deteksi Overprice per PLKK")
    pl = st.selectbox("PLKK", sorted(plkk_all), key="po_plkk")
    thr = OV_THR
    P = _price_rows(comp)
    ref = P[P["plkk"] != pl]                    # normal = median komponen yang sama di PLKK lain
    R = P[(P["plkk"] == pl) & P["ep"].isin(L["ep"])]
    if R.empty:
        return st.info("PLKK ini tidak punya komponen berharga pada filter saat ini.")
    R = _eval_price(R, ref, OV_MODES[0], thr)
    R = R.merge(L[["ep", "peserta", "tgl_s"]], on="ep", how="left")
    ok = R[R["ok"]]
    nv, ac = float((ok["norm"] * ok["volume"]).sum()), float(ok["harga"].sum())
    kpis([("Biaya Jaminan (terbanding)", fF(ac)), ("Biaya pada Harga Normal", fF(nv)),
          ("Rasio PLKK vs Normal", (f1(ac / nv) + "×") if nv else "–", "tertimbang volume"),
          ("Potensi Kelebihan", fF(R["lebih"].sum()), f"{int((R['ratio'] >= thr).sum())} baris ≥ {f1(thr)}×")])
    st.caption(f"Tiap layanan dibandingkan dengan layanan yang sama di PLKK lain. Rasio = harga per satuan di PLKK ini ÷ "
               f"median harga PLKK lain (minimal {OV_MIN_N} data). 1× = sama dengan normal; 1,6× = 60% lebih mahal. "
               "🔴 ≥ 1,5× · 🟠 1,2–1,5× · 🟢 < 1,2×. Komponen gabungan (mis. tindakan operatif/gigi) harganya bercampur, "
               "jadi rasio hanya indikatif.")
    npl = ref.groupby("kode")["plkk"].nunique()
    A = (R.assign(nv=R["norm"] * R["volume"]).groupby(["kode", "komponen", "satuan"], as_index=False)
         .agg(vol=("volume", "sum"), harga=("harga", "sum"), nv=("nv", "sum"), n=("n", "first"), ok=("ok", "all"),
              kas=("kasus", "nunique"), lebih=("lebih", "sum"), p90=("p90", "first")))
    A["unit"], A["norm"] = A["harga"] / A["vol"], A["nv"] / A["vol"]
    A["npl"] = A["kode"].map(npl).fillna(0)
    A["ratio"] = np.where(A["ok"] & (A["norm"] > 0), A["unit"] / A["norm"], np.nan)
    A["status"] = np.select([A["ratio"].isna(), A["ratio"] >= thr, A["ratio"] >= 1.2],
                            ["⚪ Pembanding kurang", "🔴 Indikasi overprice", "🟠 Waspada"], "🟢 Wajar")
    A = A.sort_values(["ratio", "harga"], ascending=False, na_position="last")
    B = A[A["ratio"].notna()].head(15)
    with st.container(border=True):
        st.markdown("**Layanan / Komponen — Harga PLKK vs Normal**")
        if not B.empty:
            lbl = [f"{k} (/ {s})" for k, s in zip(B["komponen"], B["satuan"])]
            col = ["#dc2626" if r >= thr else "#f59e0b" if r >= 1.2 else "#16a34a" for r in B["ratio"]]
            fig = go.Figure()
            fig.add_bar(y=lbl, x=B["norm"], orientation="h", name="Normal (median PLKK lain)", marker_color="#94a3b8",
                        text=[fF(v) for v in B["norm"]], textposition="outside", cliponaxis=False)
            fig.add_bar(y=lbl, x=B["unit"], orientation="h", name="PLKK ini", marker_color=col,
                        text=[f"{fF(u)} · {f1(r)}×" for u, r in zip(B["unit"], B["ratio"])], textposition="outside",
                        cliponaxis=False)
            fig.update_layout(barmode="group")
            fig.update_yaxes(autorange="reversed", automargin=True)
            vmax = float(max(B["unit"].max(), B["norm"].max()))
            fig.update_xaxes(range=[0, vmax * 1.45])
            money_axis(fig, vmax * 1.45)
            plot(fig, max(300, 62 * len(B) + 90))
        df = pd.DataFrame({"Komponen": A["komponen"], "Satuan": A["satuan"], "Kasus": A["kas"], "Volume": A["vol"],
                           "Biaya Jaminan": A["harga"], "Harga/unit PLKK": A["unit"], "Normal (median)": A["norm"],
                           "P90 PLKK lain": A["p90"], "Jml PLKK Pembanding": A["npl"], "n Pembanding": A["n"], "Rasio": A["ratio"],
                           "Potensi Kelebihan": A["lebih"], "Status": A["status"]})
        table(df, {"Kasus": fN, "Volume": fN, "Biaya Jaminan": fF, "Harga/unit PLKK": fF, "Normal (median)": fF,
                   "P90 PLKK lain": fF, "Jml PLKK Pembanding": fN, "n Pembanding": fN, "Rasio": lambda v: f1(v) + "×", "Potensi Kelebihan": fF})
    with st.container(border=True):
        st.markdown("**Kasus / Layanan yang Melebihi Normal di PLKK ini**")
        F = R[R["ratio"] >= thr].sort_values("lebih", ascending=False)
        if F.empty:
            st.success(f"Tidak ada kasus di PLKK ini yang ≥ {f1(thr)}× normal.")
        else:
            df = pd.DataFrame({"No. Kasus": F["kasus"], "Nama": F["peserta"], "Tanggal": F["tgl_s"], "Layanan": F["b"].map(dict(enumerate(BS))),
                               "Komponen": F["komponen"], "Volume": F["volume"], "Harga/unit": F["unit"], "Normal": F["norm"],
                               "Rasio": F["ratio"], "Potensi Kelebihan": F["lebih"]})
            table(df, {"Volume": fN, "Harga/unit": fF, "Normal": fF, "Rasio": lambda v: f1(v) + "×",
                       "Potensi Kelebihan": fF}, height=420)
            st.caption("Salin No. Kasus ke Case Explorer untuk melihat journey layanannya.")


def _has_width_param() -> bool:
    try:
        import inspect
        return "width" in inspect.signature(st.button).parameters
    except Exception:  # noqa: BLE001
        return False


def stretch_button(container, label: str, **kw) -> None:
    """Tombol selebar kontainer. Hasilnya sengaja tidak dikembalikan agar Streamlit tidak
    menampilkan nilai 'False' di layar."""
    if _has_width_param():
        container.button(label, width="stretch", **kw)
    else:
        container.button(label, **kw)


# ─────────────────────────────────────────────────────────────────────────────
# 7. APLIKASI UTAMA
# ─────────────────────────────────────────────────────────────────────────────
FILTER_KEYS = ["f_kw", "f_cb", "f_pl", "f_b", "f_jk", "f_sk"]


def reset_filters(dmin, dmax):
    for k in FILTER_KEYS:
        st.session_state[k] = []
    st.session_state["f_date"] = (dmin, dmax)
    st.session_state["case_sel"] = ""


def find_default_data() -> Path | None:
    """Cari file data bawaan: data/Data_Dummy.xlsx, atau Data_Dummy.xlsx di folder app.py."""
    base = Path(__file__).resolve().parent
    for p in (base / "data" / "Data_Dummy.xlsx", base / "Data_Dummy.xlsx"):
        if p.exists():
            return p
    return None


def main():
    st.set_page_config(page_title="Dashboard Utilization Review PLKK", page_icon="🏥", layout="wide")
    inject_css()
    _KEY[0] = 0
    sample = find_default_data()

    sb = st.sidebar
    sb.title("Utilization Review PLKK")
    with sb.expander("Ganti sumber data (opsional)", expanded=False):
        up = st.file_uploader("Upload file Excel (format template BMIV-01 s.d. BMIV-04)",
                              type=["xlsx", "xlsm"], key="uploader")
        st.caption("Kosongkan untuk memakai data bawaan.")
    if up is not None:
        content, src = up.getvalue(), up.name
    elif sample is not None:
        content, src = sample.read_bytes(), sample.name + " (data bawaan)"
    else:
        st.title("Dashboard Utilization Review PLKK")
        st.error("File data bawaan tidak ditemukan. Letakkan `Data_Dummy.xlsx` di folder `data/` "
                 "(sejajar dengan app.py), atau upload file Excel lewat sidebar.")
        st.stop()

    try:
        eps, comp, warns = load_data(content, CACHE_VER)
    except TemplateError as e:
        st.error("File tidak sesuai template:")
        st.markdown(str(e))
        st.caption("Pastikan workbook memiliki sheet BMIV-01, BMIV-02, BMIV-03, BMIV-04 (boleh berakhiran "
                   "mis. _Dummy) dengan nama kolom persis seperti template.")
        st.stop()
    if eps.empty:
        st.error("Tidak ada baris data yang valid pada file.")
        st.stop()

    all_periods = sorted(eps["periode"].unique())
    dmin, dmax = eps["tgl"].min().date(), eps["tgl"].max().date()
    sig = (src, len(content))
    if st.session_state.get("sig") != sig:
        reset_filters(dmin, dmax)
        st.session_state["sig"] = sig

    # ── Filter ──
    sb.caption(f"Sumber: **{src}** · {ml(all_periods[0])} – {ml(all_periods[-1])}")
    sb.date_input("Periode Tanggal", min_value=dmin, max_value=dmax, key="f_date")
    sb.multiselect("Kanwil", sorted(eps["kanwil"].unique()), key="f_kw", placeholder="Semua Kanwil")
    sb.multiselect("Cabang", sorted(eps["cabang"].unique()), key="f_cb", placeholder="Semua Cabang")
    sb.multiselect("PLKK", sorted(eps["plkk"].unique()), key="f_pl", placeholder="Semua PLKK")
    sb.multiselect("Jenis Layanan", BN, key="f_b", placeholder="Semua Layanan")
    sb.multiselect("Jenis Kasus", sorted(eps["jenis_kasus"].unique()), key="f_jk", placeholder="Semua Jenis Kasus")
    sb.multiselect("Sektor Usaha", sorted(eps["sektor"].unique()), key="f_sk", placeholder="Semua Sektor")
    stretch_button(sb, "Reset filter", on_click=reset_filters, args=(dmin, dmax))
    sb.markdown("**Parameter per kapita**")
    sb.number_input("Jumlah TK aktif JKK (jiwa)", min_value=0, step=1000, key="f_tk",
                    help="Penyebut PMPM dan Utilisasi per 1.000 TK (sheet Parameter pada template). Harus satu populasi "
                         "dengan biaya yang difilter. Kosong/0 = dipakai proxy peserta unik yang melapor klaim.")

    dr = st.session_state.get("f_date", (dmin, dmax))
    d1 = dr[0] if len(dr) >= 1 else dmin
    d2 = dr[1] if len(dr) >= 2 else dmax
    mask = (eps["tgl_s"] >= d1.isoformat()) & (eps["tgl_s"] <= d2.isoformat())
    for key, col in [("f_kw", "kanwil"), ("f_cb", "cabang"), ("f_pl", "plkk"), ("f_jk", "jenis_kasus"),
                     ("f_sk", "sektor")]:
        sel = st.session_state.get(key) or []
        if sel:
            mask &= eps[col].isin(sel)
    sel_b = [BN.index(x) for x in (st.session_state.get("f_b") or [])]
    if sel_b:
        mask &= eps["b"].isin(sel_b)
    L = eps[mask]
    keep = np.zeros(int(comp["ep"].max()) + 1 if len(comp) else 1, dtype=bool)
    keep[L["ep"].to_numpy()] = True
    Cf = comp[keep[comp["ep"].to_numpy()]]       # termasuk baris informasi (dipakai LB-ST)
    C = Cf[~Cf["info"]]                          # komponen berbiaya untuk tab analisis lain
    periods = [p for p in all_periods if d1.isoformat()[:7] <= p <= d2.isoformat()[:7]]
    m = max(1, len(periods))
    sb.caption(f"{fN(len(L))} dari {fN(len(eps))} baris terfilter")

    # ── Utilitas sidebar ──
    sb.download_button("Unduh data ternormalisasi (.xlsx)", export_xlsx(content, CACHE_VER),
                       "data_ternormalisasi_plkk.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    if sample is not None:
        sb.download_button("Unduh file contoh / template", sample.read_bytes(), sample.name,
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # ── Konten ──
    st.title("Dashboard Utilization Review PLKK")
    tabs = st.tabs(["Executive Summary", "Tren Bulanan", "BMIV-01 RJTP", "BMIV-02 RJTL", "BMIV-03 RANAP",
                    "BMIV-04 Khusus", "PLKK Performance", "LB-ST", "Unit Cost & Per Kapita", "Case Explorer"])
    with tabs[9]:
        tab_case(eps, comp)
    if L.empty:
        for t in tabs[:9]:
            with t:
                empty()
        return
    with tabs[0]:
        tab_exec(L, C, m)
    with tabs[1]:
        tab_trend(L, periods, m)
    for b in range(4):
        with tabs[2 + b]:
            tab_bmiv(L, C, b, m)
    with tabs[6]:
        tab_plkk(L, m, comp)
    with tabs[7]:
        tab_lbst(L, Cf, periods)
    with tabs[8]:
        tab_uc(L, Cf, periods)


if __name__ == "__main__":
    main()
