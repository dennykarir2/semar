"""
Dashboard interaktif Bea Cukai Semarang untuk SEMAR.

- Data dibaca dari data_dashboard_semar.xlsx (sheet "Penerimaan" dan "Penindakan").
- Dashboard diatur lewat chat: perintah dipahami oleh aturan kata kunci, dan bila
  tidak cukup, oleh model DeepSeek yang hanya diminta menghasilkan JSON perintah.
- Semua ANGKA dihitung Python dari Excel. Model tidak pernah menulis angka.
"""

from __future__ import annotations

import html as H
import json
import math
import re
import urllib.parse
from dataclasses import dataclass, field, replace
from datetime import date
from pathlib import Path

import pandas as pd

import peta_wilayah as P

VERSI_DASHBOARD = "2026-11-17"

BULAN = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli",
         "Agustus", "September", "Oktober", "November", "Desember"]
BULAN_PENDEK = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"]
NAMA_BULAN_REGEX = {
    "januari": 1, "jan": 1, "februari": 2, "feb": 2, "pebruari": 2, "maret": 3, "mar": 3, "april": 4, "apr": 4,
    "mei": 5, "juni": 6, "jun": 6, "juli": 7, "jul": 7, "agustus": 8, "agu": 8, "agt": 8, "ags": 8,
    "september": 9, "sep": 9, "sept": 9, "oktober": 10, "okt": 10, "november": 11, "nov": 11, "nopember": 11,
    "desember": 12, "des": 12,
}
WILAYAH: list[str] = list(P.WILAYAH_KERJA)
ALIAS_WILAYAH = {
    "Kota Semarang": [r"kota semarang", r"semarang kota", r"kodya semarang", r"kotamadya semarang"],
    "Kab. Semarang": [r"kabupaten semarang", r"kab\.? semarang", r"semarang kabupaten", r"ungaran"],
    "Kab. Kendal": [r"kendal"],
    "Kab. Demak": [r"demak"],
    "Kab. Grobogan": [r"grobogan", r"purwodadi"],
    "Kota Salatiga": [r"salatiga"],
}
METRIK_NAMA = {"penerimaan": "penerimaan negara", "penindakan": "penindakan pengawasan"}
GRAFIK_NAMA = {"tren": "tren bulanan", "komposisi": "komposisi jenis", "peringkat": "peringkat wilayah"}


# ===========================================================================
# FORMAT ANGKA (gaya Indonesia: titik ribuan, koma desimal)
# ===========================================================================
def _id(x: float, desimal: int = 1) -> str:
    s = f"{x:,.{desimal}f}"
    return s.replace(",", "§").replace(".", ",").replace("§", ".")


def rp(x: float | None, singkat: bool = False) -> str:
    if x is None or (isinstance(x, float) and math.isnan(x)):
        return "-"
    a = abs(x)
    if a >= 1e12:
        return f"Rp {_id(x / 1e12, 2)} " + ("T" if singkat else "triliun")
    if a >= 1e9:
        return f"Rp {_id(x / 1e9, 1)} " + ("M" if singkat else "miliar")
    if a >= 1e6:
        return f"Rp {_id(x / 1e6, 1)} " + ("jt" if singkat else "juta")
    return f"Rp {_id(x, 0)}"


def angka(x: float | None) -> str:
    return "-" if x is None else _id(x, 0)


def persen(x: float | None, tanda: bool = False) -> str:
    if x is None:
        return "-"
    return (f"{x:+.1f}%" if tanda else f"{x:.1f}%").replace(".", ",")


def nama_bulan_rentang(a: int, b: int) -> str:
    return BULAN[a - 1] if a == b else f"{BULAN[a - 1]}\u2013{BULAN[b - 1]}"


# ===========================================================================
# NORMALISASI NAMA WILAYAH
# ===========================================================================
def kanon_wilayah(nama: str) -> str:
    t = re.sub(r"\s+", " ", str(nama).strip().lower())
    for kanon, pola in ALIAS_WILAYAH.items():
        if any(re.fullmatch(p, t) for p in pola):
            return kanon
    inti = re.sub(r"^(kabupaten|kab\.?|kota madya|kotamadya|kodya|kota)\s+", "", t)
    if inti == "semarang":
        return "Kab. Semarang" if re.match(r"^(kabupaten|kab)", t) else "Kota Semarang"
    for kanon in WILAYAH:
        if inti == re.sub(r"^(kab\.|kota)\s+", "", kanon.lower()):
            return kanon
    return str(nama).strip()


# ===========================================================================
# DATA
# ===========================================================================
class DataDashboard:
    def __init__(self, path: Path):
        self.path = Path(path)
        self.mtime = None
        self.pen = pd.DataFrame()
        self.tin = pd.DataFrame()
        self.galat: list[str] = []

    # -- membaca Excel -----------------------------------------------------
    @staticmethod
    def _kolom(df: pd.DataFrame, peta: dict[str, str]) -> pd.DataFrame:
        baru = {}
        for c in df.columns:
            kecil = str(c).strip().lower()
            for awalan, tujuan in peta.items():
                if kecil.startswith(awalan) and tujuan not in baru.values():
                    baru[c] = tujuan
                    break
        return df.rename(columns=baru)

    @staticmethod
    def _bulan(v) -> float:
        if pd.isna(v):
            return math.nan
        s = str(v).strip().lower()
        if s in NAMA_BULAN_REGEX:
            return float(NAMA_BULAN_REGEX[s])
        try:
            return float(s)
        except ValueError:
            return math.nan

    def _siapkan(self, df: pd.DataFrame, peta: dict[str, str], wajib: list[str], angka_kol: list[str]) -> pd.DataFrame:
        df = self._kolom(df, peta)
        for k in wajib:
            if k not in df.columns:
                raise ValueError(f"kolom '{k}' tidak ditemukan")
        for k in ("target", "nilai", "ket"):
            if k not in df.columns:
                df[k] = math.nan if k != "ket" else ""
        df["bulan"] = df["bulan"].map(self._bulan)
        for k in ["tahun", *angka_kol]:
            df[k] = pd.to_numeric(df[k], errors="coerce")
        df["target"] = pd.to_numeric(df["target"], errors="coerce")
        df["nilai"] = pd.to_numeric(df["nilai"], errors="coerce")
        df = df.dropna(subset=["tahun", "bulan", "wilayah", "jenis", *angka_kol])
        df = df[(df["bulan"] >= 1) & (df["bulan"] <= 12)].copy()
        df["tahun"] = df["tahun"].astype(int)
        df["bulan"] = df["bulan"].astype(int)
        df["wilayah"] = df["wilayah"].map(kanon_wilayah)
        df["jenis"] = df["jenis"].astype(str).str.strip()
        df["ket"] = df["ket"].fillna("").astype(str)
        return df

    def muat(self, paksa: bool = False) -> None:
        if not self.path.exists():
            self.galat = [f"File data tidak ditemukan: {self.path.name}"]
            self.pen = self.tin = pd.DataFrame()
            return
        mtime = self.path.stat().st_mtime
        if not paksa and mtime == self.mtime:
            return
        galat, pen, tin = [], pd.DataFrame(), pd.DataFrame()
        try:
            xl = pd.ExcelFile(self.path, engine="openpyxl")
            nama = {n.strip().lower(): n for n in xl.sheet_names}
            if "penerimaan" in nama:
                try:
                    pen = self._siapkan(
                        xl.parse(nama["penerimaan"]),
                        {"tahun": "tahun", "bulan": "bulan", "wilayah": "wilayah", "jenis": "jenis",
                         "realisasi": "realisasi", "target": "target", "keterangan": "ket"},
                        ["tahun", "bulan", "wilayah", "jenis", "realisasi"], ["realisasi"])
                except Exception as e:  # noqa: BLE001
                    galat.append(f"Sheet 'Penerimaan' bermasalah: {e}")
            else:
                galat.append("Sheet 'Penerimaan' tidak ditemukan.")
            if "penindakan" in nama:
                try:
                    tin = self._siapkan(
                        xl.parse(nama["penindakan"]),
                        {"tahun": "tahun", "bulan": "bulan", "wilayah": "wilayah", "jenis": "jenis",
                         "jumlah": "jumlah", "nilai": "nilai", "keterangan": "ket"},
                        ["tahun", "bulan", "wilayah", "jenis", "jumlah"], ["jumlah"])
                except Exception as e:  # noqa: BLE001
                    galat.append(f"Sheet 'Penindakan' bermasalah: {e}")
            else:
                galat.append("Sheet 'Penindakan' tidak ditemukan.")
        except PermissionError:
            galat.append("File data sedang terkunci (Database Pengetahuan masih terbuka dan belum disimpan?). Memakai data terakhir.")
            self.galat = galat
            return
        except Exception as e:  # noqa: BLE001
            galat.append(f"Gagal membaca file data: {e}")
        self.pen, self.tin, self.galat, self.mtime = pen, tin, galat, mtime

    # -- sifat -------------------------------------------------------------
    def df(self, metrik: str) -> tuple[pd.DataFrame, str]:
        return (self.pen, "realisasi") if metrik == "penerimaan" else (self.tin, "jumlah")

    @property
    def ada_data(self) -> bool:
        return not (self.pen.empty and self.tin.empty)

    @property
    def ada_contoh(self) -> bool:
        return any(d["ket"].str.contains("contoh", case=False).any() for d in (self.pen, self.tin) if not d.empty)

    def tahun_ada(self, metrik: str | None = None) -> list[int]:
        frames = [self.df(metrik)[0]] if metrik else [self.pen, self.tin]
        th = set()
        for d in frames:
            if not d.empty:
                th |= set(d["tahun"].unique())
        return sorted(int(t) for t in th)

    def jenis_ada(self, metrik: str) -> list[str]:
        d, _ = self.df(metrik)
        return sorted(d["jenis"].unique()) if not d.empty else []

    def wilayah_ada(self) -> list[str]:
        w = set(WILAYAH)
        for d in (self.pen, self.tin):
            if not d.empty:
                w |= set(d["wilayah"].unique())
        return WILAYAH + sorted(w - set(WILAYAH))

    def bulan_terakhir(self, metrik: str, tahun: int) -> int:
        d, _ = self.df(metrik)
        if d.empty:
            return 12
        s = d[d["tahun"] == tahun]["bulan"]
        return int(s.max()) if len(s) else 12


# ===========================================================================
# KONDISI DASHBOARD
# ===========================================================================
@dataclass
class Kondisi:
    metrik: str = "penerimaan"
    tahun: int = 0
    banding: int | None = None
    bulan_awal: int = 1
    bulan_akhir: int = 12
    wilayah: list[str] = field(default_factory=list)   # kosong = semua wilayah
    jenis: list[str] = field(default_factory=list)     # kosong = semua jenis
    grafik: str = "tren"
    bulan_manual: bool = False


def kondisi_awal(data: DataDashboard) -> Kondisi:
    th = data.tahun_ada("penerimaan") or data.tahun_ada() or [date.today().year]
    tahun = max(th)
    return Kondisi(tahun=tahun, banding=(tahun - 1 if (tahun - 1) in th else None),
                   bulan_akhir=data.bulan_terakhir("penerimaan", tahun))


@dataclass
class Niat:
    metrik: str | None = None
    tahun: int | None = None
    banding: int | str | None = None         # int, "hapus", atau None (tidak disebut)
    bulan_awal: int | None = None
    bulan_akhir: int | None = None
    semua_bulan: bool = False
    wilayah: list[str] | None = None         # None = tidak disebut, [] = semua wilayah
    jenis: list[str] | None = None
    grafik: str | None = None
    reset: bool = False
    bantuan: bool = False
    tampilkan: bool = False                   # sekadar "tampilkan dashboard"
    sebab: bool = False                       # pertanyaan "mengapa ..." (penyebab tidak ada di data)

    def kosong(self) -> bool:
        return not any([self.metrik, self.tahun, self.banding, self.bulan_awal, self.semua_bulan,
                        self.wilayah is not None, self.jenis is not None, self.grafik, self.reset,
                        self.bantuan, self.tampilkan])


# ===========================================================================
# PENGURAI PERINTAH (aturan kata kunci, tanpa API)
# ===========================================================================
KATA_PENINDAKAN = r"penindakan|pengawasan|ilegal|sitaan|penyitaan"
KATA_PENERIMAAN = r"penerimaan|realisasi|pendapatan|setoran|target|capaian"
KATA_DASHBOARD = r"dashboard|papan\s*data|peta|grafik|tren|diagram"
ALIAS_JENIS = {
    "cht": ["cukai hasil tembakau"], "tembakau": ["cukai hasil tembakau"], "rokok": ["cukai hasil tembakau", "rokok ilegal"],
    "mmea": ["cukai mmea"], "alkohol": ["cukai mmea", "minuman beralkohol ilegal"],
    "miras": ["minuman beralkohol ilegal"], "kiriman": ["barang kiriman"],
}
KATA_UMUM_JENIS = {"barang", "lainnya", "hasil", "masuk", "keluar", "jenis", "cukai", "minuman", "impor", "ekspor"}
DEFINISI = re.compile(r"^\s*(apa\s+(itu|yang\s+dimaksud)|apakah\s+yang\s+dimaksud|pengertian|definisi|arti\s+dari|jelaskan\s+(apa|tentang|mengenai))\b")
SEBAB = re.compile(r"\b(mengapa|kenapa|penyebab|alasan|faktor)\b")
BANTUAN = re.compile(r"(contoh\s+perintah|apa\s+saja\s+(perintah|yang\s+bisa)|cara\s+(mengatur|memakai|menggunakan)\s+dashboard|bantuan\s+dashboard|help\s+dashboard)")
RESET = re.compile(r"\b(reset|atur\s+ulang|kembalikan|kembali\s+ke\s+(awal|semula|default)|bersihkan\s+filter|hapus\s+(semua\s+)?filter|default)\b")
BANDING_KATA = re.compile(r"\b(bandingkan|dibandingkan|membandingkan|banding|versus|vs|dibanding|komparasi|perbandingan)\b")
TANPA_BANDING = re.compile(r"(tanpa\s+(perbandingan|banding)|hapus\s+(perbandingan|banding)|matikan\s+(perbandingan|banding)|tidak\s+usah\s+(dibandingkan|banding))")
RANGE_KATA = re.compile(r"(sampai|s\.?d\.?|hingga|sampe|\bdari\b|[-\u2013\u2014])")


def _romawi(s: str) -> int | None:
    return {"1": 1, "2": 2, "3": 3, "4": 4, "i": 1, "ii": 2, "iii": 3, "iv": 4}.get(s.lower())


def _kata_jenis_unik(semua: dict[str, list[str]]) -> dict[str, str]:
    """Kata penciri yang hanya muncul di satu nama jenis -> nama jenis itu."""
    hitung: dict[str, set[str]] = {}
    for jenis_list in semua.values():
        for j in jenis_list:
            for tok in set(re.findall(r"[a-z0-9]{4,}", j.lower())):
                if tok in KATA_UMUM_JENIS:
                    continue
                hitung.setdefault(tok, set()).add(j)
    return {tok: next(iter(js)) for tok, js in hitung.items() if len(js) == 1}


def parse_aturan(teks: str, data: DataDashboard, k: Kondisi, hari_ini: date | None = None,
                 lanjutan: bool = False) -> Niat | None:
    t = " " + re.sub(r"\s+", " ", teks.lower().strip()) + " "
    hari_ini = hari_ini or date.today()
    n = Niat()

    if BANTUAN.search(t):
        n.bantuan = True
        return n
    if RESET.search(t) and re.search(r"dashboard|filter|tampilan|peta|semua", t):
        n.reset = True
        return n

    # --- metrik
    m_tin = re.search(KATA_PENINDAKAN, t) is not None
    m_pen = re.search(KATA_PENERIMAAN, t) is not None
    if m_tin and not m_pen:
        n.metrik = "penindakan"
    elif m_pen and not m_tin:
        n.metrik = "penerimaan"
    elif m_pen and m_tin:
        n.metrik = "penindakan" if t.find("penindakan") != -1 and (t.find("penerimaan") == -1 or t.find("penindakan") < t.find("penerimaan")) else "penerimaan"

    # --- tahun & perbandingan
    tahun_ada = data.tahun_ada()
    th_teks = [int(x) for x in re.findall(r"\b(20\d{2})\b", t)]
    maks = max(tahun_ada) if tahun_ada else hari_ini.year
    if TANPA_BANDING.search(t):
        n.banding = "hapus"
    elif BANDING_KATA.search(t):
        if len(set(th_teks)) >= 2:
            n.tahun, n.banding = max(th_teks), min(th_teks)
        elif len(set(th_teks)) == 1:
            n.tahun = th_teks[0]
            n.banding = th_teks[0] - 1
        else:
            dasar = n.tahun or k.tahun or maks
            n.banding = dasar - 1
    elif len(set(th_teks)) >= 2:
        n.tahun, n.banding = max(th_teks), min(th_teks)
    elif len(th_teks) == 1:
        n.tahun = th_teks[0]
    if n.tahun is None and n.banding in (None,):
        if re.search(r"tahun\s+(lalu|kemarin|sebelumnya)", t):
            n.tahun = min(max(tahun_ada), hari_ini.year - 1) if tahun_ada else hari_ini.year - 1
        elif re.search(r"tahun\s+(ini|berjalan)", t):
            n.tahun = min(maks, hari_ini.year)
        elif re.search(r"tahun\s+(terakhir|terbaru)", t):
            n.tahun = maks

    # --- bulan
    if re.search(r"(setahun\s+penuh|semua\s+bulan|seluruh\s+bulan|sepanjang\s+tahun|satu\s+tahun\s+penuh)", t):
        n.semua_bulan = True
    else:
        mt = re.search(r"\b(triwulan|kuartal|tw|q)\s*\.?\s*(i{1,3}|iv|[1-4])\b", t)
        ms = re.search(r"\bsemester\s*(ii|i|[12])\b", t)
        if mt and _romawi(mt.group(2)):
            q = _romawi(mt.group(2))
            n.bulan_awal, n.bulan_akhir = 3 * (q - 1) + 1, 3 * q
        elif ms:
            s = {"1": 1, "i": 1, "2": 2, "ii": 2}[ms.group(1)]
            n.bulan_awal, n.bulan_akhir = (1, 6) if s == 1 else (7, 12)
        else:
            urut = []
            for mm in re.finditer(r"\b(" + "|".join(sorted(NAMA_BULAN_REGEX, key=len, reverse=True)) + r")\b", t):
                nomor = NAMA_BULAN_REGEX[mm.group(1)]
                # "mar", "des", "sep" dll. hanya dianggap bulan jika bukan bagian kata lain (sudah dijamin \b)
                urut.append(nomor)
            if len(urut) >= 2 and RANGE_KATA.search(t):
                n.bulan_awal, n.bulan_akhir = urut[0], urut[1]
            elif len(urut) >= 2:
                n.bulan_awal, n.bulan_akhir = min(urut), max(urut)
            elif len(urut) == 1:
                n.bulan_awal = n.bulan_akhir = urut[0]
        if n.bulan_awal and n.bulan_akhir and n.bulan_awal > n.bulan_akhir:
            n.bulan_awal, n.bulan_akhir = n.bulan_akhir, n.bulan_awal

    # --- wilayah
    if re.search(r"(semua|seluruh)\s+(wilayah|daerah|kabupaten|kota)|se-?wilayah", t):
        n.wilayah = []
    else:
        ketemu = []
        for kanon, pola in ALIAS_WILAYAH.items():
            if any(re.search(r"\b" + p + r"\b", t) for p in pola):
                ketemu.append(kanon)
        if ketemu:
            n.wilayah = ketemu

    # --- jenis (disaring menurut metrik yang dipakai)
    semua_jenis = {"penerimaan": data.jenis_ada("penerimaan"), "penindakan": data.jenis_ada("penindakan")}
    metrik_pakai = n.metrik or k.metrik
    unik = _kata_jenis_unik(semua_jenis)
    cocok: dict[str, list[str]] = {"penerimaan": [], "penindakan": []}
    for mtr, daftar in semua_jenis.items():
        for j in daftar:
            if j.lower() in t:
                cocok[mtr].append(j)
        for tok, j in unik.items():
            if j in daftar and re.search(r"\b" + re.escape(tok) + r"\b", t) and j not in cocok[mtr]:
                cocok[mtr].append(j)
        for alias, target in ALIAS_JENIS.items():
            if re.search(r"\b" + alias + r"\b", t):
                for j in daftar:
                    if j.lower() in target and j not in cocok[mtr]:
                        cocok[mtr].append(j)
    if n.metrik is None and (cocok["penerimaan"] or cocok["penindakan"]):
        if cocok["penindakan"] and not cocok["penerimaan"]:
            n.metrik = "penindakan"
        elif cocok["penerimaan"] and not cocok["penindakan"]:
            n.metrik = "penerimaan"
        metrik_pakai = n.metrik or metrik_pakai
    if cocok[metrik_pakai]:
        n.jenis = cocok[metrik_pakai]
    elif re.search(r"(semua|seluruh)\s+jenis", t):
        n.jenis = []

    # --- grafik
    if re.search(r"\b(komposisi|proporsi|porsi|persentase\s+jenis|menurut\s+jenis|per\s+jenis)\b", t):
        n.grafik = "komposisi"
    elif re.search(r"\b(peringkat|ranking|urutan|tertinggi|terbesar|terbanyak|terendah|terkecil|per\s+wilayah)\b", t):
        n.grafik = "peringkat"
    elif re.search(r"\b(tren|trend|per\s+bulan|bulanan|garis|perkembangan)\b", t):
        n.grafik = "tren"

    n.sebab = SEBAB.search(t) is not None
    if re.search(KATA_DASHBOARD, t) or re.search(r"\btampilkan\b|\blihat\b|\bbuka\b", t):
        if n.kosong():
            n.tampilkan = bool(re.search(KATA_DASHBOARD, t))

    if n.kosong():
        return None

    # --- hanya perintah yang SINYALNYA KUAT yang dipegang dashboard. Pertanyaan umum
    #     ("syarat impor barang kiriman?", "alamat kantor di Kendal?") tetap ke basis pengetahuan.
    ada_slot = any([n.tahun, n.banding, n.bulan_awal, n.semua_bulan, n.wilayah is not None, n.jenis is not None, n.grafik])
    metrik_eksplisit = m_tin or m_pen
    verba = re.search(r"\b(tampilkan|tunjukkan|munculkan|fokus|sorot|filter|ganti|ubah|atur|bandingkan|dibandingkan|banding)\b", t)
    data_kata = re.search(r"\b(berapa|jumlah|total|statistik|angka|data|nilai|rata-rata)\b", t)
    kata_dash = re.search(KATA_DASHBOARD, t)
    kata_kb = re.search(r"\b(tarif|syarat|persyaratan|prosedur|cara|aturan|peraturan|ketentuan|dokumen|alamat|jam|biaya|"
                        r"denda|sanksi|lokasi|kontak|telepon|izin|perizinan|formulir)\b", t)
    tipe = sum(bool(x) for x in [n.tahun or n.banding, n.bulan_awal or n.semua_bulan, n.wilayah is not None,
                                  n.jenis is not None, n.grafik])
    kuat = bool(n.reset or n.bantuan or n.tampilkan or n.banding == "hapus"
                or (kata_dash and (ada_slot or metrik_eksplisit))
                or (verba and (ada_slot or metrik_eksplisit))
                or (metrik_eksplisit and not kata_kb and (ada_slot or data_kata))
                or (not kata_kb and (tipe >= 3 or (n.jenis is not None and tipe >= 2)))
                # kelanjutan singkat saat percakapan sedang membahas dashboard ("triwulan 2", "semua wilayah")
                or (lanjutan and not kata_kb and len(t.split()) <= 9 and ada_slot))
    if not kuat:
        return None
    # pertanyaan definisi murni ("apa itu penindakan?") diserahkan ke basis pengetahuan
    if DEFINISI.search(t) and not any([n.tahun, n.banding, n.bulan_awal, n.semua_bulan, n.wilayah, n.jenis, n.grafik]) \
            and not re.search(KATA_DASHBOARD, t):
        return None
    return n


def gerbang_kuat(teks: str, data: DataDashboard) -> bool:
    """Apakah pesan ini layak dikirim ke model untuk diurai (hanya dipakai bila aturan gagal)?"""
    t = teks.lower()
    if DEFINISI.search(t):
        return False
    kata = [KATA_DASHBOARD, KATA_PENERIMAAN, KATA_PENINDAKAN, r"\bwilayah\b", r"\b20\d{2}\b", r"\bbandingkan\b",
            "|".join(NAMA_BULAN_REGEX), "|".join(p for ps in ALIAS_WILAYAH.values() for p in ps)]
    return any(re.search(k_, t) for k_ in kata)


# ===========================================================================
# PENGURAI PERINTAH VIA MODEL (hanya menghasilkan JSON; angka tetap dari Excel)
# ===========================================================================
PROMPT_NIAT = """Kamu adalah pengurai perintah untuk dashboard "Bea Cukai Semarang". Tugasmu HANYA mengubah pesan pengguna menjadi JSON. Jangan menjawab pertanyaan, jangan menulis angka.

Balas dengan satu objek JSON saja (json), tanpa teks lain, dengan bidang:
{{"niat": "dashboard" | "bukan_dashboard",
 "metrik": "penerimaan" | "penindakan" | null,
 "tahun": angka | null,
 "banding": angka | "hapus" | null,
 "bulan_awal": 1-12 | null, "bulan_akhir": 1-12 | null, "semua_bulan": true | false,
 "wilayah": [nama,...] | [] | null,
 "jenis": [nama,...] | [] | null,
 "grafik": "tren" | "komposisi" | "peringkat" | null,
 "reset": true | false}}

Aturan:
- Isi hanya bidang yang disebut pengguna; sisanya null. "wilayah": [] artinya semua wilayah. "jenis": [] artinya semua jenis.
- "niat" = "bukan_dashboard" bila pesan bukan perintah atau pertanyaan tentang penerimaan negara, penindakan pengawasan, peta, wilayah, jenis, atau periode (misalnya pertanyaan definisi atau prosedur umum).
- "penerimaan" = penerimaan negara (bea masuk, bea keluar, cukai). "penindakan" = jumlah penindakan pengawasan.
- Nama wilayah HARUS salah satu dari: {wilayah}. "Semarang" saja yang tidak jelas: pilih "Kota Semarang".
- Nama jenis HARUS persis dari daftar berikut. Penerimaan: {jenis_pen}. Penindakan: {jenis_tin}.
- Tahun tersedia: {tahun}. Tahun berjalan: {sekarang}. "tahun lalu" = {lalu}.
- Keadaan dashboard saat ini: {kondisi}
"""


def _json_dari(teks: str) -> dict | None:
    teks = teks.strip()
    teks = re.sub(r"^```(?:json)?|```$", "", teks, flags=re.M).strip()
    try:
        return json.loads(teks)
    except json.JSONDecodeError:
        m = re.search(r"\{.*\}", teks, flags=re.S)
        if m:
            try:
                return json.loads(m.group(0))
            except json.JSONDecodeError:
                return None
    return None


def niat_dari_json(d: dict, data: DataDashboard, k: Kondisi) -> Niat | None:
    if not isinstance(d, dict) or d.get("niat") == "bukan_dashboard":
        return None
    n = Niat()
    if d.get("metrik") in ("penerimaan", "penindakan"):
        n.metrik = d["metrik"]
    th = data.tahun_ada()
    if isinstance(d.get("tahun"), int) and not isinstance(d.get("tahun"), bool):
        n.tahun = d["tahun"]
    if d.get("banding") == "hapus":
        n.banding = "hapus"
    elif isinstance(d.get("banding"), int) and not isinstance(d.get("banding"), bool):
        n.banding = d["banding"]
    for kunci in ("bulan_awal", "bulan_akhir"):
        v = d.get(kunci)
        if isinstance(v, int) and not isinstance(v, bool) and 1 <= v <= 12:
            setattr(n, kunci, v)
    if n.bulan_awal and not n.bulan_akhir:
        n.bulan_akhir = n.bulan_awal
    if n.bulan_akhir and not n.bulan_awal:
        n.bulan_awal = n.bulan_akhir
    n.semua_bulan = bool(d.get("semua_bulan"))
    if isinstance(d.get("wilayah"), list):
        valid = [kanon_wilayah(x) for x in d["wilayah"] if isinstance(x, str)]
        n.wilayah = [w for w in dict.fromkeys(valid) if w in data.wilayah_ada()]
        if d["wilayah"] and not n.wilayah:
            n.wilayah = None
    if isinstance(d.get("jenis"), list):
        pakai = n.metrik or k.metrik
        daftar = {j.lower(): j for j in data.jenis_ada(pakai)}
        n.jenis = [daftar[x.lower()] for x in d["jenis"] if isinstance(x, str) and x.lower() in daftar]
        if d["jenis"] and not n.jenis:
            n.jenis = None
    if d.get("grafik") in GRAFIK_NAMA:
        n.grafik = d["grafik"]
    n.reset = bool(d.get("reset"))
    return None if n.kosong() else n


def parse_llm(klien, model: str, extra_body: dict, teks: str, data: DataDashboard, k: Kondisi,
              hari_ini: date | None = None) -> Niat | None:
    hari_ini = hari_ini or date.today()
    th = data.tahun_ada()
    sistem = PROMPT_NIAT.format(
        wilayah=", ".join(WILAYAH), jenis_pen=", ".join(data.jenis_ada("penerimaan")) or "-",
        jenis_tin=", ".join(data.jenis_ada("penindakan")) or "-", tahun=", ".join(map(str, th)) or "-",
        sekarang=hari_ini.year, lalu=hari_ini.year - 1, kondisi=json.dumps(k.__dict__, ensure_ascii=False))
    resp = klien.chat.completions.create(
        model=model, temperature=0.1, max_tokens=300, extra_body=extra_body,
        response_format={"type": "json_object"},
        messages=[{"role": "system", "content": sistem}, {"role": "user", "content": teks}])
    d = _json_dari(resp.choices[0].message.content or "")
    n = niat_dari_json(d, data, k) if d else None
    if n is not None:
        n.sebab = SEBAB.search(teks.lower()) is not None
    return n


def pahami(teks: str, data: DataDashboard, k: Kondisi, klien=None, model: str = "", extra_body: dict | None = None,
           hari_ini: date | None = None, lanjutan: bool = False) -> Niat | None:
    n = parse_aturan(teks, data, k, hari_ini, lanjutan)
    if n is not None:
        return n
    if klien is not None and gerbang_kuat(teks, data):
        try:
            return parse_llm(klien, model, extra_body or {}, teks, data, k, hari_ini)
        except Exception:  # noqa: BLE001  (tanpa jaringan: aturan saja)
            return None
    return None


# ===========================================================================
# MENERAPKAN PERINTAH
# ===========================================================================
def terapkan(k: Kondisi, n: Niat, data: DataDashboard) -> tuple[Kondisi, list[str]]:
    cat: list[str] = []
    if n.reset:
        return kondisi_awal(data), ["Filter dikembalikan ke pengaturan awal."]
    b = replace(k, wilayah=list(k.wilayah), jenis=list(k.jenis))

    if n.metrik and n.metrik != b.metrik:
        b.metrik = n.metrik
        b.jenis = []                                   # jenis penerimaan dan penindakan berbeda
        if not n.jenis:
            pass
    th_metrik = data.tahun_ada(b.metrik) or data.tahun_ada()
    if n.tahun is not None:
        if n.tahun in th_metrik:
            if n.tahun != b.tahun and not b.bulan_manual and not (n.bulan_awal or n.semua_bulan):
                b.bulan_awal, b.bulan_akhir = 1, data.bulan_terakhir(b.metrik, n.tahun)
            b.tahun = n.tahun
        else:
            cat.append(f"Tahun {n.tahun} tidak ada di data (tersedia: {', '.join(map(str, th_metrik))}).")
    elif b.tahun not in th_metrik and th_metrik:
        b.tahun = max(th_metrik)
        b.bulan_akhir = data.bulan_terakhir(b.metrik, b.tahun)
    if n.banding == "hapus":
        b.banding = None
    elif isinstance(n.banding, int):
        if n.banding in th_metrik and n.banding != b.tahun:
            b.banding = n.banding
        else:
            cat.append(f"Tahun pembanding {n.banding} tidak tersedia, perbandingan tidak diubah.")
    if b.banding is not None and b.banding == b.tahun:
        b.banding = None
    if n.semua_bulan:
        b.bulan_awal, b.bulan_akhir, b.bulan_manual = 1, data.bulan_terakhir(b.metrik, b.tahun), False
    elif n.bulan_awal:
        b.bulan_awal, b.bulan_akhir, b.bulan_manual = n.bulan_awal, n.bulan_akhir or n.bulan_awal, True
    if n.wilayah is not None:
        b.wilayah = [w for w in n.wilayah if w in data.wilayah_ada()]
    if n.jenis is not None:
        b.jenis = list(n.jenis)
    if n.grafik:
        b.grafik = n.grafik
    if b.bulan_akhir < b.bulan_awal:
        b.bulan_awal, b.bulan_akhir = b.bulan_akhir, b.bulan_awal
    terakhir = data.bulan_terakhir(b.metrik, b.tahun)
    if b.bulan_awal > terakhir and not data.df(b.metrik)[0].empty:
        cat.append(f"Data {b.tahun} baru tersedia sampai {BULAN[terakhir - 1]}; bulan setelahnya kosong.")
    return b, cat


# ===========================================================================
# PERHITUNGAN
# ===========================================================================
@dataclass
class Ringkasan:
    metrik: str
    tahun: int
    banding: int | None
    ada_data: bool
    total: float = 0.0
    total_banding: float | None = None
    delta: float | None = None
    target_total: float | None = None
    capaian: float | None = None
    nilai_barang: float | None = None
    nilai_barang_banding: float | None = None
    per_bulan: list = field(default_factory=lambda: [None] * 12)
    per_bulan_banding: list | None = None
    per_wilayah: dict = field(default_factory=dict)
    per_wilayah_banding: dict | None = None
    per_jenis: dict = field(default_factory=dict)
    bulan_terakhir: int = 12
    rata_bulanan: float | None = None


def _saring(df: pd.DataFrame, k: Kondisi, tahun: int, *, bulan=True, wilayah=True, jenis=True) -> pd.DataFrame:
    if df.empty:
        return df
    d = df[df["tahun"] == tahun]
    if bulan:
        d = d[(d["bulan"] >= k.bulan_awal) & (d["bulan"] <= k.bulan_akhir)]
    if wilayah and k.wilayah:
        d = d[d["wilayah"].isin(k.wilayah)]
    if jenis and k.jenis:
        d = d[d["jenis"].isin(k.jenis)]
    return d


def ringkas(k: Kondisi, data: DataDashboard) -> Ringkasan:
    df, kol = data.df(k.metrik)
    r = Ringkasan(metrik=k.metrik, tahun=k.tahun, banding=k.banding, ada_data=not df.empty)
    if df.empty:
        return r
    r.bulan_terakhir = data.bulan_terakhir(k.metrik, k.tahun)
    inti = _saring(df, k, k.tahun)
    r.total = float(inti[kol].sum())
    n_bulan = inti["bulan"].nunique()
    r.rata_bulanan = r.total / n_bulan if n_bulan else None
    if k.metrik == "penerimaan":
        ada_t = inti.dropna(subset=["target"])
        tt = float(ada_t["target"].sum())
        if tt > 0:
            r.target_total = tt
            r.capaian = float(ada_t["realisasi"].sum()) / tt * 100
    else:
        r.nilai_barang = float(inti["nilai"].sum()) if inti["nilai"].notna().any() else None
    if k.banding is not None:
        ban = _saring(df, k, k.banding)
        r.total_banding = float(ban[kol].sum())
        if r.total_banding > 0:
            r.delta = (r.total / r.total_banding - 1) * 100
        if k.metrik == "penindakan" and ban["nilai"].notna().any():
            r.nilai_barang_banding = float(ban["nilai"].sum())

    def per_bulan(tahun: int) -> list:
        d = _saring(df, k, tahun, bulan=False)
        g = d.groupby("bulan")[kol].sum()
        ada = set(df[df["tahun"] == tahun]["bulan"].unique())
        return [float(g.get(b, 0.0)) if b in ada else None for b in range(1, 13)]

    r.per_bulan = per_bulan(k.tahun)
    if k.banding is not None:
        r.per_bulan_banding = per_bulan(k.banding)

    def per_wil(tahun: int) -> dict:
        d = _saring(df, k, tahun, wilayah=False)
        g = d.groupby("wilayah")[kol].sum()
        semua = data.wilayah_ada()
        return {w: float(g.get(w, 0.0)) for w in semua}

    r.per_wilayah = per_wil(k.tahun)
    if k.banding is not None:
        r.per_wilayah_banding = per_wil(k.banding)
    d = _saring(df, k, k.tahun, jenis=False)
    g = d.groupby("jenis")[kol].sum().sort_values(ascending=False)
    r.per_jenis = {j: float(v) for j, v in g.items()}
    return r


def teratas(d: dict, hanya: list[str] | None = None) -> tuple[str, float] | None:
    item = {a: b for a, b in d.items() if (hanya is None or a in hanya)}
    if not item or max(item.values()) <= 0:
        return None
    a = max(item, key=item.get)
    return a, item[a]


def peringkat_wilayah(d: dict, nama: str) -> tuple[int, int] | None:
    """Urutan sebuah wilayah (1 = tertinggi) di antara enam wilayah kerja."""
    item = {a: b for a, b in d.items() if a in WILAYAH}
    if nama not in item or max(item.values(), default=0) <= 0:
        return None
    urut = sorted(item, key=lambda a: -item[a])
    return urut.index(nama) + 1, len(urut)


def label_filter(k: Kondisi) -> dict[str, str]:
    return {
        "metrik": METRIK_NAMA[k.metrik].capitalize(),
        "periode": f"{nama_bulan_rentang(k.bulan_awal, k.bulan_akhir)} {k.tahun}",
        "wilayah": ", ".join(k.wilayah) if k.wilayah else "Seluruh wilayah",
        "jenis": ", ".join(k.jenis) if k.jenis else "Semua jenis",
        "banding": f"vs {k.banding}" if k.banding else "Tanpa perbandingan",
    }


def jelaskan(k: Kondisi, r: Ringkasan, data: DataDashboard, cat: list[str], n: Niat | None = None) -> str:
    """Jawaban teks untuk chat. SEMUA angka dari Excel, dihitung Python."""
    if not r.ada_data:
        return ("Data untuk dashboard belum tersedia. Pastikan file **data_dashboard_semar.xlsx** berada satu folder "
                "dengan aplikasi dan sheet-nya terisi." + ("\n\n" + "\n".join(f"- {g}" for g in data.galat) if data.galat else ""))
    lf = label_filter(k)
    baris = [f"Dashboard disesuaikan: **{lf['metrik'].lower()}**, **{lf['periode']}**, **{lf['wilayah'].lower() if not k.wilayah else lf['wilayah']}**"
             + (f", jenis **{lf['jenis']}**" if k.jenis else "") + "."]
    if k.metrik == "penerimaan":
        baris.append(f"- Total realisasi: **{rp(r.total)}**" +
                     (f" ({persen(r.delta, True)} dibanding {k.banding}: {rp(r.total_banding)})" if r.delta is not None else ""))
        if r.capaian is not None:
            baris.append(f"- Capaian terhadap target: **{persen(r.capaian)}** (target {rp(r.target_total)})")
    else:
        baris.append(f"- Total penindakan: **{angka(r.total)} kali**" +
                     (f" ({persen(r.delta, True)} dibanding {k.banding}: {angka(r.total_banding)} kali)" if r.delta is not None else ""))
        if r.nilai_barang is not None:
            baris.append(f"- Nilai barang hasil penindakan: **{rp(r.nilai_barang)}**")
    tw = teratas(r.per_wilayah, k.wilayah or WILAYAH)
    if len(k.wilayah) == 1:
        pw = peringkat_wilayah(r.per_wilayah, k.wilayah[0])
        if pw:
            baris.append(f"- Peringkat {k.wilayah[0]}: **#{pw[0]} dari {pw[1]}** wilayah kerja")
    elif tw:
        satuan = rp(tw[1]) if k.metrik == "penerimaan" else f"{angka(tw[1])} kali"
        baris.append(f"- Wilayah tertinggi{' (di antara yang dipilih)' if k.wilayah else ''}: **{tw[0]}** ({satuan})")
    tj = teratas(r.per_jenis)
    if tj and r.total > 0 and not (k.jenis and len(k.jenis) == 1):
        baris.append(f"- Jenis terbesar: **{tj[0]}** ({persen(tj[1] / sum(r.per_jenis.values()) * 100)} dari total)")
    if r.total == 0:
        baris.append("- Tidak ada data untuk kombinasi filter ini.")
    for c in cat:
        baris.append(f"- Catatan: {c}")
    if n is not None and n.sebab:
        baris.append("- Data hanya memuat angka, jadi **penyebab** perubahan tidak tersedia di basis pengetahuan SEMAR.")
    if data.ada_contoh:
        baris.append("\n_Perhatian: angka ini masih **data contoh** (bukan data resmi). Ganti isi Database Pengetahuan dengan data resmi Anda._")
    return "\n".join(baris)


TEKS_BANTUAN = """Dashboard bisa diatur lewat chat. Contoh perintah:

- **Jenis data:** "tampilkan penindakan", "tampilkan penerimaan negara"
- **Periode:** "tahun 2025", "triwulan 2", "semester 1", "Maret sampai Juni", "setahun penuh"
- **Perbandingan:** "bandingkan 2026 dengan 2025", "tanpa perbandingan"
- **Wilayah:** "fokus Kendal", "Kendal dan Demak", "Kota Semarang", "semua wilayah"
- **Jenis:** "cukai hasil tembakau", "rokok ilegal", "semua jenis"
- **Grafik:** "tampilkan peringkat wilayah", "komposisi jenis", "tren bulanan"
- **Mulai ulang:** "reset dashboard"

Perintah bisa digabung, misalnya: "penindakan rokok ilegal Kendal triwulan 3 tahun 2025"."""


# ===========================================================================
# TAMPILAN (HTML/SVG). Tema dan warna memakai variabel CSS dari app_web.py
# ===========================================================================
def _p(teks: str) -> str:
    return re.sub(r">\s+<", "><", re.sub(r"[ \t]*\n[ \t]*", " ", teks)).strip()


def _nice(x: float) -> float:
    if x <= 0:
        return 1.0
    e = 10 ** math.floor(math.log10(x))
    for m in (1, 1.5, 2, 2.5, 3, 4, 5, 6, 8, 10):
        if m * e >= x:
            return m * e
    return 10 * e


def _lerp(a: str, b: str, t: float) -> str:
    ca = [int(a[i:i + 2], 16) for i in (1, 3, 5)]
    cb = [int(b[i:i + 2], 16) for i in (1, 3, 5)]
    return "#" + "".join(f"{round(x + (y - x) * t):02X}" for x, y in zip(ca, cb))


def skala_warna(t: float) -> str:
    t = max(0.0, min(1.0, t))
    return _lerp("#DCE9FB", "#4F8FE0", t / 0.5) if t < 0.5 else _lerp("#4F8FE0", "#0A2A63", (t - 0.5) / 0.5)


def _satuan(k: Kondisi, v: float, singkat=False) -> str:
    return rp(v, singkat) if k.metrik == "penerimaan" else f"{angka(v)} kali"


IKON = {
    "koin": '<svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="9" fill="none" stroke="currentColor" stroke-width="2"/><path d="M9 9h4.5a2 2 0 0 1 0 4H9m0 0h5M9 9v7" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"/></svg>',
    "target": '<svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="9" fill="none" stroke="currentColor" stroke-width="2"/><circle cx="12" cy="12" r="5" fill="none" stroke="currentColor" stroke-width="2"/><circle cx="12" cy="12" r="1.6" fill="currentColor"/></svg>',
    "tren": '<svg viewBox="0 0 24 24"><path d="M3 17l6-6 4 4 8-8M15 7h6v6" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"/></svg>',
    "pin": '<svg viewBox="0 0 24 24"><path d="M12 21s7-6.2 7-11.5A7 7 0 0 0 5 9.5C5 14.8 12 21 12 21z" fill="none" stroke="currentColor" stroke-width="2"/><circle cx="12" cy="9.5" r="2.6" fill="currentColor"/></svg>',
    "perisai": '<svg viewBox="0 0 24 24"><path d="M12 3l8 3v6c0 5-3.6 8-8 9-4.4-1-8-4-8-9V6z" fill="none" stroke="currentColor" stroke-width="2" stroke-linejoin="round"/><path d="M8.5 12l2.5 2.5 4.5-5" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>',
    "kotak": '<svg viewBox="0 0 24 24"><path d="M4 8l8-4 8 4v8l-8 4-8-4z M4 8l8 4 8-4 M12 12v8" fill="none" stroke="currentColor" stroke-width="2" stroke-linejoin="round"/></svg>',
}


def html_judul(k: Kondisi, data: DataDashboard) -> str:
    lf = label_filter(k)
    chip = "".join(f'<span class="dsb-chip" style="--i:{i}">{H.escape(v)}</span>' for i, v in enumerate(
        [lf["metrik"], lf["periode"], lf["wilayah"], *( [lf["jenis"]] if k.jenis else [] ), lf["banding"]]))
    contoh = '<span class="dsb-contoh">DATA CONTOH</span>' if data.ada_contoh else ""
    return _p(f"""<div class="dsb-judul"><div><div class="dsb-judul-atas">Dashboard wilayah kerja</div>
<h2>KPPBC Tipe Madya Pabean A Semarang</h2></div>{contoh}</div><div class="dsb-chips">{chip}</div>""")


def html_kpi(k: Kondisi, r: Ringkasan) -> str:
    def delta(d: float | None, bd) -> str:
        if d is None:
            return ""
        arah, kls = ("▲", "naik") if d >= 0 else ("▼", "turun")
        return f'<span class="dsb-delta {kls}">{arah} {persen(abs(d))}</span> <span class="dsb-vs">vs {bd}</span>'

    kartu = []
    tw = teratas(r.per_wilayah, k.wilayah or WILAYAH)
    pw = peringkat_wilayah(r.per_wilayah, k.wilayah[0]) if len(k.wilayah) == 1 else None
    if k.metrik == "penerimaan":
        kartu.append(("koin", "Total realisasi", rp(r.total), delta(r.delta, k.banding) or f"{nama_bulan_rentang(k.bulan_awal, k.bulan_akhir)} {k.tahun}"))
        if r.capaian is not None:
            kartu.append(("target", "Capaian target", persen(r.capaian), f"target {rp(r.target_total, True)}"))
        else:
            kartu.append(("target", "Capaian target", "-", "target belum diisi di Database Pengetahuan"))
        kartu.append(("tren", "Rata-rata per bulan", rp(r.rata_bulanan), f"{len([v for v in r.per_bulan if v is not None])} bulan data"))
    else:
        kartu.append(("perisai", "Total penindakan", f"{angka(r.total)} <small>kali</small>", delta(r.delta, k.banding) or f"{nama_bulan_rentang(k.bulan_awal, k.bulan_akhir)} {k.tahun}"))
        kartu.append(("kotak", "Nilai barang", rp(r.nilai_barang) if r.nilai_barang is not None else "-", "hasil penindakan"))
        kartu.append(("tren", "Rata-rata per bulan", f"{_id(r.rata_bulanan, 1) if r.rata_bulanan is not None else '-'} <small>kali</small>", f"{len([v for v in r.per_bulan if v is not None])} bulan data"))
    if pw:
        kartu.append(("pin", "Peringkat wilayah", f"#{pw[0]} <small>dari {pw[1]}</small>", H.escape(k.wilayah[0]) + " &middot; " + _satuan(k, r.per_wilayah.get(k.wilayah[0], 0), True)))
    elif tw:
        kartu.append(("pin", "Wilayah tertinggi" + (" (terpilih)" if k.wilayah else ""), H.escape(tw[0]), _satuan(k, tw[1], True)))
    else:
        kartu.append(("pin", "Wilayah tertinggi", "-", "tidak ada data"))
    isi = "".join(
        f'<div class="dsb-kpi" style="--i:{i}"><div class="dsb-ikon">{IKON[ik]}</div><div class="dsb-kpi-isi">'
        f'<div class="dsb-kpi-label">{lab}</div><div class="dsb-kpi-nilai">{nilai}</div><div class="dsb-kpi-sub">{sub}</div></div></div>'
        for i, (ik, lab, nilai, sub) in enumerate(kartu))
    return _p(f'<div class="dsb-kpi-grid">{isi}</div>')


GESER_LABEL = {"Kab. Semarang": (332, 352), "Kota Salatiga": (398, 470)}


def html_peta(k: Kondisi, r: Ringkasan) -> str:
    w, h = P.VIEWBOX[2], P.VIEWBOX[3]
    nilai = {n: r.per_wilayah.get(n, 0.0) for n in WILAYAH}
    maks = max(nilai.values()) if nilai else 0
    aktif = set(k.wilayah)
    bagian = [f'<svg viewBox="0 0 {w} {h}" class="dsb-svg-peta" xmlns="http://www.w3.org/2000/svg" role="img" aria-label="Peta wilayah kerja Bea Cukai Semarang">',
              f'<defs><filter id="dsb-glow" x="-20%" y="-20%" width="140%" height="140%"><feDropShadow dx="0" dy="0" stdDeviation="5" flood-color="#F58220" flood-opacity=".85"/></filter></defs>']
    for nama, v in P.PETA.items():
        if not v["kerja"]:
            bagian.append(f'<path d="{v["d"]}" class="dsb-tetangga"/>'
                          f'<text x="{v["x"]}" y="{v["y"]}" class="dsb-nama-tetangga">{H.escape(nama)}</text>')
    urut = sorted(WILAYAH, key=lambda n: n in aktif)   # yang dipilih digambar terakhir (di atas)
    for i, nama in enumerate(urut):
        v = P.PETA[nama]
        t = (nilai[nama] / maks) if maks > 0 else 0
        kelas = "dsb-wil"
        if aktif:
            kelas += " pilih" if nama in aktif else " redup"
        tip = f"{nama}: {_satuan(k, nilai[nama])}"
        bagian.append(f'<path d="{v["d"]}" class="{kelas}" style="fill:{skala_warna(t)};--d:{i * 70}ms"><title>{H.escape(tip)}</title></path>')
    for nama in WILAYAH:
        v = P.PETA[nama]
        x, y = GESER_LABEL.get(nama, (v["x"], v["y"]))
        redup = " redup" if aktif and nama not in aktif else ""
        bagian.append(f'<g class="dsb-label{redup}"><text x="{x}" y="{y}" class="dsb-l-nama">{H.escape(nama)}</text>'
                      f'<text x="{x}" y="{y + 23}" class="dsb-l-nilai">{H.escape(_satuan(k, nilai[nama], True))}</text></g>')
    bagian.append(f'<g transform="translate({w - 44} 46)" class="dsb-utara"><path d="M0 -26 L9 6 L0 0 L-9 6 Z" fill="#0A1F44"/><text y="24" text-anchor="middle">U</text></g>')
    bagian.append("</svg>")
    rendah, tinggi = (rp(0, True) if k.metrik == "penerimaan" else "0"), _satuan(k, maks, True)
    legenda = (f'<div class="dsb-legenda"><span>{rendah}</span><i></i><span>{H.escape(tinggi)}</span></div>')
    lf = label_filter(k)
    return _p(f"""<div class="dsb-kartu dsb-peta-kartu"><div class="dsb-kartu-judul"><b>Peta sebaran</b>
<span>{H.escape(lf['metrik'])} &middot; {H.escape(lf['periode'])}</span></div>{''.join(bagian)}{legenda}
<div class="dsb-sumber">{H.escape(P.SUMBER_PETA)}</div></div>""")


def html_tren(k: Kondisi, r: Ringkasan) -> str:
    W, Ht, ml, mr, mt, mb = 760, 280, 62, 18, 28, 36
    a, b = r.per_bulan, r.per_bulan_banding
    semua = [v for v in a + (b or []) if v is not None]
    tinggi = _nice(max(semua) * 1.05) if semua and max(semua) > 0 else 1
    lebar = W - ml - mr
    xs = [ml + i * lebar / 11 for i in range(12)]
    y = lambda v: mt + (Ht - mt - mb) * (1 - v / tinggi)
    s = [f'<svg viewBox="0 0 {W} {Ht}" class="dsb-svg-grafik" xmlns="http://www.w3.org/2000/svg" role="img" aria-label="Tren bulanan">',
         '<defs><linearGradient id="dsb-area" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#F58220" stop-opacity=".45"/><stop offset="1" stop-color="#F58220" stop-opacity="0"/></linearGradient>'
         '<linearGradient id="dsb-garis" x1="0" y1="0" x2="1" y2="0"><stop offset="0" stop-color="#F7B733"/><stop offset="1" stop-color="#F58220"/></linearGradient></defs>']
    x0 = xs[k.bulan_awal - 1] - lebar / 22
    x1 = xs[k.bulan_akhir - 1] + lebar / 22
    s.append(f'<rect x="{x0:.1f}" y="{mt}" width="{x1 - x0:.1f}" height="{Ht - mt - mb}" class="dsb-rentang"/>')
    for i in range(5):
        v = tinggi * i / 4
        yy = y(v)
        s.append(f'<line x1="{ml}" x2="{W - mr}" y1="{yy:.1f}" y2="{yy:.1f}" class="dsb-grid"/>'
                 f'<text x="{ml - 8}" y="{yy + 4:.1f}" class="dsb-sumbu" text-anchor="end">{H.escape(rp(v, True) if k.metrik == "penerimaan" else angka(v))}</text>')
    for i, bl in enumerate(BULAN_PENDEK):
        s.append(f'<text x="{xs[i]:.1f}" y="{Ht - 12}" class="dsb-sumbu" text-anchor="middle">{bl}</text>')

    def segmen(vals):
        out, cur = [], []
        for i, v in enumerate(vals):
            if v is None:
                if cur:
                    out.append(cur)
                cur = []
            else:
                cur.append((xs[i], y(v), i, v))
        if cur:
            out.append(cur)
        return out

    if b:
        for sg in segmen(b):
            d = "M" + "L".join(f"{px:.1f} {py:.1f}" for px, py, _, _ in sg)
            s.append(f'<path d="{d}" class="dsb-banding"/>')
            for px, py, i, v in sg:
                s.append(f'<circle cx="{px:.1f}" cy="{py:.1f}" r="3" class="dsb-titik-banding"><title>{BULAN_PENDEK[i]} {k.banding}: {H.escape(_satuan(k, v))}</title></circle>')
    for sg in segmen(a):
        d = "M" + "L".join(f"{px:.1f} {py:.1f}" for px, py, _, _ in sg)
        dasar = y(0)
        s.append(f'<path d="{d}L{sg[-1][0]:.1f} {dasar:.1f}L{sg[0][0]:.1f} {dasar:.1f}Z" fill="url(#dsb-area)" class="dsb-area"/>')
        s.append(f'<path d="{d}" class="dsb-garis" pathLength="1"/>')
        for px, py, i, v in sg:
            s.append(f'<circle cx="{px:.1f}" cy="{py:.1f}" r="4.5" class="dsb-titik"><title>{BULAN_PENDEK[i]} {k.tahun}: {H.escape(_satuan(k, v))}</title></circle>')
    s.append("</svg>")
    leg = f'<span class="dsb-leg-a">{k.tahun}</span>' + (f'<span class="dsb-leg-b">{k.banding}</span>' if k.banding else "")
    sub = "Bulan terpilih diberi bayangan." + (f" Data {k.tahun} tersedia sampai {BULAN[r.bulan_terakhir - 1]}." if r.bulan_terakhir < 12 else "")
    return _p(f'<div class="dsb-kartu"><div class="dsb-kartu-judul"><b>Tren bulanan</b><span class="dsb-legenda-garis">{leg}</span></div>{"".join(s)}<div class="dsb-sumber">{H.escape(sub)}</div></div>')


PALET = ["#F58220", "#1B4F9C", "#F7B733", "#2F80ED", "#C9622A", "#0A2A63", "#7A4A21", "#5E9CE6"]


def _batang(items: list[tuple[str, float]], k: Kondisi, pilih: set[str], total: float, banding: dict | None, warna_per_baris: bool) -> str:
    if not items:
        return '<div class="dsb-kosong">Tidak ada data untuk filter ini.</div>'
    mx = max(v for _, v in items) or 1
    baris = []
    for i, (nama, v) in enumerate(items):
        lebar = v / mx * 100
        redup = " redup" if pilih and nama not in pilih else ""
        warna = PALET[i % len(PALET)] if warna_per_baris else "#1B4F9C"
        bd = ""
        if banding is not None and banding.get(nama):
            d = (v / banding[nama] - 1) * 100
            bd = f'<span class="dsb-delta {"naik" if d >= 0 else "turun"} kecil">{"▲" if d >= 0 else "▼"} {persen(abs(d))}</span>'
        pct = f'<span class="dsb-pct">{persen(v / total * 100)}</span>' if total > 0 else ""
        baris.append(
            f'<div class="dsb-baris{redup}" style="--i:{i}"><div class="dsb-baris-no">{i + 1}</div>'
            f'<div class="dsb-baris-isi"><div class="dsb-baris-atas"><span class="dsb-baris-nama">{H.escape(nama)}</span>'
            f'<span class="dsb-baris-nilai">{H.escape(_satuan(k, v))} {pct} {bd}</span></div>'
            f'<div class="dsb-trek"><div class="dsb-isi" style="width:{lebar:.1f}%;background:linear-gradient(90deg,{warna},{_lerp(warna, "#FFFFFF", .35)})"></div></div></div></div>')
    return "".join(baris)


def html_komposisi(k: Kondisi, r: Ringkasan) -> str:
    items = list(r.per_jenis.items())
    total = sum(v for _, v in items)
    return _p(f'<div class="dsb-kartu"><div class="dsb-kartu-judul"><b>Komposisi jenis</b><span>menurut {"nilai realisasi" if k.metrik == "penerimaan" else "jumlah penindakan"}</span></div>'
              f'{_batang(items, k, set(k.jenis), total, None, True)}</div>')


def html_peringkat(k: Kondisi, r: Ringkasan) -> str:
    items = sorted(r.per_wilayah.items(), key=lambda x: -x[1])
    items = [(a, b) for a, b in items if a in WILAYAH or b > 0]
    total = sum(v for _, v in items)
    return _p(f'<div class="dsb-kartu"><div class="dsb-kartu-judul"><b>Peringkat wilayah</b><span>{"▲▼ = perubahan vs " + str(k.banding) if k.banding else "tanpa perbandingan"}</span></div>'
              f'{_batang(items, k, set(k.wilayah), total, r.per_wilayah_banding, False)}</div>')


def html_grafik(k: Kondisi, r: Ringkasan) -> str:
    return {"tren": html_tren, "komposisi": html_komposisi, "peringkat": html_peringkat}[k.grafik](k, r)


# ===========================================================================
# CSS DASHBOARD
# ===========================================================================
CSS = """
.dsb-judul{display:flex;justify-content:space-between;align-items:flex-end;gap:12px;margin:.1rem 0 .5rem 0;flex-wrap:wrap;}
.dsb-judul-atas{font-size:.74rem;letter-spacing:.16em;text-transform:uppercase;font-weight:700;color:#C9622A;}
.dsb-judul h2{font-family:'Playfair Display',Georgia,serif!important;font-size:1.45rem;margin:0;padding:0;color:var(--navy);line-height:1.15;}
.dsb-contoh{font-size:.72rem;font-weight:800;letter-spacing:.12em;padding:6px 12px;border-radius:999px;color:#fff;background:linear-gradient(90deg,#C0392B,#F58220);box-shadow:0 8px 18px -8px rgba(192,57,43,.7);animation:denyut 3s ease-in-out infinite;}
.dsb-chips{display:flex;flex-wrap:wrap;gap:7px;margin-bottom:.9rem;}
.dsb-chip{font-size:.78rem;font-weight:600;padding:5px 12px;border-radius:999px;color:var(--navy);background:linear-gradient(90deg,rgba(27,79,156,.10),rgba(245,130,32,.16));border:1px solid rgba(27,79,156,.2);animation:muncul .45s ease both;animation-delay:calc(var(--i)*60ms);}
.dsb-kpi-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(215px,1fr));gap:12px;margin-bottom:1rem;}
.dsb-kpi{display:flex;gap:11px;align-items:center;padding:14px 14px;border-radius:20px;background:linear-gradient(145deg,#fff,#FFF4E2);border:1px solid rgba(10,31,68,.07);
  box-shadow:0 14px 30px -18px rgba(10,31,68,.5);animation:muncul .5s ease both;animation-delay:calc(var(--i)*80ms);transition:transform .2s,box-shadow .2s;}
.dsb-kpi:hover{transform:translateY(-4px);box-shadow:0 20px 34px -16px rgba(245,130,32,.55);}
.dsb-ikon{width:40px;height:40px;flex:0 0 auto;border-radius:13px;display:flex;align-items:center;justify-content:center;color:#fff;background:linear-gradient(135deg,#1B4F9C,#0A1F44);}
.dsb-kpi:nth-child(odd) .dsb-ikon{background:linear-gradient(135deg,#F58220,#C9622A);}
.dsb-ikon svg{width:22px;height:22px;}
.dsb-kpi-isi{min-width:0;}
.dsb-kpi-label{font-size:.68rem;font-weight:700;letter-spacing:.05em;text-transform:uppercase;color:#7A4A21;}
.dsb-kpi-nilai{font-size:1.15rem;font-weight:800;color:var(--navy);line-height:1.15;word-break:break-word;}
.dsb-kpi-nilai small{font-size:.7rem;font-weight:600;opacity:.7;}
.dsb-kpi-sub{font-size:.72rem;color:#5C6785;margin-top:2px;}
.dsb-delta{font-weight:800;font-size:.74rem;padding:1px 7px;border-radius:999px;}
.dsb-delta.naik{color:#0E7A43;background:rgba(14,122,67,.12);} .dsb-delta.turun{color:#B3261E;background:rgba(179,38,30,.12);} .dsb-delta.kecil{font-size:.66rem;}
.dsb-vs{font-size:.7rem;color:#5C6785;}
.dsb-kartu{background:rgba(255,255,255,.88);backdrop-filter:blur(6px);border:1px solid rgba(10,31,68,.07);border-radius:24px;padding:16px 18px 12px 18px;margin-bottom:1rem;box-shadow:0 18px 40px -22px rgba(10,31,68,.5);animation:muncul .55s ease both;}
.dsb-kartu-judul{display:flex;justify-content:space-between;align-items:baseline;gap:10px;margin-bottom:.3rem;flex-wrap:wrap;}
.dsb-kartu-judul b{font-family:'Playfair Display',Georgia,serif;font-size:1.15rem;color:var(--navy);}
.dsb-kartu-judul span{font-size:.76rem;color:#5C6785;}
.dsb-svg-peta{width:100%;height:auto;display:block;border-radius:16px;background:linear-gradient(180deg,#EAF2FD,#F6FAFF);}
.dsb-tetangga{fill:#DDE3EE;stroke:#fff;stroke-width:1.2;}
.dsb-nama-tetangga{font-size:15px;fill:#8C97B2;text-anchor:middle;font-weight:600;}
.dsb-wil{stroke:#fff;stroke-width:2;stroke-linejoin:round;cursor:pointer;transition:opacity .3s,filter .25s,stroke .25s;animation:muncul .7s ease both;animation-delay:var(--d);}
.dsb-wil:hover{stroke:#F58220;stroke-width:3.5;filter:brightness(1.1);}
.dsb-wil.pilih{stroke:#F58220;stroke-width:4;filter:url(#dsb-glow);}
.dsb-wil.redup{opacity:.35!important;}
.dsb-label{pointer-events:none;transition:opacity .3s;} .dsb-label.redup{opacity:.45!important;}
.dsb-l-nama{font-size:19px;font-weight:800;fill:#fff;stroke:#0A1F44;stroke-width:4.4px;paint-order:stroke;stroke-linejoin:round;text-anchor:middle;}
.dsb-l-nilai{font-size:16.5px;font-weight:700;fill:#FFE29A;stroke:#0A1F44;stroke-width:3.8px;paint-order:stroke;stroke-linejoin:round;text-anchor:middle;}
.dsb-utara text{font-size:13px;font-weight:800;fill:#0A1F44;}
.dsb-legenda{display:flex;align-items:center;gap:10px;font-size:.74rem;color:#3A4668;margin:.5rem 0 .2rem 0;}
.dsb-legenda i{flex:1;height:10px;border-radius:999px;background:linear-gradient(90deg,#DCE9FB,#4F8FE0,#0A2A63);}
.dsb-sumber{font-size:.68rem;color:#7C86A3;line-height:1.4;margin-top:.3rem;}
.dsb-svg-grafik{width:100%;height:auto;display:block;}
.dsb-grid{stroke:rgba(10,31,68,.09);stroke-width:1;} .dsb-sumbu{font-size:11px;fill:#5C6785;}
.dsb-rentang{fill:rgba(245,130,32,.09);}
.dsb-garis{fill:none;stroke:url(#dsb-garis);stroke-width:3.5;stroke-linecap:round;stroke-linejoin:round;stroke-dasharray:1;stroke-dashoffset:1;animation:gambar 1.4s ease forwards;}
.dsb-banding{fill:none;stroke:#1B4F9C;stroke-width:2.2;stroke-dasharray:6 5;opacity:.75;animation:muncul .8s ease both;}
.dsb-area{animation:muncul 1.2s ease both;}
.dsb-titik{fill:#fff;stroke:#F58220;stroke-width:2.5;cursor:pointer;transition:r .15s;animation:muncul 1.4s ease both;} .dsb-titik:hover{r:7;}
.dsb-titik-banding{fill:#1B4F9C;opacity:.8;}
.dsb-legenda-garis span{display:inline-flex;align-items:center;gap:6px;margin-left:12px;font-weight:700;}
.dsb-leg-a:before,.dsb-leg-b:before{content:"";width:22px;height:0;border-top:3.5px solid #F58220;border-radius:3px;}
.dsb-leg-b:before{border-top:3px dashed #1B4F9C;}
.dsb-baris{display:flex;gap:10px;align-items:center;margin:.55rem 0;animation:muncul .5s ease both;animation-delay:calc(var(--i)*70ms);transition:opacity .3s;} .dsb-baris.redup{opacity:.42!important;}
.dsb-baris-no{width:24px;height:24px;flex:0 0 auto;border-radius:50%;background:linear-gradient(135deg,#0A1F44,#1B4F9C);color:#fff;font-size:.72rem;font-weight:800;display:flex;align-items:center;justify-content:center;}
.dsb-baris-isi{flex:1;min-width:0;}
.dsb-baris-atas{display:flex;justify-content:space-between;gap:8px;font-size:.82rem;flex-wrap:wrap;}
.dsb-baris-nama{font-weight:700;color:var(--navy);} .dsb-baris-nilai{color:#3A4668;} .dsb-pct{font-weight:700;color:#C9622A;margin-left:4px;}
.dsb-trek{height:11px;border-radius:999px;background:rgba(10,31,68,.08);margin-top:4px;overflow:hidden;}
.dsb-isi{height:100%;border-radius:999px;animation:tumbuh 1s cubic-bezier(.2,.8,.2,1) both;animation-delay:calc(var(--i)*80ms);}
.dsb-kosong{padding:1.4rem;text-align:center;color:#5C6785;}
.panel-chat{position:sticky;top:3.2rem;}
@keyframes gambar{to{stroke-dashoffset:0}}
@keyframes tumbuh{from{width:0}}
"""
