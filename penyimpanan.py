"""
Penyimpanan 7 hari untuk Pemeriksaan Barang: foto/file dokumentasi dan baris sheet hasil.

Dua pilihan penyimpanan, memakai kode yang sama:
  - SQLite (bawaan): satu file database di folder data_semar/. Cocok untuk dijalankan di komputer sendiri.
  - Postgres: bila DATABASE_URL diisi (mis. dari Neon). Wajib untuk server seperti Streamlit Cloud,
    karena file lokal di sana hilang setiap aplikasi restart atau di-deploy ulang.

Semua yang lebih tua dari MASA_SIMPAN_HARI dihapus otomatis oleh bersihkan().
Berkas ini tidak memakai Streamlit.
"""

from __future__ import annotations

import base64
import io
import json
import os
import time
from dataclasses import asdict, is_dataclass
from datetime import time as jam
from pathlib import Path

MASA_SIMPAN_HARI = 7
MAKS_BYTE_BERKAS = 20 * 1024 * 1024     # PDF/berkas lebih besar dari ini tidak disimpan
SISI_SIMPAN = 2400                       # foto disimpan dengan sisi terpanjang segini (kualitas cukup untuk dokumentasi)
SISI_MINI = 96                           # gambar kecil di dalam sel sheet
PREFIKS = "semar_"                       # awalan nama tabel, agar aman bila database dipakai bersama aplikasi lain

_SKEMA_SQLITE = [
    f"CREATE TABLE IF NOT EXISTS {PREFIKS}baris (id INTEGER PRIMARY KEY AUTOINCREMENT, isi TEXT NOT NULL, "
    "dibuat REAL NOT NULL, diubah REAL NOT NULL)",
    f"CREATE TABLE IF NOT EXISTS {PREFIKS}berkas (id INTEGER PRIMARY KEY AUTOINCREMENT, baris_id INTEGER, slot TEXT, "
    "nama TEXT, mime TEXT, ukuran INTEGER, mini TEXT, data BLOB, dibuat REAL NOT NULL)",
    f"CREATE INDEX IF NOT EXISTS {PREFIKS}berkas_baris ON {PREFIKS}berkas (baris_id)",
    f"CREATE TABLE IF NOT EXISTS {PREFIKS}pengguna (id INTEGER PRIMARY KEY AUTOINCREMENT, nip TEXT NOT NULL UNIQUE, "
    "nama TEXT NOT NULL, jabatan TEXT, satker TEXT, seksi TEXT, sandi TEXT NOT NULL, peran TEXT NOT NULL DEFAULT 'pengguna', "
    "status TEXT NOT NULL DEFAULT 'menunggu', gagal INTEGER NOT NULL DEFAULT 0, terkunci_sampai REAL NOT NULL DEFAULT 0, "
    "dibuat REAL NOT NULL, masuk_terakhir REAL)",
]
_SKEMA_POSTGRES = [
    f"CREATE TABLE IF NOT EXISTS {PREFIKS}baris (id BIGSERIAL PRIMARY KEY, isi TEXT NOT NULL, "
    "dibuat DOUBLE PRECISION NOT NULL, diubah DOUBLE PRECISION NOT NULL)",
    f"CREATE TABLE IF NOT EXISTS {PREFIKS}berkas (id BIGSERIAL PRIMARY KEY, baris_id BIGINT, slot TEXT, "
    "nama TEXT, mime TEXT, ukuran INTEGER, mini TEXT, data BYTEA, dibuat DOUBLE PRECISION NOT NULL)",
    f"CREATE INDEX IF NOT EXISTS {PREFIKS}berkas_baris ON {PREFIKS}berkas (baris_id)",
    f"CREATE TABLE IF NOT EXISTS {PREFIKS}pengguna (id BIGSERIAL PRIMARY KEY, nip TEXT NOT NULL UNIQUE, "
    "nama TEXT NOT NULL, jabatan TEXT, satker TEXT, seksi TEXT, sandi TEXT NOT NULL, peran TEXT NOT NULL DEFAULT 'pengguna', "
    "status TEXT NOT NULL DEFAULT 'menunggu', gagal INTEGER NOT NULL DEFAULT 0, terkunci_sampai DOUBLE PRECISION NOT NULL DEFAULT 0, "
    "dibuat DOUBLE PRECISION NOT NULL, masuk_terakhir DOUBLE PRECISION)",
]
_KOLOM_META = "id, baris_id, slot, nama, mime, ukuran, mini, dibuat"     # tanpa 'data' agar ringan


class GalatPenyimpanan(Exception):
    """Galat yang pesannya aman ditampilkan kepada pengguna."""


class Penyimpanan:
    def __init__(self, url: str | None = None):
        url = (url or "").strip()
        if url:
            if url.startswith("postgres://"):
                url = "postgresql://" + url[len("postgres://"):]
            if not url.startswith("postgresql://"):
                raise GalatPenyimpanan("DATABASE_URL harus diawali postgresql:// (salin dari halaman Connect di Neon).")
            self.jenis, self._dsn, self._ph = "postgresql", url, "%s"
        else:
            folder = Path(os.getenv("SEMAR_DATA") or Path(__file__).resolve().parent / "data_semar")
            try:
                folder.mkdir(parents=True, exist_ok=True)
            except OSError:                       # folder aplikasi tidak bisa ditulis (mis. di server): pakai folder sementara
                import tempfile
                folder = Path(tempfile.gettempdir()) / "data_semar"
                folder.mkdir(parents=True, exist_ok=True)
            self.jenis, self._dsn, self._ph = "sqlite", str(folder / "dokumentasi.db"), "?"
        self._siapkan()

    # ------------------------------------------------------------------ dasar
    def _buka(self):
        if self.jenis == "sqlite":
            import sqlite3
            c = sqlite3.connect(self._dsn, timeout=30, check_same_thread=False)
            c.row_factory = sqlite3.Row
            return c
        try:
            import psycopg
            from psycopg.rows import dict_row
        except ImportError as e:
            raise GalatPenyimpanan("Paket psycopg belum terpasang. Tambahkan 'psycopg[binary]' pada requirements.txt.") from e
        return psycopg.connect(self._dsn, row_factory=dict_row, connect_timeout=20)

    def _q(self, sql: str, params: tuple = (), hasil: str | None = None):
        """Jalankan satu perintah SQL pada koneksi sekali pakai. hasil: 'semua' | 'satu' | None."""
        sql = sql.replace("?", self._ph)
        c = self._buka()
        try:
            cur = c.cursor()
            cur.execute(sql, params)
            out = cur.fetchall() if hasil == "semua" else cur.fetchone() if hasil == "satu" else None
            c.commit()
            return out
        finally:
            c.close()

    def _siapkan(self) -> None:
        for sql in (_SKEMA_SQLITE if self.jenis == "sqlite" else _SKEMA_POSTGRES):
            self._q(sql)

    @staticmethod
    def _kamus(r) -> dict:
        return {k: r[k] for k in r.keys()}

    # ------------------------------------------------------------------ baris sheet
    def simpan_baris(self, isi: str) -> int:
        t = time.time()
        r = self._q(f"INSERT INTO {PREFIKS}baris (isi, dibuat, diubah) VALUES (?, ?, ?) RETURNING id", (isi, t, t), "satu")
        return int(r["id"])

    def perbarui_baris(self, id_: int, isi: str) -> None:
        self._q(f"UPDATE {PREFIKS}baris SET isi = ?, diubah = ? WHERE id = ?", (isi, time.time(), id_))

    def hapus_baris(self, id_: int) -> None:
        self._q(f"DELETE FROM {PREFIKS}berkas WHERE baris_id = ?", (id_,))
        self._q(f"DELETE FROM {PREFIKS}baris WHERE id = ?", (id_,))

    def hapus_semua(self) -> None:
        self._q(f"DELETE FROM {PREFIKS}berkas")
        self._q(f"DELETE FROM {PREFIKS}baris")

    def muat_baris(self) -> list[dict]:
        r = self._q(f"SELECT id, isi, dibuat FROM {PREFIKS}baris ORDER BY id", (), "semua")
        return [self._kamus(x) for x in r]

    # ------------------------------------------------------------------ berkas dokumentasi
    def simpan_berkas(self, baris_id: int, slot: str, nama: str, mime: str, data: bytes, mini: str = "") -> dict:
        t = time.time()
        r = self._q(f"INSERT INTO {PREFIKS}berkas (baris_id, slot, nama, mime, ukuran, mini, data, dibuat) "
                    "VALUES (?, ?, ?, ?, ?, ?, ?, ?) RETURNING id", (baris_id, slot, nama, mime, len(data), mini, data, t), "satu")
        return {"id": int(r["id"]), "baris_id": baris_id, "slot": slot, "nama": nama, "mime": mime,
                "ukuran": len(data), "mini": mini, "dibuat": t}

    def meta_berkas(self, baris_ids: list[int]) -> dict[int, list[dict]]:
        if not baris_ids:
            return {}
        ph = ", ".join("?" for _ in baris_ids)
        r = self._q(f"SELECT {_KOLOM_META} FROM {PREFIKS}berkas WHERE baris_id IN ({ph}) ORDER BY id", tuple(baris_ids), "semua")
        hasil: dict[int, list[dict]] = {}
        for x in r:
            m = self._kamus(x)
            hasil.setdefault(int(m["baris_id"]), []).append(m)
        return hasil

    def ambil_berkas(self, id_: int) -> dict | None:
        r = self._q(f"SELECT {_KOLOM_META}, data FROM {PREFIKS}berkas WHERE id = ?", (id_,), "satu")
        if r is None:
            return None
        m = self._kamus(r)
        m["data"] = bytes(m["data"])
        return m

    # ------------------------------------------------------------------ pemeliharaan
    def bersihkan(self) -> None:
        """Hapus baris dan berkas yang lebih tua dari MASA_SIMPAN_HARI."""
        batas = time.time() - MASA_SIMPAN_HARI * 86400
        self._q(f"DELETE FROM {PREFIKS}berkas WHERE dibuat < ?", (batas,))
        self._q(f"DELETE FROM {PREFIKS}baris WHERE dibuat < ?", (batas,))
        self._q(f"DELETE FROM {PREFIKS}berkas WHERE baris_id NOT IN (SELECT id FROM {PREFIKS}baris)")

    def ringkas(self) -> dict:
        n_baris = self._q(f"SELECT COUNT(*) AS n FROM {PREFIKS}baris", (), "satu")["n"]
        r = self._q(f"SELECT COUNT(*) AS n, COALESCE(SUM(ukuran), 0) AS total FROM {PREFIKS}berkas", (), "satu")
        return {"jenis": self.jenis, "baris": int(n_baris), "berkas": int(r["n"]), "byte": int(r["total"])}


# ===========================================================================
# Persiapan berkas sebelum disimpan
# ===========================================================================
def siapkan_untuk_simpan(nama: str, data: bytes) -> tuple[str, str, bytes, str]:
    """Kembalikan (nama, mime, data, mini_base64). Foto dikecilkan ke JPEG (sisi terpanjang SISI_SIMPAN) dan diberi gambar
    kecil untuk sel sheet. PDF disimpan apa adanya."""
    if data[:5] == b"%PDF-":
        return nama, "application/pdf", data, ""
    try:
        from PIL import Image, ImageOps
        im = Image.open(io.BytesIO(data))
        asli_cukup = (im.format == "JPEG" and max(im.size) <= SISI_SIMPAN and im.getexif().get(274) in (None, 1))
        im = ImageOps.exif_transpose(im)
        if im.mode in ("RGBA", "LA", "P"):
            im = im.convert("RGBA")
            latar = Image.new("RGB", im.size, "white")
            latar.paste(im, mask=im.split()[3])
            im = latar
        else:
            im = im.convert("RGB")
        besar = im.copy()
        besar.thumbnail((SISI_SIMPAN, SISI_SIMPAN), Image.LANCZOS)
        buf = io.BytesIO()
        besar.save(buf, "JPEG", quality=88, optimize=True)
        kecil = im.copy()
        kecil.thumbnail((SISI_MINI, SISI_MINI), Image.LANCZOS)
        buf2 = io.BytesIO()
        kecil.save(buf2, "JPEG", quality=70)
        mini = base64.b64encode(buf2.getvalue()).decode()
        if asli_cukup:                       # sudah JPEG dan tidak terlalu besar: simpan apa adanya (tanpa kompres ulang)
            return nama, "image/jpeg", data, mini
        return Path(nama).stem + ".jpg", "image/jpeg", buf.getvalue(), mini
    except Exception:       # noqa: BLE001  (bukan gambar yang dikenal: simpan apa adanya)
        import mimetypes
        return nama, mimetypes.guess_type(nama)[0] or "application/octet-stream", data, ""


def simpan_berkas_baris(db: Penyimpanan, baris_id: int, daftar: list[tuple[str, str, bytes]]) -> tuple[list[dict], list[str]]:
    """Simpan banyak berkas [(slot, nama, data)] untuk satu baris. Mengembalikan (meta berkas, peringatan)."""
    meta, peringatan = [], []
    for slot, nama, data in daftar:
        if len(data) > MAKS_BYTE_BERKAS:
            peringatan.append(f"{nama} tidak disimpan di dokumentasi karena lebih dari {MAKS_BYTE_BERKAS // (1024 * 1024)} MB.")
            continue
        n, mime, isi, mini = siapkan_untuk_simpan(nama, data)
        try:
            meta.append(db.simpan_berkas(baris_id, slot, n, mime, isi, mini))
        except Exception as e:      # noqa: BLE001
            peringatan.append(f"{nama} gagal disimpan di dokumentasi ({type(e).__name__}).")
    return meta, peringatan


# ===========================================================================
# Baris sheet <-> JSON
# ===========================================================================
def _enkode(o):
    if isinstance(o, jam):
        return {"__jam__": o.strftime("%H:%M:%S")}
    if is_dataclass(o):
        return {"__pengaturan__": asdict(o)}
    return str(o)


def _dekode(d: dict):
    if "__jam__" in d:
        h, m, *s = (int(x) for x in d["__jam__"].split(":"))
        return jam(h, m, s[0] if s else 0)
    if "__pengaturan__" in d:
        import pemeriksaan as PM
        nilai = {k: v for k, v in d["__pengaturan__"].items() if k in PM.Pengaturan.__dataclass_fields__}
        return PM.Pengaturan(**nilai)
    return d


def baris_ke_json(b: dict) -> str:
    """Serialisasi baris untuk database. Foto (b64) dan id/berkas tidak ikut: foto ada di tabel berkas."""
    salin = {k: v for k, v in b.items() if k not in ("id", "berkas")}
    ctx = salin.get("_ctx")
    if isinstance(ctx, dict):
        salin["_ctx"] = {"ocr": ctx.get("ocr", []), "sumber": ctx.get("sumber", []), "gambar": [],
                         "pengaturan": ctx.get("pengaturan")}
    return json.dumps(salin, default=_enkode, ensure_ascii=False)


def baris_dari_json(isi: str, id_: int, meta: list[dict] | None = None) -> dict:
    b = json.loads(isi, object_hook=_dekode)
    b["id"] = int(id_)
    b["berkas"] = meta or []
    return b
