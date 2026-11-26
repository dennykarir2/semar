"""
Akun pengguna SEMAR: pendaftaran, masuk (login), persetujuan admin, dan pengelolaan akun.

  - Pengguna masuk dengan NIP (18 digit) dan kata sandi.
  - Pendaftaran meminta: Nama, NIP, Jabatan, Satker, Seksi, dan kata sandi. Akun baru berstatus "menunggu" sampai admin
    menyetujuinya. Admin pertama ditentukan oleh SEMAR_ADMIN_NIP (di server), atau pendaftar pertama bila dijalankan di
    komputer sendiri.
  - Kata sandi disimpan sebagai hash PBKDF2-SHA256 dengan garam acak (tidak pernah sebagai teks).
  - 5 kali salah berturut-turut mengunci akun selama 15 menit.

Berkas ini tidak memakai Streamlit. Tabel akun dibuat oleh penyimpanan.py.
"""

from __future__ import annotations

import base64
import hashlib
import hmac
import os
import re
import secrets
import time

from penyimpanan import PREFIKS, Penyimpanan

ITERASI = 600_000
MAKS_GAGAL = 5
KUNCI_MENIT = 15
MIN_SANDI = 8
PANJANG_NIP = 18
TABEL = f"{PREFIKS}pengguna"
KOLOM_PUBLIK = "id, nip, nama, jabatan, satker, seksi, peran, status, dibuat, masuk_terakhir"


# ===========================================================================
# Kata sandi
# ===========================================================================
def hash_sandi(sandi: str, garam: bytes | None = None, iterasi: int = ITERASI) -> str:
    garam = garam or os.urandom(16)
    h = hashlib.pbkdf2_hmac("sha256", sandi.encode("utf-8"), garam, iterasi)
    return f"pbkdf2_sha256${iterasi}${base64.b64encode(garam).decode()}${base64.b64encode(h).decode()}"


def cek_sandi(sandi: str, tersimpan: str) -> bool:
    try:
        skema, iterasi, garam, h = tersimpan.split("$")
        if skema != "pbkdf2_sha256":
            return False
        baru = hashlib.pbkdf2_hmac("sha256", sandi.encode("utf-8"), base64.b64decode(garam), int(iterasi))
        return hmac.compare_digest(baru, base64.b64decode(h))
    except (ValueError, TypeError):
        return False


_HASH_PALSU = None


def _hash_palsu() -> str:
    """Hash untuk NIP yang tidak ada, agar waktu respons tidak membocorkan apakah NIP terdaftar."""
    global _HASH_PALSU
    if _HASH_PALSU is None:
        _HASH_PALSU = hash_sandi("tidak-dipakai")
    return _HASH_PALSU


def sandi_sementara() -> str:
    return secrets.token_urlsafe(9)


# ===========================================================================
# Validasi
# ===========================================================================
def bersihkan_nip(teks: str) -> str:
    """Ambil angkanya saja ('19900101 201501 1 001' -> '199001012015011001')."""
    return re.sub(r"\D", "", teks or "")


def _spasi(teks: str) -> str:
    return re.sub(r"\s+", " ", teks or "").strip()


def validasi_daftar(nama: str, nip: str, jabatan: str, satker: str, seksi: str, sandi: str, ulang: str) -> list[str]:
    galat = []
    if len(_spasi(nama)) < 3:
        galat.append("Nama wajib diisi.")
    n = bersihkan_nip(nip)
    if len(n) != PANJANG_NIP:
        galat.append(f"NIP harus {PANJANG_NIP} digit (terisi {len(n)} digit).")
    for label, v in (("Jabatan", jabatan), ("Satker", satker), ("Seksi", seksi)):
        if not _spasi(v):
            galat.append(f"{label} wajib diisi.")
    if len(sandi) < MIN_SANDI:
        galat.append(f"Kata sandi minimal {MIN_SANDI} karakter.")
    elif n and n in sandi:
        galat.append("Kata sandi tidak boleh mengandung NIP.")
    if sandi != ulang:
        galat.append("Ulangi kata sandi tidak sama.")
    return galat


# ===========================================================================
# Akun
# ===========================================================================
class Akun:
    def __init__(self, db: Penyimpanan):
        self.db = db

    # ------------------------------------------------------------------ baca
    def _satu(self, nip: str) -> dict | None:
        r = self.db._q(f"SELECT * FROM {TABEL} WHERE nip = ?", (nip,), "satu")
        return self.db._kamus(r) if r is not None else None

    def jumlah_pengguna(self) -> int:
        return int(self.db._q(f"SELECT COUNT(*) AS n FROM {TABEL}", (), "satu")["n"])

    def jumlah_admin_aktif(self) -> int:
        return int(self.db._q(f"SELECT COUNT(*) AS n FROM {TABEL} WHERE peran = 'admin' AND status = 'aktif'", (), "satu")["n"])

    def daftar_pengguna(self) -> list[dict]:
        r = self.db._q(f"SELECT {KOLOM_PUBLIK} FROM {TABEL} ORDER BY id", (), "semua")
        return [self.db._kamus(x) for x in r]

    # ------------------------------------------------------------------ daftar
    def daftar(self, nama: str, nip: str, jabatan: str, satker: str, seksi: str, sandi: str, ulang: str,
               admin_nip: str = "", online: bool = False) -> tuple[bool, str, str]:
        """Daftarkan akun. Mengembalikan (berhasil, pesan, status). status: 'aktif' | 'menunggu' | ''."""
        galat = validasi_daftar(nama, nip, jabatan, satker, seksi, sandi, ulang)
        if galat:
            return False, " ".join(galat), ""
        n = bersihkan_nip(nip)
        admin_nip = bersihkan_nip(admin_nip)
        if online and not admin_nip:
            return False, "Pendaftaran belum dibuka: admin belum ditentukan (SEMAR_ADMIN_NIP di Secrets).", ""
        if self._satu(n) is not None:
            return False, "NIP ini sudah terdaftar. Silakan masuk, atau hubungi admin bila lupa kata sandi.", ""
        pertama = self.jumlah_pengguna() == 0
        admin = (admin_nip != "" and n == admin_nip) or (not online and pertama)
        peran, status = ("admin", "aktif") if admin else ("pengguna", "menunggu")
        try:
            self.db._q(f"INSERT INTO {TABEL} (nip, nama, jabatan, satker, seksi, sandi, peran, status, dibuat) "
                       "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
                       (n, _spasi(nama), _spasi(jabatan), _spasi(satker), _spasi(seksi), hash_sandi(sandi), peran, status, time.time()))
        except Exception:        # noqa: BLE001  (NIP ganda pada pendaftaran serentak, dst.)
            if self._satu(n) is not None:
                return False, "NIP ini sudah terdaftar.", ""
            raise
        if admin:
            return True, "Akun admin dibuat dan langsung aktif. Silakan masuk.", status
        return True, "Pendaftaran berhasil. Akun Anda menunggu persetujuan admin sebelum bisa dipakai.", status

    # ------------------------------------------------------------------ masuk
    def masuk(self, nip: str, sandi: str) -> tuple[dict | None, str]:
        """Periksa NIP dan kata sandi. Mengembalikan (pengguna atau None, pesan galat)."""
        n = bersihkan_nip(nip)
        umum = "NIP atau kata sandi salah."
        u = self._satu(n) if n else None
        if u is None:
            cek_sandi(sandi, _hash_palsu())
            return None, umum
        sekarang = time.time()
        if u["terkunci_sampai"] > sekarang:
            menit = int((u["terkunci_sampai"] - sekarang) // 60) + 1
            return None, f"Terlalu banyak percobaan salah. Akun dikunci sementara, coba lagi dalam {menit} menit."
        if not cek_sandi(sandi, u["sandi"]):
            gagal = u["gagal"] + 1
            if gagal >= MAKS_GAGAL:
                self.db._q(f"UPDATE {TABEL} SET gagal = 0, terkunci_sampai = ? WHERE id = ?", (sekarang + KUNCI_MENIT * 60, u["id"]))
                return None, f"Terlalu banyak percobaan salah. Akun dikunci {KUNCI_MENIT} menit."
            self.db._q(f"UPDATE {TABEL} SET gagal = ? WHERE id = ?", (gagal, u["id"]))
            return None, umum
        if u["status"] == "menunggu":
            return None, "Akun Anda masih menunggu persetujuan admin."
        if u["status"] != "aktif":
            return None, "Akun Anda dinonaktifkan. Hubungi admin."
        self.db._q(f"UPDATE {TABEL} SET gagal = 0, terkunci_sampai = 0, masuk_terakhir = ? WHERE id = ?", (sekarang, u["id"]))
        return {k: u[k] for k in ("id", "nip", "nama", "jabatan", "satker", "seksi", "peran")}, ""

    def masih_berlaku(self, id_: int) -> dict | None:
        """Data terbaru pengguna bila akunnya masih aktif (dipakai agar akun yang dinonaktifkan langsung tertutup)."""
        r = self.db._q(f"SELECT {KOLOM_PUBLIK} FROM {TABEL} WHERE id = ?", (id_,), "satu")
        if r is None:
            return None
        u = self.db._kamus(r)
        return u if u["status"] == "aktif" else None

    # ------------------------------------------------------------------ kelola
    def atur_status(self, id_: int, status: str) -> None:
        if status not in ("aktif", "menunggu", "nonaktif"):
            raise ValueError(status)
        self.db._q(f"UPDATE {TABEL} SET status = ? WHERE id = ?", (status, id_))

    def atur_peran(self, id_: int, peran: str) -> None:
        if peran not in ("admin", "pengguna"):
            raise ValueError(peran)
        self.db._q(f"UPDATE {TABEL} SET peran = ? WHERE id = ?", (peran, id_))

    def atur_sandi(self, id_: int, sandi_baru: str) -> None:
        self.db._q(f"UPDATE {TABEL} SET sandi = ?, gagal = 0, terkunci_sampai = 0 WHERE id = ?", (hash_sandi(sandi_baru), id_))

    def ganti_sandi(self, id_: int, lama: str, baru: str, ulang: str, nip: str) -> tuple[bool, str]:
        r = self.db._q(f"SELECT sandi FROM {TABEL} WHERE id = ?", (id_,), "satu")
        if r is None or not cek_sandi(lama, r["sandi"]):
            return False, "Kata sandi lama salah."
        if len(baru) < MIN_SANDI:
            return False, f"Kata sandi baru minimal {MIN_SANDI} karakter."
        if nip and nip in baru:
            return False, "Kata sandi tidak boleh mengandung NIP."
        if baru != ulang:
            return False, "Ulangi kata sandi tidak sama."
        self.atur_sandi(id_, baru)
        return True, "Kata sandi diganti."

    def hapus(self, id_: int) -> None:
        self.db._q(f"DELETE FROM {TABEL} WHERE id = ?", (id_,))
