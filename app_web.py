"""
SEMAR versi web (Streamlit) - tema tokoh Semar, Kota Semarang, dan Bea Cukai.

Jalankan:  jalankan_semar_web.bat   (Windows, klik dua kali)
atau      python -m streamlit run app_web.py --theme.base light

Mesinnya sama dengan semar.py: hanya menjawab dari tabel Excel dan website
yang tercantum di file basis_pengetahuan_semar.xlsx.

Tampilan Dashboard: awalnya dashboard penuh dengan gambar Cyborg Semar melayang di
kanan bawah (berkas semar_melayang.png, satu folder dengan app_web.py). Klik gambar
itu -> kolom chat/perintah SEMAR muncul dan dashboard bergeser berbagi ruang.
Tombol "Tutup" di kolom chat mengembalikan dashboard penuh.

Halaman (dipilih di sidebar): 1) Dashboard, 2) Semar (percakapan), 3) Pemeriksaan Barang
(sistem OCR + AI untuk pemeriksa barang; logikanya di pemeriksaan.py).

Suara: SEMAR berbicara dengan suara PRIA Indonesia "Ardi" (Microsoft, lewat pustaka edge-tts,
tanpa API key, perlu internet). Bila edge-tts tidak ada atau internet mati, dipakai suara bawaan
peramban (Web Speech API). Jawaban dibacakan otomatis; gambar Semar melayang menyapa saat
diklik ("Ada yang bisa SEMAR bantu?") dan berpamitan saat ditutup ("Senang bisa
membantu."). Saklar "Suara SEMAR" di sidebar menyalakan/mematikan semuanya.
Pengaturan suara (opsional, taruh di .env):
  SEMAR_MESIN=peramban      -> paksa suara bawaan peramban (tanpa edge-tts)
  SEMAR_SUARA_SERVER=id-ID-ArdiNeural   (satu-satunya suara pria Indonesia di edge-tts)
  SEMAR_EDGE_NADA=-6Hz  SEMAR_EDGE_LAJU=-4%   (nada/laju suara Ardi; makin negatif makin berat/pelan)
  SEMAR_SUARA, SEMAR_NADA, SEMAR_LAJU        -> pengaturan untuk suara cadangan peramban
"""

import asyncio
import base64
import hashlib
import html as H
import inspect
import io
import json
import os
import re
import time
import urllib.parse
from pathlib import Path

import pandas as pd
import streamlit as st

import semar as S

IKON_SEMAR = Path(os.getenv("SEMAR_IKON", S.BASE_DIR / "semar_ikon.webp"))   # wajah Semar bulat (banner, sidebar, avatar chat)


def _ikon_favicon():
    """Ikon tab peramban: wajah Semar bila berkasnya ada, kalau tidak emoji bawaan."""
    try:
        from PIL import Image
        return Image.open(IKON_SEMAR).convert("RGBA")
    except Exception:  # noqa: BLE001
        return "🏛️"


st.set_page_config(page_title="SEMAR", page_icon=_ikon_favicon(), layout="wide",
                   initial_sidebar_state="expanded")

# Pemeriksaan: semar.py di folder ini harus versi terbaru.
_PERLU = ("VERSI_MODEL", "EXTRA_BODY", "diagnosa_kunci", "NAMA_ENV")
_HILANG = [n for n in _PERLU if not hasattr(S, n)]
if _HILANG:
    st.error("File semar.py di folder ini masih VERSI LAMA, jadi aplikasi web belum bisa berjalan.")
    st.markdown(
        "**Cara memperbaiki:**\n"
        "1. Ambil file `semar.py` terbaru dari percakapan (yang paling akhir dikirim).\n"
        "2. Salin ke folder SEMAR dan **timpa** file `semar.py` yang lama (pilih *Replace*).\n"
        "3. Matikan SEMAR (tutup jendela hitam), lalu jalankan lagi."
    )
    st.code(f"File yang terbaca : {getattr(S, '__file__', '?')}\nBagian yang tidak ada: {', '.join(_HILANG)}", language="text")
    st.stop()

try:
    import dashboard as DB
    import pemeriksaan as PM
    import peta_wilayah as PW
except ImportError as e:  # berkas pendukung belum disalin ke folder ini
    st.error("Berkas pendukung belum lengkap, jadi aplikasi web belum bisa berjalan.")
    st.markdown(
        "Salin **tiga berkas** ini ke folder SEMAR (satu folder dengan `app_web.py`):\n"
        "- `dashboard.py`\n- `peta_wilayah.py`\n- `pemeriksaan.py`\n\nLalu jalankan ulang SEMAR.")
    st.code(f"Penyebab: {e}", language="text")
    st.stop()

DATA_PATH = Path(os.getenv("SEMAR_DATA", S.BASE_DIR / "data_dashboard_semar.xlsx"))
GAMBAR_SEMAR = Path(os.getenv("SEMAR_GAMBAR", S.BASE_DIR / "semar_melayang.png"))

# --- Suara SEMAR (diucapkan oleh peramban) ---
SAPA_BUKA = "Ada yang bisa SEMAR bantu?"      # saat gambar Semar melayang diklik (chat terbuka)
SAPA_TUTUP = "Senang bisa membantu."          # saat chat ditutup
MAKS_HURUF_SUARA = 1800                              # jawaban yang lebih panjang hanya dibacakan sampai batas ini


def _angka_env(nama: str, awal: float, bawah: float, atas: float) -> float:
    try:
        return min(atas, max(bawah, float(os.getenv(nama, str(awal)).replace(",", "."))))
    except ValueError:
        return awal


SUARA_NADA = _angka_env("SEMAR_NADA", 0.8, 0.1, 2.0)    # nada rendah: lebih berat, cocok untuk Semar
SUARA_LAJU = _angka_env("SEMAR_LAJU", 0.95, 0.5, 2.0)
SUARA_NAMA = os.getenv("SEMAR_SUARA", "").strip()        # contoh: "Ardi" atau "Andika"
# Suara utama: pria Indonesia dari server Microsoft (edge-tts). Peramban hanya cadangan.
MESIN_EDGE = os.getenv("SEMAR_MESIN", "edge").strip().lower() != "peramban"
EDGE_SUARA = os.getenv("SEMAR_SUARA_SERVER", "id-ID-ArdiNeural").strip()
EDGE_NADA = os.getenv("SEMAR_EDGE_NADA", "-6Hz").strip()
EDGE_LAJU = os.getenv("SEMAR_EDGE_LAJU", "-4%").strip()


# ===========================================================================
# ILUSTRASI (SVG orisinal): tokoh Semar, siluet Kota Semarang, motif kawung
# ===========================================================================
SEMAR_SVG = """<svg viewBox="0 0 200 200" xmlns="http://www.w3.org/2000/svg" role="img" aria-label="Semar">
<defs>
<radialGradient id="sm-bg" cx="50%" cy="35%" r="75%"><stop offset="0" stop-color="#FFE9A8"/><stop offset=".55" stop-color="#F7B733"/><stop offset="1" stop-color="#E8730C"/></radialGradient>
<linearGradient id="sm-kulit" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#F0C48A"/><stop offset="1" stop-color="#D99A52"/></linearGradient>
<pattern id="sm-poleng" width="16" height="16" patternUnits="userSpaceOnUse"><rect width="16" height="16" fill="#101827"/><rect width="8" height="8" fill="#F5EBD7"/><rect x="8" y="8" width="8" height="8" fill="#F5EBD7"/></pattern>
<clipPath id="sm-klip"><circle cx="100" cy="100" r="94"/></clipPath>
</defs>
<circle cx="100" cy="100" r="97" fill="#0A1F44"/>
<circle cx="100" cy="100" r="94" fill="url(#sm-bg)"/>
<g clip-path="url(#sm-klip)">
<g opacity=".16" fill="none" stroke="#7A3E00" stroke-width="2"><circle cx="100" cy="100" r="84"/><circle cx="100" cy="100" r="68" stroke-dasharray="3 5"/></g>
<path d="M18 206 Q26 148 100 142 Q174 148 182 206 Z" fill="url(#sm-poleng)" stroke="#101827" stroke-width="3"/>
<path d="M70 146 Q100 168 130 146 L128 138 Q100 152 72 138 Z" fill="#101827"/>
<rect x="84" y="118" width="32" height="30" rx="10" fill="#C98A47"/>
<ellipse cx="53" cy="96" rx="9" ry="13" fill="#D99A52" stroke="#5B3A1A" stroke-width="2.5"/>
<ellipse cx="147" cy="96" rx="9" ry="13" fill="#D99A52" stroke="#5B3A1A" stroke-width="2.5"/>
<ellipse cx="100" cy="92" rx="47" ry="52" fill="url(#sm-kulit)" stroke="#5B3A1A" stroke-width="3"/>
<path d="M52 82 Q56 38 100 36 Q144 38 148 82 Q134 56 100 56 Q66 56 52 82 Z" fill="#161B2E" stroke="#5B3A1A" stroke-width="2.5"/>
<path d="M100 40 C88 32 90 16 100 6 C110 16 112 32 100 40 Z" fill="#161B2E" stroke="#5B3A1A" stroke-width="2.5"/>
<path d="M97 30 Q100 20 103 30" fill="none" stroke="#F7B733" stroke-width="2" stroke-linecap="round"/>
<path d="M69 85 Q79 78 91 85" fill="none" stroke="#3B2410" stroke-width="3.4" stroke-linecap="round"/>
<path d="M109 85 Q121 78 131 85" fill="none" stroke="#3B2410" stroke-width="3.4" stroke-linecap="round"/>
<path d="M72 93 Q81 87 90 93 Q81 99 72 93 Z" fill="#FBF3E4" stroke="#3B2410" stroke-width="1.8"/>
<path d="M110 93 Q119 87 128 93 Q119 99 110 93 Z" fill="#FBF3E4" stroke="#3B2410" stroke-width="1.8"/>
<path d="M71 92 Q81 85 91 92 L91 93 Q81 90 71 93 Z" fill="#6B4423"/>
<path d="M109 92 Q119 85 129 92 L129 93 Q119 90 109 93 Z" fill="#6B4423"/>
<circle cx="81" cy="94" r="2.7" fill="#161B2E"/><circle cx="119" cy="94" r="2.7" fill="#161B2E"/>
<path d="M73 101 Q81 105 89 101" fill="none" stroke="#B9783A" stroke-width="1.6" stroke-linecap="round"/>
<path d="M111 101 Q119 105 127 101" fill="none" stroke="#B9783A" stroke-width="1.6" stroke-linecap="round"/>
<path d="M100 97 C86 99 85 119 100 120 C115 119 114 99 100 97 Z" fill="#D28A42" stroke="#5B3A1A" stroke-width="2.4"/>
<ellipse cx="95.5" cy="106" rx="3.4" ry="2.4" fill="#F3CF9E" opacity=".85"/>
<path d="M93 116 Q96 118.5 99 116 M101 116 Q104 118.5 107 116" fill="none" stroke="#7A4A21" stroke-width="1.6" stroke-linecap="round"/>
<circle cx="72" cy="112" r="9" fill="#E7A27A" opacity=".45"/><circle cx="128" cy="112" r="9" fill="#E7A27A" opacity=".45"/>
<path d="M80 127 Q100 140 120 127" fill="none" stroke="#3B2410" stroke-width="3" stroke-linecap="round"/>
<path d="M77 124 Q76 128 79 130 M123 124 Q124 128 121 130" fill="none" stroke="#B9783A" stroke-width="1.8" stroke-linecap="round"/>
</g>
<circle cx="100" cy="100" r="94" fill="none" stroke="#FFE9A8" stroke-width="2"/>
</svg>"""

SKYLINE_SVG = """<svg viewBox="0 0 1200 170" preserveAspectRatio="xMidYMax slice" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
<g fill="#071530">
<g transform="translate(125 0)"><ellipse cx="170" cy="146" rx="54" ry="8"/><rect x="140" y="128" width="60" height="18" rx="2"/><rect x="150" y="112" width="40" height="18"/>
<path d="M158 112 L162 50 L170 12 L178 50 L182 112 Z"/><path d="M170 0 C162 14 164 22 170 28 C176 22 178 14 170 0 Z"/>
<path d="M148 126 L120 112 L150 118 Z M192 126 L220 112 L190 118 Z"/></g>
<rect x="400" y="96" width="290" height="50"/><rect x="394" y="92" width="302" height="8"/>
<rect x="404" y="58" width="42" height="40"/><path d="M398 58 L425 26 L452 58 Z"/><rect x="421" y="16" width="8" height="14"/>
<rect x="644" y="58" width="42" height="40"/><path d="M638 58 L665 26 L692 58 Z"/><rect x="661" y="16" width="8" height="14"/>
<path d="M520 96 L545 66 L570 96 Z"/><rect x="522" y="96" width="46" height="50"/>
<rect x="870" y="104" width="110" height="42"/><path d="M880 104 A45 45 0 0 1 970 104 Z"/>
<rect x="922" y="48" width="6" height="14"/><rect x="918" y="54" width="14" height="4"/>
<rect x="858" y="82" width="22" height="64"/><path d="M854 82 L869 56 L884 82 Z"/>
<rect x="970" y="82" width="22" height="64"/><path d="M966 82 L981 56 L996 82 Z"/>
<path d="M1030 142 L1190 142 L1174 158 L1046 158 Z"/><rect x="1150" y="112" width="26" height="30"/><rect x="1154" y="100" width="18" height="14"/>
<rect x="1040" y="124" width="104" height="18"/>
<path d="M1000 146 L1000 90 L1016 90 L1016 146 Z M960 90 L1030 90 L1030 84 L960 84 Z"/>
</g>
<g>
<rect x="1044" y="128" width="22" height="14" fill="#F58220"/><rect x="1068" y="128" width="22" height="14" fill="#F7B733"/><rect x="1092" y="128" width="22" height="14" fill="#2F80ED"/><rect x="1116" y="128" width="22" height="14" fill="#F58220"/>
<rect x="1056" y="114" width="22" height="13" fill="#2F80ED"/><rect x="1080" y="114" width="22" height="13" fill="#F58220"/><rect x="1104" y="114" width="22" height="13" fill="#F7B733"/>
</g>
<path d="M0 150 Q60 140 120 150 T240 150 T360 150 T480 150 T600 150 T720 150 T840 150 T960 150 T1080 150 T1200 150 L1200 170 L0 170 Z" fill="#050F26" opacity=".92"/>
<path d="M0 152 Q60 144 120 152 T240 152 T360 152 T480 152 T600 152 T720 152 T840 152 T960 152 T1080 152 T1200 152" fill="none" stroke="#F7B733" stroke-opacity=".55" stroke-width="2"/>
</svg>"""


def kawung_uri(opacity: float) -> str:
    svg = (
        "<svg xmlns='http://www.w3.org/2000/svg' width='64' height='64' viewBox='0 0 64 64'>"
        f"<g fill='none' stroke='white' stroke-width='1.2' opacity='{opacity}'>"
        "<ellipse cx='32' cy='16' rx='9' ry='15'/><ellipse cx='32' cy='48' rx='9' ry='15'/>"
        "<ellipse cx='16' cy='32' rx='15' ry='9'/><ellipse cx='48' cy='32' rx='15' ry='9'/>"
        f"<circle cx='32' cy='32' r='2.4' fill='white'/></g></svg>"
    )
    return "data:image/svg+xml;utf8," + urllib.parse.quote(svg)


@st.cache_resource(show_spinner=False)
def _ikon_uri(tanda: float) -> str | None:
    """Data-URI ikon wajah Semar dari semar_ikon.webp. `tanda` = waktu ubah berkas agar cache ikut segar."""
    try:
        mime = "image/webp" if IKON_SEMAR.suffix.lower() == ".webp" else "image/png"
        return f"data:{mime};base64," + base64.b64encode(IKON_SEMAR.read_bytes()).decode()
    except Exception:  # noqa: BLE001
        return None


def _ikon_aktif() -> str | None:
    return _ikon_uri(IKON_SEMAR.stat().st_mtime if IKON_SEMAR.exists() else 0.0)


def avatar_semar() -> str:
    """Semar kecil untuk avatar jawaban: wajah dari gambar baru, cadangan: ilustrasi SVG bawaan."""
    return _ikon_aktif() or ("data:image/svg+xml;utf8," + urllib.parse.quote(SEMAR_SVG))


def ikon_semar_html() -> str:
    """Ikon Semar untuk banner dan sidebar (HTML)."""
    uri = _ikon_aktif()
    return f'<img src="{uri}" alt="Semar">' if uri else SEMAR_SVG


# ===========================================================================
# CSS: gradasi biru-laut Bea Cukai -> oranye senja Semarang, emas Semar
# ===========================================================================
CSS = """
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&family=Playfair+Display:wght@700;800&display=swap');
:root{--navy:#0A1F44;--biru:#1B4F9C;--oranye:#F58220;--emas:#F7B733;--krem:#FFF8EC;--tinta:#1B2440;}
html,body,.stApp,p,li,label,button,input,textarea,h1,h2,h3,h4,h5,h6,[data-testid="stMarkdownContainer"],[data-testid="stChatInput"] *{font-family:'Plus Jakarta Sans','Segoe UI',system-ui,sans-serif;}
[data-testid="stIconMaterial"],.material-symbols-rounded,[class*="material-symbols"]{font-family:'Material Symbols Rounded'!important;}
.stApp{background:
  radial-gradient(1100px 500px at 100% -10%,rgba(245,130,32,.14),transparent 60%),
  radial-gradient(900px 500px at -10% 30%,rgba(27,79,156,.12),transparent 60%),
  linear-gradient(180deg,#FFF8EC 0%,#FFFDF8 45%,#EEF3FB 100%);color:var(--tinta);}
[data-testid="stHeader"]{background:transparent;}
/* Sembunyikan hanya isi toolbar (menu, Deploy, dsb). Wadah stToolbar JANGAN di-display:none:
   di Streamlit baru tombol "buka sidebar" berada di dalamnya, sehingga ikut hilang. */
[data-testid="stToolbarActions"],[data-testid="stMainMenu"],[data-testid="stAppDeployButton"],.stDeployButton,#MainMenu,footer{display:none!important;}
[data-testid="stToolbar"]{background:transparent!important;}
[data-testid="stExpandSidebarButton"],[data-testid="stSidebarCollapsedControl"]{display:flex!important;visibility:visible!important;}
[data-testid="stExpandSidebarButton"]{background:linear-gradient(135deg,#F58220,#F7B733)!important;border-radius:50%!important;
  box-shadow:0 8px 18px -8px rgba(245,130,32,.9);}
[data-testid="stExpandSidebarButton"],[data-testid="stExpandSidebarButton"] *{color:#1B1204!important;fill:#1B1204!important;}
.block-container{padding-top:1.1rem;padding-bottom:7rem;max-width:1080px;}

/* ---------- HERO ---------- */
.hero{position:relative;overflow:hidden;border-radius:28px;min-height:330px;color:#fff;
  background:linear-gradient(115deg,#0A1F44 0%,#12377A 30%,#1B4F9C 52%,#C9622A 80%,#F58220 100%);
  background-size:220% 220%;animation:geser 14s ease-in-out infinite;
  box-shadow:0 24px 60px -18px rgba(10,31,68,.55);margin-bottom:1.4rem;}
.hero-batik{position:absolute;inset:0;background-image:url("__KAWUNG_HERO__");background-size:64px;}
.hero-cahaya{position:absolute;right:-60px;top:-80px;width:340px;height:340px;border-radius:50%;
  background:radial-gradient(circle,rgba(247,183,51,.55),rgba(247,183,51,0) 65%);filter:blur(4px);animation:denyut 6s ease-in-out infinite;}
.hero-skyline{position:absolute;left:0;right:0;bottom:0;height:150px;opacity:.96;}
.hero-skyline svg{width:100%;height:100%;display:block;}
.hero-isi{position:relative;z-index:2;display:flex;gap:28px;align-items:flex-start;flex-wrap:wrap;padding:34px 38px 0 38px;}
.hero-semar{width:148px;height:148px;flex:0 0 auto;filter:drop-shadow(0 12px 22px rgba(0,0,0,.45));animation:melayang 5s ease-in-out infinite;}
.hero-semar svg,.hero-semar img{width:100%;height:100%;display:block;}
.hero-label{display:inline-block;font-size:.78rem;letter-spacing:.16em;text-transform:uppercase;font-weight:700;
  padding:5px 14px;border-radius:999px;background:rgba(255,255,255,.14);border:1px solid rgba(255,255,255,.28);backdrop-filter:blur(6px);}
.hero-judul{font-family:'Playfair Display',Georgia,serif!important;font-weight:800;font-size:4.2rem;line-height:1.02;margin:.35rem 0 .1rem 0;letter-spacing:.06em;padding:0;
  background:linear-gradient(90deg,#FFF3C4,#F7B733 45%,#FFE29A 70%,#fff);background-size:200% auto;-webkit-background-clip:text;background-clip:text;
  -webkit-text-fill-color:transparent;color:transparent;animation:kilau 6s linear infinite;}
.hero-sub{font-size:1.08rem;margin:.2rem 0 .8rem 0;max-width:560px;color:rgba(255,255,255,.94);}
.hero-pil{display:flex;gap:8px;flex-wrap:wrap;padding-bottom:128px;}
.hero-pil span{font-size:.8rem;font-weight:600;padding:5px 13px;border-radius:999px;background:rgba(7,21,48,.45);
  border:1px solid rgba(255,255,255,.22);backdrop-filter:blur(6px);}

/* ---------- SAMBUTAN & PERTANYAAN CEPAT ---------- */
.sambutan{background:rgba(255,255,255,.78);backdrop-filter:blur(6px);border:1px solid rgba(10,31,68,.08);border-radius:22px;
  padding:1.3rem 1.6rem;margin-bottom:1rem;box-shadow:0 10px 30px -14px rgba(10,31,68,.25);position:relative;overflow:hidden;animation:muncul .6s ease both;}
.sambutan:before{content:"";position:absolute;left:0;top:0;bottom:0;width:6px;background:linear-gradient(180deg,var(--oranye),var(--emas));}
.sambutan h3{font-family:'Playfair Display',Georgia,serif!important;font-size:1.5rem;margin:0 0 .3rem 0;color:var(--navy);}
.sambutan p{margin:0;color:#3A4668;}
.label-cepat{font-weight:700;color:var(--navy);margin:.6rem 0 .3rem 0;font-size:.95rem;}
[data-testid="stMain"] .stButton>button{border-radius:999px;border:1.5px solid rgba(245,130,32,.55);background:#fff;color:var(--navy);
  font-weight:600;padding:.55rem 1rem;width:100%;transition:all .22s ease;box-shadow:0 4px 14px -8px rgba(10,31,68,.35);}
[data-testid="stMain"] .stButton>button:hover{background:linear-gradient(90deg,var(--oranye),var(--emas));border-color:transparent;color:#1B1204;
  transform:translateY(-3px);box-shadow:0 14px 26px -12px rgba(245,130,32,.7);}
[data-testid="stMain"] .stButton>button:active{transform:translateY(0);}

/* ---------- PERCAKAPAN ---------- */
[data-testid="stChatMessage"]{border-radius:22px;padding:1rem 1.25rem;margin-bottom:.9rem;border:1px solid rgba(10,31,68,.07);
  box-shadow:0 10px 28px -16px rgba(10,31,68,.35);animation:muncul .45s ease both;background:#fff;}
[data-testid="stChatMessage"] p,[data-testid="stChatMessage"] li{line-height:1.65;}
[data-testid="stChatMessage"]:has([data-testid="stChatMessageAvatarUser"]){background:linear-gradient(135deg,#0A1F44 0%,#1B4F9C 100%);border:none;}
[data-testid="stChatMessage"]:has([data-testid="stChatMessageAvatarUser"]) p,
[data-testid="stChatMessage"]:has([data-testid="stChatMessageAvatarUser"]) li,
[data-testid="stChatMessage"]:has([data-testid="stChatMessageAvatarUser"]) span{color:#fff!important;}
[data-testid="stChatMessage"]:not(:has([data-testid="stChatMessageAvatarUser"])){border-left:6px solid var(--oranye);
  background:linear-gradient(135deg,#FFFFFF 0%,#FFF6E6 100%);}
[data-testid="stChatMessageAvatarUser"]{background:linear-gradient(135deg,#F58220,#F7B733)!important;}
[data-testid="stChatMessageAvatarAssistant"],[data-testid="stChatMessageAvatarCustom"]{background:transparent!important;}
[data-testid="stChatMessage"] img{border-radius:50%;box-shadow:0 0 0 2px #F7B733;}
.rujuk-baris{display:flex;flex-wrap:wrap;gap:6px;margin-top:.6rem;align-items:center;}
.rujuk-judul{font-size:.76rem;font-weight:700;letter-spacing:.08em;text-transform:uppercase;color:#7A4A21;}
.rujuk{display:inline-flex;gap:6px;align-items:center;font-size:.78rem;padding:4px 11px;border-radius:999px;
  background:linear-gradient(90deg,rgba(27,79,156,.10),rgba(245,130,32,.14));border:1px solid rgba(27,79,156,.22);color:var(--navy)!important;text-decoration:none!important;transition:all .2s;}
.rujuk b{background:linear-gradient(90deg,var(--biru),var(--navy));color:#fff;border-radius:999px;padding:1px 8px;font-size:.72rem;}
a.rujuk:hover{transform:translateY(-2px);box-shadow:0 8px 16px -8px rgba(27,79,156,.6);}

/* ---------- KOTAK TANYA ---------- */
[data-testid="stBottomBlockContainer"],[data-testid="stBottom"]>div{background:linear-gradient(180deg,rgba(255,248,236,0),#FFF8EC 38%)!important;}
[data-testid="stChatInput"]{border-radius:999px!important;border:4px solid transparent!important;
  background:linear-gradient(#fff,#fff) padding-box,linear-gradient(90deg,#1B4F9C,#F58220,#F7B733) border-box!important;
  box-shadow:0 16px 36px -16px rgba(10,31,68,.45);}
[data-testid="stChatInput"] div{background:transparent!important;border:none!important;box-shadow:none!important;}
[data-testid="stChatInput"] textarea{color:var(--tinta)!important;background:transparent!important;}
[data-testid="stChatInput"] textarea::placeholder{color:#7C86A3!important;}
[data-testid="stChatInput"]:focus-within{box-shadow:0 0 0 4px rgba(245,130,32,.25),0 16px 36px -16px rgba(10,31,68,.45);}
[data-testid="stChatInputSubmitButton"]{background:linear-gradient(135deg,var(--oranye),var(--emas))!important;border-radius:50%!important;color:#1B1204!important;}

/* ---------- SIDEBAR ---------- */
[data-testid="stSidebar"]{background:url("__KAWUNG_SIDE__"),linear-gradient(190deg,#0A1F44 0%,#12377A 52%,#8A430E 135%)!important;border-right:none;}
[data-testid="stSidebar"] *{color:#EAF0FF;}
.side-merek{display:flex;align-items:center;gap:12px;margin:.2rem 0 .8rem 0;}
.side-merek .logo{width:54px;height:54px;flex:0 0 auto;filter:drop-shadow(0 6px 12px rgba(0,0,0,.4));}
.side-merek .logo svg,.side-merek .logo img{width:100%;height:100%;display:block;}
.side-merek .nama{font-family:'Playfair Display',Georgia,serif!important;font-weight:800;font-size:1.6rem;letter-spacing:.08em;line-height:1;
  background:linear-gradient(90deg,#FFF3C4,#F7B733);-webkit-background-clip:text;background-clip:text;-webkit-text-fill-color:transparent;}
.side-merek .kecil{font-size:.72rem;opacity:.85;letter-spacing:.04em;}
.stat-grid{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin:.4rem 0 .6rem 0;}
.stat{border-radius:16px;padding:12px 12px 10px 12px;background:linear-gradient(145deg,rgba(255,255,255,.16),rgba(255,255,255,.05));
  border:1px solid rgba(255,255,255,.2);backdrop-filter:blur(6px);transition:transform .2s;}
.stat:hover{transform:translateY(-3px);}
.stat .angka{font-size:1.9rem;font-weight:800;line-height:1;color:#FFE29A;}
.stat .ket{font-size:.72rem;opacity:.85;margin-top:3px;}
.kartu-model{border-radius:16px;padding:11px 13px;background:rgba(7,21,48,.42);border:1px solid rgba(255,255,255,.18);font-size:.8rem;margin-bottom:.7rem;}
.kartu-model b{color:#FFE29A;font-size:.92rem;}
.sub-judul{font-size:.72rem;font-weight:700;letter-spacing:.14em;text-transform:uppercase;opacity:.75;margin:.9rem 0 .35rem 0;}
[data-testid="stSidebar"] .stButton>button{border:none;border-radius:14px;font-weight:700;width:100%;
  background:linear-gradient(90deg,#F58220,#F7B733);box-shadow:0 10px 20px -10px rgba(245,130,32,.9);transition:all .2s;}
[data-testid="stSidebar"] .stButton>button *{color:#1B1204!important;}
[data-testid="stSidebar"] .stButton>button:hover{transform:translateY(-2px);filter:brightness(1.07);}
[data-testid="stSidebar"] .stButton>button[kind="secondary"]{background:rgba(255,255,255,.12);box-shadow:none;border:1px solid rgba(255,255,255,.28);}
[data-testid="stSidebar"] .stButton>button[kind="secondary"] *{color:#fff!important;}
.kaki{font-size:.74rem;opacity:.75;line-height:1.5;margin-top:.8rem;}

/* ---------- HERO KOMPAK & TATA LETAK DASHBOARD ---------- */
.hero.kompak{min-height:232px;margin-bottom:1rem;border-radius:24px;}
.hero.kompak .hero-semar{width:104px;height:104px;}
.hero.kompak .hero-isi{padding:24px 32px 0 32px;gap:22px;}
.hero.kompak .hero-judul{font-size:3rem;margin:.2rem 0 0 0;}
.hero.kompak .hero-sub{font-size:.98rem;margin:.1rem 0 .5rem 0;}
.hero.kompak .hero-skyline{height:110px;}
.hero.kompak .hero-pil{padding-bottom:84px;}
[data-testid="stColumn"]:has(.penanda-chat){position:sticky;top:3.4rem;align-self:flex-start;}
[data-testid="stVerticalBlockBorderWrapper"]:has(> div > [data-testid="stVerticalBlock"] .stChatMessage){scrollbar-width:thin;}
[data-testid="stColumn"]:has(.penanda-chat) [data-testid="stChatInput"]{margin-bottom:.6rem;}
.label-perintah{font-weight:700;color:var(--navy);font-size:.9rem;margin:.2rem 0 .4rem 0;}
.kepala-chat{display:flex;flex-direction:column;align-items:flex-start;gap:0;margin-bottom:.5rem;}
.kepala-chat b{font-family:'Playfair Display',Georgia,serif;font-size:1.15rem;color:var(--navy);}
.kepala-chat span{font-size:.76rem;color:#5C6785;}
[data-testid="stExpander"]{border-radius:18px!important;border:1px solid rgba(10,31,68,.09)!important;background:rgba(255,255,255,.7);}
/* ---------- SAKLAR TERSEGMEN ---------- */
[data-testid="stSidebar"] button[data-variant="segmented_control"]{background:rgba(255,255,255,.10)!important;border:1px solid rgba(255,255,255,.25)!important;}
[data-testid="stSidebar"] button[data-variant="segmented_control"] *{color:#EAF0FF!important;}
[data-testid="stSidebar"] button[data-variant="segmented_control"][aria-checked="true"]{background:linear-gradient(90deg,#F58220,#F7B733)!important;border-color:transparent!important;box-shadow:0 8px 16px -8px rgba(245,130,32,.9);}
[data-testid="stSidebar"] button[data-variant="segmented_control"][aria-checked="true"] *{color:#1B1204!important;font-weight:700;}
[data-testid="stMain"] button[data-variant="segmented_control"]{background:#fff!important;border:1px solid rgba(10,31,68,.14)!important;}
[data-testid="stMain"] button[data-variant="segmented_control"] *{color:var(--navy)!important;font-weight:600;}
[data-testid="stMain"] button[data-variant="segmented_control"][aria-checked="true"]{background:linear-gradient(90deg,#0A1F44,#1B4F9C)!important;border-color:transparent!important;}
[data-testid="stMain"] button[data-variant="segmented_control"][aria-checked="true"] *{color:#fff!important;}
/* ---------- KOLOM CHAT YANG DIBUKA LEWAT GAMBAR SEMAR ---------- */
[data-testid="stColumn"]:has(.penanda-chat){animation:geserMasuk .5s cubic-bezier(.2,.8,.2,1) both;}
.kepala-chat{padding-right:104px;}
.st-key-tutup_chat{position:absolute!important;top:0;right:0;width:auto!important;z-index:5;}
.st-key-tutup_chat .stButton>button{width:auto!important;padding:.3rem .95rem!important;font-size:.85rem;}
/* ---------- WADAH SUARA (tak terlihat, hanya membawa skrip) ---------- */
[class*="st-key-bunyi_"]{position:absolute!important;width:1px!important;height:1px!important;overflow:hidden!important;
  opacity:0;pointer-events:none;margin:0!important;padding:0!important;}
/* ---------- ANIMASI ---------- */
@keyframes geserMasuk{from{opacity:0;transform:translateX(34px)}to{opacity:1;transform:translateX(0)}}
@keyframes semarApung{0%,100%{transform:translateY(0)}50%{transform:translateY(-12px)}}
@keyframes semarGelembung{0%{opacity:0;transform:translateY(10px)}4%,86%{opacity:1;transform:none}100%{opacity:0;transform:none;visibility:hidden}}
@keyframes geser{0%{background-position:0% 50%}50%{background-position:100% 50%}100%{background-position:0% 50%}}
@keyframes kilau{to{background-position:200% center}}
@keyframes melayang{0%,100%{transform:translateY(0) rotate(-1.5deg)}50%{transform:translateY(-9px) rotate(1.5deg)}}
@keyframes denyut{0%,100%{opacity:.75;transform:scale(1)}50%{opacity:1;transform:scale(1.12)}}
@keyframes muncul{from{opacity:0;transform:translateY(10px)}to{opacity:1;transform:translateY(0)}}
@media (prefers-reduced-motion:reduce){*{animation:none!important;transition:none!important;}}
@media (max-width:700px){.hero-judul{font-size:3rem}.hero-isi{padding:26px 20px 0 20px}.hero-semar{width:104px;height:104px}.hero-skyline{height:110px}.hero-pil{padding-bottom:110px}}
"""


CSS_NAV = """
.st-key-nav_halaman{margin:.1rem 0 .5rem 0;}
.st-key-nav_halaman .stButton>button{justify-content:flex-start;text-align:left;padding:.62rem .9rem;font-size:.95rem;}
.st-key-nav_halaman .stButton>button[kind="secondary"]{background:rgba(255,255,255,.08);}
.st-key-nav_halaman .stButton>button[kind="secondary"]:hover{background:rgba(255,255,255,.2);}
"""


def padatkan(teks: str) -> str:
    """Hilangkan baris kosong/spasi berlebih supaya tidak dibaca sebagai markdown."""
    return re.sub(r">\s+<", "><", re.sub(r"[ \t]*\n[ \t]*", " ", teks)).strip()


def terapkan_tema():
    css = (CSS + DB.CSS + CSS_NAV + PM.CSS).replace("__KAWUNG_HERO__", kawung_uri(0.11)).replace("__KAWUNG_SIDE__", kawung_uri(0.07))
    st.markdown(f"<style>{re.sub(chr(10) + '+', ' ', css)}</style>", unsafe_allow_html=True)


# ===========================================================================
# SEMAR MELAYANG: gambar di kanan bawah yang berfungsi sebagai tombol pembuka chat
# ===========================================================================
TINGGI_SEMAR = 300   # tinggi gambar melayang di layar (piksel); sebelumnya 250, kini 20% lebih besar

# Tombol Streamlit biasa yang "dikenakan" gambar Semar lewat CSS, sehingga klik = rerun yang andal.
CSS_SEMAR_MELAYANG = """
.block-container{padding-bottom:20rem!important;}
.st-key-semar_fab{position:fixed!important;right:28px;bottom:12px;left:auto!important;top:auto!important;width:auto!important;
  z-index:1000;animation:semarApung 5s ease-in-out infinite;}
.st-key-semar_fab:before{content:"";position:absolute;left:-60%;right:-60%;top:14%;bottom:0;border-radius:50%;z-index:-1;pointer-events:none;
  background:radial-gradient(closest-side,rgba(247,183,51,.55),rgba(245,130,32,.2) 55%,transparent 75%);animation:denyut 4s ease-in-out infinite;}
.st-key-semar_fab button{position:relative;display:block;width:__FW__px!important;height:__FH__px!important;min-height:0!important;
  padding:0!important;margin:0!important;border:none!important;border-radius:0!important;background:transparent!important;box-shadow:none!important;
  transform:none!important;overflow:visible!important;cursor:pointer;line-height:0;}
.st-key-semar_fab button,.st-key-semar_fab button *{font-size:0!important;color:transparent!important;}
.st-key-semar_fab button:before{content:"";position:absolute;inset:0;background:url("__FOTO__") center bottom/contain no-repeat;
  filter:drop-shadow(0 16px 14px rgba(10,31,68,.45)) drop-shadow(0 0 10px rgba(247,183,51,.55));transform-origin:50% 100%;
  transition:transform .3s cubic-bezier(.2,.8,.2,1);}
.st-key-semar_fab button:hover:before{transform:scale(1.07) rotate(-1.5deg);}
.st-key-semar_fab button:active:before{transform:scale(.97);}
.st-key-semar_fab button:focus-visible{outline:3px solid #F7B733!important;outline-offset:6px;border-radius:26px!important;}
.st-key-semar_fab button:after{content:"Tanya SEMAR";position:absolute;right:calc(100% - 4px);top:12%;white-space:nowrap;
  font-size:.86rem!important;font-weight:700;line-height:1;color:#fff!important;padding:11px 17px;border-radius:18px 18px 4px 18px;
  background:linear-gradient(90deg,#0A1F44,#1B4F9C);border:1px solid rgba(247,183,51,.75);
  box-shadow:0 14px 26px -12px rgba(10,31,68,.7);animation:semarGelembung 14s ease .5s both;}
.st-key-semar_fab button:hover:after{animation:none;}
@media (max-width:700px){
  .st-key-semar_fab{right:12px;bottom:8px;}
  .st-key-semar_fab button{width:__FWK__px!important;height:__FHK__px!important;}
  .st-key-semar_fab button:after{display:none;}
}
@media (prefers-reduced-motion:reduce){.st-key-semar_fab:before,.st-key-semar_fab button:after{animation:none!important;}}
"""


# Saat chat terbuka, Semar tetap tampil tetapi transparan (tidak lagi menghilang).
# Diklik lagi -> chat tertutup dan Semar kembali normal.
CSS_SEMAR_TRANSPARAN = """
.st-key-semar_fab{opacity:.3;transition:opacity .35s ease;}
.st-key-semar_fab:hover{opacity:.65;}
.st-key-semar_fab:before{display:none!important;}
.st-key-semar_fab button:after{display:none!important;}
"""


@st.cache_resource(show_spinner=False)
def _muat_gambar_semar(tanda: float) -> tuple[str, float, bool]:
    """Kembalikan (data-URI, lebar/tinggi, berkas_asli). `tanda` = waktu ubah berkas, agar cache ikut segar.

    Gambar dipangkas ke bagian yang tidak transparan lalu dikecilkan supaya ringan dikirim ke peramban.
    Bila berkas tidak ada atau tidak terbaca, dipakai Semar SVG bawaan.
    """
    try:
        from PIL import Image
        im = Image.open(GAMBAR_SEMAR).convert("RGBA")
        kotak = im.getchannel("A").point(lambda v: 255 if v > 8 else 0).getbbox()
        if kotak:
            im = im.crop(kotak)
        im.thumbnail((380, 570), Image.LANCZOS)
        buf = io.BytesIO()
        try:
            im.save(buf, "WEBP", quality=86, method=6)
            mime = "image/webp"
        except Exception:  # noqa: BLE001  (Pillow tanpa WebP)
            buf = io.BytesIO()
            im.save(buf, "PNG", optimize=True)
            mime = "image/png"
        return f"data:{mime};base64,{base64.b64encode(buf.getvalue()).decode()}", im.width / im.height, True
    except Exception:  # noqa: BLE001
        return avatar_semar(), 1.0, False


def _buka_chat() -> None:
    st.session_state.chat_buka = True


def _tutup_chat() -> None:
    st.session_state.chat_buka = False


def tampil_semar_melayang(terbuka: bool = False) -> None:
    """Gambar Cyborg Semar melayang di kanan bawah. Diklik -> kolom chat/perintah terbuka.

    terbuka=True (chat sudah terbuka): gambar tetap ada tetapi transparan; diklik -> chat ditutup.
    """
    ada = GAMBAR_SEMAR.exists()
    uri, rasio, asli = _muat_gambar_semar(GAMBAR_SEMAR.stat().st_mtime if ada else 0.0)
    css = (CSS_SEMAR_MELAYANG.replace("__FOTO__", uri)
           .replace("__FH__", str(TINGGI_SEMAR)).replace("__FW__", str(round(TINGGI_SEMAR * rasio)))
           .replace("__FHK__", str(round(TINGGI_SEMAR * .7))).replace("__FWK__", str(round(TINGGI_SEMAR * .7 * rasio))))
    if terbuka:
        css += CSS_SEMAR_TRANSPARAN
    st.markdown(f"<style>{re.sub(chr(10) + '+', ' ', css)}</style>", unsafe_allow_html=True)
    with st.container(key="semar_fab"):
        st.button("Tutup SEMAR" if terbuka else "Buka SEMAR", key="semar_fab_btn",
                  on_click=_tutup_chat if terbuka else _buka_chat,
                  help="Klik untuk menutup chat" if terbuka else "Klik untuk bertanya atau mengatur dashboard lewat SEMAR")
    if not asli:
        st.sidebar.caption(f"Gambar {GAMBAR_SEMAR.name} tidak ditemukan di folder SEMAR, jadi dipakai gambar Semar bawaan.")


# ===========================================================================
# SUARA SEMAR: teks -> ucapan lewat Web Speech API di peramban
# ===========================================================================
def teks_untuk_suara(teks: str) -> str:
    """Ubah jawaban (markdown) menjadi teks yang enak diucapkan: tanpa tanda markdown, kode rujukan, tautan."""
    t = H.unescape(str(teks or ""))
    t = re.sub(r"```.*?```", " ", t, flags=re.S)
    t = re.sub(r"<[^>]+>", " ", t)
    t = re.sub(r"!\[[^\]]*\]\([^)]*\)", " ", t)                            # gambar markdown
    t = re.sub(r"\[([^\]]+)\]\(https?://[^)]*\)", r"\1", t)                # tautan markdown -> teksnya saja
    t = re.sub(r"https?://\S+", " ", t)                                    # alamat web telanjang
    t = re.sub(r"\[[^\[\]\s]{1,40}(?:\s*[,;]\s*[^\[\]\s]{1,40})*\]", " ", t)   # rujukan [K-001] [W1-2] [Data]
    t = re.sub(r"(?m)^\s{0,3}#{1,6}\s+", "", t)                            # judul markdown
    t = re.sub(r"#(\d+)", r"nomor \1", t)                                  # "#2 dari 6" -> "nomor 2 dari 6"
    t = re.sub(r"\*{1,3}([^*\n]+)\*{1,3}", r"\1", t)                       # tebal/miring
    t = re.sub(r"(?<!\w)_{1,2}([^_\n]+)_{1,2}(?!\w)", r"\1", t)
    t = t.replace("*", " ").replace("`", " ")
    t = re.sub(r"(?<!\w)([+\-−])(\d(?:[\d.,]*\d)?)\s*%",
               lambda m: ("plus " if m.group(1) == "+" else "minus ") + m.group(2) + " persen", t)
    t = re.sub(r"%", " persen", t)

    def _rupiah(m: re.Match) -> str:
        satuan = {"T": "triliun", "M": "miliar", "jt": "juta"}.get((m.group(2) or "").strip(), (m.group(2) or "").strip())
        return f"{m.group(1)} {satuan} rupiah".replace("  ", " ")
    t = re.sub(r"\bRp\.?\s*(-?\d(?:[\d.,]*\d)?)(?:\s*(triliun|miliar|juta|ribu|T|M|jt)\b)?", _rupiah, t)
    t = re.sub(r"\bvs\.?(?=\s)", "dibandingkan dengan", t, flags=re.I)
    t = re.sub(r"\bKab\.\s*", "Kabupaten ", t)
    t = re.sub(r"(?<=\w)\s*–\s*(?=\w)", " sampai ", t)                # Januari–Juni
    t = re.sub(r"\s*—\s*", ", ", t)
    t = re.sub(r"\bSEMAR\b", "Semar", t)                                   # agar tidak dieja huruf demi huruf
    t = re.sub(r"\b(?:ha){2,}h?\b", " ", t, flags=re.I)                   # tawa ("hahaha") tidak diucapkan
    t = t.replace("|", ", ")
    t = re.sub(r"\s*\(([^()]*)\)", r", \1,", t)                            # (keterangan) -> , keterangan,
    t = re.sub(r"[^\w\s.,;:!?'\"%+\-/]", " ", t)                           # emoji dan simbol lain

    baris = []
    for b in t.splitlines():
        b = re.sub(r"^\s*(?:[-+•]|\d{1,2}[.)])\s+", "", b).strip(" \t,;")
        if not b:
            continue
        baris.append(b if b[-1] in ".!?:" else b + ".")
    t = " ".join(baris)
    t = re.sub(r",\s*([.!?;:])", r"\1", t)
    t = re.sub(r"(?:,\s*){2,}", ", ", t)
    t = re.sub(r"\s+([,.!?;:])", r"\1", t)
    t = re.sub(r"\.{2,}", ".", t)
    return re.sub(r"\s+", " ", t).strip()


def pecah_kalimat(teks: str, maks: int = 170, batas: int = MAKS_HURUF_SUARA) -> list[str]:
    """Potong teks menjadi bagian pendek per kalimat. Peramban (terutama Chrome) bisa berhenti
    di tengah jika satu ucapan terlalu panjang, jadi tiap bagian dibuat singkat dan diantrekan."""
    kalimat: list[str] = []
    for k in re.split(r"(?<=[.!?;:])\s+", teks.strip()):
        while len(k) > maks:                                   # kalimat panjang: potong di koma, lalu di spasi
            pos = max(k.rfind(", ", 0, maks), k.rfind(" ", 0, maks))
            pos = pos if pos > 40 else maks
            kalimat.append(k[:pos + 1].strip())
            k = k[pos + 1:].strip()
        if k:
            kalimat.append(k)
    bagian: list[str] = []
    total = 0
    for k in kalimat:
        if total + len(k) > batas and bagian:
            break
        total += len(k) + 1
        if bagian and len(bagian[-1]) + len(k) + 1 <= maks:
            bagian[-1] += " " + k
        else:
            bagian.append(k)
    return bagian


@st.cache_resource
def _status_edge() -> dict:
    return {"gagal_sampai": 0.0, "pesan": "", "pernah_ok": False}


@st.cache_data(show_spinner=False, max_entries=64, ttl=6 * 3600)
def _mp3_edge(teks: str, suara: str, nada: str, laju: str, batas_detik: float) -> str:
    """Buat suara Ardi (edge-tts) -> data-URI mp3. Melempar galat bila gagal (galat tidak disimpan di cache)."""
    import edge_tts

    async def _ambil() -> bytes:
        keluar = bytearray()
        async for p in edge_tts.Communicate(teks, suara, rate=laju, pitch=nada).stream():
            if p["type"] == "audio":
                keluar += p["data"]
        return bytes(keluar)

    data = asyncio.run(asyncio.wait_for(_ambil(), batas_detik))
    if not data:
        raise RuntimeError("edge-tts mengembalikan audio kosong")
    return "data:audio/mpeg;base64," + base64.b64encode(data).decode()


def suara_server(teks: str, batas_detik: float = 20.0) -> str | None:
    """Suara pria Indonesia dari server, atau None bila tidak tersedia (lalu peramban yang bicara).
    Setelah gagal, server tidak dicoba lagi selama 2 menit agar aplikasi tidak lambat saat offline."""
    st_ = _status_edge()
    if not MESIN_EDGE or not teks.strip() or time.time() < st_["gagal_sampai"]:
        return None
    try:
        uri = _mp3_edge(teks, EDGE_SUARA, EDGE_NADA, EDGE_LAJU, batas_detik)
        st_["pernah_ok"], st_["pesan"] = True, ""
        return uri
    except ImportError:
        st_["gagal_sampai"], st_["pesan"] = time.time() + 120, "pustaka edge-tts belum terpasang"
        print("  [suara] pustaka edge-tts belum terpasang (jalankan: python -m pip install edge-tts)")
    except Exception as e:  # noqa: BLE001  (offline, diblokir, waktu habis, dll.)
        st_["gagal_sampai"], st_["pesan"] = time.time() + 120, "tidak bisa terhubung ke server suara (periksa internet)"
        print(f"  [suara] GAGAL membuat suara Ardi: {type(e).__name__}: {e}")
    return None


# Modul JavaScript yang dipasang SEKALI di halaman utama (bukan di iframe), sehingga tetap hidup
# walau Streamlit menggambar ulang halaman. __CFG__ diganti dengan pengaturan JSON.
JS_SUARA = r"""
(function () {
  var W = window, D = document, CFG = __CFG__;
  var lama = W.__semarSuara;
  if (lama && lama.bongkar) { try { lama.bongkar(); } catch (e) {} }
  var TS = W.speechSynthesis;
  var S = { versi: CFG.versi, aktif: true, nomor: 0, antrean: [], audio: null };

  function pilihSuara() {
    var v = [];
    try { v = TS.getVoices() || []; } catch (e) {}
    var kecil = function (x) { return String(x || '').toLowerCase(); };
    if (CFG.suaraNama) {
      var khusus = v.filter(function (x) { return kecil(x.name).indexOf(kecil(CFG.suaraNama)) >= 0; })[0];
      if (khusus) return khusus;
    }
    var id = v.filter(function (x) { return /^id([-_]|$)/i.test(x.lang || '') || /indonesia/i.test(x.name || ''); });
    if (!id.length) return null;
    var pria = id.filter(function (x) { return /\b(ardi|andika|male|pria)\b/i.test(x.name || ''); })[0];
    return pria || id[0];
  }

  S.pilihSuara = pilihSuara;
  S.berhenti = function () {
    S.nomor++; S.antrean = [];
    if (S.audio) { try { S.audio.pause(); } catch (e) {} S.audio = null; }
    if (TS) { try { TS.cancel(); } catch (e) {} }
  };

  // Suara cadangan: bacaan bawaan peramban (Web Speech API).
  function bicaraPeramban(bagian) {
    if (!TS || !S.aktif || !bagian || !bagian.length) return;
    S.berhenti();
    var no = S.nomor, v = pilihSuara();
    var daftar = bagian.map(function (t) {
      var u = new W.SpeechSynthesisUtterance(t);
      if (v) { try { u.voice = v; } catch (e) {} u.lang = v.lang || 'id-ID'; } else { u.lang = 'id-ID'; }
      u.pitch = CFG.nada; u.rate = CFG.laju; u.volume = 1;
      return u;
    });
    S.antrean = daftar;                     // simpan rujukan agar tidak dibersihkan sebelum selesai
    setTimeout(function () {                // jeda singkat setelah cancel(); Chrome kadang menelan ucapan baru
      if (no !== S.nomor) return;
      try { TS.resume(); } catch (e) {}
      daftar.forEach(function (u) { TS.speak(u); });
    }, 80);
  }

  // Suara utama: mp3 pria dari server (audio). Bila gagal diputar, pakai suara peramban.
  S.bicara = function (bagian, audio) {
    if (!S.aktif) return;
    if (!audio) { bicaraPeramban(bagian); return; }
    S.berhenti();
    var no = S.nomor, a = new W.Audio(audio);
    S.audio = a;
    var cadangan = function () { if (no === S.nomor) { S.audio = null; bicaraPeramban(bagian); } };
    a.onerror = cadangan;
    try { var pr = a.play(); if (pr && pr.catch) pr.catch(cadangan); } catch (e) { cadangan(); }
  };

  S.atur = function (aktif) { S.aktif = !!aktif; if (!S.aktif) S.berhenti(); };

  // Klik gambar Semar melayang (buka/tutup) dan tombol "Tutup": sapa langsung saat klik,
  // tanpa menunggu Streamlit menggambar ulang (suara hanya boleh keluar setelah ada klik pengguna).
  S.klik = function (e) {
    var el = e.target && e.target.closest ? e.target.closest('.st-key-semar_fab button, .st-key-tutup_chat button') : null;
    if (!el || !S.aktif) return;
    var buka = !!el.closest('.st-key-semar_fab') && /buka/i.test(el.textContent || '');
    S.bicara(buka ? CFG.sapaBuka : CFG.sapaTutup, buka ? CFG.sapaBukaAudio : CFG.sapaTutupAudio);
  };
  D.addEventListener('click', S.klik, true);

  S.bongkar = function () { D.removeEventListener('click', S.klik, true); S.berhenti(); };
  if (TS) { try { TS.getVoices(); } catch (e) {} }   // memancing peramban memuat daftar suara
  W.__semarSuara = S;
})();
"""

# st.html bisa menjalankan JavaScript langsung di halaman (Streamlit baru). Streamlit lama memakai iframe.
_ST_HTML_JS = "unsafe_allow_javascript" in inspect.signature(st.html).parameters


def _cfg_suara() -> dict:
    cfg = {
        "nada": SUARA_NADA, "laju": SUARA_LAJU, "suaraNama": SUARA_NAMA,
        "sapaBuka": pecah_kalimat(teks_untuk_suara(SAPA_BUKA)),
        "sapaTutup": pecah_kalimat(teks_untuk_suara(SAPA_TUTUP)),
    }
    cfg["sapaBukaAudio"] = suara_server(" ".join(cfg["sapaBuka"]), 10.0)
    cfg["sapaTutupAudio"] = suara_server(" ".join(cfg["sapaTutup"]), 10.0)
    cfg["versi"] = hashlib.md5((JS_SUARA + json.dumps(cfg, sort_keys=True)).encode("utf-8")).hexdigest()[:10]
    return cfg


def _js_aman(obj) -> str:
    """JSON untuk disisipkan di dalam <script>: mencegah penutupan tag lebih awal."""
    return json.dumps(obj).replace("</", "<\\/").replace("<!--", "<\\!--")


def _jalankan_js(kode: str) -> None:
    badan = f"<script>{kode}</script>"
    if _ST_HTML_JS:
        st.html(badan, unsafe_allow_javascript=True)
    else:
        import streamlit.components.v1 as components
        components.html(badan, height=0)


def _skrip_suara(aktif: bool, ucap: list[str] | None = None, audio: str | None = None, nomor: int = 0) -> str:
    cfg = _cfg_suara()
    modul = JS_SUARA.replace("__CFG__", _js_aman(cfg))
    panggil = f"P.__semarSuara.atur({_js_aman(bool(aktif))});"
    if ucap:
        panggil += f"P.__semarSuara.bicara({_js_aman(ucap)},{_js_aman(audio)});"
    return (
        f"/* suara-{nomor} */(function(){{var P=window.frameElement?window.parent:window;try{{"
        f"if(!P.__semarSuara||P.__semarSuara.versi!=={_js_aman(cfg['versi'])}){{"
        f"var s=P.document.createElement('script');s.textContent={_js_aman(modul)};"
        f"P.document.head.appendChild(s);s.remove();}}{panggil}"
        f"}}catch(e){{console.warn('SEMAR suara:',e);}}}})();"
    )


def ucapkan(teks: str) -> None:
    """Antrekan jawaban untuk diucapkan; dibacakan saat render_ucapan() dipanggil."""
    if not st.session_state.get("suara_aktif", True):
        return
    bagian = pecah_kalimat(teks_untuk_suara(teks))
    if bagian:
        with st.spinner("SEMAR menyiapkan suara..."):
            audio = suara_server(" ".join(bagian))
        st.session_state["_ucapan"] = (bagian, audio)


def render_ucapan() -> None:
    """Kirim ucapan yang tertunda ke peramban (sekali saja, lalu dikosongkan)."""
    ucapan = st.session_state.pop("_ucapan", None)
    if not ucapan:
        return
    bagian, audio = ucapan
    st.session_state["_nomor_ucapan"] = st.session_state.get("_nomor_ucapan", 0) + 1
    with st.container(key="bunyi_ucap"):
        _jalankan_js(_skrip_suara(st.session_state.get("suara_aktif", True), bagian, audio, st.session_state["_nomor_ucapan"]))


def pasang_suara(aktif: bool) -> None:
    """Pasang modul suara dan selaraskan status saklarnya (isi tidak berubah = tidak dijalankan ulang)."""
    with st.container(key="bunyi_atur"):
        _jalankan_js(_skrip_suara(aktif))


# Streamlit mengingat keadaan sidebar di peramban: begitu pernah ditutup, pengaturan
# initial_sidebar_state="expanded" tidak lagi dipakai. Skrip ini membuka sidebar sekali
# saja di awal sesi (layar lebar saja; di ponsel Streamlit sengaja menutupnya).
JS_SIDEBAR_AWAL = r"""
(function () {
  var P = window.frameElement ? window.parent : window;
  try {
    if (P.innerWidth <= 768) return;
    var D = P.document, coba = 0;
    var t = P.setInterval(function () {
      coba++;
      var tombol = D.querySelector('[data-testid="stExpandSidebarButton"], [data-testid="stSidebarCollapsedControl"] button');
      if (tombol) { tombol.click(); P.clearInterval(t); return; }   // tombol ini hanya ada bila sidebar tertutup
      if (coba >= 30) P.clearInterval(t);          // menyerah setelah +-4,5 detik
    }, 150);
  } catch (e) { console.warn('SEMAR sidebar:', e); }
})();
"""


# Tombol cadangan "Tampilkan sidebar". Muncul melayang di kiri atas HANYA bila sidebar tertutup
# dan tombol bawaan Streamlit tidak terlihat (versi Streamlit berbeda-beda tempat/penampakan tombolnya).
# Dipasang sekali di halaman induk, jadi tetap ada walau skrip Python berjalan ulang.
JS_TOMBOL_SIDEBAR = r"""
(function () {
  var P = window.frameElement ? window.parent : window;
  try {
    var VERSI = '2';
    if (P.__semarTombolSidebar && P.__semarTombolSidebar.versi === VERSI) return;
    if (P.__semarTombolSidebar) { try { P.__semarTombolSidebar.bersih(); } catch (e) {} }
    var D = P.document;
    var SEL_NATIF = '[data-testid="stExpandSidebarButton"], [data-testid="stSidebarCollapsedControl"] button';

    var gaya = D.createElement('style');
    gaya.id = 'semar-gaya-tombol-sidebar';
    gaya.textContent =
      '#semar-tombol-sidebar{position:fixed;top:14px;left:14px;z-index:999980;width:46px;height:46px;border-radius:50%;' +
      'border:2px solid rgba(255,255,255,.85);cursor:pointer;display:none;align-items:center;justify-content:center;padding:0;' +
      'background:linear-gradient(135deg,#F58220,#F7B733);color:#1B1204;' +
      'box-shadow:0 10px 22px -8px rgba(245,130,32,.95),0 0 0 3px rgba(10,31,68,.12);transition:transform .2s ease,filter .2s ease;}' +
      '#semar-tombol-sidebar.tampak{display:flex;animation:semarTombolMuncul .35s ease both;}' +
      '#semar-tombol-sidebar:hover{transform:translateY(-2px) scale(1.06);filter:brightness(1.07);}' +
      '#semar-tombol-sidebar:active{transform:scale(.95);}' +
      '#semar-tombol-sidebar:focus-visible{outline:3px solid #0A1F44;outline-offset:3px;}' +
      '#semar-tombol-sidebar svg{width:24px;height:24px;display:block;}' +
      '@keyframes semarTombolMuncul{from{opacity:0;transform:translateX(-12px)}to{opacity:1;transform:none}}' +
      '@media (max-width:768px){#semar-tombol-sidebar{top:8px;left:8px;width:42px;height:42px;}' +
      'html.semar-sidebar-tutup .block-container{padding-top:3.8rem!important;}}' +
      '@media (prefers-reduced-motion:reduce){#semar-tombol-sidebar{animation:none!important;transition:none!important;}}';
    D.head.appendChild(gaya);

    var tb = D.createElement('button');
    tb.id = 'semar-tombol-sidebar';
    tb.type = 'button';
    tb.title = 'Tampilkan sidebar';
    tb.setAttribute('aria-label', 'Tampilkan sidebar');
    tb.innerHTML =
      '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' +
      '<rect x="3" y="4" width="18" height="16" rx="3"></rect><path d="M9 4v16"></path><path d="M13.5 9.5L16 12l-2.5 2.5"></path></svg>';
    D.body.appendChild(tb);

    function natif() {
      var n = D.querySelector(SEL_NATIF);
      if (n) return n;
      // Streamlit baru: cari tombol "open/expand/show sidebar" lewat label aksesibilitasnya.
      var semua = D.querySelectorAll('button[aria-label], button[title]');
      for (var i = 0; i < semua.length; i++) {
        var t = (semua[i].getAttribute('aria-label') || semua[i].getAttribute('title') || '');
        if (semua[i] !== tb && /(open|expand|show|buka|tampilkan).{0,12}(sidebar|side bar|bilah)/i.test(t)) return semua[i];
      }
      return null;
    }

    function tampak(el) {
      if (!el) return false;
      var r = el.getBoundingClientRect(), cs = P.getComputedStyle(el);
      return r.width > 0 && r.height > 0 && cs.visibility !== 'hidden' && cs.display !== 'none' && parseFloat(cs.opacity || '1') > 0.05;
    }

    function sidebarTertutup() {
      var sb = D.querySelector('[data-testid="stSidebar"]');
      if (!sb) return !!natif();                              // sidebar tak ada di DOM: tertutup bila tombol buka bawaan ada
      var ae = sb.getAttribute('aria-expanded');
      if (ae === 'false') return true;
      if (ae === 'true') return false;
      if (natif()) return true;                               // tanpa aria-expanded: tombol buka bawaan hanya ada saat tertutup
      var r = sb.getBoundingClientRect();
      return r.width < 8 || r.right <= 4;
    }

    function perbarui() {
      var tutup = sidebarTertutup();
      D.documentElement.classList.toggle('semar-sidebar-tutup', tutup);
      // Tombol bawaan yang terlihat dipakai apa adanya; tombol cadangan hanya bila bawaannya tak terlihat.
      tb.classList.toggle('tampak', tutup && !tampak(natif()));
    }

    tb.addEventListener('click', function () {
      var n = natif();
      if (n) { n.click(); }
      else {
        // Cadangan terakhir: cari tombol apa pun yang menyebut sidebar.
        /* tidak ada tombol bawaan yang bisa diklik */
      }
      P.setTimeout(perbarui, 120);
      P.setTimeout(perbarui, 450);
    });

    var timer = P.setInterval(perbarui, 400);
    perbarui();

    P.__semarTombolSidebar = {
      versi: VERSI,
      bersih: function () {
        P.clearInterval(timer);
        D.documentElement.classList.remove('semar-sidebar-tutup');
        if (tb.parentNode) tb.parentNode.removeChild(tb);
        if (gaya.parentNode) gaya.parentNode.removeChild(gaya);
      }
    };
  } catch (e) { console.warn('SEMAR tombol sidebar:', e); }
})();
"""


def buka_sidebar_awal() -> None:
    """Buka sidebar sekali di awal sesi, dan pasang tombol cadangan "Tampilkan sidebar".

    Putaran berikutnya tidak menyentuhnya, jadi tombol tutup tetap bebas dipakai."""
    if st.session_state.get("_sidebar_dibuka"):
        return
    st.session_state["_sidebar_dibuka"] = True
    with st.container(key="bunyi_sidebar"):
        _jalankan_js(JS_TOMBOL_SIDEBAR + JS_SIDEBAR_AWAL)


# ===========================================================================
# KOMPONEN TAMPILAN
# ===========================================================================
def tampil_hero(kb: S.BasisPengetahuan, kompak: bool = False):
    st.markdown(padatkan(f"""
<div class="hero{' kompak' if kompak else ''}">
  <div class="hero-batik"></div><div class="hero-cahaya"></div>
  <div class="hero-skyline">{SKYLINE_SVG}</div>
  <div class="hero-isi">
    <div class="hero-semar">{ikon_semar_html()}</div>
    <div class="hero-teks">
      <span class="hero-label">Bea Cukai &middot; Kota Semarang</span>
      <h1 class="hero-judul">SEMAR</h1>
      <p class="hero-sub">Ngemong, nuntun, njawab &mdash; hanya dari sumber yang Anda percaya.</p>
      <div class="hero-pil">
        <span>{len(kb.entri)} baris pengetahuan</span>
        <span>{len(kb.situs)} website</span>
      </div>
    </div>
  </div>
</div>"""), unsafe_allow_html=True)


def lencana_rujukan(daftar: list[str]) -> str:
    if not daftar:
        return ""
    chip = []
    for r in daftar:
        m = re.match(r"^\[(.+?)\]\s+(.*)$", r)
        kode, ket = (m.group(1), m.group(2)) if m else ("", r)
        sumber = ket.split(" - ")[0].strip()
        if sumber.lower().startswith(("http://", "https://")):
            host = urllib.parse.urlparse(sumber).netloc or sumber
            chip.append(f'<a class="rujuk" href="{H.escape(sumber, quote=True)}" target="_blank" rel="noopener noreferrer">'
                        f'<b>{H.escape(kode)}</b>{H.escape(host)}</a>')
        else:
            chip.append(f'<span class="rujuk"><b>{H.escape(kode)}</b>{H.escape(ket)}</span>')
    return f'<div class="rujuk-baris"><span class="rujuk-judul">Rujukan</span>{"".join(chip)}</div>'


def ketik(teks: str):
    """Efek mengetik: jawaban muncul kata demi kata."""
    for bagian in re.split(r"(\s+)", teks):
        yield bagian
        if bagian.strip():
            time.sleep(0.018)


@st.dialog("Isi basis pengetahuan", width="large")
def dialog_basis(kb: S.BasisPengetahuan):
    st.markdown("**Tabel Pengetahuan** (dari Excel)")
    if kb.entri:
        st.dataframe(
            pd.DataFrame([{"ID": d.kode, "Kategori": d.kategori, "Topik": d.judul,
                           "Isi Pengetahuan": d.teks, "Sumber": d.sumber} for d in kb.entri]),
            width="stretch", hide_index=True)
    else:
        st.info("Belum ada baris pengetahuan aktif. Isi sheet 'Pengetahuan' di file Excel.")
    st.markdown("**Website yang boleh dibaca**")
    if kb.situs:
        st.dataframe(
            pd.DataFrame([{"No": s["no"], "Nama": s["nama"], "URL": s["url"],
                           "Potongan teks terbaca": len(kb.cache_web.get(s["url"], (0, []))[1])} for s in kb.situs]),
            width="stretch", hide_index=True)
        st.caption("Angka 0 berarti halaman gagal dibaca. Lihat jendela hitam (Command Prompt) untuk alasannya.")
    else:
        st.info("Belum ada website aktif. Isi sheet 'Website' di file Excel dan set Aktif = Ya.")


# ===========================================================================
# PERSIAPAN AGEN
# ===========================================================================
@st.cache_resource
def kb_bersama() -> S.BasisPengetahuan:
    """Satu basis pengetahuan dipakai bersama agar website tidak dibaca berulang."""
    return S.BasisPengetahuan(S.EXCEL_PATH)


def siapkan_agen() -> S.Semar:
    try:
        agen = S.Semar()
    except SystemExit as e:  # API key belum terbaca
        st.error("API key DeepSeek belum terbaca, jadi SEMAR belum bisa dipakai.")
        st.code(str(e.code), language="text")
        st.info("Perbaiki file .env, simpan, lalu muat ulang halaman ini (tekan F5).")
        st.stop()
    agen.kb = kb_bersama()
    return agen


terapkan_tema()

# Halaman yang bisa dipilih di sidebar: (nama, ikon)
HALAMAN = [
    ("Dashboard", ":material/space_dashboard:"),
    ("Semar", ":material/forum:"),
    ("Pemeriksaan Barang", ":material/inventory_2:"),
]
NAMA_HALAMAN = [h[0] for h in HALAMAN]


def _pindah_halaman(nama: str) -> None:
    st.session_state["mode"] = nama


TINGGI_CHAT = 600
CONTOH_PERINTAH = [
    "Tampilkan penindakan",
    "Fokus Kendal dan Demak",
    "Bandingkan 2026 dengan 2025",
    "Tampilkan peringkat wilayah",
    "Triwulan 2 tahun 2025",
    "Reset dashboard",
]
LABEL_METRIK = {"Penerimaan negara": "penerimaan", "Penindakan pengawasan": "penindakan"}
LABEL_GRAFIK = {"Tren bulanan": "tren", "Komposisi jenis": "komposisi", "Peringkat wilayah": "peringkat"}


@st.cache_resource
def data_bersama() -> DB.DataDashboard:
    return DB.DataDashboard(DATA_PATH)


# ---------------------------------------------------------------------------
# Kondisi dashboard <-> widget manual (kondisi = sumber kebenaran tunggal)
# ---------------------------------------------------------------------------
def sinkron_widget(k: DB.Kondisi) -> None:
    s = st.session_state
    s["w_metrik"] = "Penerimaan negara" if k.metrik == "penerimaan" else "Penindakan pengawasan"
    s["w_tahun"] = k.tahun
    s["w_banding"] = k.banding if k.banding else "Tanpa perbandingan"
    s["w_bulan_awal"] = k.bulan_awal
    s["w_bulan_akhir"] = k.bulan_akhir
    s["w_wilayah"] = list(k.wilayah)
    for m in ("penerimaan", "penindakan"):
        s[f"w_jenis_{m}"] = list(k.jenis) if k.metrik == m else []
    s["w_grafik"] = next(a for a, b in LABEL_GRAFIK.items() if b == k.grafik)


def dari_widget(data: DB.DataDashboard) -> None:
    s = st.session_state
    lama: DB.Kondisi = s.kondisi
    metrik = LABEL_METRIK.get(s.get("w_metrik"), lama.metrik)
    tahun_ok = data.tahun_ada(metrik) or data.tahun_ada()
    tahun = s.get("w_tahun", lama.tahun)
    if tahun not in tahun_ok and tahun_ok:
        tahun = max(tahun_ok)
    banding = s.get("w_banding")
    banding = banding if isinstance(banding, int) and banding in tahun_ok and banding != tahun else None
    awal = s.get("w_bulan_awal", lama.bulan_awal)
    akhir = s.get("w_bulan_akhir", lama.bulan_akhir)
    if awal > akhir:
        awal, akhir = akhir, awal
    bulan_diubah = (awal, akhir) != (lama.bulan_awal, lama.bulan_akhir)
    manual = True if bulan_diubah else lama.bulan_manual
    if tahun != lama.tahun and not manual:          # ganti tahun tanpa menyentuh bulan: ikuti data tahun itu
        awal, akhir = 1, data.bulan_terakhir(metrik, tahun)
    s.kondisi = DB.Kondisi(
        metrik=metrik, tahun=tahun, banding=banding, bulan_awal=awal, bulan_akhir=akhir,
        wilayah=list(s.get("w_wilayah", [])), jenis=list(s.get(f"w_jenis_{metrik}", [])),
        grafik=LABEL_GRAFIK.get(s.get("w_grafik"), lama.grafik),
        bulan_manual=manual)


# ---------------------------------------------------------------------------
# Persiapan
# ---------------------------------------------------------------------------
@st.cache_resource
def kb_bersama() -> S.BasisPengetahuan:
    """Satu basis pengetahuan dipakai bersama agar website tidak dibaca berulang."""
    return S.BasisPengetahuan(S.EXCEL_PATH)


def siapkan_agen() -> S.Semar:
    try:
        agen = S.Semar()
    except SystemExit as e:  # API key belum terbaca
        st.error("API key DeepSeek belum terbaca, jadi SEMAR belum bisa dipakai.")
        st.code(str(e.code), language="text")
        st.info("Perbaiki file .env, simpan, lalu muat ulang halaman ini (tekan F5).")
        st.stop()
    agen.kb = kb_bersama()
    return agen


if "agen" not in st.session_state:
    st.session_state.agen = siapkan_agen()
    st.session_state.pesan = []
    st.session_state.siap = False
    st.session_state.dash_lanjutan = False
    st.session_state.setdefault("mode", "Dashboard")
st.session_state.setdefault("chat_buka", False)    # kolom chat tertutup saat aplikasi dibuka

if "_mode_baru" in st.session_state:          # perpindahan mode diminta pada run sebelumnya
    st.session_state["mode"] = st.session_state.pop("_mode_baru")
st.session_state.setdefault("mode", "Dashboard")
if st.session_state["mode"] not in NAMA_HALAMAN:     # mis. nama lama "Percakapan"
    st.session_state["mode"] = "Semar" if st.session_state["mode"] == "Percakapan" else "Dashboard"

agen: S.Semar = st.session_state.agen
kb = agen.kb
data = data_bersama()
data.muat()

if not st.session_state.siap:
    with st.spinner("Memuat basis pengetahuan (Excel dan website)..."):
        try:
            kb.segarkan(paksa=True)
            st.session_state.siap = True
        except Exception as e:  # noqa: BLE001
            st.error(f"Gagal memuat file Excel: {e}")
            st.stop()

if "kondisi" not in st.session_state:
    st.session_state.kondisi = DB.kondisi_awal(data)
else:                                           # data Excel bisa berubah: pastikan filter masih sah
    k_ = st.session_state.kondisi
    th_ = data.tahun_ada(k_.metrik) or data.tahun_ada()
    if th_ and k_.tahun not in th_:
        st.session_state.kondisi = DB.kondisi_awal(data)


# ===========================================================================
# SIDEBAR
# ===========================================================================
with st.sidebar:
    st.markdown(padatkan(f"""
<div class="side-merek"><div class="logo">{ikon_semar_html()}</div>
<div><div class="nama">SEMAR</div><div class="kecil">Pamong informasi</div></div></div>
<div class="stat-grid">
  <div class="stat"><div class="angka">{len(kb.entri)}</div><div class="ket">baris pengetahuan</div></div>
  <div class="stat"><div class="angka">{len(kb.situs)}</div><div class="ket">website aktif</div></div>
</div>
"""), unsafe_allow_html=True)

    st.markdown('<div class="sub-judul">Halaman</div>', unsafe_allow_html=True)
    mode = st.session_state["mode"]
    with st.container(key="nav_halaman"):
        for _nama, _ikon in HALAMAN:
            st.button(_nama, key=f"nav_{_nama}", icon=_ikon, width="stretch",
                      type="primary" if _nama == mode else "secondary",
                      on_click=_pindah_halaman, args=(_nama,))
    st.markdown('<div class="sub-judul">Pengaturan</div>', unsafe_allow_html=True)
    efek_ketik = st.toggle("Efek mengetik", value=True)
    suara_aktif = st.toggle("Suara SEMAR", value=True, key="suara_aktif",
                            help="SEMAR membacakan jawaban dan menyapa saat gambar Semar diklik. "
                                 "Matikan untuk menghentikan suara seketika.")
    if suara_aktif:
        _st = _status_edge()
        # Keterangan jenis suara disembunyikan; hanya masalah (bila ada) yang tetap ditampilkan.
        if MESIN_EDGE and _st["pesan"]:
            st.caption(f"Suara pria (Ardi) belum bisa dipakai: {_st['pesan']}. Sementara dipakai suara bawaan peramban.")

    st.markdown('<div class="sub-judul">Aksi</div>', unsafe_allow_html=True)
    if st.button("Muat ulang Excel dan website", width="stretch"):
        with st.spinner("Membaca ulang..."):
            kb.segarkan(paksa=True)
            data.muat(paksa=True)
        st.rerun()
    if st.button("Lihat isi basis pengetahuan", width="stretch", type="secondary"):
        dialog_basis(kb)
    if st.button("Hapus percakapan", width="stretch", type="secondary"):
        agen.riwayat.clear()
        st.session_state.pesan = []
        st.session_state.dash_lanjutan = False
        st.rerun()


AVATAR_SEMAR = avatar_semar()


# ===========================================================================
# LOGIKA JAWABAN: perintah dashboard ATAU pertanyaan basis pengetahuan
# ===========================================================================
def jawab_gabungan(pertanyaan: str) -> tuple[str, list[str], bool]:
    s = st.session_state
    niat = DB.pahami(pertanyaan, data, s.kondisi, klien=agen.klien, model=S.MODEL,
                     extra_body=S.EXTRA_BODY, lanjutan=s.dash_lanjutan)
    if niat is None:                                   # bukan perintah dashboard -> basis pengetahuan
        s.dash_lanjutan = False
        jawaban, rujukan = agen.jawab(pertanyaan)
        return jawaban, rujukan, False
    s.dash_lanjutan = True
    if niat.bantuan:
        agen._simpan(pertanyaan, DB.TEKS_BANTUAN)
        return DB.TEKS_BANTUAN, [], False
    k, cat = DB.terapkan(s.kondisi, niat, data)
    s.kondisi = k
    jawaban = DB.jelaskan(k, DB.ringkas(k, data), data, cat, niat)
    agen._simpan(pertanyaan, jawaban)
    lembar = "Penerimaan" if k.metrik == "penerimaan" else "Penindakan"
    return jawaban, [f"[Data] {data.path.name} - sheet {lembar}"], True


def tampil_riwayat() -> None:
    for m in st.session_state.pesan:
        with st.chat_message(m["role"], avatar=AVATAR_SEMAR if m["role"] == "assistant" else None):
            st.markdown(m["content"])
            if m.get("rujukan"):
                st.markdown(lencana_rujukan(m["rujukan"]), unsafe_allow_html=True)


def proses(pertanyaan: str) -> bool:
    """Tampilkan pertanyaan + jawaban (dipanggil di dalam wadah yang sesuai). True bila dashboard berubah."""
    st.session_state.pesan.append({"role": "user", "content": pertanyaan})
    with st.chat_message("user"):
        st.markdown(pertanyaan)
    with st.chat_message("assistant", avatar=AVATAR_SEMAR):
        galat = False
        with st.spinner("SEMAR sedang menelusuri sumber..."):
            try:
                jawaban, rujukan, dash_baru = jawab_gabungan(pertanyaan)
            except Exception as e:  # noqa: BLE001
                jawaban, rujukan, dash_baru = f"Terjadi kesalahan saat menghubungi DeepSeek: {e}", [], False
                galat = True
        if not galat:                                   # pesan galat teknis tidak dibacakan
            ucapkan(jawaban)
            if not (dash_baru and mode == "Semar"):
                render_ucapan()                         # mulai bicara bersamaan dengan efek mengetik
            # (tampilan Percakapan + perintah dashboard: layar diulang dulu; ucapan menyusul di akhir skrip)
        if efek_ketik and not dash_baru:
            st.write_stream(ketik(jawaban))
        else:
            st.markdown(jawaban)
        if rujukan:
            st.markdown(lencana_rujukan(rujukan), unsafe_allow_html=True)
    st.session_state.pesan.append({"role": "assistant", "content": jawaban, "rujukan": rujukan})
    return dash_baru


def tombol_cepat(daftar: list[str], kolom: int, awalan: str, awali: str = "") -> None:
    for baris in range(0, len(daftar), kolom):
        for kol, t in zip(st.columns(kolom), daftar[baris:baris + kolom]):
            label = t if len(t) <= 34 else t[:31] + "..."
            if kol.button(label, key=f"{awalan}_{baris}_{t}", width="stretch"):
                st.session_state.tanya_cepat = f"{awali}{t}"
                st.rerun()


# ===========================================================================
# TAMPILAN DASHBOARD (kolom kiri)
# ===========================================================================
def tampil_dashboard(lebar: bool = False) -> None:
    k: DB.Kondisi = st.session_state.kondisi
    sinkron_widget(k)
    r = DB.ringkas(k, data)
    st.markdown(DB.html_judul(k, data), unsafe_allow_html=True)
    for g in data.galat:
        st.warning(g)
    if not data.ada_data:
        st.info("Data dashboard belum ada. Pastikan file **data_dashboard_semar.xlsx** berada satu folder "
                "dengan aplikasi, lalu tekan *Muat ulang Excel dan website*.")
        return
    st.markdown(DB.html_kpi(k, r), unsafe_allow_html=True)

    def grafik() -> None:
        st.segmented_control("Grafik", list(LABEL_GRAFIK), key="w_grafik", on_change=dari_widget, args=(data,),
                             label_visibility="collapsed")
        st.markdown(DB.html_grafik(k, r), unsafe_allow_html=True)

    if lebar:                                       # dashboard penuh: peta di kiri, grafik di kanan
        kol_peta, kol_grafik = st.columns([5, 6], gap="large")
        kol_peta.markdown(DB.html_peta(k, r), unsafe_allow_html=True)
        with kol_grafik:
            grafik()
    else:                                           # berbagi ruang dengan chat: ditumpuk ke bawah
        st.markdown(DB.html_peta(k, r), unsafe_allow_html=True)
        grafik()

    with st.expander("Atur manual (tanpa chat)"):
        th = data.tahun_ada(k.metrik) or data.tahun_ada()
        c1, c2, c3 = st.columns(3)
        c1.segmented_control("Data", list(LABEL_METRIK), key="w_metrik", on_change=dari_widget, args=(data,))
        c2.selectbox("Tahun", th, key="w_tahun", on_change=dari_widget, args=(data,))
        c3.selectbox("Bandingkan dengan", ["Tanpa perbandingan"] + [t for t in th if t != k.tahun],
                     key="w_banding", on_change=dari_widget, args=(data,))
        c6, c7 = st.columns(2)
        c6.selectbox("Dari bulan", list(range(1, 13)), key="w_bulan_awal", format_func=lambda i: DB.BULAN[i - 1],
                     on_change=dari_widget, args=(data,))
        c7.selectbox("Sampai bulan", list(range(1, 13)), key="w_bulan_akhir", format_func=lambda i: DB.BULAN[i - 1],
                     on_change=dari_widget, args=(data,))
        c4, c5 = st.columns(2)
        c4.multiselect("Wilayah (kosong = semua)", DB.WILAYAH, key="w_wilayah", on_change=dari_widget, args=(data,))
        c5.multiselect("Jenis (kosong = semua)", data.jenis_ada(k.metrik), key=f"w_jenis_{k.metrik}",
                       on_change=dari_widget, args=(data,))
        if st.button("Kembalikan ke pengaturan awal", key="w_reset"):
            st.session_state.kondisi = DB.kondisi_awal(data)
            st.rerun()


# ===========================================================================
# HALAMAN 3: PEMERIKSAAN BARANG (OCR + AI)
# ===========================================================================
TIPE_GAMBAR = ["png", "jpg", "jpeg", "webp", "bmp", "tif", "tiff"]
SLOT_INFO = {
    "dokumen": dict(no=1, judul="Dokumen", tipe=["pdf"] + TIPE_GAMBAR,
                    sub="BC 2.5, packing list, dokumen pelengkap (PDF atau foto)"),
    "label": dict(no=2, judul="Foto label barang", tipe=TIPE_GAMBAR,
                  sub="Label pada roll, kemasan, atau barang"),
    "barang": dict(no=3, judul="Foto barang", tipe=TIPE_GAMBAR,
                   sub="Barang dan kemasannya, dari beberapa sisi"),
}


def tampil_hero_periksa() -> None:
    st.markdown(padatkan(f"""
<div class="hero kompak">
  <div class="hero-batik"></div><div class="hero-cahaya"></div>
  <div class="hero-skyline">{SKYLINE_SVG}</div>
  <div class="hero-isi">
    <div class="hero-semar">{ikon_semar_html()}</div>
    <div class="hero-teks">
      <span class="hero-label">Bea Cukai &middot; OCR dan AI untuk pemeriksa</span>
      <h1 class="hero-judul panjang">PEMERIKSAAN BARANG</h1>
      <p class="hero-sub">Foto atau unggah dokumen, label, dan barang. SEMAR membaca (OCR), menyusun baris sheet hasil pemeriksaan, dan menyiapkannya untuk diunduh.</p>
      <div class="hero-pil"><span>OCR dokumen</span><span>AI membaca foto barang</span><span>Sheet layar penuh</span></div>
    </div>
  </div>
</div>"""), unsafe_allow_html=True)


def _init_periksa() -> None:
    ss = st.session_state
    ss.setdefault("pb_baris", [])
    ss.setdefault("pb_seq", 0)
    ss.setdefault("pb_ver", 0)
    ss.setdefault("pb_kamera", {s: [] for s in SLOT_INFO})
    ss.setdefault("pb_kamver", {s: 0 for s in SLOT_INFO})


def _simpan_foto(slot: str) -> None:
    ss = st.session_state
    foto = ss.get(f"pb_cam_{slot}_{ss.pb_kamver[slot]}")
    if foto is None:
        return
    ss.pb_kamera[slot].append({"nama": f"kamera_{slot}_{time.strftime('%H%M%S')}_{len(ss.pb_kamera[slot]) + 1}.jpg",
                               "data": foto.getvalue()})
    ss.pb_kamver[slot] += 1           # kamera dikosongkan agar siap mengambil foto berikutnya


def _hapus_foto(slot: str, i: int) -> None:
    st.session_state.pb_kamera[slot].pop(i)


def _berkas_slot(slot: str) -> list[PM.Berkas]:
    ss = st.session_state
    unggahan = ss.get(f"pb_up_{slot}_{ss.pb_ver}") or []
    hasil = [PM.Berkas(slot, f.name, f.getvalue(), "unggah") for f in unggahan]
    hasil += [PM.Berkas(slot, f["nama"], f["data"], "kamera") for f in ss.pb_kamera[slot]]
    return hasil


def _slot_masukan(slot: str) -> int:
    ss, info = st.session_state, SLOT_INFO[slot]
    with st.container(key=f"slot_{slot}"):
        kepala = st.empty()
        st.file_uploader(info["judul"], type=info["tipe"], accept_multiple_files=True,
                         key=f"pb_up_{slot}_{ss.pb_ver}", label_visibility="collapsed")
        if st.toggle("Buka kamera", key=f"pb_kamtog_{slot}"):
            foto = st.camera_input(f"Foto {info['judul'].lower()}", key=f"pb_cam_{slot}_{ss.pb_kamver[slot]}",
                                   label_visibility="collapsed")
            if foto is not None:
                st.button("Simpan foto ini", key=f"pb_simpan_{slot}_{ss.pb_kamver[slot]}", type="primary",
                          icon=":material/add_a_photo:", on_click=_simpan_foto, args=(slot,))
            st.caption("Kamera hanya jalan di localhost atau https. Di HP, tombol unggah juga menawarkan kamera.")
        if ss.pb_kamera[slot]:
            kolom = st.columns(4)
            for i, f in enumerate(ss.pb_kamera[slot]):
                with kolom[i % 4]:
                    st.image(f["data"], width="stretch")
                    st.button("Hapus", key=f"pb_hapusfoto_{slot}_{i}_{len(ss.pb_kamera[slot])}",
                              icon=":material/close:", on_click=_hapus_foto, args=(slot, i))
        n = len(_berkas_slot(slot))
        label = f"{n} berkas" if n else "kosong"
        kepala.markdown(
            f'<div class="pb-slot-kepala"><span class="pb-no">{info["no"]}</span>'
            f'<div><div class="pb-slot-judul">{info["judul"]}</div><div class="pb-slot-sub">{info["sub"]}</div></div>'
            f'<span class="pb-hitung{" ada" if n else ""}">{label}</span></div>', unsafe_allow_html=True)
    return n


def _pengaturan_periksa() -> PM.Pengaturan:
    with st.expander("Pengaturan pemeriksaan", icon=":material/tune:"):
        tingkat = st.radio("Tingkat pemeriksaan (sesuai Instruksi Pemeriksaan)", [10, 30], horizontal=True,
                           format_func=lambda x: f"{x}%", key="pb_tingkat")
        kondisi = st.radio("Kondisi barang", ["Sesuai foto (AI)", "Baik", "Rusak"], horizontal=True, key="pb_kondisi")
        kem = st.text_input("Jumlah kemasan yang diperiksa (kosongkan = hitung otomatis)", key="pb_kem")
        catatan = st.text_area("Catatan pemeriksa (opsional)", key="pb_catatan", height=80,
                               placeholder="Mis. kegunaan barang, kondisi kemasan, atau hal yang tidak terlihat di foto")
        st.caption("Waktu bawaan mengikuti template pemeriksa (bongkar muat 09:00 sampai 10:00, pemeriksaan 10:01 sampai 11:01). Ubah bila berbeda.")
        a, b = st.columns(2)
        bm1 = a.time_input("Bongkar muat: mulai", value=PM.JAM_BAWAAN["bm_mulai"], step=60, key="pb_bm1")
        bm2 = b.time_input("Bongkar muat: selesai", value=PM.JAM_BAWAAN["bm_selesai"], step=60, key="pb_bm2")
        pr1 = a.time_input("Pemeriksaan: mulai", value=PM.JAM_BAWAAN["per_mulai"], step=60, key="pb_pr1")
        pr2 = b.time_input("Pemeriksaan: selesai", value=PM.JAM_BAWAAN["per_selesai"], step=60, key="pb_pr2")
    angka = PM.baca_angka(kem) if kem.strip() else None
    if kem.strip() and (angka is None or angka < 1):
        st.warning("Jumlah kemasan diperiksa tidak terbaca; dihitung otomatis.")
        angka = None
    return PM.Pengaturan(tingkat=tingkat, kondisi="" if kondisi.startswith("Sesuai") else kondisi.lower(),
                         catatan=catatan, kemasan_override=int(angka) if angka else None,
                         bm_mulai=bm1, bm_selesai=bm2, per_mulai=pr1, per_selesai=pr2)


def _proses_periksa(p: PM.Pengaturan) -> None:
    ss = st.session_state
    berkas = [b for s in SLOT_INFO for b in _berkas_slot(s)]
    try:
        with st.status("SEMAR sedang memproses...", expanded=True) as stt:
            hasil = PM.jalankan(agen, berkas, p, progres=stt.write)
            stt.update(label="Selesai. Baris baru sudah masuk ke sheet.", state="complete", expanded=False)
    except PM.GalatPemeriksaan as e:
        st.error(str(e))
        return
    except Exception as e:  # noqa: BLE001
        st.error(f"Terjadi kesalahan yang tidak terduga: {e}")
        return
    ss.pb_seq += 1
    hasil["id"], hasil["rev"] = ss.pb_seq, 0
    ss.pb_baris.append(hasil)
    ss.pb_pilih = hasil["id"]
    ss.pb_ver += 1                                   # kosongkan kolom unggah untuk dokumen berikutnya
    ss.pb_kamera = {s: [] for s in SLOT_INFO}
    for s in SLOT_INFO:
        ss.pb_kamver[s] += 1
    if ss.get("pb_auto_buka", True):
        ss.pb_buka_sheet = True
    st.rerun()


@st.cache_resource
def _komponen_sheet():
    """Komponen sheet interaktif (klik sel = salin, klik kolom V/X = centang). Gagal -> None (tampilan statis)."""
    try:
        import streamlit.components.v1 as components
        return components.declare_component("sheet_pemeriksaan", path=PM.tulis_komponen())
    except Exception:  # noqa: BLE001
        return None


def _sheet_interaktif(kunci: str, mode: str, in_dialog: bool = False) -> None:
    """Tampilkan sheet. mode 'penuh' = 90% tinggi layar aplikasi; 'inline' = 460 px. Klik centang diteruskan ke Python."""
    ss = st.session_state
    komponen = _komponen_sheet()
    if komponen is None:
        st.markdown(PM.html_sheet(ss.pb_baris, "460px"), unsafe_allow_html=True)
        return
    aksi = komponen(html=PM.html_sheet(ss.pb_baris), mode=mode, tinggi=460, chrome=DIALOG_CHROME, key=kunci, default=None)
    kunci_seq = f"pb_seq_{kunci}"
    if isinstance(aksi, dict) and aksi.get("seq", 0) > ss.get(kunci_seq, 0):
        ss[kunci_seq] = aksi["seq"]                       # tiap klik hanya diterapkan sekali
        if PM.terapkan_aksi_sheet(ss.pb_baris, aksi):
            st.rerun(scope="fragment" if in_dialog else "app")


DIALOG_CHROME = 170   # px di atas/bawah sheet pada dialog (judul + tombol); ubah bila ada celah atau sheet terpotong
_KW_DIALOG = {"on_dismiss": "rerun"} if "on_dismiss" in inspect.signature(st.dialog).parameters else {}


@st.dialog("Sheet hasil pemeriksaan", width="large", **_KW_DIALOG)
def dialog_sheet() -> None:
    baris = st.session_state.get("pb_baris", [])
    st.markdown('<div class="xs-penuh"></div>', unsafe_allow_html=True)
    a, b = st.columns([4, 1])
    a.caption("Klik sel untuk menyalin isinya. Klik kolom Rekam LHP sampai CAP TTD untuk V, klik lagi X, klik lagi kosong.")
    b.download_button("Unduh .xlsx", data=PM.buat_xlsx(baris), key="pb_unduh_dialog", disabled=not baris,
                      file_name=f"hasil_pemeriksaan_{time.strftime('%Y%m%d_%H%M')}.xlsx", icon=":material/download:",
                      mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width="stretch")
    _sheet_interaktif("pb_sheet_penuh", "penuh", in_dialog=True)


def _kosongkan_sheet() -> None:
    st.session_state.pb_baris = []


def _hapus_baris(id_: int) -> None:
    st.session_state.pb_baris = [r for r in st.session_state.pb_baris if r["id"] != id_]


def _tab_sheet() -> None:
    ss = st.session_state
    baris = ss.pb_baris
    st.markdown('<div class="pb-sheet-judul">Sheet hasil pemeriksaan</div>'
                '<div class="pb-sheet-sub">Format kolomnya sama dengan sheet &quot;contoh Output&quot;. '
                'Setiap dokumen yang diproses menjadi satu baris. Klik sel untuk menyalin isinya; klik kolom Rekam LHP sampai CAP TTD '
                'untuk V, klik lagi X, klik lagi kosong.</div>', unsafe_allow_html=True)
    a, b, c = st.columns([1.35, 1.1, 1])
    if a.button("Buka layar penuh", key="pb_layar_penuh", type="primary", icon=":material/fullscreen:", width="stretch"):
        dialog_sheet()
    b.download_button("Unduh .xlsx", data=PM.buat_xlsx(baris), key="pb_unduh", disabled=not baris,
                      file_name=f"hasil_pemeriksaan_{time.strftime('%Y%m%d_%H%M')}.xlsx", icon=":material/download:",
                      mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width="stretch")
    with c.popover("Kosongkan", icon=":material/delete:", disabled=not baris, width="stretch"):
        st.write("Hapus semua baris di sheet? Unduh dulu bila masih diperlukan.")
        st.button("Ya, hapus semua baris", key="pb_kosongkan", on_click=_kosongkan_sheet)
    st.toggle("Buka layar penuh otomatis setelah hasil jadi", value=True, key="pb_auto_buka")
    _sheet_interaktif("pb_sheet_inline", "inline")


def _pilih_baris(kunci: str):
    """Selectbox pemilih baris; mengembalikan baris terpilih (atau None)."""
    ss = st.session_state
    baris = ss.pb_baris
    if not baris:
        st.info("Belum ada baris. Proses dokumen terlebih dahulu.")
        return None
    ids = [r["id"] for r in baris]
    if ss.get(kunci) not in ids:
        ss[kunci] = ids[-1]

    def nama(i: int) -> str:
        r = next(x for x in baris if x["id"] == i)
        return f"#{i}  {PM.nilai_tampil(r, 'no_dok') or '(tanpa nomor)'}  {r.get('perusahaan', '')}"
    id_ = st.selectbox("Baris", ids, format_func=nama, key=kunci)
    return next(x for x in baris if x["id"] == id_)


def _tab_koreksi() -> None:
    r = _pilih_baris("pb_pilih")
    if r is None:
        return
    for w in r.get("peringatan", []):
        st.markdown(f'<div class="pb-peringatan">{H.escape(w)}</div>', unsafe_allow_html=True)
    pfx = f"pb_f_{r['id']}_{r.get('rev', 0)}_"
    with st.form(f"pb_form_{r['id']}_{r.get('rev', 0)}"):
        a, b = st.columns(2)
        no_dok = a.text_input("No Dokumen", value=PM.nilai_tampil(r, "no_dok"), key=pfx + "no")
        tgl = b.text_input("tgl dok (DD-MM-YYYY)", value=r.get("tgl_dok", ""), key=pfx + "tgl")
        perusahaan = st.text_input("Nama Perusahaan", value=r.get("perusahaan", ""), key=pfx + "per")
        st.caption("Status pekerjaan (V = selesai, X = tidak/belum)")
        judul_cek = {k: j for k, j, t, *_ in PM.KOLOM if t == "cek"}
        kc = st.columns(3)
        cek = {k: kc[i % 3].selectbox(judul_cek[k], ["", "V", "X"], index=["", "V", "X"].index(PM.norm_cek(r.get(k))),
                                      format_func=lambda x: x or "(kosong)", key=pfx + k)
               for i, k in enumerate(PM.KUNCI_CEK)}
        uraian = st.text_area("TEMPLATE URAIAN BARANG", value=r.get("uraian", ""), height=230, key=pfx + "ur")
        a, b = st.columns(2)
        nomor_pl = a.text_input("Nomor packing list", value=r.get("nomor_pl", ""), key=pfx + "pl")
        tingkat = b.radio("Tingkat pemeriksaan", [10, 30], index=0 if r.get("tingkat", 10) == 10 else 1,
                          horizontal=True, format_func=lambda x: f"{x}%", key=pfx + "ti")
        kesimpulan = st.text_area("URAIAN KESIMPULAN (otomatis dari nomor packing list dan tingkat)",
                                  value=r.get("kesimpulan", ""), height=200, key=pfx + "ke")
        a, b, c = st.columns(3)
        kem_p = a.text_input("Jumlah Kemasan yg Diperiksa", value="" if r.get("kemasan_periksa") is None else str(r["kemasan_periksa"]), key=pfx + "kp")
        bar_p = b.text_input("Jumlah Barang yang Diperiksa", value=PM.nilai_tampil(r, "barang_periksa"), key=pfx + "bp")
        kgm = c.text_input("Konversi ke KGM", value=PM.nilai_tampil(r, "kgm"), key=pfx + "kg")
        a, b = st.columns(2)
        kem_t = a.text_input("Jumlah Kemasan Total se-Dokumen", value=r.get("kemasan_total", ""), key=pfx + "kt")
        bar_t = b.text_input("Jumlah Barang Total se-Dokumen", value=r.get("barang_total", ""), key=pfx + "bt")
        hitung = st.checkbox("Hitung ulang jumlah barang diperiksa dan KGM dari jumlah kemasan diperiksa di atas",
                             value=False, key=pfx + "hit", disabled=not r.get("_hit"))
        st.caption("Waktu")
        wk = st.columns(4)
        judul_wk = ["Bongkar muat mulai", "Bongkar muat selesai", "Pemeriksaan mulai", "Pemeriksaan selesai"]
        waktu = {k: wk[i].time_input(judul_wk[i], value=r.get(k), step=60, key=pfx + k)
                 for i, k in enumerate(PM.KUNCI_WAKTU)}
        simpan = st.form_submit_button("Simpan perubahan", type="primary")
    if simpan:
        digit = re.sub(r"\D", "", no_dok)
        r["no_dok"] = int(digit) if digit else None
        r["tgl_dok"], r["perusahaan"], r["uraian"] = PM._tgl_valid(tgl), perusahaan.strip(), uraian.strip()
        r.update(cek)
        r.update(waktu)
        k_diedit = kesimpulan.strip() != (r.get("kesimpulan") or "").strip()
        berubah = nomor_pl.strip() != r.get("nomor_pl", "") or tingkat != r.get("tingkat", 10)
        r["nomor_pl"], r["tingkat"] = nomor_pl.strip(), tingkat
        if k_diedit:
            r["kesimpulan"], r["k_auto"] = kesimpulan.strip(), False
        elif (r.get("k_auto") or not r.get("kesimpulan")) and (berubah or not r.get("kesimpulan")):
            r["kesimpulan"], r["k_auto"] = PM.susun_kesimpulan(r["nomor_pl"], tingkat), True
        L = PM.baca_angka(kem_p)
        r["kemasan_periksa"] = int(L) if L and L >= 1 else None
        if hitung and r.get("_hit") and r["kemasan_periksa"]:
            r.update(PM.rumus_barang(r["_hit"], min(r["kemasan_periksa"], r["_hit"]["N"])))
        else:
            if bar_p.strip() != PM.nilai_tampil(r, "barang_periksa"):
                r["barang_periksa"], r["bp_tampil"], r["bp_fmt"] = bar_p.strip(), bar_p.strip(), None
            if kgm.strip() != PM.nilai_tampil(r, "kgm"):
                r["kgm"] = PM.baca_angka(kgm)
        r["kemasan_total"], r["barang_total"] = kem_t.strip(), bar_t.strip()
        r["rev"] = r.get("rev", 0) + 1
        st.rerun()
    st.button("Hapus baris ini", key=f"pb_hapus_{r['id']}", icon=":material/delete:", on_click=_hapus_baris, args=(r["id"],))
    st.markdown('<div class="pb-rumus"><b>Cara hitung.</b> Kemasan diperiksa = tingkat % dari kemasan total, dibulatkan ke atas, '
                'minimal 2 (tidak melebihi total). Barang diperiksa = kemasan diperiksa / kemasan total x total barang, dibulatkan ke atas '
                'bila berkoma (total semua seri; uraian fokus pada seri terbanyak). Konversi KGM = berat bersih dokumen x kemasan diperiksa / '
                'kemasan total (satuan KGM dan TNE dikonversi langsung).</div>', unsafe_allow_html=True)


def _tab_ocr() -> None:
    from dataclasses import replace
    ss = st.session_state
    r = _pilih_baris("pb_pilih_ocr")
    if r is None:
        return
    ctx = r.get("_ctx") or {}
    pfx = f"pb_o_{r['id']}_{r.get('rev', 0)}_"
    st.caption("Teks hasil OCR yang dibaca AI. Bila ada salah baca, perbaiki di sini lalu susun ulang baris. "
               "Susun ulang menimpa kolom hasil AI; centang dan waktu tetap dipertahankan.")
    for i, o in enumerate(ctx.get("ocr", [])):
        status = f" (gagal: {o['galat']})" if o["galat"] else (f" ({o['halaman']} halaman)" if o["halaman"] else "")
        with st.expander(f"{PM.SLOT_JUDUL[o['slot']]}: {o['nama']}{status}"):
            st.text_area("Teks OCR", value=o["teks"], height=240, key=f"{pfx}{i}", label_visibility="collapsed")
    gambar = ctx.get("gambar", [])
    if gambar:
        with st.expander(f"Foto yang dibaca AI ({len(gambar)})"):
            kol = st.columns(4)
            for i, g in enumerate(gambar):
                kol[i % 4].image(base64.b64decode(g["b64"]), caption=g["nama"], width="stretch")
    p0 = ctx.get("pengaturan") or PM.Pengaturan()
    catatan = st.text_area("Catatan pemeriksa", value=p0.catatan, height=70, key=pfx + "cat")
    kond = st.radio("Kondisi barang", ["Sesuai foto (AI)", "Baik", "Rusak"], horizontal=True, key=pfx + "kon",
                    index={"": 0, "baik": 1, "rusak": 2}.get(p0.kondisi, 0))
    if st.button("Susun ulang baris dari teks ini", key=pfx + "ulang", icon=":material/refresh:"):
        ocr_baru = [dict(o, teks=ss.get(f"{pfx}{i}", o["teks"])) for i, o in enumerate(ctx.get("ocr", []))]
        p_baru = replace(p0, catatan=catatan, kondisi="" if kond.startswith("Sesuai") else kond.lower())
        try:
            with st.spinner("SEMAR menyusun ulang baris..."):
                baru = PM.analisis(agen, ocr_baru, gambar, p_baru)
        except PM.GalatPemeriksaan as e:
            st.error(str(e))
            return
        except Exception as e:  # noqa: BLE001
            st.error(f"Terjadi kesalahan yang tidak terduga: {e}")
            return
        ss.pb_baris[ss.pb_baris.index(r)] = PM.gabung_ulang(r, baru)
        st.rerun()


def tampil_pemeriksaan() -> None:
    _init_periksa()
    st.markdown("<style>.block-container{max-width:1380px!important;}</style>", unsafe_allow_html=True)
    tampil_hero_periksa()
    kiri, kanan = st.columns([5, 7], gap="large")
    with kiri:
        jumlah = {s: _slot_masukan(s) for s in SLOT_INFO}
        p = _pengaturan_periksa()
        siap = jumlah["dokumen"] > 0
        if st.button("Proses OCR dan AI", key="pb_proses", type="primary", icon=":material/auto_awesome:",
                     width="stretch", disabled=not siap):
            _proses_periksa(p)
        if not siap:
            st.caption("Tambahkan minimal satu dokumen untuk memulai.")
        elif not (jumlah["label"] and jumlah["barang"]):
            st.caption("Foto label dan foto barang belum lengkap. Uraian fisik di sheet akan lebih sedikit tanpa keduanya.")
    with kanan:
        t1, t2, t3 = st.tabs(["Sheet hasil", "Koreksi dan AI", "Teks OCR"])
        with t1:
            _tab_sheet()
        with t2:
            _tab_koreksi()
        with t3:
            _tab_ocr()
    if st.session_state.pop("pb_buka_sheet", False):
        dialog_sheet()


# ===========================================================================
# PENGATUR HALAMAN (sesuai pilihan di sidebar)
# ===========================================================================
if mode == "Dashboard":
    st.markdown("<style>.block-container{max-width:1380px!important;}</style>", unsafe_allow_html=True)
    tampil_hero(kb, kompak=True)

    if not st.session_state.chat_buka:
        # Tampilan awal: dashboard penuh + Semar melayang. Klik Semar -> chat terbuka.
        tampil_dashboard(lebar=True)
        tampil_semar_melayang()
    else:
        # Chat terbuka: dashboard (kiri) berbagi ruang dengan kolom chat/perintah/jawaban (kanan).
        kiri, kanan = st.columns([3, 2], gap="large")

        # Kolom kanan diproses LEBIH DULU supaya perintah chat sudah mengubah dashboard sebelum digambar.
        with kanan:
            st.markdown('<div class="penanda-chat"></div><div class="kepala-chat"><b>Tanya atau atur lewat SEMAR</b>'
                        '<span>ketik perintah, dashboard ikut berubah</span></div>', unsafe_allow_html=True)
            with st.container(key="tutup_chat"):
                st.button("Tutup", key="tutup_chat_btn", icon=":material/close:", on_click=_tutup_chat)
            # Kolom ketik di ATAS, lalu riwayat/sambutan/"Coba perintah" di bawahnya.
            # (wadah baru diisi setelah ini, jadi perintah tetap diproses sebelum dashboard digambar)
            pertanyaan = st.chat_input("Contoh: tampilkan penindakan di Kendal tahun 2025", key="ci_dashboard")
            if not pertanyaan:
                pertanyaan = st.session_state.pop("tanya_cepat", None)
            wadah = st.container(height=TINGGI_CHAT, border=False)
            with wadah:
                tampil_riwayat()
                if not st.session_state.pesan and not pertanyaan:
                    st.markdown(padatkan("""
<div class="sambutan"><h3>Sugeng rawuh</h3><p>Atur dashboard dengan perintah singkat, atau tanyakan hal lain.
SEMAR menjawab hanya dari Excel dan website yang ditentukan.</p></div>"""), unsafe_allow_html=True)
                    st.markdown('<div class="label-perintah">Coba perintah:</div>', unsafe_allow_html=True)
                    tombol_cepat(CONTOH_PERINTAH, 1, "perintah")
                if pertanyaan:
                    proses(pertanyaan)

        with kiri:
            tampil_dashboard()

        tampil_semar_melayang(terbuka=True)             # Semar tetap ada, tetapi transparan

elif mode == "Pemeriksaan Barang":
    tampil_pemeriksaan()

else:  # ---------------------------- halaman Semar / percakapan (seperti sebelumnya) ----------------------------
    tampil_hero(kb)
    pertanyaan = st.chat_input("Tanyakan sesuatu pada SEMAR...", key="ci_percakapan")
    if not pertanyaan:
        pertanyaan = st.session_state.pop("tanya_cepat", None)
    tampil_riwayat()
    if not st.session_state.pesan and not pertanyaan:
        st.markdown(padatkan("""
<div class="sambutan"><h3>Sugeng rawuh, silakan bertanya</h3>
<p>SEMAR menjawab hanya dari tabel Excel dan website yang ditentukan. Kalau informasinya tidak ada di sana,
SEMAR akan mengatakannya terus terang, tanpa mengarang. Dashboard bisa diatur lewat chat: coba
&ldquo;tampilkan penindakan di Kendal&rdquo;.</p></div>"""), unsafe_allow_html=True)
        topik = list(dict.fromkeys(d.judul for d in kb.entri if d.judul))[:6]
        if topik:
            st.markdown('<div class="label-cepat">Coba tanyakan:</div>', unsafe_allow_html=True)
            tombol_cepat(topik, 3, "cepat", "Jelaskan: ")
    if pertanyaan:
        dash_baru = proses(pertanyaan)
        if dash_baru:                                   # perintah dashboard: pindah ke tampilan Dashboard
            st.session_state["_mode_baru"] = "Dashboard"
            st.session_state.chat_buka = True           # percakapan tetap terlihat di samping dashboard
            st.rerun()


# ===========================================================================
# SUARA: dijalankan paling akhir pada setiap putaran skrip
# ===========================================================================
render_ucapan()                  # ucapan yang tertunda (mis. setelah st.rerun)
pasang_suara(suara_aktif)        # pasang modul suara + selaraskan saklar "Suara SEMAR"
buka_sidebar_awal()              # sidebar terbuka di awal sesi (sekali saja)
