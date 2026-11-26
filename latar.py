"""
Gambar latar aplikasi: Semar berseragam Bea Cukai, minimalis, transparan 20%, di atas latar putih bergradasi.

  - Bawaan: ilustrasi vektor (SVG) yang digambar di sini.
  - Bisa diganti dengan gambar sendiri: simpan satu berkas bernama latar_semar.png / .webp / .jpg di folder aplikasi.
    Gambar itu otomatis dibuat transparan 20% dan dikecilkan agar ringan.
Berkas ini tidak memakai Streamlit.
"""

from __future__ import annotations

import base64
import io
from pathlib import Path

OPASITAS = 0.20
NAMA_BERKAS = ("latar_semar.png", "latar_semar.webp", "latar_semar.jpg", "latar_semar.jpeg")
TINGGI_MAKS = 900          # gambar pengganti dikecilkan sampai setinggi ini (agar CSS tetap ringan)


def svg_semar_seragam(opasitas: float = OPASITAS) -> str:
    return f"""<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 560 760" width="560" height="760">
<defs>
  <radialGradient id="aura" cx="50%" cy="48%" r="52%"><stop offset="0" stop-color="#F7B733" stop-opacity=".55"/><stop offset=".55" stop-color="#F58220" stop-opacity=".18"/><stop offset="1" stop-color="#F58220" stop-opacity="0"/></radialGradient>
  <linearGradient id="kulit" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#EDC095"/><stop offset="1" stop-color="#CC9262"/></linearGradient>
  <linearGradient id="rambut" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#FFFFFF"/><stop offset="1" stop-color="#DAD5CB"/></linearGradient>
  <linearGradient id="seragam" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#2358A8"/><stop offset="1" stop-color="#0A1F44"/></linearGradient>
  <linearGradient id="seragam2" x1="0" y1="0" x2="1" y2="0"><stop offset="0" stop-color="#1B4F9C"/><stop offset="1" stop-color="#12377A"/></linearGradient>
  <linearGradient id="celana" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#12306A"/><stop offset="1" stop-color="#0A1F44"/></linearGradient>
  <linearGradient id="emas" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#F7B733"/><stop offset="1" stop-color="#F58220"/></linearGradient>
  <linearGradient id="sepatu" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#2A3556"/><stop offset="1" stop-color="#0E1530"/></linearGradient>
</defs>
<g opacity="{opasitas}">
  <circle cx="280" cy="385" r="300" fill="url(#aura)"/>
  <ellipse cx="280" cy="728" rx="165" ry="15" fill="#0A1F44" opacity=".25"/>

  <!-- kaki dan sepatu -->
  <path d="M212 600 H268 V700 H212 Z" fill="url(#celana)"/><path d="M292 600 H348 V700 H292 Z" fill="url(#celana)"/>
  <path d="M196 700 C196 680 270 680 272 700 C276 726 192 726 196 700 Z" fill="url(#sepatu)"/>
  <path d="M288 700 C290 680 364 680 364 700 C368 726 284 726 288 700 Z" fill="url(#sepatu)"/>

  <!-- lengan kiri (turun) -->
  <path d="M158 342 C98 372 78 462 96 530 L146 524 C146 462 164 408 186 352 Z" fill="url(#seragam2)"/>
  <path d="M94 518 L148 510 L150 528 L96 538 Z" fill="url(#emas)"/>
  <circle cx="118" cy="556" r="27" fill="url(#kulit)"/>

  <!-- badan -->
  <path d="M150 352 C150 318 206 298 280 298 C354 298 410 318 410 352 C432 430 440 520 424 568 C410 612 360 628 280 628 C200 628 150 612 136 568 C120 520 128 430 150 352 Z" fill="url(#seragam)"/>
  <path d="M234 300 L276 338 L246 470 L196 380 Z" fill="#2E67BA" opacity=".55"/>
  <path d="M326 300 L284 338 L314 470 L364 380 Z" fill="#2E67BA" opacity=".55"/>
  <path d="M250 296 L280 366 L310 296 Z" fill="#F4F7FC"/>
  <path d="M268 322 H292 L298 372 L280 398 L262 372 Z" fill="url(#emas)"/>
  <circle cx="280" cy="448" r="8" fill="url(#emas)"/><circle cx="280" cy="498" r="8" fill="url(#emas)"/><circle cx="280" cy="548" r="8" fill="url(#emas)"/>
  <rect x="192" y="402" width="62" height="20" rx="4" fill="#F4F7FC"/><rect x="199" y="408" width="36" height="3" rx="1.5" fill="#9AA6C2"/><rect x="199" y="414" width="24" height="3" rx="1.5" fill="#9AA6C2"/>
  <path d="M322 388 H364 V420 C364 440 346 450 343 452 C340 450 322 440 322 420 Z" fill="url(#emas)"/>
  <path d="M343 396 L347 408 L360 408 L350 416 L354 429 L343 421 L332 429 L336 416 L326 408 L339 408 Z" fill="#fff"/>

  <!-- bahu -->
  <g transform="rotate(-16 186 334)"><rect x="150" y="322" width="72" height="22" rx="8" fill="url(#emas)"/><rect x="164" y="322" width="5" height="22" fill="#0A1F44" opacity=".55"/><rect x="178" y="322" width="5" height="22" fill="#0A1F44" opacity=".55"/></g>
  <g transform="rotate(16 374 334)"><rect x="338" y="322" width="72" height="22" rx="8" fill="url(#emas)"/><rect x="377" y="322" width="5" height="22" fill="#0A1F44" opacity=".55"/><rect x="391" y="322" width="5" height="22" fill="#0A1F44" opacity=".55"/></g>

  <!-- lengan kanan (melambai) -->
  <path d="M404 340 C468 332 508 282 500 214 L456 222 C462 268 442 298 402 312 Z" fill="url(#seragam2)"/>
  <path d="M452 220 L504 212 L508 232 L456 242 Z" fill="url(#emas)" transform="rotate(-4 480 226)"/>
  <circle cx="480" cy="190" r="28" fill="url(#kulit)"/><ellipse cx="452" cy="176" rx="9" ry="17" fill="url(#kulit)" transform="rotate(-30 452 176)"/>

  <!-- kepala -->
  <ellipse cx="194" cy="214" rx="14" ry="20" fill="url(#kulit)"/><ellipse cx="366" cy="214" rx="14" ry="20" fill="url(#kulit)"/>
  <ellipse cx="280" cy="176" rx="96" ry="74" fill="url(#rambut)"/>
  <path d="M258 132 C246 100 266 82 282 84 C300 82 316 100 304 132 Z" fill="url(#rambut)"/>
  <circle cx="281" cy="94" r="10" fill="url(#emas)"/><circle cx="281" cy="94" r="4" fill="#fff" opacity=".7"/>
  <ellipse cx="280" cy="222" rx="84" ry="78" fill="url(#kulit)"/>
  <path d="M226 190 Q246 178 266 190" stroke="#F4F1EA" stroke-width="9" stroke-linecap="round" fill="none"/>
  <path d="M294 190 Q314 178 334 190" stroke="#F4F1EA" stroke-width="9" stroke-linecap="round" fill="none"/>
  <ellipse cx="248" cy="212" rx="7" ry="8.5" fill="#1B2440"/><ellipse cx="312" cy="212" rx="7" ry="8.5" fill="#1B2440"/>
  <circle cx="250.5" cy="209" r="2.4" fill="#fff"/><circle cx="314.5" cy="209" r="2.4" fill="#fff"/>
  <ellipse cx="280" cy="236" rx="15" ry="11" fill="#C98857" opacity=".75"/>
  <circle cx="232" cy="238" r="13" fill="#F58220" opacity=".16"/><circle cx="328" cy="238" r="13" fill="#F58220" opacity=".16"/>
  <path d="M246 258 Q280 286 314 258" stroke="#7A4A2A" stroke-width="5" stroke-linecap="round" fill="none"/>
  <path d="M262 276 C262 306 272 322 280 330 C288 322 298 306 298 276 Q280 288 262 276 Z" fill="url(#rambut)"/>
</g>
</svg>"""


def _data_uri_svg(svg: str) -> str:
    return "data:image/svg+xml;base64," + base64.b64encode(svg.encode("utf-8")).decode()


def _data_uri_pengganti(berkas: Path) -> str | None:
    """Gambar buatan sendiri: dibuat transparan 20% dan dikecilkan (WebP) agar ringan."""
    try:
        from PIL import Image
        im = Image.open(berkas).convert("RGBA")
        if im.height > TINGGI_MAKS:
            im = im.resize((max(1, round(im.width * TINGGI_MAKS / im.height)), TINGGI_MAKS), Image.LANCZOS)
        alfa = im.getchannel("A").point(lambda v: int(v * OPASITAS))
        im.putalpha(alfa)
        buf = io.BytesIO()
        im.save(buf, "WEBP", quality=70, method=6)
        return "data:image/webp;base64," + base64.b64encode(buf.getvalue()).decode()
    except Exception:       # noqa: BLE001  (berkas rusak: pakai gambar bawaan)
        return None


def latar_uri(folder: Path) -> str:
    """URI gambar latar: gambar pengganti (latar_semar.*) bila ada, kalau tidak ilustrasi bawaan."""
    for nama in NAMA_BERKAS:
        f = Path(folder) / nama
        if f.exists():
            uri = _data_uri_pengganti(f)
            if uri:
                return uri
    return _data_uri_svg(svg_semar_seragam())
