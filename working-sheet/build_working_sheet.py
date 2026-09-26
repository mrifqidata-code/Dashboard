"""Bangun working sheet Excel "Slot Map Weekly Meeting" dari tab Reguler Class Availability.

Pemakaian:
    python build_working_sheet.py <feed.csv> <output.xlsx> [--week "W40 · 28 Sep – 4 Okt 2026"] [--asof 2026-09-26]

<feed.csv> = CSV dari Sheet "Feed Slot Map (dari Exboard)" (IMPORTRANGE tab Reguler Class Availability).
"""
import argparse
import csv
import re
from collections import defaultdict
from datetime import date

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule, ColorScaleRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

import os
import sys
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "exboard-slot-map"))
from build_slot_map_tab import RCA, build_rca_copy, build_slot_map, num  # noqa: E402

HELP_START = 40            # kolom rumus bantu slot map (AN ke kanan, disembunyikan)
ST_FIRST, ST_LAST = get_column_letter(HELP_START + 8), get_column_letter(HELP_START + 14)   # kolom status
LIST_IDX_COL = 17          # Q: nomor baris Reguler Class Availability untuk Daftar Diskusi (disembunyikan)

CENTERS = ["KLM", "KWC", "PML", "TMP", "BTU", "HIB"]
CENTER_NAME = {"KLM": "Kalimalang", "KWC": "Karawaci", "PML": "Pamulang", "TMP": "TMP", "BTU": "BTU", "HIB": "HIB"}
DAYS = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu"]
AGES = ["Baby", "Tiny", "Little", "Kids", "Star"]
AGE_RANGE = {"Baby": "6–17 bulan", "Tiny": "18–36 bulan", "Little": "3–5 tahun", "Kids": "5–8+ tahun", "Star": "9–12 tahun"}
SESSION_GAP_MIN = 45  # dua sesi dianggap bentrok bila jaraknya < 45 menit
DECISIONS = ["Jual minggu ini", "Buka kelas baru", "Ganti age group", "Tahan", "Tutup slot"]

SLOT_RE = re.compile(r"^(.*?)-\s*Lane\s*(\d+)\s*-\s*([A-Za-z]+)\s*-\s*(\d{1,2})[:.](\d{2})")
F = {"age": 0, "day": 1, "time": 2, "slot": 3, "cap": 4, "used": 5, "avail": 6, "coach": 7}

# ---------- style ----------
FONT = "Arial"
C_INK, C_MUTED = "1F2D33", "6B7C82"
C_ACCENT, C_ACCENT_SOFT = "0B6475", "DCEDF0"
C_INPUT = "FFF4B8"
STATUS_FILL = {
    "Belum terisi": "F8C9C0",      # kelas ada, 0 murid  -> fokus diskusi
    "Slot kosong": "FFE3A6",       # kapasitas ada, belum ada kelas
    "Sebagian": "E3F1F4",
    "Penuh": "BFE3C8",
    "Lebih kapasitas": "E2D4F5",
    "Tutup": "E4E4E4",
}
AGE_FILL = {"Baby": "F7DBE6", "Tiny": "FDE7C4", "Little": "D6EFDC", "Kids": "DBE4FB", "Star": "F1E2FA"}
thin = Side(style="thin", color="D5DDDF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)


def font(bold=False, size=10, color=C_INK, italic=False):
    return Font(name=FONT, bold=bold, size=size, color=color, italic=italic)


def fill(hex_):
    return PatternFill("solid", start_color=hex_, end_color=hex_)


# ---------- parse ----------
def cell_num(s):
    if s is None:
        return None
    s = re.sub(r"[\s  ]", "", str(s))
    if not s or s.startswith("#"):
        return None
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})$", s)
    if m:
        return (date(int(m[1]), int(m[2]), int(m[3])) - date(1899, 12, 30)).days
    try:
        n = float(s.replace(",", "."))
        return int(n) if n == int(n) else n
    except ValueError:
        return None


def clean(s):
    return re.sub(r"[  ]", " ", s or "").strip()


def title_case(s):
    return re.sub(r"(^|[\s-])(\w)", lambda m: m[1] + m[2].upper(), s.lower())


def norm_age(s):
    s = clean(s)
    if not s or re.fullmatch(r"[-–—.]+", s):
        return ""
    for a in AGES:
        if a.lower() == s.lower():
            return a
    return title_case(s)


def norm_coach(s):
    s = re.sub(r"\s+", " ", clean(s))
    if not s or s == "-" or re.fullmatch(r"(not )?available", s, re.I):
        return ""
    return title_case(s)


HEAD_MATCH = {
    "age": lambda h: re.fullmatch(r"age\s*group", h), "slot": lambda h: h == "slot",
    "cap": lambda h: re.fullmatch(r"cap(acity|asitas)?", h), "used": lambda h: h in ("used", "terisi"),
    "avail": lambda h: h.startswith("availab"), "asst": lambda h: "asst" in h or "assist" in h,
    "coach": lambda h: h == "coach",
}
DEFAULT_COLS = {"age": 1, "slot": 4, "cap": 5, "used": 6, "avail": 7, "coach": 8, "asst": None, "end": 9}


def header_cols(r):
    cols, end = {}, len(r)
    for k, raw in enumerate(r[:20]):
        h = clean(raw).lower().replace(".", "")
        if k > 0 and h == "center":
            end = k
            break
        for f, test in HEAD_MATCH.items():
            if f not in cols and test(h):
                cols[f] = k
        if (h == "sesi" or h.startswith("key")) and "coach" in cols:
            end = min(end, k)
    if "slot" not in cols or "cap" not in cols:
        return None
    cols.setdefault("asst", None)
    cols["end"] = end
    return cols


def parse_feed(path):
    """Sama dengan parser dashboard: posisi kolom dari header; pergeseran per baris dari posisi kolom Slot."""
    slots, cols = [], dict(DEFAULT_COLS)
    with open(path, encoding="utf-8") as fh:
        for r in csv.reader(fh):
            first = clean(r[0] if r else "")
            if first.lower() == "center":
                hc = header_cols(r)
                if hc:
                    cols = hc
                continue
            center = first.upper()
            if not re.fullmatch(r"[A-Z]{2,5}", center):
                continue
            j, m = -1, None
            for k in range(1, min(cols["slot"] + 7, len(r))):
                mm = SLOT_RE.match(clean(r[k]))
                if mm:
                    j, m = k, mm
                    break
            if not m:
                continue
            shift = j - cols["slot"]

            def get(f):
                if cols.get(f) is None:
                    return ""
                pos = cols[f] + shift
                if pos < 1 or pos >= len(r) or (shift and pos >= cols["end"]):
                    return None
                return r[pos]

            day = next((d for d in DAYS if d.lower() == m[3].lower()), None)
            if not day:
                continue
            slots.append({
                "center": center, "lane": int(m[2]), "day": day,
                "time": f"{int(m[4]):02d}:{m[5]}",
                "age": norm_age(re.sub(r"-\s*$", "", m[1])) or norm_age(get("age")),
                "cap": cell_num(get("cap")), "used": cell_num(get("used")),
                "coach": norm_coach(get("coach")), "asst": norm_coach(get("asst")),
            })
    return slots


def status_of(s):
    cap, used, age, coach = s["cap"], s["used"], s["age"], s["coach"]
    if cap is None or used is None:
        return "Data tidak lengkap"
    if cap == 0 and used == 0:
        return "Tutup" if (age or coach or s.get("asst")) else "Tidak dibuka"
    if used > cap:
        return "Lebih kapasitas"
    if used == cap:
        return "Penuh"
    if used == 0:
        return "Belum terisi" if age else "Slot kosong"
    return "Sebagian"


def minutes(t):
    h, m = t.split(":")
    return int(h) * 60 + int(m)


# ---------- coach availability ----------
def coach_book(slots):
    """Per center: jadwal tiap coach (hari -> set menit), age group yang diajar."""
    book = defaultdict(lambda: defaultdict(lambda: {"busy": defaultdict(set), "ages": set(), "sessions": 0, "name": ""}))
    for s in slots:
        for role in ("coach", "asst"):
            if not s.get(role):
                continue
            c = book[s["center"]][s[role].lower()]
            c["name"] = c["name"] or s[role]
            c["busy"][s["day"]].add(minutes(s["time"]))
            if s["age"]:
                c["ages"].add(s["age"])
            c["sessions"] += 1
            c["assist"] = c.get("assist", 0) + (role == "asst")
    return book


def available_coaches(book_center, day, time, exclude=""):
    """Coach yang bertugas di center itu pada hari yang sama, tidak sedang mengajar dalam ±45 menit."""
    t = minutes(time)
    out = []
    for key, c in book_center.items():
        if key == exclude.lower():
            continue
        busy = c["busy"].get(day)
        if not busy:
            continue
        if all(abs(t - b) >= SESSION_GAP_MIN for b in busy):
            out.append(c)
    order = {a: i for i, a in enumerate(AGES)}
    out.sort(key=lambda c: (-c["sessions"], c["name"]))
    return out, order


def ages_str(ages):
    order = {a: i for i, a in enumerate(AGES)}
    return ", ".join(sorted(ages, key=lambda a: (order.get(a, 99), a)))


# ---------- workbook ----------
def rng(letter):
    return f"{RCA}!${letter}:${letter}"


def build(feed, out_path, week_label, asof):
    slots = parse_feed(feed)
    wb = Workbook()
    book = coach_book(slots)
    centers = [c for c in CENTERS if any(s["center"] == c for s in slots)]
    day_idx = {d: i for i, d in enumerate(DAYS)}

    cp = wb.active
    cp.title = "Cara Pakai"
    build_guide(cp, week_label, asof)
    rk = wb.create_sheet("Ringkasan")
    build_summary(rk, centers, week_label, asof)
    for c in centers:
        ws = wb.create_sheet(c)
        cs = sorted([s for s in slots if s["center"] == c], key=lambda s: (s["lane"], day_idx[s["day"]], s["time"]))
        build_center(ws, c, cs, book[c], week_label, asof)
    build_rca_copy(wb, feed)
    rca = wb["Reguler Class Availability"]
    rca["N1"] = ("Salinan tab Reguler Class Availability dari Exboard (kolom A:L). Minggu berikutnya: paste data terbaru "
                 "ke A:L di tab ini, semua tab center ikut menghitung ulang.")
    rca.freeze_panes = "A2"
    for c in rca[1][:12]:
        c.font = font(True, 10, "FFFFFF")
        c.fill = fill(C_ACCENT)
    for letter, w in zip("ABCDEFGHIJKL", [8, 10, 18, 7, 30, 6, 6, 14, 14, 14, 8, 30]):
        rca.column_dimensions[letter].width = w
    wb.active = wb.sheetnames.index("Ringkasan")
    wb.save(out_path)


def build_guide(ws, week_label, asof):
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 26
    ws.column_dimensions["C"].width = 90
    ws.sheet_view.showGridLines = False
    r = 2
    ws.cell(r, 2, "Slot Map Weekly Meeting · Center Manager × Senior Coach").font = font(True, 16, C_ACCENT)
    r += 1
    ws.cell(r, 2, f"Data per {asof} · untuk meeting {week_label} · sumber: Exboard › Reguler Class Availability").font = font(size=10, color=C_MUTED)
    r += 2
    ws.cell(r, 2, "Alur meeting").font = font(True, 12)
    r += 1
    steps = [
        ("1. Buka tab center", "Satu tab per center (KLM, KWC, PML, TMP, BTU, HIB). Mulai dari ringkasan di atas: utilisasi, jumlah kelas belum terisi, slot kosong."),
        ("2. Lihat Slot Map", "Tampilan sama dengan tab Slot Map di Exboard, tanpa pilihan center. Grid per lane: baris = jam, kolom = hari. Tiap sel berisi age group, terisi/kapasitas (%), coach, dan assist coach. Merah = kelas sudah ada tapi 0 murid. Oranye = kapasitas ada tapi belum ada kelas."),
        ("3. Putuskan di Daftar Diskusi", "Di bawah slot map ada daftar semua slot merah & oranye. Isi kolom kuning: keputusan, age group yang dijual, coach PIC, target murid, catatan."),
        ("4. Cek coach", "Kolom 'Coach available' menyarankan coach yang bertugas hari itu dan tidak sedang mengajar di jam tersebut, beserta age group yang biasa mereka ajar. Tabel Coach di bagian bawah tab menunjukkan beban tiap coach per hari."),
        ("5. Kirim ke Student Advisor", "Filter kolom Keputusan = 'Jual minggu ini' atau 'Buka kelas baru'. Itulah daftar kelas yang dijual SA minggu itu."),
    ]
    for a, b in steps:
        ws.cell(r, 2, a).font = font(True)
        ws.cell(r, 3, b).font = font()
        ws.cell(r, 3).alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(r, 2).alignment = Alignment(vertical="top")
        r += 1
    r += 1
    ws.cell(r, 2, "Arti warna status").font = font(True, 12)
    r += 1
    meaning = [
        ("Belum terisi", "Kelas sudah punya age group, kapasitas > 0, tapi 0 murid. Prioritas untuk dijual."),
        ("Slot kosong", "Kapasitas > 0 tapi belum ada age group/kelas. Bisa dibuka jadi kelas baru."),
        ("Sebagian", "Sudah ada murid tapi belum penuh."),
        ("Penuh", "Terisi = kapasitas."),
        ("Lebih kapasitas", "Terisi melebihi kapasitas (termasuk slot kapasitas 0 yang masih ada murid). Perlu dicek."),
        ("Tutup", "Kapasitas 0 dan tidak ada murid, tapi slot masih punya age group/coach."),
    ]
    for st, txt in meaning:
        ws.cell(r, 2, st).fill = fill(STATUS_FILL[st])
        ws.cell(r, 2).font = font(True)
        ws.cell(r, 2).border = BORDER
        ws.cell(r, 3, txt).font = font()
        r += 1
    ws.cell(r, 2, "Tidak dibuka").font = font(True, color=C_MUTED)
    ws.cell(r, 3, "Kapasitas 0, tanpa kelas. Ditampilkan '–'.").font = font()
    r += 2
    ws.cell(r, 2, "Sel yang diisi").font = font(True, 12)
    r += 1
    ws.cell(r, 2, "Kuning").fill = fill(C_INPUT)
    ws.cell(r, 2).font = font(True)
    ws.cell(r, 2).border = BORDER
    ws.cell(r, 3, "Hanya sel kuning yang diisi saat meeting (minggu, peserta, keputusan, age group dijual, coach PIC, target, catatan). Sel lain berisi rumus.").font = font()
    r += 1
    ws.cell(r, 2, "Data sumber").font = font(True, color=C_ACCENT)
    ws.cell(r, 3, "Tab Reguler Class Availability = salinan tab yang sama di Exboard (kolom A:L). Semua rumus membaca tab ini.").font = font()
    r += 2
    ws.cell(r, 2, "Contoh pengisian").font = font(True, 12)
    r += 1
    ex_h = ["Hari · Jam · Lane", "Status", "Coach available", "Keputusan", "Age group dijual", "Coach PIC", "Target murid", "Catatan"]
    ex_v = ["Selasa · 09:45 · Lane 2", "Belum terisi", "Danti (Baby, Little)", "Jual minggu ini", "Tiny", "Danti", 3, "Tawarkan ke leads trial Tiny minggu lalu"]
    for i, (h, v) in enumerate(zip(ex_h, ex_v)):
        ws.cell(r + i, 2, h).font = font(True, color=C_MUTED)
        ws.cell(r + i, 3, v).font = font(italic=True)
        ws.cell(r + i, 3).alignment = Alignment(horizontal="left")
    r += len(ex_h) + 1
    ws.cell(r, 2, "Catatan").font = font(True, 12)
    r += 1
    notes = [
        "Coach available = coach yang punya jadwal di center itu pada hari yang sama dan tidak mengajar dalam rentang ±45 menit dari slot tersebut. "
        "Age group di dalam kurung = age group yang dia ajar minggu ini di center itu. Ini saran; cek ulang ke senior coach.",
        "Coach yang tidak punya jadwal sama sekali di hari itu tidak muncul sebagai available (bisa saja libur).",
        "Minggu berikutnya: copy kolom A:L dari tab Reguler Class Availability di Exboard, paste ke tab dengan nama sama di file ini. "
        "Slot map, ringkasan, status di Daftar Diskusi dan tabel Coach ikut menghitung ulang. Baris di Daftar Diskusi dan saran coach "
        "available dibuat saat file digenerate, jadi untuk daftar yang benar-benar baru, generate ulang file-nya.",
    ]
    for n in notes:
        ws.cell(r, 3, n).font = font()
        ws.cell(r, 3).alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(r, 2, "•").font = font(True)
        ws.cell(r, 2).alignment = Alignment(horizontal="right", vertical="top")
        r += 1


def build_summary(ws, centers, week_label, asof):
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 3
    ws.cell(2, 2, "Ringkasan Utilisasi per Center").font = font(True, 16, C_ACCENT)
    ws.cell(3, 2, f"Data per {asof} · meeting {week_label}").font = font(size=10, color=C_MUTED)
    heads = ["Center", "Utilisasi", "Kursi terisi", "Kapasitas", "Belum terisi\n(kelas 0 murid)", "Slot kosong\n(belum ada kelas)",
             "Sebagian", "Penuh", "Lebih kapasitas", "Buka tab"]
    widths = [14, 11, 12, 11, 15, 16, 10, 9, 12, 11]
    for i, (h, w) in enumerate(zip(heads, widths)):
        c = ws.cell(5, 2 + i, h)
        c.font = font(True, 10, "FFFFFF")
        c.fill = fill(C_ACCENT)
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        ws.column_dimensions[get_column_letter(2 + i)].width = w
    ws.row_dimensions[5].height = 30
    for k, cen in enumerate(centers):
        r = 6 + k
        ws.cell(r, 2, cen).font = font(True)
        ws.cell(r, 3, f'=IFERROR(D{r}/E{r},"")').number_format = "0%"
        ws.cell(r, 4, f'=SUMIFS({rng("G")},{rng("A")},B{r},{rng("F")},">0")')
        ws.cell(r, 5, f'=SUMIFS({rng("F")},{rng("A")},B{r},{rng("F")},">0")')
        st_rng = f"'{cen}'!${ST_FIRST}$1:${ST_LAST}$1000"
        for j, (st, keys) in enumerate([("Belum terisi", ["Belum terisi"]), ("Slot kosong", ["Kosong"]),
                                        ("Sebagian", ["Sebagian", "Tinggi"]), ("Penuh", ["Penuh"]), ("Lebih kapasitas", ["Lebih"])]):
            cell = ws.cell(r, 6 + j, "=" + "+".join(f'COUNTIF({st_rng},"{k}")' for k in keys))
            cell.fill = fill(STATUS_FILL[st]) if st in ("Belum terisi", "Slot kosong") else PatternFill()
        link = ws.cell(r, 11, f"→ {cen}")
        link.hyperlink = f"#'{cen}'!A1"
        link.font = Font(name=FONT, size=10, color=C_ACCENT, underline="single")
        for col in range(2, 12):
            ws.cell(r, col).border = BORDER
            if col != 11:
                ws.cell(r, col).font = font(col == 2 or col == 3)
            ws.cell(r, col).alignment = Alignment(horizontal="center" if col > 2 else "left")
    tr = 6 + len(centers)
    ws.cell(tr, 2, "Total").font = font(True)
    ws.cell(tr, 3, f'=IFERROR(D{tr}/E{tr},"")').number_format = "0%"
    for col in range(4, 11):
        L = get_column_letter(col)
        ws.cell(tr, col, f"=SUM({L}6:{L}{tr-1})")
    for col in range(2, 11):
        ws.cell(tr, col).font = font(True)
        ws.cell(tr, col).fill = fill(C_ACCENT_SOFT)
        ws.cell(tr, col).border = BORDER
        ws.cell(tr, col).alignment = Alignment(horizontal="center" if col > 2 else "left")
    ws.cell(tr + 2, 2, "Utilisasi = kursi terisi ÷ kapasitas, hanya slot dengan kapasitas > 0.").font = font(size=9, color=C_MUTED)
    ws.conditional_formatting.add(f"C6:C{tr-1}", ColorScaleRule(start_type="num", start_value=0, start_color="F8C9C0",
                                                                 mid_type="num", mid_value=0.5, mid_color="FFF4B8",
                                                                 end_type="num", end_value=1, end_color="BFE3C8"))


def build_center(ws, center, slots, book_c, week_label, asof):
    # Blok Slot Map: sama dengan tab Slot Map Exboard, center tetap (tanpa dropdown)
    r = build_slot_map(ws, center=center, helper_start=HELP_START)
    ws.sheet_view.zoomScale = 90
    ws.column_dimensions[get_column_letter(LIST_IDX_COL)].hidden = True
    for k, v in {"I": 40, "J": 18, "K": 16, "L": 14, "M": 14, "N": 11, "O": 34}.items():
        ws.column_dimensions[k].width = v
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws["A2"].value = ws["A2"].value + f"  ·  Data per {asof}"
    # Info meeting (kuning = diisi)
    for c_lab, c_val, lab, val in (("C", "D", "Minggu", week_label), ("E", "F", "Center Manager", None), ("G", "H", "Senior Coach", None)):
        ws[f"{c_lab}4"] = lab
        ws[f"{c_lab}4"].font = font(True, 9, C_MUTED)
        ws[f"{c_lab}4"].alignment = Alignment(horizontal="right")
        ws[f"{c_val}4"] = val
        ws[f"{c_val}4"].fill = fill(C_INPUT)
        ws[f"{c_val}4"].border = BORDER
        ws[f"{c_val}4"].font = font()
    nav_row, slotmap_row = 11, 12

    # ----- Daftar Diskusi -----
    r += 1
    diskusi_row = r
    ws.cell(r, 1, "DAFTAR DISKUSI · kelas yang belum terisi & slot kosong").font = font(True, 13, "FFFFFF")
    for c in range(1, 15):
        ws.cell(r, c).fill = fill("B03A2E")
    r += 1
    ws.cell(r, 1, "Isi kolom kuning saat meeting. Filter kolom Keputusan untuk daftar yang dijual Student Advisor minggu ini.").font = font(size=9, color=C_MUTED, italic=True)
    r += 1
    hdr = ["No", "Hari", "Jam", "Lane", "Status", "Age group terjadwal", "Kapasitas", "Coach terjadwal",
           "Coach available\n(age group yang diajar)", "Age group tersedia", "Keputusan", "Age group dijual", "Coach PIC", "Target murid", "Catatan"]
    # columns: A..O  (A=No, B=Hari, ... O=Catatan)
    for i, h in enumerate(hdr):
        c = ws.cell(r, 1 + i, h)
        c.font = font(True, 9, "FFFFFF")
        c.fill = fill(C_INK) if i < 10 else fill("7A5C00")
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        c.border = BORDER
    ws.column_dimensions["O"].width = 34
    ws.column_dimensions["N"].width = 11
    ws.column_dimensions["M"].width = 14
    ws.row_dimensions[r].height = 32
    list_head = r
    r += 1
    day_idx = {d: i for i, d in enumerate(DAYS)}
    todo = [s for s in slots if status_of(s) in ("Belum terisi", "Slot kosong")]
    todo.sort(key=lambda s: (day_idx[s["day"]], s["time"], s["lane"]))
    coach_names = sorted({c["name"] for c in book_c.values()})
    first_item = r
    for n, s in enumerate(todo, 1):
        n_ref = f"$Q{r}"
        ws.cell(r, LIST_IDX_COL, f'=IFERROR(MATCH($B$4&"|- Lane "&D{r}&" - "&B{r}&"|"&C{r},{rng("L")},0),"")')
        age = f'INDEX({rng("B")},{n_ref})&""'
        cap = num(f'INDEX({rng("F")},{n_ref})')
        used = num(f'INDEX({rng("G")},{n_ref})')
        coach = f'INDEX({rng("I")},{n_ref})&""'
        asst = f'INDEX({rng("J")},{n_ref})&""'
        avail, _ = available_coaches(book_c, s["day"], s["time"], exclude=s["coach"])
        avail = [c for c in avail if c["name"].lower() != (s.get("asst") or "").lower()]
        avail_txt = "; ".join(f'{c["name"]} ({ages_str(c["ages"]) or "–"})' for c in avail[:6])
        if len(avail) > 6:
            avail_txt += f"; +{len(avail) - 6} lainnya"
        pool = set()
        if s["age"]:
            pool.add(s["age"])
        if s["coach"] and s["coach"].lower() in book_c:
            pool |= book_c[s["coach"].lower()]["ages"]
        for c in avail:
            pool |= c["ages"]
        row = [n, s["day"], s["time"], s["lane"],
               (f'=IF({n_ref}="","Tidak ada di data",IF(AND({cap}=0,{used}=0),"Tutup",IF({used}>{cap},"Lebih",IF({used}={cap},"Penuh",'
                f'IF({used}=0,IF({age}="","Kosong","Belum terisi"),IF({used}/{cap}>=0.6,"Tinggi","Sebagian"))))))'),
               f'=IF({n_ref}="","",PROPER({age}))',
               f'=IF({n_ref}="","",{cap})',
               f'=IF({n_ref}="","",PROPER({coach})&IF({asst}="",""," + asst "&PROPER({asst})))',
               avail_txt or "Tidak ada coach lain yang bertugas & kosong",
               ages_str(pool) or "–",
               None, None, None, None, None]
        for i, v in enumerate(row):
            c = ws.cell(r, 1 + i, v)
            c.font = font(size=9, color=C_MUTED if (i == 8 and not avail_txt) else C_INK)
            c.border = BORDER
            c.alignment = Alignment(wrap_text=True, vertical="top", horizontal="center" if i in (0, 2, 3, 6, 13) else "left")
            if i >= 10:
                c.fill = fill(C_INPUT)
        lines = max(1, -(-len(avail_txt) // 44))
        ws.row_dimensions[r].height = max(16, 12.5 * lines + 3)
        r += 1
    last_item = r - 1
    if todo:
        st_col = f"E{first_item}:E{last_item}"
        ws.conditional_formatting.add(st_col, FormulaRule(formula=[f'E{first_item}="Belum terisi"'], fill=fill(STATUS_FILL["Belum terisi"])))
        ws.conditional_formatting.add(st_col, FormulaRule(formula=[f'E{first_item}="Kosong"'], fill=fill(STATUS_FILL["Slot kosong"])))
        ws.conditional_formatting.add(f"A{first_item}:O{last_item}", FormulaRule(formula=[f'$K{first_item}="Jual minggu ini"'], font=Font(name=FONT, bold=True, color="0D6B3A")))
        ws.auto_filter.ref = f"A{list_head}:O{last_item}"
        dv1 = DataValidation(type="list", formula1='"' + ",".join(DECISIONS) + '"', allow_blank=True)
        dv2 = DataValidation(type="list", formula1='"' + ",".join(AGES) + '"', allow_blank=True)
        dv4 = DataValidation(type="whole", operator="between", formula1="0", formula2="20", allow_blank=True,
                             error="Isi angka 0–20", errorTitle="Target murid")
        for dv in (dv1, dv2, dv4):
            ws.add_data_validation(dv)
        dv1.add(f"K{first_item}:K{last_item}")
        dv2.add(f"L{first_item}:L{last_item}")
        dv4.add(f"N{first_item}:N{last_item}")
    else:
        ws.cell(r, 1, "Tidak ada kelas 0 murid atau slot kosong di center ini.").font = font(italic=True, color=C_MUTED)
        r += 1

    # ----- Coach -----
    r += 2
    coach_row = r
    ws.cell(r, 1, "COACH · jadwal minggu ini di center ini").font = font(True, 13, "FFFFFF")
    for c in range(1, 15):
        ws.cell(r, c).fill = fill(C_ACCENT)
    r += 1
    ws.cell(r, 1, "Angka per hari = jumlah sesi (sebagai coach + assist). Dipakai untuk memilih coach PIC.").font = font(size=9, color=C_MUTED, italic=True)
    r += 1
    ch = ["No", "Coach", "Age group diajar"] + DAYS + ["Sesi coach", "Sesi assist", "Jam pertama–terakhir"]
    for i, h in enumerate(ch):
        c = ws.cell(r, 1 + i, h)
        c.font = font(True, 9, "FFFFFF")
        c.fill = fill(C_INK)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER
    r += 1
    c_first = r
    coaches = sorted(book_c.values(), key=lambda c: (-c["sessions"], c["name"]))
    for n, cinfo in enumerate(coaches, 1):
        allm = sorted(m for s in cinfo["busy"].values() for m in s)
        span = f"{allm[0]//60:02d}:{allm[0]%60:02d}–{allm[-1]//60:02d}:{allm[-1]%60:02d}" if allm else "–"
        ws.cell(r, 1, n)
        ws.cell(r, 2, cinfo["name"])
        ws.cell(r, 3, ages_str(cinfo["ages"]) or "–")
        for i, d in enumerate(DAYS):
            day = f'"* - "&{get_column_letter(4+i)}${c_first-1}'
            ws.cell(r, 4 + i, f'=COUNTIFS({rng("A")},$B$4,{rng("I")},$B{r},{rng("C")},{day})'
                                 f'+COUNTIFS({rng("A")},$B$4,{rng("J")},$B{r},{rng("C")},{day})')
        ws.cell(r, 11, f'=COUNTIFS({rng("A")},$B$4,{rng("I")},$B{r})')
        ws.cell(r, 12, f'=COUNTIFS({rng("A")},$B$4,{rng("J")},$B{r})')
        ws.cell(r, 13, span)
        for col in range(1, 14):
            cell = ws.cell(r, col)
            cell.font = font(size=9, bold=(col == 2))
            cell.border = BORDER
            cell.alignment = Alignment(horizontal="left" if col in (2, 3) else "center")
        r += 1
    c_last = r - 1
    if coaches:
        ws.conditional_formatting.add(f"D{c_first}:J{c_last}", ColorScaleRule(start_type="num", start_value=0, start_color="FFFFFF",
                                                                              end_type="max", end_color="8EC6D0"))
        dv3 = DataValidation(type="list", formula1=f"=$B${c_first}:$B${c_last}", allow_blank=True)
        ws.add_data_validation(dv3)
        if todo:
            dv3.add(f"M{first_item}:M{last_item}")
    else:
        ws.cell(r, 1, "Belum ada coach tercatat.").font = font(italic=True, color=C_MUTED)

    # Navigation links
    ws.cell(nav_row, 1, "Lompat ke").font = font(True, 9, C_MUTED)
    for i, (lab, row_) in enumerate([("Slot Map", slotmap_row), ("Daftar Diskusi", diskusi_row), ("Coach", coach_row)]):
        c = ws.cell(nav_row, 2 + i, f"→ {lab}")
        c.hyperlink = f"#'{center}'!A{row_}"
        c.font = Font(name=FONT, size=10, color=C_ACCENT, underline="single")
    ws.freeze_panes = "B1"


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("feed")
    ap.add_argument("out")
    ap.add_argument("--week", default="")
    ap.add_argument("--asof", default=date.today().isoformat())
    a = ap.parse_args()
    build(a.feed, a.out, a.week, a.asof)
