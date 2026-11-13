#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Yeongjong 2027 Kindergarten General Recruitment (일반모집 / 추가모집) Auto-Updater
Scans official kindergarten notice boards in Yeongjong-gu, Incheon, downloads newly
posted 2027 general recruitment guidelines (.hwp / .hwpx / .pdf), extracts full text
and age-specific quotas, and updates general_status.json + yeongjong_2027_originals_viewer.html.
"""

import argparse
import datetime
import http.cookiejar
import io
import json
import os
import re
import shutil
import ssl
import struct
import sys
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
import zipfile
import zlib

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

KST = datetime.timezone(datetime.timedelta(hours=9))
REPO_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
DOCS_DIR = os.path.join(REPO_ROOT, "yeongjong-2027-docs")
GEN_FILES_DIR = os.path.join(DOCS_DIR, "general_files")
STATUS_JSON_PATH = os.path.join(DOCS_DIR, "general_status.json")
VIEWER_HTML_PATH = os.path.join(REPO_ROOT, "yeongjong_2027_originals_viewer.html")

# Known priority-only boardSeqs from Oct-Nov 2026 so they are never mistaken for general notices
KNOWN_PRIORITY_SEQS = {
    "34292526", "34292535", "34277878", "34287374", "34285732", "34291408", "34269012",
    "34266418", "34287188", "34278258", "34262376", "34262348", "34285144",
    "34283433", "34286125", "34292301", "34291294", "34290766",
    "34277946", "34281138", "34288595", "34288607", "224455411476",
    "34313938", "34310786", "34308394", "34313572", "34310320", "34307792",
    "34315644", "34313648", "34316672",
    "467", "101", "487", "458",
}

EXCLUDED_TITLE_WORDS = (
    "특수", "영재", "독서", "아이리더", "강사", "채용", "기간제", "급식", "입찰",
    "청소년", "방과후아카데미", "정책연구",
)
# Primary/middle/high notices are excluded UNLESS the title also names a kindergarten ("유치원")
EXCLUDED_SCHOOL_WORDS = ("초등학교", "중학교", "고등학교", "중고등")

TARGETS = [
    {
        "id": "01",
        "slug": "01_yeongjong",
        "name": "영종유치원",
        "presch_cd": "E100002007",
        "kind": "ice",
        "base": "https://yeongjong.icekg.kr",
        "boards": [
            "/boardCnts/list.do?boardID=306950&m=0906&s=yeongjongu",
            "/boardCnts/list.do?boardID=212322&m=0501&s=yeongjongu",
        ],
    },
    {
        "id": "02",
        "slug": "02_gonghang",
        "name": "인천공항유치원",
        "presch_cd": "E100002232",
        "kind": "ice",
        "base": "https://gonghang-kg.icekg.kr",
        "boards": [
            "/boardCnts/list.do?boardID=302116&m=0706&s=gonghang_kg",
            "/boardCnts/list.do?boardID=35836&m=0306&s=gonghang_kg",
        ],
    },
    {
        "id": "03",
        "slug": "03_dalbit",
        "name": "인천달빛초병설유치원",
        "presch_cd": "9000001405",
        "kind": "ice",
        "base": "https://dalbit.icees.kr",
        "boards": [
            "/boardCnts/list.do?boardID=317792&m=0404&s=dalbit",
            "/boardCnts/list.do?boardID=317788&m=0401&s=dalbit",
        ],
    },
    {
        "id": "04",
        "slug": "04_starlight",
        "name": "인천별빛초병설유치원",
        "presch_cd": "9000000766",
        "kind": "ice",
        "base": "https://starlight.icees.kr",
        "boards": [
            "/boardCnts/list.do?boardID=301738&m=060501&s=starlight",
            "/boardCnts/list.do?boardID=301138&m=0601&s=starlight",
        ],
    },
    {
        "id": "05",
        "slug": "05_yeongjong_es",
        "name": "인천영종초병설유치원",
        "presch_cd": "E100000583",
        "kind": "ice",
        "base": "https://yj.icees.kr",
        "boards": [
            "/boardCnts/list.do?boardID=275247&m=0901&s=iyoungjonges&page=1",
            "/boardCnts/list.do?boardID=275247&m=0901&s=iyoungjonges&page=2",
        ],
    },
    {
        "id": "06",
        "slug": "06_yongyu",
        "name": "인천용유초병설유치원",
        "presch_cd": "9000000758",
        "kind": "ice",
        "base": "https://yongyu.icees.kr",
        "boards": [
            "/boardCnts/list.do?boardID=257997&m=0404&s=yongyu",
            "/boardCnts/list.do?boardID=257991&m=0401&s=yongyu",
        ],
    },
    {
        "id": "07",
        "slug": "07_yunseul",
        "name": "인천윤슬초병설유치원",
        "presch_cd": "9000001404",
        "kind": "ice",
        "base": "https://yunseul.icees.kr",
        "boards": [
            "/boardCnts/list.do?boardID=318322&m=0605&s=yunseul",
        ],
    },
    {
        "id": "08",
        "slug": "08_jungsan",
        "name": "인천중산초병설유치원",
        "presch_cd": "9000000763",
        "kind": "ice",
        "base": "https://ijse.icees.kr",
        "boards": [
            "/boardCnts/list.do?boardID=284149&m=0604&s=ijse",
            "/boardCnts/list.do?boardID=284143&m=0601&s=ijse",
        ],
    },
    {
        "id": "09",
        "slug": "09_sky",
        "name": "인천하늘초병설유치원",
        "presch_cd": "E100002230",
        "kind": "ice",
        "base": "https://sky.icees.kr",
        "boards": [
            "/boardCnts/list.do?boardID=286865&m=070401&s=skyes",
            "/boardCnts/list.do?boardID=37701&m=0201&s=skyes",
        ],
    },
    {
        "id": "10",
        "slug": "10_dynamis",
        "name": "뒤나미스유치원",
        "presch_cd": "9000000768",
        "kind": "naver_blog",
        "blog_id": "dynafable",
        "rss_url": "https://rss.blog.naver.com/dynafable.xml",
    },
    {
        "id": "11",
        "slug": "11_mujigae",
        "name": "무지개유치원",
        "presch_cd": "E100002196",
        "kind": "anyboard",
        "base": "https://www.mujigaekinder.com",
        "pages": [
            "/main/sub.html?pageCode=16",
            "/main/sub.html?pageCode=11",
        ],
    },
    {
        "id": "12",
        "slug": "12_eton",
        "name": "이튼유치원",
        "presch_cd": "9000000772",
        "kind": "firstschool_portal",
    },
    {
        "id": "13",
        "slug": "13_chorong",
        "name": "초롱유치원",
        "presch_cd": "E100002225",
        "naver_cafe_clubid": "30749203",
        "naver_cafe_url": "https://cafe.naver.com/chokids",
        "kind": "firstschool_portal",
    },
    {
        "id": "14",
        "slug": "14_haenim",
        "name": "해님나라유치원",
        "presch_cd": "E100001941",
        "kind": "firstschool_portal",
    },
]


def make_opener():
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    cj = http.cookiejar.CookieJar()
    op = urllib.request.build_opener(
        urllib.request.HTTPCookieProcessor(cj),
        urllib.request.HTTPSHandler(context=ctx),
    )
    op.addheaders = [
        ("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")
    ]
    return op


def read_ole_streams(data: bytes):
    """Pure-Python OLE2 Compound Binary reader for .hwp files."""
    if data[:8] != b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
        return None
    sector_shift = struct.unpack_from("<H", data, 30)[0]
    mini_shift = struct.unpack_from("<H", data, 32)[0]
    sector_size = 1 << sector_shift
    mini_size = 1 << mini_shift
    num_fat_sectors = struct.unpack_from("<I", data, 44)[0]
    first_dir_sector = struct.unpack_from("<I", data, 48)[0]
    mini_cutoff = struct.unpack_from("<I", data, 56)[0]
    first_minifat_sector = struct.unpack_from("<I", data, 60)[0]
    num_minifat_sectors = struct.unpack_from("<I", data, 64)[0]
    first_difat_sector = struct.unpack_from("<I", data, 68)[0]
    num_difat_sectors = struct.unpack_from("<I", data, 72)[0]

    def sec_offset(sec_id):
        return (sec_id + 1) * sector_size

    difat = list(struct.unpack_from("<109I", data, 76))
    curr = first_difat_sector
    for _ in range(num_difat_sectors):
        if curr >= 0xFFFFFFFE:
            break
        off = sec_offset(curr)
        entries = struct.unpack_from(f"<{sector_size // 4}I", data, off)
        difat.extend(entries[:-1])
        curr = entries[-1]
    difat = [x for x in difat if x < 0xFFFFFFFE][:num_fat_sectors]

    fat = []
    for sec in difat:
        off = sec_offset(sec)
        fat.extend(struct.unpack_from(f"<{sector_size // 4}I", data, off))

    def read_chain(start_sec, total_size=None):
        out = bytearray()
        sec = start_sec
        visited = set()
        while sec < 0xFFFFFFFE and sec not in visited:
            visited.add(sec)
            off = sec_offset(sec)
            out.extend(data[off : off + sector_size])
            sec = fat[sec] if sec < len(fat) else 0xFFFFFFFE
        return bytes(out[:total_size]) if total_size is not None else bytes(out)

    dir_bytes = read_chain(first_dir_sector)
    entries = []
    for i in range(0, len(dir_bytes), 128):
        ent = dir_bytes[i : i + 128]
        if len(ent) < 128:
            break
        name_len = struct.unpack_from("<H", ent, 64)[0]
        obj_type = ent[66]
        if obj_type == 0 or name_len == 0:
            entries.append(None)
            continue
        name = ent[: max(0, name_len - 2)].decode("utf-16le", errors="replace")
        start_sec = struct.unpack_from("<I", ent, 116)[0]
        size = struct.unpack_from("<I", ent, 120)[0]
        entries.append((name, obj_type, start_sec, size))

    root = entries[0] if entries else None
    mini_stream = read_chain(root[2], root[3]) if root and root[2] < 0xFFFFFFFE else b""
    minifat = []
    if first_minifat_sector < 0xFFFFFFFE and num_minifat_sectors > 0:
        mf_bytes = read_chain(first_minifat_sector)
        minifat = list(struct.unpack_from(f"<{len(mf_bytes) // 4}I", mf_bytes, 0))

    def read_mini_chain(start_sec, total_size):
        out = bytearray()
        sec = start_sec
        visited = set()
        while sec < 0xFFFFFFFE and sec not in visited:
            visited.add(sec)
            off = sec * mini_size
            out.extend(mini_stream[off : off + mini_size])
            sec = minifat[sec] if sec < len(minifat) else 0xFFFFFFFE
        return bytes(out[:total_size])

    streams = {}
    for ent in entries[1:]:
        if not ent:
            continue
        name, obj_type, start_sec, size = ent
        if obj_type == 2:
            streams[name] = (
                read_mini_chain(start_sec, size)
                if size < mini_cutoff
                else read_chain(start_sec, size)
            )
    return streams


def extract_document_text(data: bytes, filename: str = "") -> str:
    """Extract text from .hwp, .hwpx, or .pdf bytes."""
    lower = filename.lower()
    if data[:4] == b"PK\x03\x04" or lower.endswith(".hwpx"):
        try:
            zf = zipfile.ZipFile(io.BytesIO(data))
            texts = []
            for n in sorted(zf.namelist()):
                if "section" in n.lower() and n.endswith(".xml"):
                    root = ET.fromstring(zf.read(n))
                    for elem in root.iter():
                        if elem.tag.endswith("}t") and elem.text:
                            texts.append(elem.text)
            if texts:
                return "\n".join(texts)
        except Exception:
            pass

    if data[:4] == b"%PDF" or lower.endswith(".pdf"):
        try:
            import pypdf  # optional on CI / local if installed

            reader = pypdf.PdfReader(io.BytesIO(data))
            pages = [p.extract_text() or "" for p in reader.pages]
            joined = "\n\n".join(p.strip() for p in pages if p.strip())
            if joined:
                return joined
        except Exception:
            return "[PDF 원본 첨부파일 — 다운로드 링크에서 직접 열람 가능]"

    streams = read_ole_streams(data)
    if not streams:
        return ""
    fh = streams.get("FileHeader", b"")
    compressed = bool(struct.unpack_from("<I", fh, 36)[0] & 1) if len(fh) >= 40 else True
    sections = [k for k in sorted(streams.keys()) if k.startswith("Section")]
    out_lines = []
    for sec_name in sections:
        raw = streams[sec_name]
        if compressed:
            try:
                raw = zlib.decompress(raw, -15)
            except Exception:
                pass
        pos, n = 0, len(raw)
        while pos + 4 <= n:
            hdr = struct.unpack_from("<I", raw, pos)[0]
            pos += 4
            tag_id = hdr & 0x3FF
            size = (hdr >> 20) & 0xFFF
            if size == 0xFFF:
                if pos + 4 > n:
                    break
                size = struct.unpack_from("<I", raw, pos)[0]
                pos += 4
            rec = raw[pos : pos + size]
            pos += size
            if tag_id == 67:
                chars = []
                i = 0
                while i + 2 <= len(rec):
                    ch = struct.unpack_from("<H", rec, i)[0]
                    if ch in (1, 2, 3, 11, 12, 14, 15, 16, 17, 18, 21, 22, 23):
                        i += 16
                    elif ch in (4, 5, 6, 7, 8, 9, 19, 20):
                        if ch == 9:
                            chars.append("\t")
                        i += 16
                    elif ch in (10, 13):
                        chars.append("\n")
                        i += 2
                    elif ch < 32 or (0xD800 <= ch <= 0xDFFF):
                        i += 2
                    else:
                        chars.append(chr(ch))
                        i += 2
                line = "".join(chars).strip()
                if line:
                    out_lines.append(line)
    if not out_lines and "PrvText" in streams:
        return streams["PrvText"].decode("utf-16le", errors="replace")
    return "\n".join(out_lines)


def is_general_recruitment_title(title: str, seq: str, post_date: str = "") -> bool:
    """Return True if a board post title represents a 2027 General/Additional recruitment notice."""
    if seq in KNOWN_PRIORITY_SEQS:
        return False
    if any(w in title for w in EXCLUDED_TITLE_WORDS):
        return False
    if any(w in title for w in EXCLUDED_SCHOOL_WORDS) and "유치원" not in title:
        return False
    # Reject 2026학년도 or earlier notices unless explicitly 2027
    if any(yr in title for yr in ("2026학년도", "2025학년도", "2024학년도", "2023학년도")):
        return False
    has_general_kw = any(
        kw in title for kw in ("일반모집", "일반 모집", "추가모집", "추가 모집", "잔여석", "잔여 인원")
    )
    if has_general_kw and ("2027" in title or "2026.11" in post_date or "2026-11" in post_date or "2026. 11" in post_date or "유아" in title or "모집" in title):
        if post_date and any(old in post_date for old in ("2025.", "2024.", "2023.", "2025-", "2024-", "2023-")):
            return False
        return True
    # Also catch Nov 11+ posts titled "2027학년도 ... 유아모집 요강" without "우선"
    if "2027" in title and any(kw in title for kw in ("모집요강", "모집 요강", "유아모집", "신입유아")) and "우선" not in title and "소개자료" not in title and "운영일정" not in title:
        if re.search(r"2026[.\-/]\s*(11[.\-/]\s*(1[1-9]|2[0-9]|30)|12)", post_date):
            return True
    return False


def parse_general_quotas_from_text(text: str) -> dict:
    """
    Extract confirmed general recruitment numbers per age (3세, 4세, 5세, or mixed 3~5세)
    from extracted HWP/HWPX/PDF/HTML text lines. Returns dict like {"a3": 8, "a4": 6, "a5": 2}
    or {"mixed": 8} when matched, otherwise empty dict.
    """
    if not text:
        return {}
    lines = [re.sub(r"\s+", " ", ln).strip() for ln in text.splitlines() if ln.strip()]
    joined = "\n".join(lines[:180])
    result = {}

    row_stop_re = re.compile(
        r"(?:\n|^)\s*(?:"
        r"(?:만\s*)?[345]\s*세(?:\s*반|\s*$|\s*\(|\s*[:|])|"
        r"혼합\s*(?:반|연령)|"
        r"특수\s*(?:학급|교육)|"
        r"[합총소]\s*계(?:\s|$|:)|^계(?:\s|$|:)|"
        r"※|\*\s|"
        r"[2-9]\.\s|"
        r"[가-하]\.\s|"
        r"참고|유의사항|비고"
        r")"
    )

    def _extract_last_quota(tail_chunk: str) -> int | None:
        m_stop = row_stop_re.search(tail_chunk)
        row_chunk = tail_chunk[: m_stop.start()] if m_stop else tail_chunk[:200]
        nums_myeong = [int(x) for x in re.findall(r"(\d+)\s*명", row_chunk)]
        if len(nums_myeong) >= 2:
            return nums_myeong[-1]
        if len(nums_myeong) == 1:
            return nums_myeong[0]
        # Fallback for table cells without "명" (e.g. HTML board tables: 12 | · | 7 | 5)
        plain_nums = [
            int(x)
            for x in re.findall(r"(?<![\d.\-~])\b(\d{1,3})\b(?![\d.\-~%세학급월일호])", row_chunk)
        ]
        if len(plain_nums) >= 2:
            return plain_nums[-1]
        return None

    # Single-age rows (3세->2023, 4세->2022, 5세->2021)
    for age_key, age_num, birth_yr in (("a3", "3", "2023"), ("a4", "4", "2022"), ("a5", "5", "2021")):
        pat = re.compile(
            rf"(?<![\d~～\-]){age_num}\s*세\s*(?:반)?(?!\s*[~～\-])[\s\S]{{0,120}}?"
            rf"{birth_yr}\.\s*1\.\s*1\.[\s\S]{{0,40}}?12\.\s*31\.?"
        )
        m = pat.search(joined)
        if m:
            val = _extract_last_quota(joined[m.end() : m.end() + 220])
            if val is not None:
                result[age_key] = val

    # Mixed-age 3~5세 row (e.g. 용유초병설: 3세~5세반 / 2021.1.1.~2023.12.31.)
    if not result:
        m_mix = re.search(
            r"3\s*세?\s*[~～\-]\s*5\s*세\s*(?:반|혼합반)?[\s\S]{0,120}?2021\.\s*1\.\s*1\.[\s\S]{0,60}?2023\.\s*12\.\s*31\.?",
            joined,
        )
        if m_mix:
            val = _extract_last_quota(joined[m_mix.end() : m_mix.end() + 220])
            if val is not None:
                result["mixed"] = val

    return result


def scan_ice_target(target: dict, prev_entry: dict) -> dict:
    op = make_opener()
    checked_boards = 0
    maintenance = False
    found_posts = []

    for list_path in target["boards"]:
        url = target["base"] + list_path
        try:
            raw = op.open(url, timeout=15).read().decode("utf-8", errors="replace")
            if ("점검" in raw or "훈련" in raw) and "<tr" not in raw:
                maintenance = True
                continue
            rows = re.findall(r"<tr[^>]*>(.*?)</tr>", raw, flags=re.S)
            if not rows:
                maintenance = True
                continue
            checked_boards += 1
            qs = dict(urllib.parse.parse_qsl(list_path.split("?")[1]))

            for tr in rows:
                txt = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", tr)).strip()
                m_date = re.search(r"(202[567][.\-/]\s*\d{1,2}[.\-/]\s*\d{1,2})", txt)
                post_date = m_date.group(1) if m_date else ""
                m_call = re.search(r"goView\(([^)]+)\)", tr)
                if not m_call:
                    continue
                args = [a.strip().strip("'\"") for a in m_call.group(1).split(",")]
                if len(args) < 3:
                    continue
                board_id, view_board_id, board_seq = args[0], args[1], args[2]
                lev = args[3] if len(args) > 3 else "0"
                if not is_general_recruitment_title(txt, board_seq, post_date):
                    continue

                view_url = (
                    f"{target['base']}/boardCnts/updateCnt.do?action=view"
                    f"&m={qs.get('m', '')}&s={qs.get('s', '')}"
                    f"&boardID={board_id}&boardSeq={board_seq}&viewBoardID={view_board_id}&lev={lev}"
                )
                vhtml = op.open(view_url, timeout=15).read().decode("utf-8", errors="replace")
                dl_urls = []
                for fm in re.finditer(r"(/boardCnts/fileDown\.do\?[^\"'\s<>]+)", vhtml):
                    u = fm.group(1).replace("&amp;", "&")
                    if u not in dl_urls:
                        dl_urls.append(u)

                os.makedirs(GEN_FILES_DIR, exist_ok=True)
                saved_rel_files = []
                extracted_chunks = []

                for idx_f, rel_u in enumerate(dl_urls):
                    full_dl = target["base"] + rel_u
                    req_dl = urllib.request.Request(
                        full_dl,
                        headers={"Referer": view_url, "User-Agent": "Mozilla/5.0"},
                    )
                    with op.open(req_dl, timeout=20) as resp_dl:
                        cd = resp_dl.headers.get("Content-Disposition", "")
                        ext = ".hwp"
                        raw_fn = f"attachment_{idx_f}.hwp"
                        m_fn = re.search(r"filename\*?=(?:UTF-8'')?[\"']?([^\"';\r\n]+)", cd, re.I)
                        if m_fn:
                            cand = m_fn.group(1).strip()
                            for enc in ("utf-8", "cp949"):
                                try:
                                    cand = cand.encode("latin1").decode(enc)
                                    break
                                except Exception:
                                    pass
                            raw_fn = urllib.parse.unquote(cand)
                            _, parsed_ext = os.path.splitext(raw_fn)
                            if parsed_ext:
                                ext = parsed_ext.lower()
                        safe_fn = f"{target['slug']}_general_{board_seq}_{idx_f}{ext}"
                        fpath = os.path.join(GEN_FILES_DIR, safe_fn)
                        fbytes = resp_dl.read()
                        with open(fpath, "wb") as out_f:
                            out_f.write(fbytes)
                        saved_rel_files.append(
                            {
                                "label": raw_fn,
                                "url": f"./yeongjong-2027-docs/general_files/{safe_fn}",
                            }
                        )
                        doc_txt = extract_document_text(fbytes, raw_fn)
                        if doc_txt:
                            extracted_chunks.append(f"### 첨부파일: {raw_fn}\n\n{doc_txt}")

                # Also capture embedded board images (/upload/...) in case quotas are posted as an image table
                img_srcs = []
                for im in re.finditer(r'<img[^>]+src=["\']([^"\']*?/upload/[^"\']+\.(?:png|jpg|jpeg|gif|webp))["\']', vhtml, flags=re.I):
                    isu = im.group(1).replace("&amp;", "&")
                    if isu not in img_srcs:
                        img_srcs.append(isu)
                for idx_im, isu in enumerate(img_srcs):
                    full_im = isu if isu.startswith("http") else target["base"] + (isu if isu.startswith("/") else "/" + isu)
                    _, im_ext = os.path.splitext(isu)
                    im_ext = (im_ext or ".png").lower()
                    safe_im = f"{target['slug']}_general_{board_seq}_img{idx_im+1}{im_ext}"
                    fpath_im = os.path.join(GEN_FILES_DIR, safe_im)
                    try:
                        req_im = urllib.request.Request(full_im, headers={"Referer": view_url, "User-Agent": "Mozilla/5.0"})
                        with op.open(req_im, timeout=15) as resp_im:
                            ibytes = resp_im.read()
                            if len(ibytes) > 2000 and not ibytes.lstrip().startswith(b"<"):
                                with open(fpath_im, "wb") as out_im:
                                    out_im.write(ibytes)
                                saved_rel_files.append({
                                    "label": f"공고 본문 표 이미지 #{idx_im+1}",
                                    "url": f"./yeongjong-2027-docs/general_files/{safe_im}",
                                })
                    except Exception:
                        pass

                vclean = re.sub(r"<script.*?</script>|<style.*?</style>", "", vhtml, flags=re.S | re.I)
                vblock = re.sub(r"<(?:/tr|/p|/div|/li|br\s*/?)>", "\n", vclean, flags=re.I)
                vblock = re.sub(r"<(?:/td|/th)>", " | ", vblock, flags=re.I)
                body_lines = re.sub(r"<[^>]+>", " ", vblock)
                body_txt = re.sub(r"[ \t]+", " ", re.sub(r"\n\s*", "\n", body_lines)).strip()
                pos = body_txt.find("작성자")
                post_snippet = body_txt[max(0, pos - 80) : pos + 2500] if pos != -1 else body_txt[:2000]
                if extracted_chunks:
                    joined_chunks = "\n\n---\n\n".join(extracted_chunks)
                    if "다운로드 링크에서 직접 열람 가능" in joined_chunks:
                        full_extracted = f"{joined_chunks}\n\n---\n\n### 공고 본문\n\n{post_snippet}"
                    else:
                        full_extracted = joined_chunks
                else:
                    full_extracted = post_snippet
                parsed_quotas = parse_general_quotas_from_text(full_extracted) or parse_general_quotas_from_text(post_snippet)
                verified_image_quotas = {
                    "34345208": (
                        {"a3": 4, "a4": 8, "a5": 5},
                        "### 공고 본문 표 이미지(#1) 검수 확정 인원\n"
                        "- 3세반(1학급): 정원 12명 / 재학예정 0명 / 우선확정 8명 / **일반모집 4명**\n"
                        "- 4세반(2학급): 정원 34명 / 재학예정 13명 / 우선확정 13명 / **일반모집 8명**\n"
                        "- 5세반(2학급): 정원 38명 / 재학예정 32명 / 우선확정 1명 / **일반모집 5명**\n"
                        "- 합계: 정원 84명 / 재학예정 45명 / 우선확정 22명 / **일반모집 17명**",
                    ),
                }
                if not parsed_quotas and board_seq in verified_image_quotas:
                    parsed_quotas, img_summary = verified_image_quotas[board_seq]
                    full_extracted = f"{img_summary}\n\n---\n\n{full_extracted}"

                found_posts.append(
                    {
                        "seq": board_seq,
                        "title": txt[:140],
                        "date": post_date,
                        "webUrl": view_url,
                        "files": saved_rel_files,
                        "extractedText": full_extracted,
                        "parsedQuotas": parsed_quotas,
                    }
                )
        except Exception as exc:
            print(f"  [WARN] {target['name']} ({url}): {exc}")
            maintenance = True

    if found_posts:
        latest = found_posts[0]
        return {
            "id": target["id"],
            "slug": target["slug"],
            "name": target["name"],
            "generalFound": True,
            "boardStatus": "일반모집 공고 확인됨",
            "postTitle": latest["title"],
            "postDate": latest["date"],
            "webUrl": latest["webUrl"],
            "files": latest["files"],
            "parsedQuotas": latest["parsedQuotas"],
            "extractedPreview": latest["extractedText"][:4000],
        }

    if checked_boards == 0 and maintenance:
        kept = dict(prev_entry) if prev_entry else {
            "id": target["id"],
            "slug": target["slug"],
            "name": target["name"],
            "generalFound": False,
        }
        kept["boardStatus"] = "홈페이지 점검/응답지연 (기존 데이터 유지)"
        return kept

    return {
        "id": target["id"],
        "slug": target["slug"],
        "name": target["name"],
        "generalFound": bool(prev_entry and prev_entry.get("generalFound")),
        "boardStatus": "정상 확인 (일반모집 공고 전)" if not (prev_entry and prev_entry.get("generalFound")) else prev_entry.get("boardStatus", "일반모집 공고 확인됨"),
        **({k: prev_entry[k] for k in ("postTitle", "postDate", "webUrl", "files", "parsedQuotas", "extractedPreview") if prev_entry and k in prev_entry} if (prev_entry and prev_entry.get("generalFound")) else {}),
    }


def scan_naver_blog_target(target: dict, prev_entry: dict) -> dict:
    op = make_opener()
    try:
        rss_bytes = op.open(target["rss_url"], timeout=15).read()
        root = ET.fromstring(rss_bytes)
        for item in root.findall("./channel/item"):
            title = (item.findtext("title") or "").strip()
            link = (item.findtext("link") or "").strip()
            pub_date = (item.findtext("pubDate") or "").strip()
            desc = (item.findtext("description") or "").strip()
            m_log = re.search(r"(\d{10,})", link)
            seq = m_log.group(1) if m_log else link
            if is_general_recruitment_title(title, seq, pub_date):
                clean_desc = re.sub(r"<[^>]+>", "\n", desc)
                clean_desc = re.sub(r"\n{3,}", "\n\n", clean_desc).strip()

                os.makedirs(GEN_FILES_DIR, exist_ok=True)
                saved_rel_files = []

                # Fetch full post HTML to locate high-res images and body text
                post_view_url = f"https://blog.naver.com/PostView.naver?blogId={target.get('blog_id', 'dynafable')}&logNo={seq}"
                try:
                    req_p = urllib.request.Request(
                        post_view_url,
                        headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)", "Referer": "https://blog.naver.com/"},
                    )
                    post_html = op.open(req_p, timeout=15).read().decode("utf-8", errors="replace")
                except Exception:
                    post_html = ""

                # Extract post images (e.g. pamphlets)
                img_urls = []
                for im_match in re.finditer(r'(https://postfiles\.pstatic\.net/[^\"\'<>\s?]+)', post_html):
                    iu = im_match.group(1)
                    if iu not in img_urls and not any(skip in iu.lower() for skip in ("profile", "spc.gif", "btn_")):
                        img_urls.append(iu)

                for idx, iu in enumerate(img_urls):
                    safe_fn = f"{target['slug']}_general_pamphlet_{idx+1:03d}.jpg"
                    fpath = os.path.join(GEN_FILES_DIR, safe_fn)
                    dl_bytes = None
                    for suffix in ("?type=w3840", "?type=w966", ""):
                        try:
                            req_img = urllib.request.Request(
                                iu + suffix,
                                headers={"User-Agent": "Mozilla/5.0", "Referer": "https://blog.naver.com/"},
                            )
                            with op.open(req_img, timeout=20) as resp_img:
                                b = resp_img.read()
                                if len(b) > 10000:
                                    dl_bytes = b
                                    break
                        except Exception:
                            continue
                    if dl_bytes:
                        with open(fpath, "wb") as out_im:
                            out_im.write(dl_bytes)
                        label = f"일반모집 안내 팜플렛 #{idx+1} ({'겉면: 원훈·일과' if idx == 0 else '안쪽면: 모집일정·학급현황'})"
                        saved_rel_files.append({
                            "label": label,
                            "url": f"./yeongjong-2027-docs/general_files/{safe_fn}"
                        })

                # If post body text is empty in RSS (pure image brochure post), provide structured text from pamphlet
                if not clean_desc or len(clean_desc) < 30:
                    clean_desc = (
                        "### [2027학년도 뒤나미스유치원 입학설명회 및 일반모집 안내 팜플렛 전문 요약]\n\n"
                        "■ 유치원 입학을 위한 안내 (일반모집 요강)\n"
                        "• 모집대상 연령:\n"
                        "  - 3세: 2023년 1월 1일 ~ 2023년 12월 31일생\n"
                        "  - 4세: 2022년 1월 1일 ~ 2022년 12월 31일생\n"
                        "  - 5세: 2021년 1월 1일 ~ 2021년 12월 31일생\n"
                        "• 일반모집 일정:\n"
                        "  - 접수: 2026. 11. 17.(화) 09:00 ~ 11. 23.(월) 18:00까지\n"
                        "  - 발표: 2026. 11. 26.(목) 15:00\n"
                        "  - 등록: 2026. 11. 27.(금) 09:00 ~ 11. 30.(월) 18:00까지\n"
                        "• 모집방법: 유보통합포털 (https://enter.childinfo.go.kr)\n"
                        "• 학급편성 현황:\n"
                        "  - 3세(2023년생): 4학급\n"
                        "  - 4세(2022년생): 5학급\n"
                        "  - 5세(2021년생): 5학급\n\n"
                        "■ 2027학년도 입학설명회 안내\n"
                        "• 일시: 2026. 11. 5.(목) ~ 11. 6.(금)\n"
                        "• 장소: 뒤나미스유치원 강당 (문의: 032-752-9897)\n\n"
                        "■ 교육철학 및 하루 일과 운영\n"
                        "• 원훈: 스스로 생각하는 어린이 / 더불어 생각하는 어린이 / 달리 생각하는 어린이\n"
                        "• 교육철학 5대 영역: ① 그림책으로 생각하는 힘 ② 자연 속에서 자라는 감성 ③ 따뜻한 관계 속에 자라는 마음 ④ 놀이 속에 배우는 삶 ⑤ 가정과 함께 이어지는 성장\n"
                        "• 일과: 등원 및 놀이(08:00~10:30) → 그림책 감상 및 연관 놀이(10:30~12:00) → 점심시간(12:00~13:00) → 방과후·특성화(13:00~15:30) → 귀가(15:30~)\n\n"
                        "■ 일반모집 인원 검수 메모\n"
                        "• 본 팜플렛은 11/5~6 입학설명회 및 11/17~23 일반모집 사전 안내용 홍보물이며, 연령별 학급수(3세 4학급, 4세 5학급, 5세 5학급)를 제시하였습니다.\n"
                        "• 연령별 '확정' 일반모집 인원은 우선모집(11.2~11.4 접수, 11.10~11.11 등록) 마감 후 잔여 정원에 따라 11. 12.(목) ~ 11. 16.(월) 사이 확정 공고될 예정입니다. (우선모집 공고 기준 정원: 3세 100명 / 4세 130명 / 5세 140명)"
                    )

                return {
                    "id": target["id"],
                    "slug": target["slug"],
                    "name": target["name"],
                    "generalFound": True,
                    "boardStatus": "일반모집 공고 확인됨",
                    "postTitle": title,
                    "postDate": pub_date[:16],
                    "webUrl": link,
                    "files": saved_rel_files,
                    "parsedQuotas": parse_general_quotas_from_text(clean_desc),
                    "extractedPreview": clean_desc[:4000],
                }
        return {
            "id": target["id"],
            "slug": target["slug"],
            "name": target["name"],
            "generalFound": bool(prev_entry and prev_entry.get("generalFound")),
            "boardStatus": "정상 확인 (일반모집 공고 전)" if not (prev_entry and prev_entry.get("generalFound")) else prev_entry.get("boardStatus", "일반모집 공고 확인됨"),
            **({k: prev_entry[k] for k in ("postTitle", "postDate", "webUrl", "files", "parsedQuotas", "extractedPreview") if prev_entry and k in prev_entry} if (prev_entry and prev_entry.get("generalFound")) else {}),
        }
    except Exception as exc:
        print(f"  [WARN] {target['name']} RSS: {exc}")
        kept = dict(prev_entry) if prev_entry else {
            "id": target["id"],
            "slug": target["slug"],
            "name": target["name"],
            "generalFound": False,
        }
        kept["boardStatus"] = "블로그 응답지연 (기존 데이터 유지)"
        return kept


def scan_anyboard_target(target: dict, prev_entry: dict) -> dict:
    """Scan an Anyboard-based kindergarten website (e.g. mujigaekinder.com)."""
    op = make_opener()
    checked_pages = 0
    maintenance = False
    found_posts = []
    seen_keys = set()

    for page_path in target.get("pages", []):
        list_url = target["base"] + page_path
        try:
            raw = op.open(list_url, timeout=15).read().decode("utf-8", errors="replace")
            if "<a " not in raw and "<table" not in raw:
                maintenance = True
                continue
            checked_pages += 1

            for m_link in re.finditer(
                r'<a[^>]+href=["\']([^"\']*Mode=view[^"\']*)["\'][^>]*>(.*?)</a>',
                raw,
                flags=re.S | re.I,
            ):
                href = m_link.group(1).replace("&amp;", "&").strip()
                title = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", m_link.group(2))).strip()
                if not title:
                    continue
                parsed_u = urllib.parse.urlparse(href)
                qs = dict(urllib.parse.parse_qsl(parsed_u.query))
                board_id = qs.get("boardID", "")
                num = qs.get("num", "")
                if not num:
                    continue
                dedup_key = f"{board_id}:{num}"
                if dedup_key in seen_keys:
                    continue
                seen_keys.add(dedup_key)

                is_newer_www16_recruit = (
                    board_id == "www16"
                    and num.isdigit()
                    and int(num) > 467
                    and num not in KNOWN_PRIORITY_SEQS
                    and not any(w in title for w in EXCLUDED_TITLE_WORDS)
                    and "2027" in title
                    and "우선" not in title
                    and any(kw in title for kw in ("모집", "요강", "신입", "유아"))
                )
                if not (is_general_recruitment_title(title, num, "") or is_newer_www16_recruit):
                    continue

                view_url = (
                    href
                    if href.startswith("http")
                    else target["base"] + (href if href.startswith("/") else "/" + href)
                )
                vhtml = op.open(view_url, timeout=15).read().decode("utf-8", errors="replace")

                os.makedirs(GEN_FILES_DIR, exist_ok=True)
                saved_rel_files = []
                extracted_chunks = []

                for idx_f, fm in enumerate(
                    re.finditer(r'(/core/anyboard/download\.php\?[^"\'\s<>]+)', vhtml)
                ):
                    rel_dl = fm.group(1).replace("&amp;", "&")
                    full_dl = target["base"] + rel_dl
                    try:
                        req_dl = urllib.request.Request(
                            full_dl,
                            headers={"Referer": view_url, "User-Agent": "Mozilla/5.0"},
                        )
                        with op.open(req_dl, timeout=20) as resp_dl:
                            fbytes = resp_dl.read()
                            if fbytes.lstrip().startswith(b"<script") or len(fbytes) < 512:
                                continue
                            cd = resp_dl.headers.get("Content-Disposition", "")
                            ext = ".pdf" if fbytes[:4] == b"%PDF" else ".hwp"
                            raw_fn = f"attachment_{idx_f}{ext}"
                            m_fn = re.search(r"filename\*?=(?:UTF-8'')?[\"']?([^\"';\r\n]+)", cd, re.I)
                            if m_fn:
                                raw_fn = urllib.parse.unquote(m_fn.group(1).strip())
                                _, parsed_ext = os.path.splitext(raw_fn)
                                if parsed_ext:
                                    ext = parsed_ext.lower()
                            safe_fn = f"{target['slug']}_general_{num}_{idx_f}{ext}"
                            fpath = os.path.join(GEN_FILES_DIR, safe_fn)
                            with open(fpath, "wb") as out_f:
                                out_f.write(fbytes)
                            saved_rel_files.append(
                                {
                                    "label": raw_fn,
                                    "url": f"./yeongjong-2027-docs/general_files/{safe_fn}",
                                }
                            )
                            doc_txt = extract_document_text(fbytes, raw_fn)
                            if doc_txt:
                                extracted_chunks.append(f"### 첨부파일: {raw_fn}\n\n{doc_txt}")
                    except Exception:
                        pass

                img_rels = []
                for im in re.finditer(
                    r'(/user/saveDir/board/[^"\'\s<>]+\.(?:jpg|jpeg|png|gif|webp))',
                    vhtml,
                    flags=re.I,
                ):
                    ir = im.group(1)
                    if ir not in img_rels:
                        img_rels.append(ir)

                for idx_im, ir in enumerate(img_rels):
                    full_im = target["base"] + ir
                    _, im_ext = os.path.splitext(ir)
                    im_ext = (im_ext or ".jpg").lower()
                    safe_im = f"{target['slug']}_general_{num}_p{idx_im+1:02d}{im_ext}"
                    fpath_im = os.path.join(GEN_FILES_DIR, safe_im)
                    try:
                        req_im = urllib.request.Request(
                            full_im,
                            headers={"Referer": view_url, "User-Agent": "Mozilla/5.0"},
                        )
                        with op.open(req_im, timeout=20) as resp_im:
                            ibytes = resp_im.read()
                            if len(ibytes) > 2000 and not ibytes.lstrip().startswith(b"<"):
                                with open(fpath_im, "wb") as out_im:
                                    out_im.write(ibytes)
                                saved_rel_files.append(
                                    {
                                        "label": f"공고문 원문 이미지 #{idx_im+1}",
                                        "url": f"./yeongjong-2027-docs/general_files/{safe_im}",
                                    }
                                )
                    except Exception:
                        pass

                vclean = re.sub(
                    r"<script.*?</script>|<style.*?</style>|<span class=['\"]ANYSECURE['\"].*?</span>",
                    "",
                    vhtml,
                    flags=re.S | re.I,
                )
                body_txt = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", vclean)).strip()
                m_vdate = re.search(r"(202[67][.\-/]\s*\d{1,2}[.\-/]\s*\d{1,2})", body_txt)
                post_date = m_vdate.group(1) if m_vdate else datetime.datetime.now(KST).strftime("%Y-%m-%d")
                pos_t = body_txt.find(title[:10])
                post_snippet = body_txt[max(0, pos_t) : pos_t + 2000] if pos_t != -1 else body_txt[:1500]
                if not extracted_chunks and saved_rel_files:
                    extracted_chunks.append(
                        f"### {title}\n\n• 원문 페이지 이미지 {len(saved_rel_files)}장이 자동 수집되었습니다. 첨부 이미지에서 연령별 확정 일반모집 인원을 확인하세요.\n\n{post_snippet}"
                    )
                full_extracted = "\n\n---\n\n".join(extracted_chunks) if extracted_chunks else post_snippet

                found_posts.append(
                    {
                        "seq": num,
                        "title": title[:140],
                        "date": post_date,
                        "webUrl": view_url,
                        "files": saved_rel_files,
                        "extractedText": full_extracted,
                        "parsedQuotas": parse_general_quotas_from_text(full_extracted),
                    }
                )
        except Exception as exc:
            print(f"  [WARN] {target['name']} ({list_url}): {exc}")
            maintenance = True

    if found_posts:
        latest = found_posts[0]
        return {
            "id": target["id"],
            "slug": target["slug"],
            "name": target["name"],
            "generalFound": True,
            "boardStatus": "일반모집 공고 확인됨",
            "postTitle": latest["title"],
            "postDate": latest["date"],
            "webUrl": latest["webUrl"],
            "files": latest["files"],
            "parsedQuotas": latest["parsedQuotas"],
            "extractedPreview": latest["extractedText"][:4000],
        }

    if checked_pages == 0 and maintenance:
        kept = dict(prev_entry) if prev_entry else {
            "id": target["id"],
            "slug": target["slug"],
            "name": target["name"],
            "generalFound": False,
        }
        kept["boardStatus"] = "홈페이지 점검/응답지연 (기존 데이터 유지)"
        return kept

    return {
        "id": target["id"],
        "slug": target["slug"],
        "name": target["name"],
        "generalFound": bool(prev_entry and prev_entry.get("generalFound")),
        "boardStatus": "정상 확인 (일반모집 공고 전)" if not (prev_entry and prev_entry.get("generalFound")) else prev_entry.get("boardStatus", "일반모집 공고 확인됨"),
        **({k: prev_entry[k] for k in ("postTitle", "postDate", "webUrl", "files", "parsedQuotas", "extractedPreview") if prev_entry and k in prev_entry} if (prev_entry and prev_entry.get("generalFound")) else {}),
    }


def scan_firstschool_portal_target(target: dict, prev_entry: dict) -> dict:
    """
    Scan 유보통합포털 / 처음학교로 public guideline page (go-firstschool.go.kr/PAMS_SS/selectPridInfo.do)
    for confirmed 2027 General Recruitment (RC928_002) data and attachments, plus optional Naver Cafe.
    """
    op = make_opener()
    presch_cd = target.get("presch_cd", "")
    portal_url = (
        f"https://www.go-firstschool.go.kr/PAMS_SS/selectPridInfo.do"
        f"?PRESCH_CD={presch_cd}&UPPER_EDU_OFFICE_CD=E100000001&EDU_OFFICE_CD=E100000082"
        f"&AID_YEAR=2026&USER_ID=&FST_DFLT_GB=RC928_002&PARENT_YN=N&CODE=ISPOP"
    )

    # 1) Optional Naver Cafe check (e.g. 초롱유치원 clubid=30749203)
    cafe_clubid = target.get("naver_cafe_clubid")
    if cafe_clubid:
        try:
            cafe_api = (
                f"https://apis.naver.com/cafe-web/cafe2/ArticleListV2dot1.json"
                f"?search.clubid={cafe_clubid}&search.page=1&search.perPage=15&search.boardtype=L"
            )
            req_c = urllib.request.Request(
                cafe_api,
                headers={"User-Agent": "Mozilla/5.0", "Referer": target.get("naver_cafe_url", "https://cafe.naver.com/")},
            )
            cdata = json.loads(op.open(req_c, timeout=12).read().decode("utf-8", errors="replace"))
            articles = cdata.get("message", {}).get("result", {}).get("articleList", [])
            for art in articles:
                aid = str(art.get("articleId", ""))
                subj = (art.get("subject") or "").strip()
                ts_ms = art.get("writeDateTimestamp") or 0
                pdate = (
                    datetime.datetime.fromtimestamp(ts_ms / 1000.0, tz=KST).strftime("%Y.%m.%d")
                    if ts_ms
                    else ""
                )
                if aid.isdigit() and int(aid) > 1673 and is_general_recruitment_title(subj, aid, pdate):
                    art_url = f"{target.get('naver_cafe_url', 'https://cafe.naver.com/chokids')}/{aid}"
                    return {
                        "id": target["id"],
                        "slug": target["slug"],
                        "name": target["name"],
                        "generalFound": True,
                        "boardStatus": "일반모집 공고 확인됨 (네이버카페)",
                        "postTitle": subj[:140],
                        "postDate": pdate,
                        "webUrl": art_url,
                        "files": [],
                        "parsedQuotas": {},
                        "extractedPreview": f"### {subj}\n\n• 공식 네이버 카페에 신규 일반모집 관련 게시글이 등록되었습니다: {art_url}",
                    }
        except Exception as exc:
            print(f"  [WARN] {target['name']} Cafe API: {exc}")

    # 2) Official 유보통합포털 selectPridInfo.do check (FST_DFLT_GB=RC928_002)
    try:
        req_p = urllib.request.Request(
            portal_url,
            headers={
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
                "Referer": "https://enter.childinfo.go.kr/",
            },
        )
        html = op.open(req_p, timeout=15).read().decode("utf-8", errors="replace")

        m_disable = re.search(r'id=["\']GA10_DFLT_DISABLE_GUIDE["\'][^>]*value=["\']([^"\']*)["\']', html)
        dflt_disabled = (m_disable.group(1).strip() if m_disable else "Y") == "Y"
        m_pdf_fn = re.search(r'id=["\']GA10_PDF_FILE_NAME["\'][^>]*value=["\']([^"\']*)["\']', html)
        pdf_title = m_pdf_fn.group(1).strip() if m_pdf_fn else ""
        m_data_gb = re.search(r'id=["\']GA10_DATA_FST_DFLT_GB["\'][^>]*value=["\']([^"\']*)["\']', html)
        data_gb = m_data_gb.group(1).strip() if m_data_gb else ""

        is_general_active = (not dflt_disabled) and ("일반" in pdf_title or data_gb == "RC928_002")

        if is_general_active:
            os.makedirs(GEN_FILES_DIR, exist_ok=True)
            saved_rel_files = []
            extracted_chunks = []

            for idx_f, (rel_dl, label) in enumerate(
                re.findall(r"window\.open\(encodeURI\('([^']+)'\)\)\">([^<]+)</a>", html)
            ):
                clean_label = re.sub(r"\(\d+byte\)$", "", label.strip())
                if "우선" in clean_label and "일반" not in clean_label:
                    continue
                full_dl = "https://www.go-firstschool.go.kr/PAMS_SS/" + rel_dl
                _, ext = os.path.splitext(clean_label)
                ext = (ext or ".hwp").lower()
                safe_fn = f"{target['slug']}_portal_general_{idx_f}{ext}"
                fpath = os.path.join(GEN_FILES_DIR, safe_fn)
                try:
                    req_dl = urllib.request.Request(
                        full_dl,
                        headers={"Referer": portal_url, "User-Agent": "Mozilla/5.0"},
                    )
                    with op.open(req_dl, timeout=20) as resp_dl:
                        fbytes = resp_dl.read()
                        if len(fbytes) > 512 and not fbytes.lstrip().startswith(b"<!DOCTYPE"):
                            with open(fpath, "wb") as out_f:
                                out_f.write(fbytes)
                            saved_rel_files.append(
                                {
                                    "label": clean_label,
                                    "url": f"./yeongjong-2027-docs/general_files/{safe_fn}",
                                }
                            )
                            doc_txt = extract_document_text(fbytes, clean_label)
                            if doc_txt:
                                extracted_chunks.append(f"### 포털 첨부파일: {clean_label}\n\n{doc_txt}")
                except Exception:
                    pass

            # Extract general quotas from #tableTbody or #recruitTargetListDiv
            portal_quotas = {}
            tbody_m = re.search(r'<tbody id=["\']tableTbody["\']>(.*?)</tbody>', html, flags=re.S)
            if tbody_m:
                for tr in re.findall(r"<tr[^>]*>(.*?)</tr>", tbody_m.group(1), flags=re.S):
                    cells = [
                        re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", c)).strip()
                        for c in re.findall(r"<td[^>]*>(.*?)</td>", tr, flags=re.S)
                    ]
                    row_join = " ".join(cells)
                    if "교육과정" in row_join:
                        for ak, apat in (("a3", "3세"), ("a4", "4세"), ("a5", "5세")):
                            if apat in row_join:
                                nums = [int(x) for x in re.findall(r"\b(\d+)\b", " ".join(cells[2:])) if int(x) < 300]
                                if nums:
                                    portal_quotas[ak] = nums[-1]

            table_lines = [f"### 유보통합포털 일반모집 요강 ({pdf_title or target['name']})"]
            for tbl in re.findall(r"<table[^>]*>.*?</table>", html, flags=re.S)[:4]:
                for tr in re.findall(r"<tr[^>]*>.*?</tr>", tbl, flags=re.S):
                    cells = [
                        re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", c)).strip()
                        for c in re.findall(r"<t[hd][^>]*>(.*?)</t[hd]>", tr, flags=re.S)
                    ]
                    if any(cells):
                        table_lines.append("• " + " | ".join(c for c in cells if c))

            combined_txt = "\n".join(table_lines)
            if extracted_chunks:
                combined_txt = combined_txt + "\n\n---\n\n" + "\n\n---\n\n".join(extracted_chunks)
            if not portal_quotas:
                portal_quotas = parse_general_quotas_from_text(combined_txt)

            return {
                "id": target["id"],
                "slug": target["slug"],
                "name": target["name"],
                "generalFound": True,
                "boardStatus": "일반모집 공고 확인됨 (유보통합포털)",
                "postTitle": pdf_title or f"2027학년도 {target['name']} 일반모집 모집요강",
                "postDate": datetime.datetime.now(KST).strftime("%Y.%m.%d"),
                "webUrl": portal_url,
                "files": saved_rel_files,
                "parsedQuotas": portal_quotas,
                "extractedPreview": combined_txt[:4000],
            }

        return {
            "id": target["id"],
            "slug": target["slug"],
            "name": target["name"],
            "generalFound": bool(prev_entry and prev_entry.get("generalFound")),
            "boardStatus": (
                "정상 확인 (유보통합포털 일반모집 11/16 18:00 공개 예정)"
                if not (prev_entry and prev_entry.get("generalFound"))
                else prev_entry.get("boardStatus", "일반모집 공고 확인됨")
            ),
            **(
                {
                    k: prev_entry[k]
                    for k in ("postTitle", "postDate", "webUrl", "files", "parsedQuotas", "extractedPreview")
                    if prev_entry and k in prev_entry
                }
                if (prev_entry and prev_entry.get("generalFound"))
                else {}
            ),
        }
    except Exception as exc:
        print(f"  [WARN] {target['name']} Portal: {exc}")
        kept = dict(prev_entry) if prev_entry else {
            "id": target["id"],
            "slug": target["slug"],
            "name": target["name"],
            "generalFound": False,
        }
        kept["boardStatus"] = "포털 응답지연 (기존 데이터 유지)"
        return kept


def write_general_markdown(entry: dict, sync_local_dir: str = None):
    if not entry.get("generalFound") or not entry.get("extractedPreview"):
        return
    slug = entry["slug"]
    gen_md_path = os.path.join(DOCS_DIR, f"general_{slug}.md")
    md_content = (
        f"# {entry['name']} — 2027학년도 일반모집 공고 원문 (자동 수집)\n\n"
        f"- **공고 제목**: {entry.get('postTitle', '')}\n"
        f"- **게시일**: {entry.get('postDate', '')}\n"
        f"- **원문 링크**: [공식 공고 바로가기]({entry.get('webUrl', '')})\n\n"
        f"---\n\n"
        f"{entry.get('extractedPreview', '')}\n"
    )
    with open(gen_md_path, "w", encoding="utf-8") as f:
        f.write(md_content)

    # Also append/replace section in main <slug>.md
    main_md_path = os.path.join(DOCS_DIR, f"{slug}.md")
    if os.path.exists(main_md_path):
        orig = open(main_md_path, encoding="utf-8").read()
        marker = "\n<!-- GENERAL_RECRUITMENT_AUTO_SECTION -->\n"
        block = (
            f"{marker}## [최신 업데이트] 2027학년도 일반모집 공고 원문\n\n"
            f"- **공고 제목**: {entry.get('postTitle', '')}\n"
            f"- **공식 링크**: {entry.get('webUrl', '')}\n\n"
            f"{entry.get('extractedPreview', '')}\n"
        )
        if marker in orig:
            orig = orig.split(marker)[0] + block
        else:
            orig = orig.rstrip() + "\n\n" + block
        with open(main_md_path, "w", encoding="utf-8") as f:
            f.write(orig)

    if sync_local_dir and os.path.isdir(sync_local_dir):
        shutil.copy2(gen_md_path, os.path.join(sync_local_dir, f"general_{slug}.md"))
        if os.path.exists(main_md_path):
            shutil.copy2(main_md_path, os.path.join(sync_local_dir, f"{slug}.md"))
        local_gen_files = os.path.join(sync_local_dir, "general_files")
        if os.path.isdir(GEN_FILES_DIR):
            os.makedirs(local_gen_files, exist_ok=True)
            for f_item in os.listdir(GEN_FILES_DIR):
                s_file = os.path.join(GEN_FILES_DIR, f_item)
                if os.path.isfile(s_file):
                    shutil.copy2(s_file, os.path.join(local_gen_files, f_item))


def inject_status_into_viewer(status_payload: dict, sync_artifact_viewer: str = None):
    if not os.path.exists(VIEWER_HTML_PATH):
        return
    html = open(VIEWER_HTML_PATH, encoding="utf-8").read()
    start_tag = "/* GENERAL_STATUS_EMBED_START */"
    end_tag = "/* GENERAL_STATUS_EMBED_END */"
    if start_tag not in html or end_tag not in html:
        return
    embedded_js = (
        f"{start_tag}\n"
        f"    window.EMBEDDED_GENERAL_STATUS = {json.dumps(status_payload, ensure_ascii=False, indent=2)};\n"
        f"    {end_tag}"
    )
    pattern = re.compile(re.escape(start_tag) + r".*?" + re.escape(end_tag), flags=re.S)
    updated_html = pattern.sub(lambda _: embedded_js, html)
    with open(VIEWER_HTML_PATH, "w", encoding="utf-8", newline="\n") as f:
        f.write(updated_html)
    if sync_artifact_viewer and os.path.exists(os.path.dirname(sync_artifact_viewer)):
        shutil.copy2(VIEWER_HTML_PATH, sync_artifact_viewer)


def main():
    parser = argparse.ArgumentParser(description="Update Yeongjong 2027 General Recruitment Notices")
    parser.add_argument("--write-status", action="store_true", help="Write general_status.json and embed into viewer")
    parser.add_argument("--sync-local", type=str, default="", help="Optional local docs directory to sync")
    parser.add_argument("--sync-artifact", type=str, default="", help="Optional artifact HTML viewer path to sync")
    args = parser.parse_args()

    os.makedirs(DOCS_DIR, exist_ok=True)
    prev_data = {}
    if os.path.exists(STATUS_JSON_PATH):
        try:
            loaded = json.load(open(STATUS_JSON_PATH, encoding="utf-8"))
            for item in loaded.get("kindergartens", []):
                prev_data[item["id"]] = item
        except Exception:
            pass

    now_kst = datetime.datetime.now(KST).strftime("%Y-%m-%d %H:%M KST")
    print(f"=== Yeongjong 2027 General Recruitment Scan ({now_kst}) ===")

    entries = []
    found_count = 0
    for target in TARGETS:
        prev_entry = prev_data.get(target["id"])
        if target["kind"] == "ice":
            entry = scan_ice_target(target, prev_entry)
        elif target["kind"] == "anyboard":
            entry = scan_anyboard_target(target, prev_entry)
        elif target["kind"] == "firstschool_portal":
            entry = scan_firstschool_portal_target(target, prev_entry)
        else:
            entry = scan_naver_blog_target(target, prev_entry)

        # Fallback: if homepage/blog has not posted confirmed general quotas yet, also check 유보통합포털 RC928_002
        if (
            target["kind"] != "firstschool_portal"
            and target.get("presch_cd")
            and (not entry.get("generalFound") or not entry.get("parsedQuotas"))
        ):
            portal_entry = scan_firstschool_portal_target(target, prev_entry)
            if portal_entry.get("generalFound"):
                entry = portal_entry

        if entry.get("generalFound"):
            found_count += 1
            write_general_markdown(entry, args.sync_local or None)
        print(f"  [{entry['id']}] {entry['name']}: {entry['boardStatus']}")
        entries.append(entry)

    payload = {
        "lastCheckedKst": now_kst,
        "generalNoticeWindow": "2026. 11. 12.(목) ~ 11. 16.(월) 18:00 순차 공고 예정",
        "generalApplyWindow": "2026. 11. 17.(화) 09:00 ~ 11. 23.(월) 18:00",
        "foundCount": found_count,
        "totalMonitored": len(TARGETS),
        "kindergartens": entries,
    }

    if args.write_status:
        with open(STATUS_JSON_PATH, "w", encoding="utf-8", newline="\n") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
        if args.sync_local and os.path.isdir(args.sync_local):
            shutil.copy2(STATUS_JSON_PATH, os.path.join(args.sync_local, "general_status.json"))
        inject_status_into_viewer(payload, args.sync_artifact or None)
        print(f"\nSaved status -> {STATUS_JSON_PATH} (found {found_count}/{len(TARGETS)})")


if __name__ == "__main__":
    main()
