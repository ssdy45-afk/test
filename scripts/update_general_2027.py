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

# Known priority-only boardSeqs from Oct 2026 so they are never mistaken for general notices
KNOWN_PRIORITY_SEQS = {
    "34292526", "34292535", "34287374", "34291408", "34269012",
    "34283433", "34286125", "34292301", "34291294", "34290766",
    "34277946", "34288595", "34288607", "224455411476",
}

EXCLUDED_TITLE_WORDS = ("특수", "영재", "독서", "아이리더", "강사", "채용", "기간제", "급식", "입찰")

TARGETS = [
    {
        "id": "01",
        "slug": "01_yeongjong",
        "name": "영종유치원",
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
        "kind": "ice",
        "base": "https://starlight.icees.kr",
        "boards": [
            "/boardCnts/list.do?boardID=301738&m=060501&s=starlight",
            "/boardCnts/list.do?boardID=301730&m=0601&s=starlight",
        ],
    },
    {
        "id": "05",
        "slug": "05_yeongjong_es",
        "name": "인천영종초병설유치원",
        "kind": "ice",
        "base": "https://yj.icees.kr",
        "boards": [
            "/boardCnts/list.do?boardID=275247&m=0901&s=iyoungjonges",
        ],
    },
    {
        "id": "06",
        "slug": "06_yongyu",
        "name": "인천용유초병설유치원",
        "kind": "ice",
        "base": "https://yongyu.icees.kr",
        "boards": [
            "/boardCnts/list.do?boardID=257997&m=0404&s=yongyu",
            "/boardCnts/list.do?boardID=257995&m=0403&s=yongyu",
        ],
    },
    {
        "id": "07",
        "slug": "07_yunseul",
        "name": "인천윤슬초병설유치원",
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
        "kind": "naver_blog",
        "blog_id": "dynafable",
        "rss_url": "https://rss.blog.naver.com/dynafable.xml",
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
    Attempt to extract confirmed general recruitment numbers per age (3세, 4세, 5세)
    from extracted HWP/HWPX/PDF text lines. Returns dict like {"a3": 15, "a4": 8, "a5": 4}
    when unambiguously matched, otherwise empty dict (full raw text is always preserved).
    """
    if not text:
        return {}
    lines = [re.sub(r"\s+", " ", ln).strip() for ln in text.splitlines() if ln.strip()]
    joined = "\n".join(lines[:120])
    result = {}
    for age_key, age_pat in (("a3", r"3세"), ("a4", r"4세"), ("a5", r"5세")):
        # Match patterns like: 3세반 | 4 | 2023.1.1.~2023.12.31. | 48명 | 35명 | 13명
        m = re.search(
            rf"{age_pat}(?:반)?[^\n]{{0,80}}?202[123]\.\s*1\.\s*1\.[^\n]{{0,120}}",
            joined,
            flags=re.S,
        )
        if m:
            chunk = joined[m.start() : m.end() + 100]
            nums = [int(x) for x in re.findall(r"(\d+)\s*명", chunk)]
            # Standard table columns: [정원, 재학(우선등록포함), 일반모집인원] -> last number in row
            if len(nums) >= 3:
                result[age_key] = nums[2]
            elif len(nums) == 1:
                result[age_key] = nums[0]
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

                vclean = re.sub(r"<script.*?</script>|<style.*?</style>", "", vhtml, flags=re.S)
                body_txt = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", vclean)).strip()
                pos = body_txt.find("작성자")
                post_snippet = body_txt[max(0, pos - 80) : pos + 2000] if pos != -1 else body_txt[:1500]
                full_extracted = "\n\n---\n\n".join(extracted_chunks) if extracted_chunks else post_snippet
                parsed_quotas = parse_general_quotas_from_text(full_extracted)

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
                return {
                    "id": target["id"],
                    "slug": target["slug"],
                    "name": target["name"],
                    "generalFound": True,
                    "boardStatus": "일반모집 공고 확인됨",
                    "postTitle": title,
                    "postDate": pub_date[:16],
                    "webUrl": link,
                    "files": [],
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
        else:
            entry = scan_naver_blog_target(target, prev_entry)
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
        inject_status_into_viewer(payload, args.sync_artifact or None)
        print(f"\nSaved status -> {STATUS_JSON_PATH} (found {found_count}/{len(TARGETS)})")


if __name__ == "__main__":
    main()
