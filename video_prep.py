#!/usr/bin/env python3
"""
Google Photos Pre-2013 Photo & Video Prep Tool (Windows / macOS / Linux)

4-step browser wizard & CLI using Python standard library + FFmpeg/FFprobe:
  1. scan    : Recursively scan videos & images, probe codecs/EXIF/dates, extract thumbnails.
  2. review  : Launch local browser UI to preview, rotate, delete to trash, or merge clips.
  3. convert : Batch convert/rotate videos & images preserving dates/EXIF and move deleted files to trash.
"""

from __future__ import annotations

import argparse
from concurrent.futures import ThreadPoolExecutor, as_completed
import copy
from datetime import datetime, timezone
import hashlib
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import json
import os
from pathlib import Path
import platform
import re
import shutil
import struct
import subprocess
import sys
import threading
import time
from typing import Any, Callable
from urllib.parse import urlparse
import uuid
import webbrowser

# Ensure UTF-8 output on Windows consoles
for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, "reconfigure"):
        try:
            _stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

MODERN_CONTAINERS = {".mp4", ".m4v", ".mov"}
LEGACY_CONTAINERS = {
    ".avi", ".3gp", ".3g2", ".wmv", ".asf", ".mpg", ".mpeg", ".mpe",
    ".vob", ".flv", ".f4v", ".rm", ".rmvb", ".divx", ".xvid",
    ".mts", ".m2ts", ".ts", ".mkv", ".webm", ".ogv", ".mod", ".tod",
}
VIDEO_EXTENSIONS = MODERN_CONTAINERS | LEGACY_CONTAINERS

MODERN_IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".webp", ".gif", ".heic", ".heif"}
LEGACY_IMAGE_EXTENSIONS = {".bmp", ".tif", ".tiff", ".jfif"}
IMAGE_EXTENSIONS = MODERN_IMAGE_EXTENSIONS | LEGACY_IMAGE_EXTENSIONS
BROWSER_NATIVE_IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".webp", ".gif"}
IMAGE_MIME_TYPES = {
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".jfif": "image/jpeg",
    ".png": "image/png",
    ".webp": "image/webp",
    ".gif": "image/gif",
    ".bmp": "image/bmp",
}

SKIP_DIR_NAMES = {
    "_gphotos_work",
    "_converted",
    "_original_backup",
    "deleted_trash",
    "$recycle.bin",
    "system volume information",
}

THUMB_OFFSETS = [0.2, 0.5, 0.75, 0.05]
VALID_ROTATIONS = {0, 90, 180, 270}
THUMB_FILENAME_RE = re.compile(r"^[A-Za-z0-9_-]+\.jpg$")

DATE_TAG_KEYS = (
    "creation_time",
    "com.apple.quicktime.creationdate",
    "date",
)

FILENAME_DATETIME_PATTERNS = (
    # 20110514_153022 or 20110514-153022 or 20110514153022
    re.compile(
        r"(?<!\d)(?P<y>(?:199\d|20[0-3]\d))[-_.]?(?P<m>0[1-9]|1[0-2])[-_.]?(?P<d>0[1-9]|[12]\d|3[01])"
        r"[ _T-]?(?P<H>[01]\d|2[0-3])[-_.:]?(?P<M>[0-5]\d)[-_.:]?(?P<S>[0-5]\d)(?!\d)"
    ),
)

FILENAME_DATE_ONLY_PATTERN = re.compile(
    r"(?<!\d)(?P<y>(?:199\d|20[0-3]\d))[-_.]?(?P<m>0[1-9]|1[0-2])[-_.]?(?P<d>0[1-9]|[12]\d|3[01])(?!\d)"
)


def local_tz():
    """Return the system's current local timezone."""
    return datetime.now().astimezone().tzinfo


class JobCancelled(Exception):
    """Raised when a background scan or convert job is cancelled by the user."""


def locate_ffmpeg_tools() -> tuple[str | None, str | None]:
    """Return (ffmpeg_path, ffprobe_path) if present on PATH."""
    return shutil.which("ffmpeg"), shutil.which("ffprobe")


def check_ffmpeg_tools() -> tuple[str, str]:
    """Verify ffmpeg and ffprobe are available on PATH."""
    ffmpeg_bin, ffprobe_bin = locate_ffmpeg_tools()
    if not ffmpeg_bin or not ffprobe_bin:
        missing = []
        if not ffmpeg_bin:
            missing.append("ffmpeg")
        if not ffprobe_bin:
            missing.append("ffprobe")
        print(
            f"[오류] PATH에서 {', '.join(missing)} 를 찾을 수 없습니다.\n"
            "Windows PowerShell에서 아래 명령어로 설치한 뒤 터미널을 새로 열어 다시 시도하세요:\n"
            "  winget install -e --id Gyan.FFmpeg\n",
            file=sys.stderr,
        )
        sys.exit(1)
    return ffmpeg_bin, ffprobe_bin


def get_env_status() -> dict[str, Any]:
    """Check installation status of Python, FFmpeg/FFprobe, and Git."""
    ffmpeg_bin, ffprobe_bin = locate_ffmpeg_tools()
    ffmpeg_ver = ""
    if ffmpeg_bin:
        try:
            res = subprocess.run(
                [ffmpeg_bin, "-version"],
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=5,
            )
            if res.returncode == 0 and res.stdout:
                ffmpeg_ver = res.stdout.splitlines()[0].strip()
        except Exception:
            pass

    git_bin = shutil.which("git")
    git_ver = ""
    if git_bin:
        try:
            res = subprocess.run(
                [git_bin, "--version"],
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=5,
            )
            if res.returncode == 0 and res.stdout:
                git_ver = res.stdout.splitlines()[0].strip()
        except Exception:
            pass

    ffmpeg_ok = bool(ffmpeg_bin and ffprobe_bin)
    return {
        "python": {
            "installed": True,
            "version": f"Python {platform.python_version()}",
            "required": True,
        },
        "ffmpeg": {
            "installed": ffmpeg_ok,
            "version": ffmpeg_ver or ("설치됨 (ffprobe 포함)" if ffmpeg_ok else ""),
            "required": True,
        },
        "git": {
            "installed": bool(git_bin),
            "version": git_ver or ("설치됨" if git_bin else ""),
            "required": False,
        },
        "ready": ffmpeg_ok,
    }


def run_cmd(cmd: list[str], timeout: int = 120) -> subprocess.CompletedProcess[str]:
    """Run a subprocess safely without shell=True and with UTF-8 decoding."""
    return subprocess.run(
        cmd,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=timeout,
    )


def make_item_id(rel_path_posix: str) -> str:
    """Create a stable short ID from the relative POSIX path."""
    return hashlib.sha1(rel_path_posix.encode("utf-8")).hexdigest()[:12]


def is_valid_capture_year(year: int) -> bool:
    """Reject sentinel/invalid capture years like 1904 or far-future dates."""
    max_year = datetime.now().year + 1
    return 1995 <= year <= max_year


def parse_tag_datetime(raw: str) -> datetime | None:
    """Parse common ffprobe metadata timestamp strings into timezone-aware UTC datetime."""
    text = (raw or "").strip()
    if not text:
        return None

    # Normalize common variants:
    # - Trailing 'Z' -> '+00:00'
    # - 'YYYY:MM:DD HH:MM:SS' -> 'YYYY-MM-DD HH:MM:SS'
    # - '+0900' -> '+09:00'
    if re.match(r"^\d{4}:\d{2}:\d{2}", text):
        text = text[:10].replace(":", "-") + text[10:]
    if text.endswith(("Z", "z")):
        text = text[:-1] + "+00:00"
    text = re.sub(r"([+-]\d{2})(\d{2})$", r"\1:\2", text)

    dt: datetime | None = None
    try:
        dt = datetime.fromisoformat(text)
    except ValueError:
        for fmt in (
            "%Y-%m-%d %H:%M:%S",
            "%Y-%m-%dT%H:%M:%S",
            "%Y-%m-%d %H:%M",
            "%Y-%m-%d",
        ):
            try:
                dt = datetime.strptime(text, fmt)
                break
            except ValueError:
                continue

    if dt is None or not is_valid_capture_year(dt.year):
        return None

    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=local_tz())
    return dt.astimezone(timezone.utc)


def extract_filename_datetime(filename: str, mtime_epoch: float) -> datetime | None:
    """Try to infer capture datetime from filename patterns."""
    tz = local_tz()
    for pat in FILENAME_DATETIME_PATTERNS:
        m = pat.search(filename)
        if m:
            try:
                dt = datetime(
                    int(m.group("y")),
                    int(m.group("m")),
                    int(m.group("d")),
                    int(m.group("H")),
                    int(m.group("M")),
                    int(m.group("S")),
                    tzinfo=tz,
                )
                if is_valid_capture_year(dt.year):
                    return dt.astimezone(timezone.utc)
            except ValueError:
                pass

    m = FILENAME_DATE_ONLY_PATTERN.search(filename)
    if m:
        try:
            y, mo, d = int(m.group("y")), int(m.group("m")), int(m.group("d"))
            if is_valid_capture_year(y):
                mtime_local = datetime.fromtimestamp(mtime_epoch, tz=tz)
                if (mtime_local.year, mtime_local.month, mtime_local.day) == (y, mo, d):
                    dt = mtime_local
                else:
                    dt = datetime(y, mo, d, 12, 0, 0, tzinfo=tz)
                return dt.astimezone(timezone.utc)
        except ValueError:
            pass

    return None


def detect_best_date(
    probe_data: dict[str, Any],
    filename: str,
    mtime_epoch: float,
) -> tuple[str, float, str]:
    """
    Determine the best shooting timestamp in priority order:
      1. Container/stream metadata tags (valid year >= 1995)
      2. Filename date/time pattern
      3. File modification time (st_mtime)
    Returns (utc_iso_string, epoch_seconds, date_source).
    """
    tag_sources: list[dict[str, str]] = []
    fmt_tags = probe_data.get("format", {}).get("tags") or {}
    if isinstance(fmt_tags, dict):
        tag_sources.append({str(k).lower(): str(v) for k, v in fmt_tags.items()})

    for stream in probe_data.get("streams") or []:
        s_tags = stream.get("tags") or {}
        if isinstance(s_tags, dict):
            tag_sources.append({str(k).lower(): str(v) for k, v in s_tags.items()})

    for tags in tag_sources:
        for key in DATE_TAG_KEYS:
            if key in tags:
                dt = parse_tag_datetime(tags[key])
                if dt is not None:
                    return (
                        dt.strftime("%Y-%m-%dT%H:%M:%SZ"),
                        dt.timestamp(),
                        "metadata",
                    )

    fn_dt = extract_filename_datetime(filename, mtime_epoch)
    if fn_dt is not None:
        return (
            fn_dt.strftime("%Y-%m-%dT%H:%M:%SZ"),
            fn_dt.timestamp(),
            "filename",
        )

    mtime_dt = datetime.fromtimestamp(mtime_epoch, tz=local_tz()).astimezone(timezone.utc)
    return (
        mtime_dt.strftime("%Y-%m-%dT%H:%M:%SZ"),
        mtime_dt.timestamp(),
        "mtime",
    )


def classify_video(
    ext: str,
    video_codec: str,
    has_audio: bool,
    audio_codec: str,
) -> tuple[bool, str]:
    """
    Option B classification:
      - Legacy container -> convert
      - Modern container with legacy video or audio codec -> convert
      - Modern container with H.264/HEVC + AAC/MP3 (or no audio) -> keep unless rotated
    """
    ext_lower = ext.lower()
    v_codec = (video_codec or "").lower()
    a_codec = (audio_codec or "").lower()

    if ext_lower not in MODERN_CONTAINERS:
        return True, f"구형 포맷 ({ext_lower})"

    bad_video = v_codec not in {"h264", "hevc"}
    bad_audio = has_audio and (a_codec not in {"aac", "mp3"})
    if bad_video or bad_audio:
        parts = []
        if bad_video:
            parts.append(v_codec or "unknown")
        if bad_audio:
            parts.append(a_codec or "unknown")
        return True, f"구형 코덱 ({'/'.join(parts)})"

    return False, "호환 포맷 (유지)"


def parse_rotation_from_stream(v_stream: dict[str, Any]) -> int:
    """Read existing display rotation tag if present (informational)."""
    for sd in v_stream.get("side_data_list") or []:
        if "rotation" in sd:
            try:
                return int(round(float(sd["rotation"])))
            except (TypeError, ValueError):
                pass
    tags = {str(k).lower(): str(v) for k, v in (v_stream.get("tags") or {}).items()}
    if "rotate" in tags:
        try:
            return int(round(float(tags["rotate"])))
        except (TypeError, ValueError):
            pass
    return 0


def probe_video(file_path: Path, ffprobe_bin: str = "ffprobe") -> dict[str, Any]:
    """Run ffprobe on a video file and return parsed JSON."""
    cmd = [
        ffprobe_bin,
        "-v", "error",
        "-show_format",
        "-show_streams",
        "-print_format", "json",
        str(file_path),
    ]
    res = run_cmd(cmd, timeout=45)
    if res.returncode != 0:
        raise RuntimeError(res.stderr.strip() or f"ffprobe exited with {res.returncode}")
    return json.loads(res.stdout or "{}")


def extract_thumbnail(
    file_path: Path,
    thumb_path: Path,
    duration: float,
    offset_ratio: float = 0.2,
    ffmpeg_bin: str = "ffmpeg",
) -> bool:
    """
    Extract a representative JPEG thumbnail with FFmpeg's default auto-rotation enabled
    so the browser preview matches how Google Photos / modern players orient the video.
    """
    thumb_path.parent.mkdir(parents=True, exist_ok=True)
    if duration > 1.0:
        if offset_ratio == 0.2:
            seek_sec = min(duration * 0.2, 3.0)
        else:
            seek_sec = max(0.0, min(duration * offset_ratio, max(0.0, duration - 0.3)))
    else:
        seek_sec = 0.0

    attempts = [seek_sec]
    if seek_sec > 0.0:
        attempts.append(0.0)

    vf = "scale='min(320,iw)':'min(320,ih)':force_original_aspect_ratio=decrease"
    for seek in attempts:
        cmd = [
            ffmpeg_bin,
            "-y",
            "-ss", f"{seek:.3f}",
            "-i", str(file_path),
            "-frames:v", "1",
            "-vf", vf,
            "-q:v", "5",
            str(thumb_path),
        ]
        try:
            res = run_cmd(cmd, timeout=45)
        except Exception:
            res = None
        if res and res.returncode == 0 and thumb_path.exists() and thumb_path.stat().st_size > 0:
            return True
    return False


JPEG_SOF_MARKERS = {
    0xC0, 0xC1, 0xC2, 0xC3,
    0xC5, 0xC6, 0xC7,
    0xC9, 0xCA, 0xCB,
    0xCD, 0xCE, 0xCF,
}
EXIF_ORIENTATION_TO_ROT = {6: 90, 3: 180, 8: 270}


def _read_tiff_ascii(tiff: bytes, endian: str, typ: int, cnt: int, val_field_off: int) -> str:
    """Read an ASCII string field from TIFF/EXIF bytes."""
    if typ != 2 or cnt <= 0 or val_field_off + 4 > len(tiff):
        return ""
    if cnt <= 4:
        raw = tiff[val_field_off : val_field_off + cnt]
    else:
        data_off = struct.unpack_from(endian + "I", tiff, val_field_off)[0]
        if data_off + cnt > len(tiff):
            return ""
        raw = tiff[data_off : data_off + cnt]
    return raw.split(b"\x00", 1)[0].decode("ascii", errors="ignore").strip()


def _parse_exif_tiff_block(tiff: bytes) -> tuple[int, str]:
    """Parse Orientation rotation and best capture date string from TIFF bytes."""
    if len(tiff) < 8:
        return 0, ""
    byte_order = tiff[:2]
    if byte_order == b"II":
        endian = "<"
    elif byte_order == b"MM":
        endian = ">"
    else:
        return 0, ""

    magic, ifd0_off = struct.unpack_from(endian + "HI", tiff, 2)
    if magic != 42 or ifd0_off + 2 > len(tiff):
        return 0, ""

    meta_rotation = 0
    dt_ifd0 = ""
    dt_orig = ""
    dt_digi = ""
    exif_ifd_off = 0

    count0 = struct.unpack_from(endian + "H", tiff, ifd0_off)[0]
    for i in range(count0):
        entry_off = ifd0_off + 2 + i * 12
        if entry_off + 12 > len(tiff):
            break
        tag, typ, cnt = struct.unpack_from(endian + "HHI", tiff, entry_off)
        val_off = entry_off + 8
        if tag == 0x0112 and typ == 3 and cnt >= 1:
            orient = struct.unpack_from(endian + "H", tiff, val_off)[0]
            meta_rotation = EXIF_ORIENTATION_TO_ROT.get(orient, 0)
        elif tag == 0x0132:
            dt_ifd0 = _read_tiff_ascii(tiff, endian, typ, cnt, val_off)
        elif tag == 0x8769 and cnt >= 1:
            exif_ifd_off = struct.unpack_from(endian + "I", tiff, val_off)[0]

    if exif_ifd_off and exif_ifd_off + 2 <= len(tiff):
        count_sub = struct.unpack_from(endian + "H", tiff, exif_ifd_off)[0]
        for i in range(count_sub):
            entry_off = exif_ifd_off + 2 + i * 12
            if entry_off + 12 > len(tiff):
                break
            tag, typ, cnt = struct.unpack_from(endian + "HHI", tiff, entry_off)
            val_off = entry_off + 8
            if tag == 0x9003:
                dt_orig = _read_tiff_ascii(tiff, endian, typ, cnt, val_off)
            elif tag == 0x9004:
                dt_digi = _read_tiff_ascii(tiff, endian, typ, cnt, val_off)

    best_date = ""
    for candidate in (dt_orig, dt_digi, dt_ifd0):
        if candidate and not candidate.startswith("0000:00:00"):
            best_date = candidate
            break
    return meta_rotation, best_date


def parse_jpeg_info(file_path: Path) -> dict[str, Any]:
    """Read dimensions, EXIF orientation rotation, and capture date from a JPEG file."""
    info: dict[str, Any] = {"width": 0, "height": 0, "meta_rotation": 0, "date_str": ""}
    try:
        with file_path.open("rb") as f:
            if f.read(2) != b"\xff\xd8":
                return info
            while True:
                b = f.read(1)
                if not b:
                    break
                while b == b"\xff":
                    nxt = f.read(1)
                    if not nxt:
                        b = b""
                        break
                    if nxt != b"\xff":
                        b = nxt
                        break
                if not b:
                    break
                m = b[0]
                if m in (0xD9, 0xDA):
                    break
                if m == 0x00 or m == 0x01 or (0xD0 <= m <= 0xD7):
                    continue
                seg_len_raw = f.read(2)
                if len(seg_len_raw) < 2:
                    break
                seg_len = struct.unpack(">H", seg_len_raw)[0]
                if seg_len < 2:
                    break
                payload_len = seg_len - 2
                if m == 0xE1 or m in JPEG_SOF_MARKERS:
                    payload = f.read(payload_len)
                    if len(payload) < payload_len:
                        break
                    if m == 0xE1 and payload.startswith(b"Exif\x00\x00") and not info["date_str"] and info["meta_rotation"] == 0:
                        rot, dt_str = _parse_exif_tiff_block(payload[6:])
                        info["meta_rotation"] = rot
                        if dt_str:
                            info["date_str"] = dt_str
                    elif m in JPEG_SOF_MARKERS and len(payload) >= 5 and info["width"] == 0:
                        h, w = struct.unpack_from(">HH", payload, 1)
                        info["width"] = int(w)
                        info["height"] = int(h)
                else:
                    f.seek(payload_len, os.SEEK_CUR)
    except Exception:
        pass
    return info


def read_image_basic_info(file_path: Path) -> dict[str, Any]:
    """Read image width, height, EXIF rotation, and EXIF date without external libs."""
    ext = file_path.suffix.lower()
    if ext in {".jpg", ".jpeg", ".jfif"}:
        return parse_jpeg_info(file_path)

    info: dict[str, Any] = {"width": 0, "height": 0, "meta_rotation": 0, "date_str": ""}
    try:
        with file_path.open("rb") as f:
            head = f.read(64)
        if ext == ".png" and len(head) >= 24 and head[:8] == b"\x89PNG\r\n\x1a\n" and head[12:16] == b"IHDR":
            w, h = struct.unpack_from(">II", head, 16)
            info["width"], info["height"] = int(w), int(h)
        elif ext == ".gif" and len(head) >= 10 and head[:6] in (b"GIF87a", b"GIF89a"):
            w, h = struct.unpack_from("<HH", head, 6)
            info["width"], info["height"] = int(w), int(h)
        elif ext == ".bmp" and len(head) >= 26 and head[:2] == b"BM":
            w, h = struct.unpack_from("<ii", head, 18)
            info["width"], info["height"] = abs(int(w)), abs(int(h))
        elif ext == ".webp" and len(head) >= 30 and head[:4] == b"RIFF" and head[8:12] == b"WEBP":
            chunk = head[12:16]
            if chunk == b"VP8X":
                w = 1 + int.from_bytes(head[24:27], "little")
                h = 1 + int.from_bytes(head[27:30], "little")
                info["width"], info["height"] = w, h
            elif chunk == b"VP8 " and head[23:26] == b"\x9d\x01\x2a":
                w, h = struct.unpack_from("<HH", head, 26)
                info["width"], info["height"] = w & 0x3FFF, h & 0x3FFF
            elif chunk == b"VP8L" and head[20] == 0x2F:
                b0, b1, b2, b3 = head[21:25]
                w = 1 + (((b1 & 0x3F) << 8) | b0)
                h = 1 + (((b3 & 0x0F) << 10) | (b2 << 2) | ((b1 & 0xC0) >> 6))
                info["width"], info["height"] = w, h
    except Exception:
        pass
    return info


def extract_and_patch_jpeg_exif(src_path: Path) -> bytes | None:
    """
    Extract the APP1 Exif payload from a JPEG and reset Orientation (0x0112) to 1
    so physically rotated output pixels are never double-rotated by viewers.
    """
    try:
        with src_path.open("rb") as f:
            if f.read(2) != b"\xff\xd8":
                return None
            raw_payload: bytes | None = None
            while True:
                b = f.read(1)
                if not b:
                    break
                while b == b"\xff":
                    nxt = f.read(1)
                    if not nxt:
                        b = b""
                        break
                    if nxt != b"\xff":
                        b = nxt
                        break
                if not b:
                    break
                m = b[0]
                if m in (0xD9, 0xDA):
                    break
                if m == 0x00 or m == 0x01 or (0xD0 <= m <= 0xD7):
                    continue
                seg_len_raw = f.read(2)
                if len(seg_len_raw) < 2:
                    break
                seg_len = struct.unpack(">H", seg_len_raw)[0]
                if seg_len < 2:
                    break
                payload_len = seg_len - 2
                if m == 0xE1:
                    payload = f.read(payload_len)
                    if len(payload) == payload_len and payload.startswith(b"Exif\x00\x00"):
                        raw_payload = payload
                        break
                else:
                    f.seek(payload_len, os.SEEK_CUR)
    except Exception:
        return None

    if raw_payload is None or len(raw_payload) < 14:
        return None

    buf = bytearray(raw_payload)
    tiff_base = 6
    byte_order = bytes(buf[tiff_base : tiff_base + 2])
    if byte_order == b"II":
        endian = "<"
    elif byte_order == b"MM":
        endian = ">"
    else:
        return None

    magic, ifd0_off = struct.unpack_from(endian + "HI", buf, tiff_base + 2)
    if magic != 42 or tiff_base + ifd0_off + 2 > len(buf):
        return None

    def patch_ifd(ifd_rel_off: int) -> int:
        if ifd_rel_off <= 0 or tiff_base + ifd_rel_off + 2 > len(buf):
            return 0
        cnt_entries = struct.unpack_from(endian + "H", buf, tiff_base + ifd_rel_off)[0]
        for idx in range(cnt_entries):
            e_off = tiff_base + ifd_rel_off + 2 + idx * 12
            if e_off + 12 > len(buf):
                return 0
            tag, typ, cnt = struct.unpack_from(endian + "HHI", buf, e_off)
            if tag == 0x0112 and typ == 3 and cnt == 1:
                struct.pack_into(endian + "H", buf, e_off + 8, 1)
        next_ptr_off = tiff_base + ifd_rel_off + 2 + cnt_entries * 12
        if next_ptr_off + 4 <= len(buf):
            return struct.unpack_from(endian + "I", buf, next_ptr_off)[0]
        return 0

    ifd1_off = patch_ifd(ifd0_off)
    if ifd1_off:
        patch_ifd(ifd1_off)
    return bytes(buf)


def build_minimal_exif_payload(best_date_utc: str) -> bytes:
    """Build a minimal little-endian TIFF/EXIF APP1 payload with Orientation=1 and capture date."""
    dt = parse_tag_datetime(best_date_utc)
    if dt is None:
        dt = datetime.now(timezone.utc)
    dt_local = dt.astimezone(local_tz())
    dt_bytes = dt_local.strftime("%Y:%m:%d %H:%M:%S").encode("ascii") + b"\x00"  # 20 bytes

    tiff_header = b"II" + struct.pack("<HI", 42, 8)
    ifd0 = bytearray()
    ifd0 += struct.pack("<H", 3)
    ifd0 += struct.pack("<HHIHH", 0x0112, 3, 1, 1, 0)       # Orientation = 1
    ifd0 += struct.pack("<HHII", 0x0132, 2, 20, 50)         # DateTime
    ifd0 += struct.pack("<HHII", 0x8769, 4, 1, 70)          # ExifIFDPointer
    ifd0 += struct.pack("<I", 0)

    exif_sub = bytearray()
    exif_sub += struct.pack("<H", 2)
    exif_sub += struct.pack("<HHII", 0x9003, 2, 20, 100)    # DateTimeOriginal
    exif_sub += struct.pack("<HHII", 0x9004, 2, 20, 120)    # DateTimeDigitized
    exif_sub += struct.pack("<I", 0)

    return b"Exif\x00\x00" + tiff_header + bytes(ifd0) + dt_bytes + bytes(exif_sub) + dt_bytes + dt_bytes


def inject_jpeg_exif(src_path: Path, dst_path: Path, best_date_utc: str) -> None:
    """
    Inject patched source EXIF (with Orientation reset to 1) or minimal capture-date EXIF
    into a newly written JPEG file atomically.
    """
    try:
        dst_bytes = dst_path.read_bytes()
        if len(dst_bytes) < 4 or not dst_bytes.startswith(b"\xff\xd8"):
            return

        exif_payload: bytes | None = None
        if src_path.suffix.lower() in {".jpg", ".jpeg", ".jfif"}:
            exif_payload = extract_and_patch_jpeg_exif(src_path)
        if exif_payload is None:
            exif_payload = build_minimal_exif_payload(best_date_utc)

        app1_segment = b"\xff\xe1" + struct.pack(">H", len(exif_payload) + 2) + exif_payload
        out = bytearray(b"\xff\xd8")
        inserted = False
        pos = 2
        n = len(dst_bytes)

        while pos < n:
            if dst_bytes[pos] != 0xFF:
                if not inserted:
                    out.extend(app1_segment)
                    inserted = True
                out.extend(dst_bytes[pos:])
                break
            marker_start = pos
            while pos < n and dst_bytes[pos] == 0xFF:
                pos += 1
            if pos >= n:
                if not inserted:
                    out.extend(app1_segment)
                    inserted = True
                out.extend(dst_bytes[marker_start:])
                break
            m = dst_bytes[pos]
            pos += 1

            if m in (0xD9, 0xDA):
                if not inserted:
                    out.extend(app1_segment)
                    inserted = True
                out.extend(dst_bytes[marker_start:])
                break
            if m == 0x00 or m == 0x01 or (0xD0 <= m <= 0xD7):
                out.extend(dst_bytes[marker_start:pos])
                continue
            if pos + 2 > n:
                if not inserted:
                    out.extend(app1_segment)
                    inserted = True
                out.extend(dst_bytes[marker_start:])
                break
            seg_len = struct.unpack_from(">H", dst_bytes, pos)[0]
            seg_end = pos + seg_len
            if seg_len < 2 or seg_end > n:
                if not inserted:
                    out.extend(app1_segment)
                    inserted = True
                out.extend(dst_bytes[marker_start:])
                break

            payload = dst_bytes[pos + 2 : seg_end]
            if m == 0xE0:
                out.extend(dst_bytes[marker_start:seg_end])
                if not inserted:
                    out.extend(app1_segment)
                    inserted = True
            elif m == 0xE1 and payload.startswith(b"Exif\x00\x00"):
                if not inserted:
                    out.extend(app1_segment)
                    inserted = True
            else:
                if not inserted:
                    out.extend(app1_segment)
                    inserted = True
                out.extend(dst_bytes[marker_start:seg_end])
            pos = seg_end

        tmp_path = dst_path.with_name(dst_path.name + ".exif.tmp")
        tmp_path.write_bytes(bytes(out))
        os.replace(tmp_path, dst_path)
    except Exception:
        pass


def classify_image(ext: str) -> tuple[bool, str]:
    """Classify image extension into compatible vs legacy/conversion-needed."""
    ext_lower = ext.lower()
    if ext_lower in {".bmp", ".tif", ".tiff"}:
        return True, f"구형 이미지 ({ext_lower} → .jpg)"
    if ext_lower == ".jfif":
        return True, "호환 변환 (.jfif → .jpg)"
    return False, "호환 이미지 (유지)"


def extract_image_thumbnail(
    file_path: Path,
    thumb_path: Path,
    ffmpeg_bin: str = "ffmpeg",
    max_dim: int = 320,
    quality: int = 5,
) -> bool:
    """Create a JPEG thumbnail or preview for an image using FFmpeg."""
    thumb_path.parent.mkdir(parents=True, exist_ok=True)
    vf = f"scale='min({int(max_dim)},iw)':'min({int(max_dim)},ih)':force_original_aspect_ratio=decrease"
    cmd = [
        ffmpeg_bin,
        "-y",
        "-i", str(file_path),
        "-frames:v", "1",
        "-vf", vf,
        "-q:v", str(int(quality)),
        str(thumb_path),
    ]
    try:
        res = run_cmd(cmd, timeout=45)
        return bool(res.returncode == 0 and thumb_path.exists() and thumb_path.stat().st_size > 0)
    except Exception:
        return False


def build_ffmpeg_image_cmd(
    src_path: Path,
    dst_path: Path,
    user_rotation: int,
    ffmpeg_bin: str = "ffmpeg",
) -> list[str]:
    """Build FFmpeg command to rotate/convert a single image file."""
    filters: list[str] = []
    if user_rotation == 90:
        filters.append("transpose=1")
    elif user_rotation == 180:
        filters.append("hflip,vflip")
    elif user_rotation == 270:
        filters.append("transpose=2")

    cmd = [ffmpeg_bin, "-y", "-i", str(src_path)]
    if filters:
        cmd.extend(["-vf", ",".join(filters)])
    ext_lower = dst_path.suffix.lower()
    if ext_lower != ".gif":
        cmd.extend(["-frames:v", "1"])
    if ext_lower in {".jpg", ".jpeg"}:
        cmd.extend(["-q:v", "2"])
    elif ext_lower == ".webp":
        cmd.extend(["-quality", "90"])
    cmd.append(str(dst_path))
    return cmd


def verify_output_image(out_path: Path) -> None:
    """Verify converted image exists and is non-empty."""
    if not out_path.exists() or out_path.stat().st_size <= 0:
        raise RuntimeError("출력 이미지 파일이 비어 있거나 생성되지 않았습니다.")


def build_ffmpeg_convert_cmd(
    src_path: Path,
    dst_path: Path,
    user_rotation: int,
    has_audio: bool,
    creation_time_utc: str,
    crf: int = 20,
    preset: str = "medium",
    ffmpeg_bin: str = "ffmpeg",
) -> list[str]:
    """Build FFmpeg command for H.264/AAC MP4 conversion with deinterlace, rotation, and metadata."""
    filters = ["yadif=mode=0:parity=-1:deint=1"]
    if user_rotation == 90:
        filters.append("transpose=1")
    elif user_rotation == 180:
        filters.append("hflip,vflip")
    elif user_rotation == 270:
        filters.append("transpose=2")
    filters.append("scale=trunc(iw/2)*2:trunc(ih/2)*2")

    cmd = [
        ffmpeg_bin,
        "-y",
        "-fflags", "+genpts",
        "-i", str(src_path),
        "-map", "0:v:0",
        "-map", "0:a:0?",
        "-vf", ",".join(filters),
        "-c:v", "libx264",
        "-pix_fmt", "yuv420p",
        "-crf", str(crf),
        "-preset", preset,
    ]
    if has_audio:
        cmd.extend(["-c:a", "aac", "-b:a", "128k", "-ar", "44100"])
    else:
        cmd.append("-an")

    cmd.extend([
        "-movflags", "+faststart",
        "-max_muxing_queue_size", "1024",
        "-map_metadata", "0",
        "-metadata", f"creation_time={creation_time_utc}",
        "-metadata:s:v:0", "rotate=",
        str(dst_path),
    ])
    return cmd


def build_ffmpeg_date_fix_cmd(
    src_path: Path,
    dst_path: Path,
    creation_time_utc: str,
    ffmpeg_bin: str = "ffmpeg",
) -> list[str]:
    """Build lossless remux command to embed creation_time into modern MP4/MOV."""
    return [
        ffmpeg_bin,
        "-y",
        "-i", str(src_path),
        "-map", "0",
        "-c", "copy",
        "-movflags", "+faststart",
        "-map_metadata", "0",
        "-metadata", f"creation_time={creation_time_utc}",
        str(dst_path),
    ]


def effective_dimensions(item: dict[str, Any]) -> tuple[int, int]:
    """Return even display dimensions after accounting for metadata and user rotation."""
    w = int(item.get("width") or 0)
    h = int(item.get("height") or 0)
    if w <= 0 or h <= 0:
        return (1280, 720)
    meta_rot = int(item.get("meta_rotation") or 0)
    if abs(meta_rot) % 180 == 90:
        w, h = h, w
    user_rot = int(item.get("user_rotation") or 0)
    if user_rot in (90, 270):
        w, h = h, w
    w = max(2, (w // 2) * 2)
    h = max(2, (h // 2) * 2)
    return (w, h)


def compute_group_canvas(group_items: list[dict[str, Any]]) -> tuple[int, int]:
    """Compute common even canvas size (max effective width and height) for a merge group."""
    dims = [effective_dimensions(it) for it in group_items]
    max_w = max((d[0] for d in dims), default=1280)
    max_h = max((d[1] for d in dims), default=720)
    return (max(2, (max_w // 2) * 2), max(2, (max_h // 2) * 2))


def build_ffmpeg_segment_cmd(
    src_path: Path,
    dst_ts_path: Path,
    user_rotation: int,
    has_audio: bool,
    group_has_audio: bool,
    target_w: int,
    target_h: int,
    crf: int = 20,
    preset: str = "medium",
    ffmpeg_bin: str = "ffmpeg",
) -> list[str]:
    """Build FFmpeg command to encode one normalized MPEG-TS segment for concatenation."""
    cmd = [
        ffmpeg_bin,
        "-y",
        "-fflags", "+genpts",
        "-i", str(src_path),
    ]
    if group_has_audio and not has_audio:
        cmd.extend([
            "-f", "lavfi",
            "-i", "anullsrc=channel_layout=stereo:sample_rate=44100",
            "-map", "0:v:0",
            "-map", "1:a:0",
            "-shortest",
        ])
    else:
        cmd.extend(["-map", "0:v:0"])
        if has_audio:
            cmd.extend(["-map", "0:a:0?"])

    filters = ["yadif=mode=0:parity=-1:deint=1"]
    if user_rotation == 90:
        filters.append("transpose=1")
    elif user_rotation == 180:
        filters.append("hflip,vflip")
    elif user_rotation == 270:
        filters.append("transpose=2")
    filters.append(
        f"scale={target_w}:{target_h}:force_original_aspect_ratio=decrease,"
        f"pad={target_w}:{target_h}:(ow-iw)/2:(oh-ih)/2,setsar=1"
    )

    cmd.extend([
        "-vf", ",".join(filters),
        "-c:v", "libx264",
        "-pix_fmt", "yuv420p",
        "-crf", str(crf),
        "-preset", preset,
        "-video_track_timescale", "90000",
    ])
    if group_has_audio:
        cmd.extend(["-c:a", "aac", "-b:a", "128k", "-ar", "44100", "-ac", "2"])
    else:
        cmd.append("-an")

    cmd.extend([
        "-max_muxing_queue_size", "1024",
        "-f", "mpegts",
        str(dst_ts_path),
    ])
    return cmd


def build_ffmpeg_concat_cmd(
    concat_list_path: Path,
    dst_mp4_path: Path,
    group_has_audio: bool,
    creation_time_utc: str,
    ffmpeg_bin: str = "ffmpeg",
) -> list[str]:
    """Build lossless concat demuxer command combining normalized MPEG-TS segments into MP4."""
    cmd = [
        ffmpeg_bin,
        "-y",
        "-f", "concat",
        "-safe", "0",
        "-i", str(concat_list_path),
        "-c", "copy",
    ]
    if group_has_audio:
        cmd.extend(["-bsf:a", "aac_adtstoasc"])
    else:
        cmd.append("-an")
    cmd.extend([
        "-movflags", "+faststart",
        "-metadata", f"creation_time={creation_time_utc}",
        str(dst_mp4_path),
    ])
    return cmd


def sanitize_merged_filename(raw_name: str, fallback_stem: str) -> str:
    """Sanitize a user-provided merged filename and ensure a .mp4 suffix."""
    base = Path((raw_name or "").replace("\\", "/")).name.strip()
    base = re.sub(r'[<>:"/\\|?*]+', "_", base).strip(" .")
    if not base:
        base = f"{fallback_stem}_merged"
    if base.lower().endswith(".mp4"):
        stem = base[:-4].strip(" .") or f"{fallback_stem}_merged"
    else:
        stem = Path(base).stem.strip(" .") or f"{fallback_stem}_merged"
    return f"{stem}.mp4"


def prune_merge_groups(
    merge_groups: list[dict[str, Any]],
    valid_ids: set[str],
) -> list[dict[str, Any]]:
    """Ensure every item belongs to at most one valid merge group of >= 2 items."""
    seen_items: set[str] = set()
    cleaned: list[dict[str, Any]] = []
    for grp in merge_groups or []:
        new_ids: list[str] = []
        for iid in grp.get("item_ids") or []:
            if iid in valid_ids and iid not in seen_items:
                seen_items.add(iid)
                new_ids.append(iid)
        if len(new_ids) >= 2:
            grp_copy = dict(grp)
            grp_copy["item_ids"] = new_ids
            grp_copy.setdefault("status", "pending")
            grp_copy.setdefault("output_rel_path", "")
            cleaned.append(grp_copy)
    return cleaned


def auto_detect_merge_groups(
    plan: dict[str, Any],
    max_gap_sec: float = 15.0,
) -> list[dict[str, Any]]:
    """
    Detect consecutive clips in the same folder whose start time is within max_gap_sec
    of the previous clip's end time (for clips not already in a merge group).
    """
    grouped_ids = {
        iid
        for grp in plan.get("merge_groups") or []
        for iid in grp.get("item_ids") or []
    }
    by_folder: dict[str, list[dict[str, Any]]] = {}
    for item in plan.get("items") or []:
        if item.get("media_type", "video") != "video" or item.get("delete"):
            continue
        if item["id"] in grouped_ids:
            continue
        folder = Path(item["rel_path"]).parent.as_posix()
        by_folder.setdefault(folder, []).append(item)

    created: list[dict[str, Any]] = []
    for folder in sorted(by_folder.keys()):
        clips = sorted(
            by_folder[folder],
            key=lambda x: (float(x.get("best_date_epoch") or 0.0), x["filename"].lower()),
        )
        if len(clips) < 2:
            continue
        run: list[dict[str, Any]] = [clips[0]]
        for nxt in clips[1:]:
            prev = run[-1]
            prev_end = float(prev.get("best_date_epoch") or 0.0) + max(float(prev.get("duration") or 0.0), 0.0)
            nxt_start = float(nxt.get("best_date_epoch") or 0.0)
            gap = nxt_start - prev_end
            if -5.0 <= gap <= max_gap_sec:
                run.append(nxt)
            else:
                if len(run) >= 2:
                    created.append({
                        "id": f"mg_{uuid.uuid4().hex[:8]}",
                        "name": sanitize_merged_filename("", Path(run[0]["filename"]).stem),
                        "item_ids": [c["id"] for c in run],
                        "output_rel_path": "",
                        "status": "pending",
                    })
                run = [nxt]
        if len(run) >= 2:
            created.append({
                "id": f"mg_{uuid.uuid4().hex[:8]}",
                "name": sanitize_merged_filename("", Path(run[0]["filename"]).stem),
                "item_ids": [c["id"] for c in run],
                "output_rel_path": "",
                "status": "pending",
            })
    return created


def atomic_write_json(path: Path, data: Any) -> None:
    """Atomically write JSON to disk to prevent corruption on crash."""
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp_path = path.with_suffix(".json.tmp")
    with tmp_path.open("w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    os.replace(tmp_path, path)


def load_plan(work_dir: Path) -> dict[str, Any]:
    """Load plan.json from work_dir and normalize item defaults for backward compatibility."""
    plan_path = work_dir / "plan.json"
    if not plan_path.exists():
        raise FileNotFoundError(f"plan.json을 찾을 수 없습니다: {plan_path}")
    with plan_path.open("r", encoding="utf-8") as f:
        data = json.load(f)
    data.setdefault("merge_groups", [])
    for item in data.get("items") or []:
        item.setdefault("media_type", "video")
        item.setdefault("delete", False)
    return data


def find_media_files(
    source_dir: Path,
    work_dir: Path,
    include_images: bool = True,
) -> list[tuple[Path, str]]:
    """Recursively find video (and optionally image) files while skipping work/trash/system dirs."""
    work_resolved = work_dir.resolve()
    results: list[tuple[Path, str]] = []
    for root, dirs, files in os.walk(source_dir):
        root_path = Path(root)
        kept_dirs = []
        for d in sorted(dirs):
            d_path = root_path / d
            if d.lower() in SKIP_DIR_NAMES:
                continue
            try:
                if d_path.resolve() == work_resolved:
                    continue
            except OSError:
                continue
            kept_dirs.append(d)
        dirs[:] = kept_dirs

        for fname in sorted(files):
            fpath = root_path / fname
            ext = fpath.suffix.lower()
            if ext in VIDEO_EXTENSIONS:
                results.append((fpath, "video"))
            elif include_images and ext in IMAGE_EXTENSIONS:
                results.append((fpath, "image"))
    return results


def find_video_files(source_dir: Path, work_dir: Path) -> list[Path]:
    """Backward-compatible helper returning only video file paths."""
    return [fpath for fpath, _ in find_media_files(source_dir, work_dir, include_images=False)]


def scan_single_video(
    file_path: Path,
    source_dir: Path,
    thumbs_dir: Path,
    existing_by_rel: dict[str, dict[str, Any]],
    ffmpeg_bin: str,
    ffprobe_bin: str,
) -> dict[str, Any]:
    """Probe and thumbnail a single video, preserving user choices if unchanged."""
    rel_posix = file_path.relative_to(source_dir).as_posix()
    stat = file_path.stat()
    file_size = stat.st_size
    file_mtime = stat.st_mtime

    prev = existing_by_rel.get(rel_posix)
    item_id = prev["id"] if prev and prev.get("id") else make_item_id(rel_posix)
    thumb_rel = f"thumbs/{item_id}.jpg"
    thumb_path = thumbs_dir / f"{item_id}.jpg"

    if (
        prev is not None
        and prev.get("size_bytes") == file_size
        and abs(float(prev.get("mtime_epoch", 0)) - file_mtime) < 1e-3
        and thumb_path.exists()
        and thumb_path.stat().st_size > 0
        and not prev.get("probe_error")
    ):
        prev.setdefault("media_type", "video")
        prev.setdefault("delete", False)
        return prev

    try:
        probe = probe_video(file_path, ffprobe_bin=ffprobe_bin)
        streams = probe.get("streams") or []
        v_streams = [s for s in streams if s.get("codec_type") == "video"]
        a_streams = [s for s in streams if s.get("codec_type") == "audio"]

        if not v_streams:
            raise RuntimeError("비디오 스트림이 없습니다.")

        v0 = v_streams[0]
        video_codec = str(v0.get("codec_name") or "unknown").lower()
        width = int(v0.get("width") or 0)
        height = int(v0.get("height") or 0)
        meta_rotation = parse_rotation_from_stream(v0)

        has_audio = len(a_streams) > 0
        audio_codec = str(a_streams[0].get("codec_name") or "").lower() if has_audio else ""

        fmt = probe.get("format") or {}
        try:
            duration = float(fmt.get("duration") or v0.get("duration") or 0.0)
        except (TypeError, ValueError):
            duration = 0.0

        utc_iso, utc_epoch, date_source = detect_best_date(probe, file_path.name, file_mtime)
        needs_convert, reason = classify_video(file_path.suffix, video_codec, has_audio, audio_codec)
        thumb_ok = extract_thumbnail(
            file_path,
            thumb_path,
            duration=duration,
            offset_ratio=THUMB_OFFSETS[0],
            ffmpeg_bin=ffmpeg_bin,
        )
        probe_error = ""
    except Exception as exc:
        video_codec = "unknown"
        audio_codec = ""
        has_audio = False
        width = 0
        height = 0
        duration = 0.0
        meta_rotation = 0
        utc_dt = datetime.fromtimestamp(file_mtime, tz=local_tz()).astimezone(timezone.utc)
        utc_iso = utc_dt.strftime("%Y-%m-%dT%H:%M:%SZ")
        utc_epoch = utc_dt.timestamp()
        date_source = "mtime"
        needs_convert = True
        reason = "분석 오류"
        thumb_ok = False
        probe_error = str(exc)

    user_rotation = int(prev.get("user_rotation", 0)) if prev else 0
    if user_rotation not in VALID_ROTATIONS:
        user_rotation = 0
    delete_flag = bool(prev.get("delete", False)) if prev else False
    convert_flag = bool(prev["convert"]) if (prev and "convert" in prev) else needs_convert
    if user_rotation != 0:
        convert_flag = True
    reviewed = bool(prev.get("reviewed", False)) if prev else False
    status = prev.get("status", "pending") if prev else "pending"
    if prev and prev.get("size_bytes") != file_size:
        status = "pending"

    return {
        "id": item_id,
        "media_type": "video",
        "rel_path": rel_posix,
        "filename": file_path.name,
        "ext": file_path.suffix.lower(),
        "size_bytes": file_size,
        "mtime_epoch": file_mtime,
        "duration": round(duration, 2),
        "width": width,
        "height": height,
        "video_codec": video_codec,
        "audio_codec": audio_codec,
        "has_audio": has_audio,
        "meta_rotation": meta_rotation,
        "best_date_utc": utc_iso,
        "best_date_epoch": utc_epoch,
        "date_source": date_source,
        "needs_convert": needs_convert,
        "reason": reason,
        "user_rotation": user_rotation,
        "convert": convert_flag,
        "delete": delete_flag,
        "reviewed": reviewed,
        "status": status,
        "thumb": thumb_rel if thumb_ok else "",
        "thumb_offset_idx": prev.get("thumb_offset_idx", 0) if prev else 0,
        "probe_error": probe_error,
    }


def scan_single_image(
    file_path: Path,
    source_dir: Path,
    thumbs_dir: Path,
    existing_by_rel: dict[str, dict[str, Any]],
    ffmpeg_bin: str,
    ffprobe_bin: str,
) -> dict[str, Any]:
    """Inspect image dimensions/EXIF/date and create a JPEG thumbnail, preserving user choices."""
    rel_posix = file_path.relative_to(source_dir).as_posix()
    stat = file_path.stat()
    file_size = stat.st_size
    file_mtime = stat.st_mtime

    prev = existing_by_rel.get(rel_posix)
    item_id = prev["id"] if prev and prev.get("id") else make_item_id(rel_posix)
    thumb_rel = f"thumbs/{item_id}.jpg"
    thumb_path = thumbs_dir / f"{item_id}.jpg"

    if (
        prev is not None
        and prev.get("size_bytes") == file_size
        and abs(float(prev.get("mtime_epoch", 0)) - file_mtime) < 1e-3
        and thumb_path.exists()
        and thumb_path.stat().st_size > 0
        and not prev.get("probe_error")
    ):
        prev.setdefault("media_type", "image")
        prev.setdefault("delete", False)
        return prev

    probe_error = ""
    info = read_image_basic_info(file_path)
    width = int(info.get("width") or 0)
    height = int(info.get("height") or 0)
    meta_rotation = int(info.get("meta_rotation") or 0)
    probe_data: dict[str, Any] = {}

    if width <= 0 or height <= 0:
        try:
            probe_data = probe_video(file_path, ffprobe_bin=ffprobe_bin)
            v_streams = [s for s in probe_data.get("streams") or [] if s.get("codec_type") == "video"]
            if v_streams:
                v0 = v_streams[0]
                width = int(v0.get("width") or 0)
                height = int(v0.get("height") or 0)
                if meta_rotation == 0:
                    meta_rotation = parse_rotation_from_stream(v0)
        except Exception as exc:
            probe_error = str(exc)

    exif_dt = parse_tag_datetime(info.get("date_str") or "")
    if exif_dt is not None:
        utc_iso = exif_dt.strftime("%Y-%m-%dT%H:%M:%SZ")
        utc_epoch = exif_dt.timestamp()
        date_source = "metadata"
    elif probe_data:
        utc_iso, utc_epoch, date_source = detect_best_date(probe_data, file_path.name, file_mtime)
    else:
        fn_dt = extract_filename_datetime(file_path.name, file_mtime)
        if fn_dt is not None:
            utc_iso = fn_dt.strftime("%Y-%m-%dT%H:%M:%SZ")
            utc_epoch = fn_dt.timestamp()
            date_source = "filename"
        else:
            mtime_dt = datetime.fromtimestamp(file_mtime, tz=local_tz()).astimezone(timezone.utc)
            utc_iso = mtime_dt.strftime("%Y-%m-%dT%H:%M:%SZ")
            utc_epoch = mtime_dt.timestamp()
            date_source = "mtime"

    needs_convert, reason = classify_image(file_path.suffix)
    thumb_ok = extract_image_thumbnail(
        file_path,
        thumb_path,
        ffmpeg_bin=ffmpeg_bin,
        max_dim=320,
        quality=5,
    )

    user_rotation = int(prev.get("user_rotation", 0)) if prev else 0
    if user_rotation not in VALID_ROTATIONS:
        user_rotation = 0
    delete_flag = bool(prev.get("delete", False)) if prev else False
    convert_flag = bool(prev["convert"]) if (prev and "convert" in prev) else needs_convert
    if user_rotation != 0:
        convert_flag = True
    reviewed = bool(prev.get("reviewed", False)) if prev else False
    status = prev.get("status", "pending") if prev else "pending"
    if prev and prev.get("size_bytes") != file_size:
        status = "pending"

    return {
        "id": item_id,
        "media_type": "image",
        "rel_path": rel_posix,
        "filename": file_path.name,
        "ext": file_path.suffix.lower(),
        "size_bytes": file_size,
        "mtime_epoch": file_mtime,
        "duration": 0.0,
        "width": width,
        "height": height,
        "video_codec": file_path.suffix.lstrip(".").lower() or "image",
        "audio_codec": "",
        "has_audio": False,
        "meta_rotation": meta_rotation,
        "best_date_utc": utc_iso,
        "best_date_epoch": utc_epoch,
        "date_source": date_source,
        "needs_convert": needs_convert,
        "reason": reason,
        "user_rotation": user_rotation,
        "convert": convert_flag,
        "delete": delete_flag,
        "reviewed": reviewed,
        "status": status,
        "thumb": thumb_rel if thumb_ok else "",
        "thumb_offset_idx": 0,
        "probe_error": probe_error,
    }


def run_scan_job(
    source_dir: Path,
    work_dir: Path,
    workers: int = 4,
    ffmpeg_bin: str = "ffmpeg",
    ffprobe_bin: str = "ffprobe",
    progress_cb: Callable[[int, int, str], None] | None = None,
    cancel_event: threading.Event | None = None,
    include_images: bool = True,
) -> dict[str, Any]:
    """Scan source_dir for videos (and optionally images), extract thumbnails, and write plan.json."""
    thumbs_dir = work_dir / "thumbs"
    thumbs_dir.mkdir(parents=True, exist_ok=True)

    existing_by_rel: dict[str, dict[str, Any]] = {}
    existing_merge_groups: list[dict[str, Any]] = []
    plan_path = work_dir / "plan.json"
    if plan_path.exists():
        try:
            old_plan = load_plan(work_dir)
            for item in old_plan.get("items", []):
                if "rel_path" in item:
                    existing_by_rel[item["rel_path"]] = item
            existing_merge_groups = old_plan.get("merge_groups", [])
        except Exception as exc:
            print(f"[경고] 기존 plan.json 읽기 실패 (새로 생성합니다): {exc}", file=sys.stderr)

    files = find_media_files(source_dir, work_dir, include_images=include_images)
    total = len(files)

    if total == 0:
        plan_data = {
            "version": 2,
            "source_dir": str(source_dir),
            "work_dir": str(work_dir),
            "include_images": include_images,
            "scanned_at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
            "items": [],
            "merge_groups": [],
        }
        atomic_write_json(plan_path, plan_data)
        if progress_cb:
            progress_cb(0, 0, "발견된 영상/이미지 파일이 없습니다.")
        return plan_data

    vid_cnt = sum(1 for _, mt in files if mt == "video")
    img_cnt = total - vid_cnt
    items: list[dict[str, Any]] = []
    done_count = 0
    if progress_cb:
        progress_cb(0, total, f"영상 {vid_cnt}개 · 이미지 {img_cnt}개 스캔을 시작합니다...")

    with ThreadPoolExecutor(max_workers=max(1, int(workers))) as pool:
        futures = {}
        for fpath, media_type in files:
            fn = scan_single_image if media_type == "image" else scan_single_video
            fut = pool.submit(
                fn,
                fpath,
                source_dir,
                thumbs_dir,
                existing_by_rel,
                ffmpeg_bin,
                ffprobe_bin,
            )
            futures[fut] = fpath

        for fut in as_completed(futures):
            if cancel_event and cancel_event.is_set():
                for pending_f in futures:
                    pending_f.cancel()
                raise JobCancelled("사용자에 의해 스캔이 중단되었습니다.")
            done_count += 1
            item = fut.result()
            items.append(item)
            if progress_cb:
                progress_cb(done_count, total, item["filename"])

    items.sort(key=lambda x: x["rel_path"].lower())
    valid_merge_ids = {
        x["id"] for x in items
        if x.get("media_type", "video") == "video" and not x.get("delete")
    }
    pruned_groups = prune_merge_groups(existing_merge_groups, valid_merge_ids)
    plan_data = {
        "version": 2,
        "source_dir": str(source_dir),
        "work_dir": str(work_dir),
        "include_images": include_images,
        "scanned_at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "items": items,
        "merge_groups": pruned_groups,
    }
    atomic_write_json(plan_path, plan_data)
    return plan_data


def cmd_scan(args: argparse.Namespace) -> int:
    """Execute the 'scan' subcommand."""
    ffmpeg_bin, ffprobe_bin = check_ffmpeg_tools()
    source_dir = Path(args.source_dir).expanduser().resolve()
    if not source_dir.is_dir():
        print(f"[오류] 폴더를 찾을 수 없습니다: {source_dir}", file=sys.stderr)
        return 1

    work_dir = Path(args.work_dir).expanduser().resolve() if args.work_dir else (source_dir / "_gphotos_work")
    include_images = not bool(getattr(args, "no_images", False))
    print(f"스캔 대상 폴더: {source_dir}")
    print(f"작업 폴더:      {work_dir}")
    print(f"이미지 포함:    {'예' if include_images else '아니오 (--no-images)'}")

    def on_progress(done: int, total: int, msg: str) -> None:
        if done == 0 and total > 0:
            print(msg)
        elif total == 0:
            print("발견된 영상/이미지 파일: 0개")
        elif done % 20 == 0 or done == total:
            print(f"  진행률: {done}/{total} ({msg})")

    plan_data = run_scan_job(
        source_dir=source_dir,
        work_dir=work_dir,
        workers=int(args.workers),
        ffmpeg_bin=ffmpeg_bin,
        ffprobe_bin=ffprobe_bin,
        progress_cb=on_progress,
        include_images=include_images,
    )
    items = plan_data.get("items", [])
    if not items:
        return 0

    vid_cnt = sum(1 for x in items if x.get("media_type", "video") == "video")
    img_cnt = sum(1 for x in items if x.get("media_type") == "image")
    convert_cnt = sum(1 for x in items if x["convert"])
    missing_date_cnt = sum(1 for x in items if x["date_source"] != "metadata")
    print("\n=== 스캔 완료 ===")
    print(f"전체 항목:         {len(items)}개 (동영상 {vid_cnt} · 이미지 {img_cnt})")
    print(f"변환/정리 예정:    {convert_cnt}개 (구형 포맷/코덱)")
    print(f"날짜 메타데이터 없음: {missing_date_cnt}개 (파일명/수정일로 보완 예정)")
    print(f"\n다음 단계: 아래 명령으로 브라우저에서 미리보기/회전/삭제/병합을 진행하세요:")
    print(f'  python "{Path(__file__).name}" review --work-dir "{work_dir}"')
    return 0


APP_HTML = r"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>Google Photos 옛날 사진·영상 정리 마법사</title>
<style>
  :root {
    --bg: #0f1117;
    --panel: #181c26;
    --card: #1f2431;
    --border: #2e3648;
    --text: #eef1f6;
    --muted: #9aa4b8;
    --accent: #4f8cff;
    --accent-hover: #3b78eb;
    --green: #2ecc71;
    --orange: #f59e0b;
    --red: #ef4444;
    --purple: #7c3aed;
  }
  * { box-sizing: border-box; }
  body {
    margin: 0;
    background: var(--bg);
    color: var(--text);
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "Malgun Gothic", sans-serif;
  }
  .topbar {
    position: sticky;
    top: 0;
    z-index: 30;
    background: rgba(15, 17, 23, 0.96);
    backdrop-filter: blur(8px);
    border-bottom: 1px solid var(--border);
    padding: 12px 20px;
  }
  .topbar-inner {
    max-width: 1440px;
    margin: 0 auto;
    display: flex;
    flex-direction: column;
    gap: 10px;
  }
  .header-top {
    display: flex;
    flex-wrap: wrap;
    align-items: center;
    justify-content: space-between;
    gap: 12px;
  }
  h1 { margin: 0; font-size: 18px; font-weight: 700; }
  .save-badge {
    font-size: 13px;
    padding: 4px 10px;
    border-radius: 999px;
    background: rgba(46, 204, 113, 0.14);
    color: var(--green);
    border: 1px solid rgba(46, 204, 113, 0.35);
  }
  .stepper {
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    gap: 8px;
  }
  .step-tab {
    background: var(--panel);
    color: var(--muted);
    border: 1px solid var(--border);
    border-radius: 10px;
    padding: 9px 12px;
    text-align: left;
    font-size: 13px;
    font-weight: 600;
    cursor: pointer;
    display: flex;
    align-items: center;
    gap: 8px;
  }
  .step-tab.active {
    background: rgba(79, 140, 255, 0.16);
    border-color: var(--accent);
    color: #fff;
  }
  .step-tab.done {
    border-color: rgba(46, 204, 113, 0.45);
    color: #b7f5d0;
  }
  .step-tab:disabled {
    opacity: 0.42;
    cursor: not-allowed;
  }
  .step-num {
    width: 22px;
    height: 22px;
    border-radius: 999px;
    background: var(--card);
    border: 1px solid var(--border);
    display: inline-flex;
    align-items: center;
    justify-content: center;
    font-size: 12px;
    flex-shrink: 0;
  }
  .step-tab.active .step-num {
    background: var(--accent);
    border-color: var(--accent);
    color: #fff;
  }
  main {
    max-width: 1440px;
    margin: 20px auto 48px auto;
    padding: 0 20px;
  }
  .step-view { display: none; }
  .step-view.active { display: block; }
  .panel-box {
    background: var(--panel);
    border: 1px solid var(--border);
    border-radius: 12px;
    padding: 18px 20px;
    margin-bottom: 18px;
  }
  .panel-title {
    margin: 0 0 10px 0;
    font-size: 16px;
    font-weight: 700;
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 10px;
    flex-wrap: wrap;
  }
  .muted { color: var(--muted); font-size: 13px; }
  .env-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
    gap: 14px;
    margin-top: 12px;
  }
  .env-card {
    background: var(--card);
    border: 1px solid var(--border);
    border-radius: 10px;
    padding: 14px;
    display: flex;
    flex-direction: column;
    gap: 8px;
  }
  .env-card.ok { border-color: rgba(46, 204, 113, 0.45); }
  .env-card.missing { border-color: rgba(239, 68, 68, 0.55); }
  .env-head {
    display: flex;
    align-items: center;
    justify-content: space-between;
    font-weight: 700;
    font-size: 14px;
  }
  .badge-ok { color: var(--green); font-size: 12px; }
  .badge-miss { color: var(--red); font-size: 12px; }
  .badge-opt { color: var(--orange); font-size: 12px; }
  .cmd-box {
    background: #090b10;
    border: 1px solid var(--border);
    border-radius: 7px;
    padding: 7px 10px;
    font-family: Consolas, monospace;
    font-size: 12px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 8px;
    word-break: break-all;
  }
  .alert-warn {
    background: rgba(239, 68, 68, 0.12);
    border: 1px solid rgba(239, 68, 68, 0.45);
    color: #fecaca;
    padding: 10px 14px;
    border-radius: 8px;
    font-size: 13px;
    margin-top: 12px;
  }
  .alert-info {
    background: rgba(79, 140, 255, 0.12);
    border: 1px solid rgba(79, 140, 255, 0.45);
    color: #dbeafe;
    padding: 12px 14px;
    border-radius: 8px;
    font-size: 13px;
    margin-top: 12px;
    display: flex;
    flex-wrap: wrap;
    align-items: center;
    justify-content: space-between;
    gap: 10px;
  }
  .folder-row {
    display: flex;
    flex-wrap: wrap;
    gap: 8px;
    margin-top: 10px;
  }
  .text-input {
    flex: 1;
    min-width: 260px;
    background: var(--card);
    color: var(--text);
    border: 1px solid var(--border);
    border-radius: 8px;
    padding: 9px 12px;
    font-size: 14px;
  }
  button.btn {
    background: var(--panel);
    color: var(--text);
    border: 1px solid var(--border);
    border-radius: 8px;
    padding: 7px 11px;
    font-size: 13px;
    cursor: pointer;
  }
  button.btn:hover:not(:disabled) { border-color: var(--accent); }
  button.btn:disabled { opacity: 0.45; cursor: not-allowed; }
  button.btn.selected {
    background: var(--orange);
    color: #111;
    border-color: var(--orange);
    font-weight: 700;
  }
  .btn-accent {
    background: var(--accent) !important;
    border-color: var(--accent) !important;
    color: #fff !important;
    font-weight: 700;
    padding: 9px 16px !important;
  }
  .btn-primary {
    background: var(--purple) !important;
    border-color: var(--purple) !important;
    color: #fff !important;
    font-weight: 600;
  }
  .btn-green {
    background: #16a34a !important;
    border-color: #16a34a !important;
    color: #fff !important;
    font-weight: 700;
    padding: 10px 18px !important;
  }
  .btn-danger {
    background: #b91c1c !important;
    border-color: #b91c1c !important;
    color: #fff !important;
    font-weight: 600;
  }
  .progress-wrap {
    margin-top: 14px;
    background: #090b10;
    border: 1px solid var(--border);
    border-radius: 999px;
    height: 16px;
    overflow: hidden;
  }
  .progress-fill {
    height: 100%;
    width: 0%;
    background: linear-gradient(90deg, var(--accent), var(--green));
    transition: width 0.25s ease;
  }
  .log-box {
    margin-top: 12px;
    background: #090b10;
    border: 1px solid var(--border);
    border-radius: 10px;
    padding: 12px;
    font-family: Consolas, monospace;
    font-size: 12px;
    line-height: 1.5;
    max-height: 300px;
    overflow-y: auto;
    white-space: pre-wrap;
    word-break: break-all;
  }
  .log-line-ok { color: #4ade80; }
  .log-line-err { color: #f87171; }
  .log-line-warn { color: #fbbf24; }
  .log-line-info { color: #cbd5e1; }
  .stats-bar {
    display: flex;
    flex-wrap: wrap;
    gap: 10px;
    margin-bottom: 12px;
    font-size: 13px;
    color: var(--muted);
  }
  .stat-chip {
    background: var(--panel);
    border: 1px solid var(--border);
    padding: 6px 11px;
    border-radius: 8px;
  }
  .stat-chip strong { color: var(--text); }
  .controls {
    display: flex;
    flex-wrap: wrap;
    align-items: center;
    justify-content: space-between;
    gap: 10px;
    margin-bottom: 10px;
  }
  .filters { display: flex; flex-wrap: wrap; gap: 6px; }
  .pill {
    background: var(--panel);
    color: var(--muted);
    border: 1px solid var(--border);
    padding: 6px 12px;
    border-radius: 999px;
    font-size: 13px;
    cursor: pointer;
  }
  .pill.active {
    background: var(--accent);
    color: #fff;
    border-color: var(--accent);
    font-weight: 600;
  }
  .search-box {
    background: var(--panel);
    color: var(--text);
    border: 1px solid var(--border);
    border-radius: 8px;
    padding: 7px 12px;
    font-size: 13px;
    min-width: 240px;
  }
  .merge-toolbar {
    display: flex;
    flex-wrap: wrap;
    align-items: center;
    justify-content: space-between;
    gap: 10px;
    padding: 10px 14px;
    background: var(--panel);
    border: 1px solid var(--border);
    border-radius: 10px;
    margin-bottom: 14px;
    font-size: 13px;
  }
  .merge-toolbar-group {
    display: flex;
    flex-wrap: wrap;
    align-items: center;
    gap: 8px;
  }
  .gap-input {
    width: 64px;
    background: var(--card);
    color: var(--text);
    border: 1px solid var(--border);
    border-radius: 6px;
    padding: 4px 8px;
    font-size: 13px;
  }
  .merge-groups-panel {
    background: var(--panel);
    border: 1px solid var(--border);
    border-radius: 12px;
    padding: 12px 14px;
    margin-bottom: 18px;
  }
  .merge-groups-header {
    display: flex;
    align-items: center;
    justify-content: space-between;
    cursor: pointer;
    font-weight: 700;
    font-size: 14px;
  }
  .merge-group-card {
    background: var(--card);
    border: 1px solid #4c3b80;
    border-radius: 10px;
    padding: 10px 12px;
    margin-top: 10px;
  }
  .merge-group-top {
    display: flex;
    flex-wrap: wrap;
    align-items: center;
    justify-content: space-between;
    gap: 8px;
    margin-bottom: 10px;
  }
  .merge-name-input {
    background: var(--panel);
    color: var(--text);
    border: 1px solid var(--border);
    border-radius: 6px;
    padding: 5px 8px;
    font-size: 13px;
    min-width: 240px;
  }
  .merge-clips-row { display: flex; flex-wrap: wrap; gap: 10px; }
  .merge-clip-chip {
    background: var(--panel);
    border: 1px solid var(--border);
    border-radius: 8px;
    padding: 6px;
    width: 150px;
    display: flex;
    flex-direction: column;
    gap: 5px;
    font-size: 11px;
  }
  .merge-clip-thumb {
    width: 100%;
    height: 82px;
    background: #080a0f;
    display: flex;
    align-items: center;
    justify-content: center;
    overflow: hidden;
    border-radius: 4px;
  }
  .merge-clip-thumb img { max-width: 76px; max-height: 76px; object-fit: contain; }
  .grid {
    display: grid;
    grid-template-columns: repeat(auto-fill, minmax(260px, 1fr));
    gap: 16px;
  }
  .card {
    background: var(--card);
    border: 1px solid var(--border);
    border-radius: 12px;
    overflow: hidden;
    display: flex;
    flex-direction: column;
    transition: border-color 0.15s;
  }
  .card.rotated { border-color: var(--orange); box-shadow: 0 0 0 1px rgba(245, 158, 11, 0.35); }
  .card.deleted { border-color: var(--red); box-shadow: 0 0 0 1px rgba(239, 68, 68, 0.4); }
  .card.deleted .thumb-stage img { opacity: 0.42; filter: grayscale(70%); }
  .thumb-stage {
    width: 100%;
    height: 220px;
    background: #080a0f;
    display: flex;
    align-items: center;
    justify-content: center;
    position: relative;
    overflow: hidden;
  }
  .thumb-stage img {
    max-width: 204px;
    max-height: 204px;
    object-fit: contain;
    transition: transform 0.2s ease;
    border-radius: 4px;
    cursor: zoom-in;
  }
  .merge-select-box {
    position: absolute;
    top: 8px;
    left: 8px;
    z-index: 3;
    display: inline-flex;
    align-items: center;
    gap: 4px;
    background: rgba(15, 17, 23, 0.85);
    border: 1px solid var(--border);
    border-radius: 6px;
    padding: 3px 7px;
    font-size: 11px;
    cursor: pointer;
  }
  .badges-top-right {
    position: absolute;
    top: 8px;
    right: 8px;
    z-index: 3;
    display: flex;
    flex-direction: column;
    align-items: flex-end;
    gap: 4px;
  }
  .rot-badge {
    background: rgba(245, 158, 11, 0.92);
    color: #111;
    font-weight: 700;
    font-size: 11px;
    padding: 3px 7px;
    border-radius: 6px;
  }
  .del-badge {
    background: rgba(239, 68, 68, 0.94);
    color: #fff;
    font-weight: 700;
    font-size: 11px;
    padding: 3px 7px;
    border-radius: 6px;
  }
  .media-badge {
    position: absolute;
    bottom: 8px;
    right: 8px;
    z-index: 3;
    background: rgba(15, 17, 23, 0.88);
    border: 1px solid var(--border);
    color: #e2e8f0;
    font-weight: 600;
    font-size: 11px;
    padding: 3px 7px;
    border-radius: 6px;
  }
  .merge-badge {
    position: absolute;
    bottom: 8px;
    left: 8px;
    z-index: 3;
    background: rgba(124, 58, 237, 0.92);
    color: #fff;
    font-weight: 700;
    font-size: 11px;
    padding: 3px 7px;
    border-radius: 6px;
  }
  .media-tabs {
    display: flex;
    flex-wrap: wrap;
    gap: 8px;
    margin-bottom: 10px;
  }
  .media-tab {
    background: var(--panel);
    color: var(--muted);
    border: 1px solid var(--border);
    padding: 7px 14px;
    border-radius: 10px;
    font-size: 13px;
    font-weight: 600;
    cursor: pointer;
  }
  .media-tab.active {
    background: rgba(79, 140, 255, 0.18);
    border-color: var(--accent);
    color: #fff;
  }
  .btn-del-active {
    background: rgba(239, 68, 68, 0.22) !important;
    border-color: var(--red) !important;
    color: #fecaca !important;
    font-weight: 700;
  }
  .modal-backdrop {
    position: fixed;
    inset: 0;
    z-index: 100;
    background: rgba(6, 8, 13, 0.95);
    display: none;
    flex-direction: column;
  }
  .modal-backdrop.open { display: flex; }
  .modal-header {
    display: flex;
    flex-wrap: wrap;
    align-items: center;
    justify-content: space-between;
    gap: 10px;
    padding: 12px 18px;
    border-bottom: 1px solid var(--border);
    background: rgba(15, 17, 23, 0.95);
  }
  .modal-stage {
    flex: 1;
    min-height: 0;
    position: relative;
    display: flex;
    align-items: center;
    justify-content: center;
    padding: 16px 64px;
    overflow: hidden;
  }
  .modal-img-wrap {
    position: relative;
    display: flex;
    align-items: center;
    justify-content: center;
    max-width: 100%;
    max-height: 100%;
  }
  .modal-img-wrap img {
    max-width: 82vw;
    max-height: 72vh;
    object-fit: contain;
    border-radius: 6px;
    transition: transform 0.2s ease;
    box-shadow: 0 12px 40px rgba(0, 0, 0, 0.65);
  }
  .modal-img-wrap img.rot90 {
    max-width: min(72vw, 72vh);
    max-height: min(72vw, 72vh);
  }
  .modal-img-wrap.deleted img {
    opacity: 0.45;
    filter: grayscale(70%);
  }
  .modal-deleted-overlay {
    position: absolute;
    padding: 8px 16px;
    border-radius: 8px;
    background: rgba(185, 28, 28, 0.92);
    color: #fff;
    font-weight: 700;
    font-size: 15px;
    pointer-events: none;
  }
  .modal-nav {
    position: absolute;
    top: 50%;
    transform: translateY(-50%);
    z-index: 4;
    background: rgba(24, 28, 38, 0.88);
    color: #fff;
    border: 1px solid var(--border);
    border-radius: 999px;
    width: 44px;
    height: 44px;
    font-size: 18px;
    cursor: pointer;
  }
  .modal-nav:hover:not(:disabled) { border-color: var(--accent); }
  .modal-nav:disabled { opacity: 0.3; cursor: not-allowed; }
  .modal-nav.prev { left: 14px; }
  .modal-nav.next { right: 14px; }
  .modal-footer {
    display: flex;
    flex-wrap: wrap;
    align-items: center;
    justify-content: space-between;
    gap: 10px;
    padding: 12px 18px;
    border-top: 1px solid var(--border);
    background: rgba(15, 17, 23, 0.95);
  }
  .modal-shortcuts { color: var(--muted); font-size: 12px; }
  .no-thumb { color: var(--muted); font-size: 12px; padding: 12px; text-align: center; }
  .card-body { padding: 12px; display: flex; flex-direction: column; gap: 8px; flex: 1; }
  .fname { font-size: 13px; font-weight: 600; word-break: break-all; line-height: 1.3; }
  .fpath { font-size: 11px; color: var(--muted); word-break: break-all; }
  .meta-row { display: flex; flex-wrap: wrap; gap: 5px; font-size: 11px; }
  .tag {
    background: var(--panel);
    border: 1px solid var(--border);
    padding: 2px 6px;
    border-radius: 5px;
    color: var(--muted);
  }
  .tag.warn { border-color: rgba(245, 158, 11, 0.45); color: #fbbf24; }
  .tag.ok { border-color: rgba(46, 204, 113, 0.4); color: #4ade80; }
  .rot-btns { display: grid; grid-template-columns: repeat(4, 1fr); gap: 5px; }
  .rot-btns button.btn { padding: 6px 4px; font-size: 12px; }
  .card-actions {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 6px;
    margin-top: auto;
    padding-top: 6px;
    border-top: 1px solid var(--border);
    font-size: 12px;
  }
  .chk-label { display: inline-flex; align-items: center; gap: 5px; cursor: pointer; user-select: none; }
  .mini-btns { display: flex; gap: 5px; }
  .mini-btns button.btn { padding: 5px 7px; font-size: 12px; }
  .pagination {
    display: flex;
    justify-content: center;
    align-items: center;
    gap: 12px;
    margin-top: 24px;
    font-size: 14px;
  }
  .mode-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
    gap: 14px;
    margin-top: 10px;
  }
  .mode-card {
    background: var(--card);
    border: 2px solid var(--border);
    border-radius: 10px;
    padding: 14px;
    cursor: pointer;
    display: flex;
    gap: 10px;
    align-items: flex-start;
  }
  .mode-card.active {
    border-color: var(--accent);
    background: rgba(79, 140, 255, 0.1);
  }
  .preview-list {
    max-height: 360px;
    overflow-y: auto;
    border: 1px solid var(--border);
    border-radius: 10px;
    background: #10131b;
  }
  .preview-row {
    display: flex;
    flex-wrap: wrap;
    align-items: center;
    justify-content: space-between;
    gap: 10px;
    padding: 10px 14px;
    border-bottom: 1px solid var(--border);
    font-size: 13px;
  }
  .preview-row:last-child { border-bottom: none; }
  .action-footer {
    display: flex;
    flex-wrap: wrap;
    align-items: center;
    justify-content: space-between;
    gap: 12px;
    margin-top: 16px;
  }
  @media (max-width: 780px) {
    .stepper { grid-template-columns: repeat(2, 1fr); }
  }
</style>
</head>
<body>
<div class="topbar">
  <div class="topbar-inner">
    <div class="header-top">
      <h1>Google Photos 옛날 사진·영상 정리 마법사</h1>
      <div id="saveBadge" class="save-badge">준비됨</div>
    </div>
    <div class="stepper">
      <button class="step-tab active" data-step="1"><span class="step-num">1</span>사전 준비 &amp; 폴더 선택</button>
      <button class="step-tab" data-step="2" disabled><span class="step-num">2</span>사진·영상 정리 (미리보기/회전/삭제/병합)</button>
      <button class="step-tab" data-step="3" disabled><span class="step-num">3</span>실행 전 확인</button>
      <button class="step-tab" data-step="4" disabled><span class="step-num">4</span>작업 진행</button>
    </div>
  </div>
</div>

<main>
  <!-- STEP 1 -->
  <section id="step1" class="step-view active">
    <div class="panel-box">
      <div class="panel-title">
        <span>1. PC 필수 프로그램 설치 확인</span>
        <button id="recheckEnvBtn" class="btn">🔄 설치 상태 다시 확인</button>
      </div>
      <div class="muted">아래 3가지가 설치되어 있으면 준비 완료입니다. 미설치 항목이 있으면 PowerShell에 명령어를 복사해 설치한 뒤 창을 새로 열거나 다시 확인을 누르세요.</div>
      <div id="envGrid" class="env-grid"></div>
      <div id="envWarn" class="alert-warn" style="display:none;">
        필수 프로그램(<strong>FFmpeg / FFprobe</strong>)이 아직 감지되지 않았습니다. PowerShell에서 <code>winget install -e --id Gyan.FFmpeg</code> 실행 후 이 프로그램을 다시 시작하거나 <strong>설치 상태 다시 확인</strong>을 눌러주세요.
      </div>
    </div>

    <div class="panel-box">
      <div class="panel-title"><span>2. HDD 사진·영상 폴더 선택 &amp; 스캔</span></div>
      <div class="muted">정리할 사진·영상이 들어 있는 폴더 경로(예: <code>D:\OldPhotos</code>)를 입력하거나 폴더 찾기를 누르세요.</div>
      <div class="folder-row">
        <input id="sourceDirInput" class="text-input" type="text" placeholder="예: D:\OldPhotos" />
        <button id="browseBtn" class="btn">📂 폴더 찾기</button>
        <button id="startScanBtn" class="btn btn-accent">🔍 사진·영상 스캔 시작</button>
      </div>
      <div style="margin-top:10px;font-size:13px;">
        <label class="chk-label"><input type="checkbox" id="optIncludeImages" checked /> <span>이미지 파일도 함께 스캔 (사진 크게 보기 · 회전 · 불필요한 파일 삭제 정리)</span></label>
      </div>
      <div id="existingPlanBox" class="alert-info" style="display:none;"></div>

      <div id="scanProgressWrap" style="display:none;">
        <div class="progress-wrap"><div id="scanProgressFill" class="progress-fill"></div></div>
        <div id="scanProgressText" class="muted" style="margin-top:8px;">준비 중...</div>
        <div style="margin-top:8px;">
          <button id="cancelScanBtn" class="btn btn-danger">⏹ 스캔 중단</button>
        </div>
        <div id="scanLogBox" class="log-box"></div>
      </div>
    </div>
  </section>

  <!-- STEP 2 -->
  <section id="step2" class="step-view">
    <div class="panel-title" style="margin-bottom:10px;">
      <div id="statsBar" class="stats-bar" style="margin-bottom:0;"></div>
      <button id="goStep3Btn" class="btn btn-accent">다음: 실행 전 작업 내용 확인 →</button>
    </div>

    <div id="mediaTabs" class="media-tabs"></div>

    <div class="controls">
      <div id="filterPills" class="filters"></div>
      <input id="searchInput" class="search-box" type="search" placeholder="파일명 또는 폴더 검색..." />
    </div>

    <div class="merge-toolbar">
      <div class="merge-toolbar-group">
        <span>선택: <strong id="selectedCount">0</strong>개 (Shift+클릭 범위 선택 / 페이지 이동 후 이어 선택 가능)</span>
        <button id="batchDeleteBtn" class="btn btn-danger" disabled>🗑 선택 삭제 표시</button>
        <button id="batchUndeleteBtn" class="btn" disabled>↩ 선택 삭제 해제</button>
        <button id="createMergeBtn" class="btn btn-primary" disabled>🔗 선택한 영상 병합 그룹 만들기</button>
        <button id="clearSelectBtn" class="btn">선택 해제</button>
      </div>
      <div class="merge-toolbar-group">
        <label>간격(초): <input id="autoGapInput" class="gap-input" type="number" min="1" max="3600" value="15" /></label>
        <button id="autoMergeBtn" class="btn">⚡ 연속 영상 자동 묶기</button>
      </div>
    </div>

    <div id="mergeGroupsPanel" class="merge-groups-panel" style="display:none;"></div>
    <div id="grid" class="grid"></div>
    <div id="pagination" class="pagination"></div>
  </section>

  <!-- STEP 3 -->
  <section id="step3" class="step-view">
    <div class="panel-box">
      <div class="panel-title"><span>1. 저장 모드 선택</span></div>
      <div class="mode-grid">
        <label class="mode-card active" id="modeCardOutput">
          <input type="radio" name="convMode" value="output" checked style="margin-top:3px;" />
          <div>
            <div style="font-weight:700;font-size:15px;">🛡️ 세이프 모드 (원본 100% 유지)</div>
            <div class="muted" style="margin-top:4px;">원본 파일은 전혀 건드리지 않고, 변환·회전된 영상(MP4)·이미지 결과물만 <code>_gphotos_work\converted\</code> 폴더에 따로 저장합니다. <em>(삭제 표시는 원본에 적용되지 않습니다)</em></div>
          </div>
        </label>
        <label class="mode-card" id="modeCardReplace">
          <input type="radio" name="convMode" value="replace-backup" style="margin-top:3px;" />
          <div>
            <div style="font-weight:700;font-size:15px;">📦 원본 백업 후 교체 모드 (업로드 최적화)</div>
            <div class="muted" style="margin-top:4px;">변환·회전 원본은 <code>_gphotos_work\original_backup\</code>, 삭제 표시 파일은 <code>_gphotos_work\deleted_trash\</code>로 안전하게 옮기고, 정리된 결과물을 원래 폴더 자리에 배치합니다.</div>
          </div>
        </label>
      </div>

      <div style="display:flex;flex-wrap:wrap;gap:16px;margin-top:14px;align-items:center;font-size:13px;">
        <label class="chk-label"><input type="checkbox" id="optFixDates" /> <span>촬영일 메타데이터가 없는 정상 MP4/MOV에도 무손실로 날짜 기록</span></label>
        <label class="chk-label"><input type="checkbox" id="optRetryFailed" /> <span>이전에 실패했던 항목 재시도</span></label>
        <label>영상 화질(CRF):
          <select id="optCrf" class="gap-input" style="width:auto;">
            <option value="18">18 (고화질)</option>
            <option value="20" selected>20 (표준 추천)</option>
            <option value="23">23 (용량 절약)</option>
          </select>
        </label>
      </div>
    </div>

    <div class="panel-box">
      <div class="panel-title">
        <span>2. 실행 전 작업 내용 리뷰</span>
        <button id="refreshPreviewBtn" class="btn">🔄 목록 새로고침</button>
      </div>
      <div id="deleteSafeWarn" class="alert-warn" style="display:none;margin-bottom:12px;"></div>
      <div id="previewStats" class="stats-bar"></div>
      <div id="previewList" class="preview-list"></div>

      <div class="action-footer">
        <button id="backToStep2Btn" class="btn">← 편집 화면으로</button>
        <button id="startConvertBtn" class="btn btn-green">🚀 정리/변환 시작하기</button>
      </div>
    </div>
  </section>

  <!-- STEP 4 -->
  <section id="step4" class="step-view">
    <div class="panel-box">
      <div class="panel-title">
        <span id="convertTitle">작업 진행 상태</span>
        <button id="cancelConvertBtn" class="btn btn-danger" style="display:none;">⏹ 현재 작업 후 안전 중단</button>
      </div>
      <div id="convertCurrentLabel" class="muted">대기 중...</div>
      <div class="progress-wrap"><div id="convertProgressFill" class="progress-fill"></div></div>
      <div id="convertCounters" class="stats-bar" style="margin-top:12px;"></div>
      <div id="convertLogBox" class="log-box"></div>

      <div class="action-footer" id="convertDoneActions" style="display:none;">
        <div style="display:flex;flex-wrap:wrap;gap:8px;">
          <button id="openOutputFolderBtn" class="btn btn-accent">📂 결과 폴더 열기</button>
          <button id="openBackupFolderBtn" class="btn" style="display:none;">📂 원본 백업 폴더 열기</button>
          <button id="openTrashFolderBtn" class="btn" style="display:none;">📂 삭제 휴지통 폴더 열기</button>
        </div>
        <div style="display:flex;flex-wrap:wrap;gap:8px;">
          <button id="step4BackEditBtn" class="btn">← 편집 화면으로</button>
          <button id="retryFailedBtn" class="btn btn-primary" style="display:none;">🔄 실패 항목 재시도</button>
        </div>
      </div>
    </div>
  </section>
</main>

<!-- PREVIEW LIGHTBOX MODAL -->
<div id="previewModal" class="modal-backdrop" role="dialog" aria-modal="true">
  <div class="modal-header">
    <div>
      <div id="modalTitle" style="font-weight:700;font-size:15px;"></div>
      <div id="modalSub" class="muted" style="font-size:12px;margin-top:2px;"></div>
    </div>
    <div style="display:flex;align-items:center;gap:10px;">
      <span id="modalIndex" class="muted"></span>
      <button id="modalCloseBtn" class="btn">✕ 닫기 (Esc)</button>
    </div>
  </div>
  <div class="modal-stage">
    <button id="modalPrevBtn" class="modal-nav prev" title="이전 (←)">◀</button>
    <div id="modalImgWrap" class="modal-img-wrap">
      <img id="modalImg" src="" alt="" />
      <div id="modalDeletedOverlay" class="modal-deleted-overlay" style="display:none;">🗑 삭제 예정 (휴지통 이동 대상)</div>
    </div>
    <button id="modalNextBtn" class="modal-nav next" title="다음 (→)">▶</button>
  </div>
  <div class="modal-footer">
    <div style="display:flex;flex-wrap:wrap;gap:6px;align-items:center;">
      <button id="modalRotCCW" class="btn">↺ 90° (L)</button>
      <button id="modalRotCW" class="btn">↻ 90° (R)</button>
      <button id="modalRot180" class="btn">↕ 180°</button>
      <button id="modalRotReset" class="btn">초기화 (0)</button>
      <button id="modalDelBtn" class="btn">🗑 삭제 표시 (D)</button>
      <button id="modalRethumbBtn" class="btn" style="display:none;">🖼 다른장면</button>
      <button id="modalOpenOsBtn" class="btn">📂 기본 앱으로 열기</button>
    </div>
    <div class="modal-shortcuts">단축키: ←/→ 이동 · L/R 회전 · 0 초기화 · D/Del 삭제 토글 · Esc 닫기</div>
  </div>
</div>

<script>
const PAGE_SIZE = 60;
let appState = { has_plan: false, source_dir: '', work_dir: '', initial_step: 1 };
let envStatus = { ready: false };
let plan = { items: [], merge_groups: [] };
let currentStep = 1;
let mediaFilter = 'all';
let currentFilter = 'all';
let searchQuery = '';
let currentPage = 1;
let selectedIds = [];
let lastClickedId = null;
let mergePanelCollapsed = false;
let scanPollTimer = null;
let convertPollTimer = null;
let modalItemId = null;

const FILTERS = [
  { key: 'all', label: '전체' },
  { key: 'convert', label: '변환/회전 예정' },
  { key: 'rotated', label: '회전 지정됨' },
  { key: 'deleted', label: '🗑 삭제 예정' },
  { key: 'merged', label: '병합 그룹' },
  { key: 'legacy', label: '구형 포맷/코덱' },
  { key: 'keep', label: '정상 유지' },
  { key: 'nodate', label: '날짜 메타데이터 없음' }
];

function escapeHtml(str) {
  return String(str || '').replace(/[&<>"']/g, s => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
  }[s]));
}

function setSaveStatus(text, isError) {
  const el = document.getElementById('saveBadge');
  el.textContent = text;
  el.style.color = isError ? '#ef4444' : '#2ecc71';
}

function fmtSize(bytes) {
  if (bytes >= 1024 * 1024 * 1024) return (bytes / (1024 ** 3)).toFixed(1) + ' GB';
  if (bytes >= 1024 * 1024) return (bytes / (1024 ** 2)).toFixed(1) + ' MB';
  return Math.max(1, Math.round(bytes / 1024)) + ' KB';
}

function fmtDur(sec) {
  const s = Math.round(sec || 0);
  return Math.floor(s / 60) + ':' + String(s % 60).padStart(2, '0');
}

function fmtDate(utcStr) {
  if (!utcStr) return '-';
  const d = new Date(utcStr);
  if (isNaN(d.getTime())) return utcStr;
  return d.toLocaleDateString('ko-KR') + ' ' + d.toLocaleTimeString('ko-KR', { hour: '2-digit', minute: '2-digit' });
}

function goStep(step) {
  currentStep = step;
  document.querySelectorAll('.step-view').forEach(el => el.classList.remove('active'));
  document.getElementById('step' + step).classList.add('active');
  document.querySelectorAll('.step-tab').forEach(tab => {
    const s = Number(tab.dataset.step);
    tab.classList.toggle('active', s === step);
  });
  if (step === 2) {
    refreshPlan().then(() => {
      renderStats();
      renderMediaTabs();
      renderMergePanel();
      renderGrid();
    });
  } else if (step === 3) {
    loadPreview();
  }
  window.scrollTo(0, 0);
}

function updateStepperLocks() {
  const hasPlan = Boolean(appState.has_plan);
  const convDoneOrRunning = appState.convert_job && appState.convert_job.status !== 'idle';
  document.querySelectorAll('.step-tab').forEach(tab => {
    const s = Number(tab.dataset.step);
    if (s === 1) tab.disabled = false;
    else if (s === 2 || s === 3) tab.disabled = !hasPlan;
    else if (s === 4) tab.disabled = !hasPlan && !convDoneOrRunning;
  });
}

async function copyText(text, btn) {
  try {
    await navigator.clipboard.writeText(text);
    const prev = btn.textContent;
    btn.textContent = '복사됨!';
    setTimeout(() => { btn.textContent = prev; }, 1200);
  } catch (e) {}
}

async function checkEnv() {
  const res = await fetch('/api/env');
  envStatus = await res.json();
  const defs = [
    { key: 'python', title: '1. Python 3', cmd: 'winget install -e --id Python.Python.3.12' },
    { key: 'ffmpeg', title: '2. FFmpeg / FFprobe (필수)', cmd: 'winget install -e --id Gyan.FFmpeg' },
    { key: 'git', title: '3. Git (업데이트용 권장)', cmd: 'winget install -e --id Git.Git' }
  ];
  const grid = document.getElementById('envGrid');
  grid.innerHTML = '';
  for (const d of defs) {
    const info = envStatus[d.key] || {};
    const ok = Boolean(info.installed);
    const card = document.createElement('div');
    card.className = 'env-card ' + (ok ? 'ok' : (info.required ? 'missing' : ''));
    const statusHtml = ok
      ? `<span class="badge-ok">✅ 설치됨</span>`
      : (info.required ? `<span class="badge-miss">❌ 미설치 (필수)</span>` : `<span class="badge-opt">⚠ 미설치 (선택)</span>`);
    card.innerHTML = `
      <div class="env-head"><span>${d.title}</span>${statusHtml}</div>
      <div class="muted">${ok ? escapeHtml(info.version) : '아래 명령어를 PowerShell에서 실행하세요.'}</div>
      <div class="cmd-box">
        <code>${escapeHtml(d.cmd)}</code>
        <button class="btn btn-copy" style="padding:3px 8px;font-size:11px;">복사</button>
      </div>
    `;
    card.querySelector('.btn-copy').onclick = (e) => copyText(d.cmd, e.target);
    grid.appendChild(card);
  }
  document.getElementById('envWarn').style.display = envStatus.ready ? 'none' : 'block';
  document.getElementById('startScanBtn').disabled = !envStatus.ready;
}

async function checkFolderInput() {
  const sourceDir = document.getElementById('sourceDirInput').value.trim();
  const box = document.getElementById('existingPlanBox');
  if (!sourceDir) {
    box.style.display = 'none';
    return;
  }
  const res = await fetch('/api/check-folder', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ source_dir: sourceDir })
  });
  const data = await res.json();
  if (!data.exists) {
    box.style.display = 'flex';
    box.innerHTML = `<span style="color:#fca5a5;">⚠ 해당 폴더를 찾을 수 없습니다: ${escapeHtml(sourceDir)}</span>`;
    return;
  }
  if (data.has_existing_plan) {
    box.style.display = 'flex';
    box.innerHTML = `
      <span>✅ 이전 스캔 기록을 발견했습니다! (항목 <strong>${data.existing_count}</strong>개 · 작업 폴더: <code>${escapeHtml(data.work_dir)}</code>)</span>
      <button class="btn btn-accent" id="loadExistingBtn">📂 기존 검수 이어하기 →</button>
    `;
    document.getElementById('loadExistingBtn').onclick = () => loadExistingPlan(sourceDir);
  } else {
    box.style.display = 'none';
  }
}

async function loadExistingPlan(sourceDir) {
  const res = await fetch('/api/load-existing', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ source_dir: sourceDir })
  });
  const data = await res.json();
  if (!res.ok) {
    alert(data.error || '불러오기 실패');
    return;
  }
  appState.has_plan = true;
  appState.source_dir = data.source_dir;
  appState.work_dir = data.work_dir;
  updateStepperLocks();
  goStep(2);
}

async function browseFolder() {
  const res = await fetch('/api/browse-folder', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: '{}' });
  const data = await res.json();
  if (data.ok && data.path) {
    document.getElementById('sourceDirInput').value = data.path;
    checkFolderInput();
  } else if (data.error) {
    alert(data.error);
  }
}

async function startScan() {
  const sourceDir = document.getElementById('sourceDirInput').value.trim();
  if (!sourceDir) { alert('사진·영상 폴더 경로를 입력하세요.'); return; }
  const includeImages = document.getElementById('optIncludeImages').checked;
  const res = await fetch('/api/scan/start', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ source_dir: sourceDir, include_images: includeImages })
  });
  const data = await res.json();
  if (!res.ok) { alert(data.error || '스캔 시작 실패'); return; }
  document.getElementById('scanProgressWrap').style.display = 'block';
  pollScanStatus();
}

async function pollScanStatus() {
  if (scanPollTimer) clearTimeout(scanPollTimer);
  const res = await fetch('/api/scan/status');
  const st = await res.json();
  document.getElementById('scanProgressWrap').style.display = st.status === 'idle' ? 'none' : 'block';
  const pct = st.total > 0 ? Math.round((st.done / st.total) * 100) : (st.status === 'done' ? 100 : 0);
  document.getElementById('scanProgressFill').style.width = pct + '%';
  document.getElementById('scanProgressText').textContent =
    st.status === 'running' ? `스캔 중: ${st.done} / ${st.total} (${st.current_file || ''})` :
    st.status === 'done' ? `스캔 완료: 총 ${st.total}개 항목` :
    st.status === 'cancelled' ? '스캔 중단됨' :
    st.status === 'error' ? `오류: ${st.error}` : '';

  const logBox = document.getElementById('scanLogBox');
  logBox.innerHTML = (st.logs || []).map(l => `<div class="log-line-${l.level || 'info'}">${escapeHtml(l.text)}</div>`).join('');
  logBox.scrollTop = logBox.scrollHeight;
  document.getElementById('cancelScanBtn').style.display = st.status === 'running' ? 'inline-block' : 'none';

  if (st.status === 'running') {
    scanPollTimer = setTimeout(pollScanStatus, 500);
  } else if (st.status === 'done') {
    appState.has_plan = true;
    appState.work_dir = st.work_dir;
    updateStepperLocks();
    setTimeout(() => goStep(2), 400);
  }
}

/* STEP 2 FUNCTIONS */
function getMergeMembership() {
  const map = {};
  (plan.merge_groups || []).forEach((grp, gIdx) => {
    (grp.item_ids || []).forEach((id, pos) => {
      map[id] = { groupNum: gIdx + 1, pos: pos + 1, total: grp.item_ids.length, group: grp };
    });
  });
  return map;
}

function matchesFilter(item, membership) {
  const mType = item.media_type || 'video';
  if (mediaFilter !== 'all' && mType !== mediaFilter) return false;
  if (currentFilter === 'deleted') {
    if (!item.delete) return false;
  } else {
    if (currentFilter === 'convert' && (item.delete || (!item.convert && !membership[item.id]))) return false;
    if (currentFilter === 'rotated' && (item.delete || item.user_rotation === 0)) return false;
    if (currentFilter === 'merged' && !membership[item.id]) return false;
    if (currentFilter === 'legacy' && !item.needs_convert) return false;
    if (currentFilter === 'keep' && (item.delete || item.convert || membership[item.id])) return false;
    if (currentFilter === 'nodate' && item.date_source === 'metadata') return false;
  }
  if (searchQuery) {
    const q = searchQuery.toLowerCase();
    if (!item.rel_path.toLowerCase().includes(q)) return false;
  }
  return true;
}

function getFilteredItems() {
  const membership = getMergeMembership();
  return (plan.items || []).filter(it => matchesFilter(it, membership));
}

function renderMediaTabs() {
  const items = plan.items || [];
  const total = items.length;
  const vidCnt = items.filter(x => (x.media_type || 'video') === 'video').length;
  const imgCnt = items.filter(x => x.media_type === 'image').length;
  const tabs = [
    { key: 'all', label: `전체 (${total})` },
    { key: 'video', label: `🎬 동영상 (${vidCnt})` },
    { key: 'image', label: `🖼 이미지 (${imgCnt})` }
  ];
  const wrap = document.getElementById('mediaTabs');
  wrap.innerHTML = '';
  for (const t of tabs) {
    const btn = document.createElement('button');
    btn.className = 'media-tab' + (mediaFilter === t.key ? ' active' : '');
    btn.textContent = t.label;
    btn.onclick = () => {
      mediaFilter = t.key;
      currentPage = 1;
      renderMediaTabs();
      renderGrid();
    };
    wrap.appendChild(btn);
  }
}

function renderStats() {
  const items = plan.items || [];
  const groups = plan.merge_groups || [];
  const total = items.length;
  const vidCnt = items.filter(x => (x.media_type || 'video') === 'video').length;
  const imgCnt = items.filter(x => x.media_type === 'image').length;
  const convertCnt = items.filter(x => !x.delete && x.convert).length;
  const rotCnt = items.filter(x => !x.delete && x.user_rotation !== 0).length;
  const delCnt = items.filter(x => x.delete).length;
  const noDateCnt = items.filter(x => x.date_source !== 'metadata').length;
  document.getElementById('statsBar').innerHTML = `
    <div class="stat-chip">전체 <strong>${total}</strong> (영상 ${vidCnt} · 사진 ${imgCnt})</div>
    <div class="stat-chip">변환/회전 <strong>${convertCnt}</strong></div>
    <div class="stat-chip">회전 지정 <strong>${rotCnt}</strong></div>
    <div class="stat-chip" style="${delCnt > 0 ? 'border-color:rgba(239,68,68,0.6);color:#fecaca;' : ''}">삭제 예정 <strong>${delCnt}</strong></div>
    <div class="stat-chip">병합 그룹 <strong>${groups.length}</strong></div>
    <div class="stat-chip">촬영일 추정 <strong>${noDateCnt}</strong></div>
  `;
}

function renderFilters() {
  const wrap = document.getElementById('filterPills');
  wrap.innerHTML = '';
  for (const f of FILTERS) {
    const btn = document.createElement('button');
    btn.className = 'pill' + (currentFilter === f.key ? ' active' : '');
    btn.textContent = f.label;
    btn.onclick = () => {
      currentFilter = f.key;
      currentPage = 1;
      renderFilters();
      renderGrid();
    };
    wrap.appendChild(btn);
  }
}

function updateSelectionUI() {
  document.getElementById('selectedCount').textContent = String(selectedIds.length);
  const byId = {};
  (plan.items || []).forEach(it => { byId[it.id] = it; });
  const selectedVideos = selectedIds.filter(id => byId[id] && (byId[id].media_type || 'video') === 'video' && !byId[id].delete);
  document.getElementById('createMergeBtn').disabled = selectedVideos.length < 2;
  document.getElementById('batchDeleteBtn').disabled = selectedIds.length === 0;
  document.getElementById('batchUndeleteBtn').disabled = selectedIds.length === 0;
}

async function refreshPlan() {
  const res = await fetch('/api/plan');
  plan = await res.json();
  plan.items = plan.items || [];
  plan.merge_groups = plan.merge_groups || [];
}

async function updateItem(id, patch) {
  setSaveStatus('저장 중...', false);
  try {
    const res = await fetch('/api/update', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ id, ...patch })
    });
    if (!res.ok) throw new Error('저장 실패');
    const data = await res.json();
    const idx = plan.items.findIndex(x => x.id === id);
    if (idx >= 0) plan.items[idx] = data.item;
    if (Array.isArray(data.merge_groups)) plan.merge_groups = data.merge_groups;
    setSaveStatus('자동 저장됨', false);
    renderStats();
    renderMediaTabs();
    updateSelectionUI();
    renderMergePanel();
    renderGrid();
    if (modalItemId === id) renderPreviewModal();
  } catch (err) {
    setSaveStatus('저장 오류!', true);
  }
}

async function batchUpdateItems(itemIds, patch) {
  if (!itemIds || itemIds.length === 0) return;
  setSaveStatus('일괄 저장 중...', false);
  try {
    const res = await fetch('/api/batch-update', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ item_ids: itemIds, ...patch })
    });
    if (!res.ok) throw new Error('저장 실패');
    const data = await res.json();
    const updatedMap = {};
    (data.items || []).forEach(it => { updatedMap[it.id] = it; });
    plan.items = (plan.items || []).map(it => updatedMap[it.id] || it);
    if (Array.isArray(data.merge_groups)) plan.merge_groups = data.merge_groups;
    selectedIds = [];
    lastClickedId = null;
    updateSelectionUI();
    setSaveStatus(`자동 저장됨 (${data.updated}개)`, false);
    renderStats();
    renderMediaTabs();
    renderMergePanel();
    renderGrid();
  } catch (err) {
    setSaveStatus('저장 오류!', true);
  }
}

async function createMergeGroup() {
  if (selectedIds.length < 2) return;
  setSaveStatus('병합 그룹 생성 중...', false);
  try {
    const res = await fetch('/api/merge/create', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ item_ids: selectedIds })
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || '생성 실패');
    plan.merge_groups = data.merge_groups || [];
    selectedIds = [];
    lastClickedId = null;
    updateSelectionUI();
    setSaveStatus('자동 저장됨', false);
    renderStats();
    renderMergePanel();
    renderGrid();
  } catch (err) {
    setSaveStatus('병합 오류!', true);
  }
}

async function updateMergeGroup(groupId, payload) {
  setSaveStatus('저장 중...', false);
  try {
    const res = await fetch('/api/merge/update', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ group_id: groupId, ...payload })
    });
    if (!res.ok) throw new Error('수정 실패');
    await refreshPlan();
    setSaveStatus('자동 저장됨', false);
    renderStats();
    renderMergePanel();
    renderGrid();
  } catch (err) {
    setSaveStatus('병합 수정 오류!', true);
  }
}

async function deleteMergeGroup(groupId) {
  setSaveStatus('해제 중...', false);
  try {
    const res = await fetch('/api/merge/delete', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ group_id: groupId })
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || '해제 실패');
    plan.merge_groups = data.merge_groups || [];
    setSaveStatus('자동 저장됨', false);
    renderStats();
    renderMergePanel();
    renderGrid();
  } catch (err) {
    setSaveStatus('해제 오류!', true);
  }
}

async function autoDetectMerge() {
  const gap = Number(document.getElementById('autoGapInput').value || 15);
  setSaveStatus('연속 영상 탐색 중...', false);
  try {
    const res = await fetch('/api/merge/auto', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ max_gap_sec: gap })
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || '자동 묶기 실패');
    plan.merge_groups = data.merge_groups || [];
    setSaveStatus(`자동 저장됨 (새 그룹 ${data.created}개)`, false);
    renderStats();
    renderMergePanel();
    renderGrid();
  } catch (err) {
    setSaveStatus('자동 묶기 오류!', true);
  }
}

function renderMergePanel() {
  const panel = document.getElementById('mergeGroupsPanel');
  const groups = plan.merge_groups || [];
  if (groups.length === 0) {
    panel.style.display = 'none';
    panel.innerHTML = '';
    return;
  }
  panel.style.display = 'block';
  const byId = {};
  (plan.items || []).forEach(it => { byId[it.id] = it; });

  panel.innerHTML = `
    <div class="merge-groups-header" id="mergePanelToggle">
      <span>🔗 병합 그룹 목록 (${groups.length}개) — 같은 그룹 영상은 순서대로 회전 적용 후 하나의 MP4로 합쳐집니다</span>
      <span>${mergePanelCollapsed ? '▼ 펼치기' : '▲ 접기'}</span>
    </div>
    <div id="mergePanelBody" style="display:${mergePanelCollapsed ? 'none' : 'block'};"></div>
  `;
  document.getElementById('mergePanelToggle').onclick = () => {
    mergePanelCollapsed = !mergePanelCollapsed;
    renderMergePanel();
  };

  const body = document.getElementById('mergePanelBody');
  groups.forEach((grp, idx) => {
    const card = document.createElement('div');
    card.className = 'merge-group-card';
    const members = (grp.item_ids || []).map(id => byId[id]).filter(Boolean);
    const totalDur = members.reduce((s, m) => s + (m.duration || 0), 0);

    card.innerHTML = `
      <div class="merge-group-top">
        <div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;">
          <strong>#${idx + 1}</strong>
          <input class="merge-name-input" type="text" value="${escapeHtml(grp.name)}" />
          <button class="btn btn-save-name">이름 저장</button>
          <span class="muted">${members.length}개 클립 · 총 ${fmtDur(totalDur)}</span>
        </div>
        <div style="display:flex;align-items:center;gap:5px;flex-wrap:wrap;">
          <button class="btn" data-grot="270">전체 ↺ 90°</button>
          <button class="btn" data-grot="90">전체 ↻ 90°</button>
          <button class="btn" data-grot="180">전체 ↕ 180°</button>
          <button class="btn" data-grot="0">전체 초기화</button>
          <button class="btn btn-del-grp" style="color:#fca5a5;">그룹 해제</button>
        </div>
      </div>
      <div class="merge-clips-row"></div>
    `;

    const nameInp = card.querySelector('.merge-name-input');
    card.querySelector('.btn-save-name').onclick = () => updateMergeGroup(grp.id, { name: nameInp.value });
    card.querySelectorAll('[data-grot]').forEach(btn => {
      btn.onclick = () => updateMergeGroup(grp.id, { rotate_all: Number(btn.dataset.grot) });
    });
    card.querySelector('.btn-del-grp').onclick = () => deleteMergeGroup(grp.id);

    const row = card.querySelector('.merge-clips-row');
    members.forEach((m, pos) => {
      const chip = document.createElement('div');
      chip.className = 'merge-clip-chip';
      chip.innerHTML = `
        <div class="merge-clip-thumb">
          ${m.thumb ? `<img src="/${m.thumb}" style="transform: rotate(${m.user_rotation}deg);" alt="" />` : '<span>없음</span>'}
        </div>
        <div style="font-weight:600;word-break:break-all;">${pos + 1}. ${escapeHtml(m.filename)}</div>
        <div style="color:var(--muted);">${fmtDur(m.duration)}${m.user_rotation ? ' · ' + m.user_rotation + '°' : ''}</div>
        <div style="display:flex;gap:4px;margin-top:auto;">
          <button class="btn btn-left" ${pos === 0 ? 'disabled' : ''}>◀</button>
          <button class="btn btn-right" ${pos === members.length - 1 ? 'disabled' : ''}>▶</button>
          <button class="btn btn-remove" style="margin-left:auto;color:#fca5a5;">✕</button>
        </div>
      `;
      chip.querySelector('.btn-left').onclick = () => {
        const arr = [...grp.item_ids];
        [arr[pos - 1], arr[pos]] = [arr[pos], arr[pos - 1]];
        updateMergeGroup(grp.id, { item_ids: arr });
      };
      chip.querySelector('.btn-right').onclick = () => {
        const arr = [...grp.item_ids];
        [arr[pos + 1], arr[pos]] = [arr[pos], arr[pos + 1]];
        updateMergeGroup(grp.id, { item_ids: arr });
      };
      chip.querySelector('.btn-remove').onclick = () => {
        const arr = grp.item_ids.filter(id => id !== m.id);
        if (arr.length < 2) deleteMergeGroup(grp.id);
        else updateMergeGroup(grp.id, { item_ids: arr });
      };
      row.appendChild(chip);
    });
    body.appendChild(card);
  });
}

async function rethumb(id) {
  setSaveStatus('썸네일 갱신 중...', false);
  try {
    const res = await fetch('/api/rethumb', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ id })
    });
    const data = await res.json();
    if (data.thumb_url) {
      const img = document.getElementById('img-' + id);
      if (img) img.src = data.thumb_url;
      if (modalItemId === id) {
        const mImg = document.getElementById('modalImg');
        if (mImg) mImg.src = data.thumb_url;
      }
    }
    setSaveStatus('자동 저장됨', false);
  } catch (e) {
    setSaveStatus('썸네일 오류', true);
  }
}

async function openVideo(id) {
  await fetch('/api/open', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ id })
  });
}

function toggleSelectItem(id, checked, shiftKey, pageIds) {
  if (shiftKey && lastClickedId && pageIds.includes(lastClickedId)) {
    const a = pageIds.indexOf(lastClickedId);
    const b = pageIds.indexOf(id);
    const [lo, hi] = a <= b ? [a, b] : [b, a];
    for (let i = lo; i <= hi; i++) {
      const tid = pageIds[i];
      if (checked && !selectedIds.includes(tid)) selectedIds.push(tid);
      if (!checked) selectedIds = selectedIds.filter(x => x !== tid);
    }
  } else {
    if (checked && !selectedIds.includes(id)) selectedIds.push(id);
    if (!checked) selectedIds = selectedIds.filter(x => x !== id);
  }
  lastClickedId = id;
  updateSelectionUI();
  renderGrid();
}

function renderGrid() {
  const membership = getMergeMembership();
  const filtered = getFilteredItems();
  const totalPages = Math.max(1, Math.ceil(filtered.length / PAGE_SIZE));
  if (currentPage > totalPages) currentPage = totalPages;
  const start = (currentPage - 1) * PAGE_SIZE;
  const pageItems = filtered.slice(start, start + PAGE_SIZE);
  const pageIds = pageItems.map(x => x.id);

  const grid = document.getElementById('grid');
  grid.innerHTML = '';

  if (pageItems.length === 0) {
    grid.innerHTML = '<div style="color:#9aa4b8;padding:30px;">조건에 맞는 항목이 없습니다.</div>';
  }

  for (const item of pageItems) {
    const isImage = item.media_type === 'image';
    const isDeleted = Boolean(item.delete);
    const card = document.createElement('div');
    card.className = 'card' + (isDeleted ? ' deleted' : (item.user_rotation ? ' rotated' : ''));

    const dateBadgeClass = item.date_source === 'metadata' ? 'ok' : 'warn';
    const dateLabel = item.date_source === 'metadata' ? (isImage ? 'EXIF' : '메타') : (item.date_source === 'filename' ? '파일명' : '수정일');
    const isSel = selectedIds.includes(item.id);
    const mem = membership[item.id];
    const thumbHtml = item.thumb
      ? `<img id="img-${item.id}" src="/${item.thumb}" style="transform: rotate(${item.user_rotation}deg);" loading="lazy" title="클릭하여 크게 보기" alt="" />`
      : (isImage
          ? `<img id="img-${item.id}" src="/preview/${item.id}" style="transform: rotate(${item.user_rotation}deg);" loading="lazy" title="클릭하여 크게 보기" alt="" />`
          : `<div class="no-thumb">썸네일 없음<br/>${escapeHtml(item.probe_error || '')}</div>`);

    const dimTag = (item.width && item.height) ? `<span class="tag">${item.width}x${item.height}</span>` : '';
    const durTag = isImage ? '' : `<span class="tag">${fmtDur(item.duration)}</span>`;
    const convLabel = isImage ? '변환/저장' : 'MP4 변환';

    card.innerHTML = `
      <div class="thumb-stage">
        <label class="merge-select-box">
          <input type="checkbox" class="merge-chk" ${isSel ? 'checked' : ''} />
          <span>선택</span>
        </label>
        ${thumbHtml}
        <div class="badges-top-right">
          ${isDeleted ? `<div class="del-badge">🗑 삭제 예정</div>` : ''}
          ${item.user_rotation ? `<div class="rot-badge">회전 ${item.user_rotation}°</div>` : ''}
        </div>
        ${mem ? `<div class="merge-badge">🔗 병합 #${mem.groupNum} (${mem.pos}/${mem.total})</div>` : ''}
        <div class="media-badge">${isImage ? '🖼 사진' : '🎬 영상'}</div>
      </div>
      <div class="card-body">
        <div class="fname">${escapeHtml(item.filename)}</div>
        <div class="fpath">${escapeHtml(item.rel_path)}</div>
        <div class="meta-row">
          <span class="tag ${item.needs_convert ? 'warn' : 'ok'}">${escapeHtml(item.reason)}</span>
          ${dimTag}
          ${durTag}
          <span class="tag">${fmtSize(item.size_bytes)}</span>
          <span class="tag ${dateBadgeClass}">${dateLabel}: ${fmtDate(item.best_date_utc)}</span>
        </div>
        <div class="rot-btns">
          <button class="btn ${item.user_rotation === 270 ? 'selected' : ''}" data-rot="270">↺ 90°</button>
          <button class="btn ${item.user_rotation === 90 ? 'selected' : ''}" data-rot="90">↻ 90°</button>
          <button class="btn ${item.user_rotation === 180 ? 'selected' : ''}" data-rot="180">↕ 180°</button>
          <button class="btn ${item.user_rotation === 0 ? 'selected' : ''}" data-rot="0">초기화</button>
        </div>
        <div class="card-actions">
          <label class="chk-label">
            <input type="checkbox" class="conv-chk" ${item.convert ? 'checked' : ''} />
            <span>${convLabel}</span>
          </label>
          <div class="mini-btns">
            <button class="btn btn-del-toggle ${isDeleted ? 'btn-del-active' : ''}" title="휴지통으로 이동 표시">${isDeleted ? '↩ 취소' : '🗑 삭제'}</button>
            <button class="btn btn-zoom" title="크게 미리보기">🔍</button>
            ${isImage ? '' : '<button class="btn btn-rethumb" title="다른 시점 썸네일 추출">🖼</button>'}
            <button class="btn btn-open" title="로컬 기본 앱으로 열기">${isImage ? '📂' : '▶'}</button>
          </div>
        </div>
      </div>
    `;

    const thumbImg = card.querySelector('.thumb-stage img');
    if (thumbImg) thumbImg.onclick = () => openPreviewModal(item.id);

    const mChk = card.querySelector('.merge-chk');
    mChk.onclick = (e) => toggleSelectItem(item.id, mChk.checked, e.shiftKey, pageIds);

    card.querySelectorAll('[data-rot]').forEach(btn => {
      btn.onclick = () => updateItem(item.id, { user_rotation: Number(btn.dataset.rot), reviewed: true });
    });
    const chk = card.querySelector('.conv-chk');
    chk.onchange = () => updateItem(item.id, { convert: chk.checked, reviewed: true });
    card.querySelector('.btn-del-toggle').onclick = () => updateItem(item.id, { delete: !isDeleted, reviewed: true });
    card.querySelector('.btn-zoom').onclick = () => openPreviewModal(item.id);
    const rethumbBtn = card.querySelector('.btn-rethumb');
    if (rethumbBtn) rethumbBtn.onclick = () => rethumb(item.id);
    card.querySelector('.btn-open').onclick = () => openVideo(item.id);

    grid.appendChild(card);
  }

  const pag = document.getElementById('pagination');
  pag.innerHTML = '';
  if (totalPages > 1) {
    const prev = document.createElement('button');
    prev.className = 'btn';
    prev.textContent = '◀ 이전';
    prev.disabled = currentPage <= 1;
    prev.onclick = () => { currentPage--; renderGrid(); window.scrollTo(0, 0); };

    const info = document.createElement('span');
    info.textContent = `${currentPage} / ${totalPages} 페이지 (총 ${filtered.length}개)`;

    const next = document.createElement('button');
    next.className = 'btn';
    next.textContent = '다음 ▶';
    next.disabled = currentPage >= totalPages;
    next.onclick = () => { currentPage++; renderGrid(); window.scrollTo(0, 0); };

    pag.append(prev, info, next);
  }
}

/* LIGHTBOX PREVIEW MODAL */
function openPreviewModal(itemId) {
  modalItemId = itemId;
  document.getElementById('previewModal').classList.add('open');
  renderPreviewModal();
}

function closePreviewModal() {
  modalItemId = null;
  document.getElementById('previewModal').classList.remove('open');
}

function renderPreviewModal() {
  if (!modalItemId) return;
  const filtered = getFilteredItems();
  let idx = filtered.findIndex(x => x.id === modalItemId);
  let item = idx >= 0 ? filtered[idx] : (plan.items || []).find(x => x.id === modalItemId);
  if (!item) {
    closePreviewModal();
    return;
  }
  const isImage = item.media_type === 'image';
  const isDeleted = Boolean(item.delete);
  const rot = Number(item.user_rotation || 0);

  document.getElementById('modalTitle').textContent = `${isImage ? '🖼' : '🎬'} ${item.filename}`;
  document.getElementById('modalSub').textContent =
    `${item.rel_path} · ${item.width}x${item.height} · ${fmtSize(item.size_bytes)} · ${fmtDate(item.best_date_utc)}` +
    (rot ? ` · 회전 ${rot}°` : '');
  document.getElementById('modalIndex').textContent = idx >= 0 ? `${idx + 1} / ${filtered.length}` : '';

  const img = document.getElementById('modalImg');
  const wantSrc = `/preview/${item.id}`;
  if (img.dataset.itemId !== item.id) {
    img.dataset.itemId = item.id;
    img.src = wantSrc;
  }
  img.style.transform = `rotate(${rot}deg)`;
  img.classList.toggle('rot90', rot === 90 || rot === 270);

  document.getElementById('modalImgWrap').classList.toggle('deleted', isDeleted);
  document.getElementById('modalDeletedOverlay').style.display = isDeleted ? 'block' : 'none';

  const delBtn = document.getElementById('modalDelBtn');
  delBtn.textContent = isDeleted ? '↩ 삭제 취소 (D)' : '🗑 삭제 표시 (D)';
  delBtn.classList.toggle('btn-del-active', isDeleted);

  document.getElementById('modalRethumbBtn').style.display = isImage ? 'none' : 'inline-block';
  document.getElementById('modalOpenOsBtn').textContent = isImage ? '📂 기본 뷰어로 열기' : '▶ 로컬 플레이어로 재생';

  document.getElementById('modalPrevBtn').disabled = idx <= 0;
  document.getElementById('modalNextBtn').disabled = idx < 0 || idx >= filtered.length - 1;
}

function navigateModal(delta) {
  if (!modalItemId) return;
  const filtered = getFilteredItems();
  const idx = filtered.findIndex(x => x.id === modalItemId);
  if (idx < 0) return;
  const nextIdx = idx + delta;
  if (nextIdx >= 0 && nextIdx < filtered.length) {
    modalItemId = filtered[nextIdx].id;
    renderPreviewModal();
  }
}

function rotateModalBy(deltaDeg) {
  if (!modalItemId) return;
  const item = (plan.items || []).find(x => x.id === modalItemId);
  if (!item) return;
  const nextRot = (Number(item.user_rotation || 0) + deltaDeg + 360) % 360;
  updateItem(item.id, { user_rotation: nextRot, reviewed: true });
}

/* STEP 3 FUNCTIONS */
function selectedMode() {
  const el = document.querySelector('input[name="convMode"]:checked');
  return el ? el.value : 'output';
}

async function loadPreview() {
  const mode = selectedMode();
  document.getElementById('modeCardOutput').classList.toggle('active', mode === 'output');
  document.getElementById('modeCardReplace').classList.toggle('active', mode === 'replace-backup');

  const res = await fetch('/api/convert/preview', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      mode,
      fix_missing_dates: document.getElementById('optFixDates').checked,
      retry_failed: document.getElementById('optRetryFailed').checked
    })
  });
  const data = await res.json();
  if (!res.ok) return;

  const s = data.summary;
  const delWarn = document.getElementById('deleteSafeWarn');
  if (s.delete_marked_count > 0 && mode === 'output') {
    delWarn.style.display = 'flex';
    delWarn.style.alignItems = 'center';
    delWarn.style.justifyContent = 'space-between';
    delWarn.style.flexWrap = 'wrap';
    delWarn.style.gap = '10px';
    delWarn.innerHTML = `
      <span>⚠ 현재 <strong>세이프 모드</strong>이므로 삭제 표시한 <strong>${s.delete_marked_count}</strong>개 파일이 원본 폴더에서 지워지지 않습니다. 휴지통(<code>_gphotos_work\\deleted_trash</code>)으로 옮기려면 교체 모드를 선택하세요.</span>
      <button class="btn btn-accent" id="switchReplaceBtn">📦 원본 백업 후 교체 모드로 전환</button>
    `;
    document.getElementById('switchReplaceBtn').onclick = () => {
      const r = document.querySelector('input[name="convMode"][value="replace-backup"]');
      if (r) { r.checked = true; loadPreview(); }
    };
  } else {
    delWarn.style.display = 'none';
  }

  document.getElementById('previewStats').innerHTML = `
    <div class="stat-chip">실행 예정 작업 <strong>${s.total_will_run}</strong>건</div>
    <div class="stat-chip">영상 변환 <strong>${s.standalone_video_count}</strong>개</div>
    <div class="stat-chip">이미지 회전/변환 <strong>${s.standalone_image_count}</strong>개</div>
    <div class="stat-chip">병합 그룹 <strong>${s.merge_group_count}</strong>개 (${s.merge_clip_count}클립)</div>
    <div class="stat-chip">휴지통 이동(삭제) <strong>${s.delete_will_run_count}</strong>개</div>
    <div class="stat-chip">변경 없이 유지 <strong>${s.keep_count}</strong>개</div>
  `;

  const list = document.getElementById('previewList');
  if (!data.items || data.items.length === 0) {
    list.innerHTML = '<div style="padding:20px;color:#9aa4b8;">변환·회전·병합·삭제할 작업이 없습니다. 2단계에서 항목을 편집해 주세요.</div>';
    document.getElementById('startConvertBtn').disabled = true;
    return;
  }
  document.getElementById('startConvertBtn').disabled = s.total_will_run === 0;
  list.innerHTML = data.items.map(it => {
    const kindLabel =
      it.kind === 'merge' ? '🔗 병합' :
      it.kind === 'image' ? '🖼 이미지' :
      it.kind === 'delete' ? '🗑 휴지통 이동' : '🎬 영상';
    const badgeClass = (it.kind === 'merge' || it.kind === 'delete') ? 'warn' : 'ok';
    return `
      <div class="preview-row">
        <div>
          <span class="tag ${badgeClass}" style="margin-right:6px;">${kindLabel}</span>
          <strong>${escapeHtml(it.src)}</strong> → <code>${escapeHtml(it.dst)}</code>
        </div>
        <div style="display:flex;gap:6px;align-items:center;">
          <span class="tag">${escapeHtml(it.action_label)}</span>
          <span class="tag ${it.will_run ? 'ok' : ''}">${it.will_run ? '실행 예정' : escapeHtml(it.skip_reason)}</span>
        </div>
      </div>
    `;
  }).join('');
}

async function startConvert(retryFailedOverride) {
  const mode = selectedMode();
  if (mode === 'replace-backup') {
    if (!confirm('원본 백업 후 교체 모드입니다.\n변환/회전 원본은 _gphotos_work/original_backup 으로, 삭제 표시 파일은 _gphotos_work/deleted_trash 로 이동하고, 정리된 결과물이 원래 폴더 자리에 저장됩니다.\n계속 진행할까요?')) {
      return;
    }
  }
  const retryFlag = retryFailedOverride !== undefined ? retryFailedOverride : document.getElementById('optRetryFailed').checked;
  const res = await fetch('/api/convert/start', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      mode,
      crf: Number(document.getElementById('optCrf').value || 20),
      preset: 'medium',
      fix_missing_dates: document.getElementById('optFixDates').checked,
      retry_failed: retryFlag
    })
  });
  const data = await res.json();
  if (!res.ok) { alert(data.error || '작업 시작 실패'); return; }
  appState.convert_job = { status: 'running' };
  updateStepperLocks();
  goStep(4);
  pollConvertStatus();
}

/* STEP 4 FUNCTIONS */
async function pollConvertStatus() {
  if (convertPollTimer) clearTimeout(convertPollTimer);
  const res = await fetch('/api/convert/status');
  const st = await res.json();
  appState.convert_job = st;

  const pct = st.total_units > 0 ? Math.round((st.done_units / st.total_units) * 100) : (st.status === 'done' ? 100 : 0);
  document.getElementById('convertProgressFill').style.width = pct + '%';
  document.getElementById('convertTitle').textContent =
    st.status === 'running' ? `작업 진행 중... (${pct}%)` :
    st.status === 'done' ? '✅ 정리 및 변환 작업이 완료되었습니다!' :
    st.status === 'cancelled' ? '⏹ 작업이 안전하게 중단되었습니다' :
    st.status === 'error' ? '❌ 작업 중 오류가 발생했습니다' : '대기 중';

  document.getElementById('convertCurrentLabel').textContent = st.current_label || '';
  document.getElementById('convertCounters').innerHTML = `
    <div class="stat-chip">변환/회전 성공 <strong>${st.converted_cnt}</strong></div>
    <div class="stat-chip">병합 성공 <strong>${st.merged_cnt}</strong></div>
    <div class="stat-chip">휴지통 이동 <strong>${st.deleted_cnt || 0}</strong></div>
    <div class="stat-chip">건너뜀 <strong>${st.skipped_cnt}</strong></div>
    <div class="stat-chip">실패 <strong>${st.failed_cnt}</strong></div>
  `;

  const box = document.getElementById('convertLogBox');
  box.innerHTML = (st.logs || []).map(l => `<div class="log-line-${l.level || 'info'}">${escapeHtml(l.text)}</div>`).join('');
  box.scrollTop = box.scrollHeight;

  const running = st.status === 'running';
  document.getElementById('cancelConvertBtn').style.display = running ? 'inline-block' : 'none';
  document.getElementById('convertDoneActions').style.display = running ? 'none' : 'flex';
  document.getElementById('openBackupFolderBtn').style.display = st.backup_folder ? 'inline-block' : 'none';
  document.getElementById('openTrashFolderBtn').style.display = st.trash_folder ? 'inline-block' : 'none';
  document.getElementById('retryFailedBtn').style.display = st.failed_cnt > 0 ? 'inline-block' : 'none';

  if (running) {
    convertPollTimer = setTimeout(pollConvertStatus, 600);
  } else {
    updateStepperLocks();
    refreshPlan();
  }
}

async function openFolderTarget(target) {
  const res = await fetch('/api/open-folder', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ target })
  });
  if (!res.ok) {
    const d = await res.json();
    alert(d.error || '폴더 열기 실패');
  }
}

async function init() {
  renderFilters();
  document.querySelectorAll('.step-tab').forEach(tab => {
    tab.onclick = () => { if (!tab.disabled) goStep(Number(tab.dataset.step)); };
  });
  document.getElementById('recheckEnvBtn').onclick = checkEnv;
  document.getElementById('browseBtn').onclick = browseFolder;
  document.getElementById('startScanBtn').onclick = startScan;
  document.getElementById('cancelScanBtn').onclick = () => fetch('/api/scan/cancel', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: '{}' });
  document.getElementById('sourceDirInput').addEventListener('change', checkFolderInput);
  document.getElementById('sourceDirInput').addEventListener('blur', checkFolderInput);

  document.getElementById('searchInput').addEventListener('input', e => {
    searchQuery = e.target.value.trim();
    currentPage = 1;
    renderGrid();
  });
  document.getElementById('batchDeleteBtn').onclick = () => batchUpdateItems(selectedIds, { delete: true, reviewed: true });
  document.getElementById('batchUndeleteBtn').onclick = () => batchUpdateItems(selectedIds, { delete: false, reviewed: true });
  document.getElementById('createMergeBtn').onclick = createMergeGroup;
  document.getElementById('clearSelectBtn').onclick = () => {
    selectedIds = [];
    lastClickedId = null;
    updateSelectionUI();
    renderGrid();
  };
  document.getElementById('autoMergeBtn').onclick = autoDetectMerge;
  document.getElementById('goStep3Btn').onclick = () => goStep(3);
  document.getElementById('backToStep2Btn').onclick = () => goStep(2);
  document.getElementById('refreshPreviewBtn').onclick = loadPreview;
  document.querySelectorAll('input[name="convMode"]').forEach(r => r.onchange = loadPreview);
  document.getElementById('optFixDates').onchange = loadPreview;
  document.getElementById('optRetryFailed').onchange = loadPreview;
  document.getElementById('startConvertBtn').onclick = () => startConvert();
  document.getElementById('cancelConvertBtn').onclick = () => fetch('/api/convert/cancel', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: '{}' });
  document.getElementById('openOutputFolderBtn').onclick = () => openFolderTarget('output');
  document.getElementById('openBackupFolderBtn').onclick = () => openFolderTarget('backup');
  document.getElementById('openTrashFolderBtn').onclick = () => openFolderTarget('trash');
  document.getElementById('step4BackEditBtn').onclick = () => goStep(2);
  document.getElementById('retryFailedBtn').onclick = () => startConvert(true);

  document.getElementById('modalCloseBtn').onclick = closePreviewModal;
  document.getElementById('previewModal').onclick = (e) => {
    if (e.target.id === 'previewModal' || e.target.classList.contains('modal-stage')) closePreviewModal();
  };
  document.getElementById('modalPrevBtn').onclick = () => navigateModal(-1);
  document.getElementById('modalNextBtn').onclick = () => navigateModal(1);
  document.getElementById('modalRotCCW').onclick = () => rotateModalBy(270);
  document.getElementById('modalRotCW').onclick = () => rotateModalBy(90);
  document.getElementById('modalRot180').onclick = () => {
    const it = (plan.items || []).find(x => x.id === modalItemId);
    if (it) updateItem(it.id, { user_rotation: it.user_rotation === 180 ? 0 : 180, reviewed: true });
  };
  document.getElementById('modalRotReset').onclick = () => {
    if (modalItemId) updateItem(modalItemId, { user_rotation: 0, reviewed: true });
  };
  document.getElementById('modalDelBtn').onclick = () => {
    const it = (plan.items || []).find(x => x.id === modalItemId);
    if (it) updateItem(it.id, { delete: !it.delete, reviewed: true });
  };
  document.getElementById('modalRethumbBtn').onclick = () => { if (modalItemId) rethumb(modalItemId); };
  document.getElementById('modalOpenOsBtn').onclick = () => { if (modalItemId) openVideo(modalItemId); };

  document.addEventListener('keydown', (e) => {
    if (!modalItemId) return;
    const tag = (e.target && e.target.tagName) ? e.target.tagName.toUpperCase() : '';
    if (tag === 'INPUT' || tag === 'TEXTAREA' || tag === 'SELECT') return;
    if (e.key === 'Escape') { closePreviewModal(); }
    else if (e.key === 'ArrowLeft') { e.preventDefault(); navigateModal(-1); }
    else if (e.key === 'ArrowRight') { e.preventDefault(); navigateModal(1); }
    else if (e.key === 'l' || e.key === 'L' || e.key === '[') { rotateModalBy(270); }
    else if (e.key === 'r' || e.key === 'R' || e.key === ']') { rotateModalBy(90); }
    else if (e.key === '0') { updateItem(modalItemId, { user_rotation: 0, reviewed: true }); }
    else if (e.key === 'd' || e.key === 'D' || e.key === 'Delete' || e.key === 'Backspace') {
      e.preventDefault();
      const it = (plan.items || []).find(x => x.id === modalItemId);
      if (it) updateItem(it.id, { delete: !it.delete, reviewed: true });
    }
  });

  await checkEnv();
  const stRes = await fetch('/api/state');
  appState = await stRes.json();
  if (appState.source_dir) {
    document.getElementById('sourceDirInput').value = appState.source_dir;
  }
  updateStepperLocks();
  updateSelectionUI();

  if (appState.convert_job && appState.convert_job.status === 'running') {
    goStep(4);
    pollConvertStatus();
  } else if (appState.scan_job && appState.scan_job.status === 'running') {
    goStep(1);
    pollScanStatus();
  } else if (appState.initial_step === 2 && appState.has_plan) {
    goStep(2);
  } else {
    goStep(1);
    if (appState.source_dir) checkFolderInput();
  }
}
init();
</script>
</body>
</html>
"""


def open_path_in_os(target: Path) -> None:
    """Open a local file or folder in the OS default application."""
    if sys.platform == "win32":
        os.startfile(str(target))  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(target)])
    else:
        subprocess.Popen(["xdg-open", str(target)])


class AppContext:
    """Thread-safe shared state for the 4-step All-in-One web app server."""

    def __init__(
        self,
        work_dir: Path | None = None,
        ffmpeg_bin: str | None = None,
        ffprobe_bin: str | None = None,
        initial_step: int = 1,
    ) -> None:
        self.lock = threading.Lock()
        self.work_dir: Path | None = work_dir
        self.ffmpeg_bin = ffmpeg_bin
        self.ffprobe_bin = ffprobe_bin
        self.initial_step = initial_step
        self.plan: dict[str, Any] = {"items": [], "merge_groups": []}
        self.by_id: dict[str, dict[str, Any]] = {}

        if work_dir is not None and (work_dir / "plan.json").exists():
            self.load_work_dir_locked(work_dir)

        self.scan_cancel = threading.Event()
        self.scan_state: dict[str, Any] = {
            "status": "idle",
            "done": 0,
            "total": 0,
            "current_file": "",
            "work_dir": str(work_dir) if work_dir else "",
            "error": "",
            "logs": [],
        }

        self.convert_cancel = threading.Event()
        self.convert_state: dict[str, Any] = {
            "status": "idle",
            "done_units": 0,
            "total_units": 0,
            "current_label": "",
            "converted_cnt": 0,
            "merged_cnt": 0,
            "deleted_cnt": 0,
            "skipped_cnt": 0,
            "failed_cnt": 0,
            "cancelled": False,
            "output_folder": "",
            "backup_folder": "",
            "trash_folder": "",
            "error": "",
            "logs": [],
        }

    @property
    def plan_path(self) -> Path | None:
        return (self.work_dir / "plan.json") if self.work_dir else None

    @property
    def thumbs_dir(self) -> Path | None:
        return (self.work_dir / "thumbs") if self.work_dir else None

    @property
    def previews_dir(self) -> Path | None:
        return (self.work_dir / "previews") if self.work_dir else None

    def load_work_dir_locked(self, work_dir: Path) -> None:
        plan = load_plan(work_dir)
        plan.setdefault("merge_groups", [])
        self.work_dir = work_dir
        self.plan = plan
        self.by_id = {it["id"]: it for it in plan.get("items", [])}

    def save_locked(self) -> None:
        if self.plan_path is not None:
            atomic_write_json(self.plan_path, self.plan)

    def is_busy_locked(self) -> bool:
        return (
            self.scan_state["status"] == "running"
            or self.convert_state["status"] == "running"
        )

    def append_scan_log_locked(self, text: str, level: str = "info") -> None:
        logs = self.scan_state["logs"]
        logs.append({"text": text, "level": level})
        if len(logs) > 300:
            del logs[:-300]

    def append_convert_log_locked(self, text: str, level: str = "info") -> None:
        logs = self.convert_state["logs"]
        logs.append({"text": text, "level": level})
        if len(logs) > 400:
            del logs[:-400]


def make_app_handler(ctx: AppContext):
    """Create an HTTP RequestHandler bound to AppContext."""

    class AppHandler(BaseHTTPRequestHandler):
        def log_message(self, fmt: str, *args: Any) -> None:
            return

        def _send_json(self, code: int, payload: Any) -> None:
            body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
            self.send_response(code)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            self.wfile.write(body)

        def _read_json_body(self) -> dict[str, Any]:
            length = int(self.headers.get("Content-Length") or 0)
            raw = self.rfile.read(length) if length > 0 else b"{}"
            return json.loads(raw.decode("utf-8"))

        def do_GET(self) -> None:
            parsed = urlparse(self.path)
            if parsed.path == "/":
                body = APP_HTML.encode("utf-8")
                self.send_response(HTTPStatus.OK)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.send_header("Content-Length", str(len(body)))
                self.end_headers()
                self.wfile.write(body)
                return

            if parsed.path == "/api/env":
                self._send_json(HTTPStatus.OK, get_env_status())
                return

            if parsed.path == "/api/state":
                with ctx.lock:
                    self._send_json(HTTPStatus.OK, {
                        "has_plan": bool(ctx.work_dir and (ctx.work_dir / "plan.json").exists()),
                        "source_dir": ctx.plan.get("source_dir", ""),
                        "work_dir": str(ctx.work_dir) if ctx.work_dir else "",
                        "initial_step": ctx.initial_step,
                        "scan_job": dict(ctx.scan_state),
                        "convert_job": dict(ctx.convert_state),
                    })
                return

            if parsed.path == "/api/plan":
                with ctx.lock:
                    self._send_json(HTTPStatus.OK, ctx.plan)
                return

            if parsed.path == "/api/scan/status":
                with ctx.lock:
                    self._send_json(HTTPStatus.OK, dict(ctx.scan_state))
                return

            if parsed.path == "/api/convert/status":
                with ctx.lock:
                    self._send_json(HTTPStatus.OK, dict(ctx.convert_state))
                return

            if parsed.path.startswith("/thumbs/"):
                fname = parsed.path[len("/thumbs/"):]
                if not THUMB_FILENAME_RE.match(fname):
                    self.send_error(HTTPStatus.FORBIDDEN, "Invalid thumbnail filename")
                    return
                with ctx.lock:
                    tdir = ctx.thumbs_dir
                if tdir is None:
                    self.send_error(HTTPStatus.NOT_FOUND, "No project loaded")
                    return
                thumbs_root = tdir.resolve()
                target = (thumbs_root / fname).resolve()
                if target.parent != thumbs_root or not target.is_file():
                    self.send_error(HTTPStatus.NOT_FOUND, "Thumbnail not found")
                    return
                data = target.read_bytes()
                self.send_response(HTTPStatus.OK)
                self.send_header("Content-Type", "image/jpeg")
                self.send_header("Content-Length", str(len(data)))
                self.send_header("Cache-Control", "no-cache")
                self.end_headers()
                self.wfile.write(data)
                return

            if parsed.path.startswith("/media/") or parsed.path.startswith("/preview/"):
                is_preview = parsed.path.startswith("/preview/")
                item_id = parsed.path[len("/preview/" if is_preview else "/media/"):]
                if not re.match(r"^[A-Za-z0-9_-]+$", item_id):
                    self.send_error(HTTPStatus.FORBIDDEN, "Invalid item id")
                    return
                with ctx.lock:
                    item = ctx.by_id.get(item_id)
                    src_root_str = ctx.plan.get("source_dir", "")
                    wdir = ctx.work_dir
                    tdir = ctx.thumbs_dir
                    pdir = ctx.previews_dir
                    ffmpeg_bin = ctx.ffmpeg_bin or "ffmpeg"
                if not item or not src_root_str:
                    self.send_error(HTTPStatus.NOT_FOUND, "Item not found")
                    return
                source_root = Path(src_root_str).resolve()
                target = (source_root / Path(item["rel_path"])).resolve()
                if not target.is_relative_to(source_root):
                    self.send_error(HTTPStatus.FORBIDDEN, "Invalid path")
                    return

                media_type = item.get("media_type", "video")
                if is_preview and media_type == "video":
                    if tdir is None:
                        self.send_error(HTTPStatus.NOT_FOUND, "No thumbnail")
                        return
                    thumb_file = (tdir / f"{item_id}.jpg").resolve()
                    if not thumb_file.is_file():
                        self.send_error(HTTPStatus.NOT_FOUND, "Thumbnail not found")
                        return
                    data = thumb_file.read_bytes()
                    self.send_response(HTTPStatus.OK)
                    self.send_header("Content-Type", "image/jpeg")
                    self.send_header("Content-Length", str(len(data)))
                    self.send_header("Cache-Control", "no-cache")
                    self.end_headers()
                    self.wfile.write(data)
                    return

                if not target.is_file() and wdir is not None:
                    if item.get("status") == "deleted":
                        trash_rel = item.get("trash_rel_path") or item["rel_path"]
                        trash_root = (wdir / "deleted_trash").resolve()
                        trash_cand = (trash_root / Path(trash_rel)).resolve()
                        if trash_cand.is_relative_to(trash_root) and trash_cand.is_file():
                            target = trash_cand
                    else:
                        bak_root = (wdir / "original_backup").resolve()
                        bak_cand = (bak_root / Path(item["rel_path"])).resolve()
                        if bak_cand.is_relative_to(bak_root) and bak_cand.is_file():
                            target = bak_cand

                if not target.is_file():
                    self.send_error(HTTPStatus.NOT_FOUND, "Media file not found")
                    return

                ext_lower = target.suffix.lower()
                if is_preview and media_type == "image" and ext_lower not in BROWSER_NATIVE_IMAGE_EXTENSIONS:
                    if pdir is None:
                        self.send_error(HTTPStatus.NOT_FOUND, "No preview dir")
                        return
                    preview_file = (pdir / f"{item_id}.jpg").resolve()
                    if not preview_file.is_file() or preview_file.stat().st_size == 0:
                        ok = extract_image_thumbnail(
                            target,
                            preview_file,
                            ffmpeg_bin=ffmpeg_bin,
                            max_dim=1600,
                            quality=3,
                        )
                        if not ok:
                            self.send_error(HTTPStatus.INTERNAL_SERVER_ERROR, "Preview generation failed")
                            return
                    data = preview_file.read_bytes()
                    self.send_response(HTTPStatus.OK)
                    self.send_header("Content-Type", "image/jpeg")
                    self.send_header("Content-Length", str(len(data)))
                    self.send_header("Cache-Control", "private, max-age=300")
                    self.end_headers()
                    self.wfile.write(data)
                    return

                mime = IMAGE_MIME_TYPES.get(ext_lower, "application/octet-stream")
                data = target.read_bytes()
                self.send_response(HTTPStatus.OK)
                self.send_header("Content-Type", mime)
                self.send_header("Content-Length", str(len(data)))
                self.send_header("Cache-Control", "private, max-age=300")
                self.end_headers()
                self.wfile.write(data)
                return

            self.send_error(HTTPStatus.NOT_FOUND)

        def do_POST(self) -> None:
            parsed = urlparse(self.path)
            try:
                body = self._read_json_body()
            except Exception:
                self._send_json(HTTPStatus.BAD_REQUEST, {"error": "Invalid JSON"})
                return

            if parsed.path == "/api/check-folder":
                raw_src = str(body.get("source_dir") or "").strip()
                if not raw_src:
                    self._send_json(HTTPStatus.OK, {"exists": False, "has_existing_plan": False})
                    return
                src = Path(raw_src).expanduser().resolve()
                if not src.is_dir():
                    self._send_json(HTTPStatus.OK, {"exists": False, "has_existing_plan": False})
                    return
                wdir = src / "_gphotos_work"
                ppath = wdir / "plan.json"
                has_plan = ppath.is_file()
                cnt = 0
                if has_plan:
                    try:
                        cnt = len(load_plan(wdir).get("items", []))
                    except Exception:
                        has_plan = False
                self._send_json(HTTPStatus.OK, {
                    "exists": True,
                    "source_dir": str(src),
                    "work_dir": str(wdir),
                    "has_existing_plan": has_plan,
                    "existing_count": cnt,
                })
                return

            if parsed.path == "/api/browse-folder":
                try:
                    import tkinter as tk
                    from tkinter import filedialog
                    root = tk.Tk()
                    root.withdraw()
                    root.attributes("-topmost", True)
                    chosen = filedialog.askdirectory(title="영상 폴더 선택")
                    root.destroy()
                    if chosen:
                        self._send_json(HTTPStatus.OK, {"ok": True, "path": str(Path(chosen))})
                    else:
                        self._send_json(HTTPStatus.OK, {"ok": False})
                except Exception:
                    self._send_json(HTTPStatus.OK, {
                        "ok": False,
                        "error": "폴더 선택 창을 열 수 없습니다. 경로를 직접 입력해 주세요.",
                    })
                return

            if parsed.path == "/api/load-existing":
                raw_work = str(body.get("work_dir") or "").strip()
                raw_src = str(body.get("source_dir") or "").strip()
                if raw_work:
                    wdir = Path(raw_work).expanduser().resolve()
                elif raw_src:
                    wdir = Path(raw_src).expanduser().resolve() / "_gphotos_work"
                else:
                    self._send_json(HTTPStatus.BAD_REQUEST, {"error": "폴더 경로가 필요합니다."})
                    return
                with ctx.lock:
                    if ctx.is_busy_locked():
                        self._send_json(HTTPStatus.CONFLICT, {"error": "작업이 실행 중입니다."})
                        return
                    try:
                        ctx.load_work_dir_locked(wdir)
                    except FileNotFoundError:
                        self._send_json(HTTPStatus.NOT_FOUND, {"error": "기존 plan.json을 찾을 수 없습니다."})
                        return
                    self._send_json(HTTPStatus.OK, {
                        "ok": True,
                        "source_dir": ctx.plan.get("source_dir", ""),
                        "work_dir": str(wdir),
                    })
                return

            if parsed.path == "/api/scan/start":
                raw_src = str(body.get("source_dir") or "").strip()
                if not raw_src:
                    self._send_json(HTTPStatus.BAD_REQUEST, {"error": "영상 폴더 경로를 입력하세요."})
                    return
                src_dir = Path(raw_src).expanduser().resolve()
                if not src_dir.is_dir():
                    self._send_json(HTTPStatus.BAD_REQUEST, {"error": f"폴더를 찾을 수 없습니다: {src_dir}"})
                    return
                raw_work = str(body.get("work_dir") or "").strip()
                wdir = Path(raw_work).expanduser().resolve() if raw_work else (src_dir / "_gphotos_work")
                workers = max(1, min(16, int(body.get("workers") or 4)))
                include_images = bool(body.get("include_images", True))

                try:
                    ffmpeg_bin, ffprobe_bin = check_ffmpeg_tools()
                except SystemExit:
                    self._send_json(HTTPStatus.BAD_REQUEST, {"error": "FFmpeg/FFprobe가 설치되어 있지 않습니다."})
                    return

                with ctx.lock:
                    if ctx.is_busy_locked():
                        self._send_json(HTTPStatus.CONFLICT, {"error": "이미 다른 작업이 진행 중입니다."})
                        return
                    ctx.ffmpeg_bin = ffmpeg_bin
                    ctx.ffprobe_bin = ffprobe_bin
                    ctx.scan_cancel.clear()
                    ctx.scan_state = {
                        "status": "running",
                        "done": 0,
                        "total": 0,
                        "current_file": "",
                        "work_dir": str(wdir),
                        "error": "",
                        "logs": [{"text": f"스캔 시작: {src_dir} (이미지 포함: {'예' if include_images else '아니오'})", "level": "info"}],
                    }

                def scan_worker() -> None:
                    def on_prog(done: int, total: int, msg: str) -> None:
                        with ctx.lock:
                            ctx.scan_state["done"] = done
                            ctx.scan_state["total"] = total
                            ctx.scan_state["current_file"] = msg
                            if done == 0:
                                ctx.append_scan_log_locked(msg, "info")
                            elif done % 10 == 0 or done == total:
                                ctx.append_scan_log_locked(f"[{done}/{total}] {msg}", "info")

                    try:
                        plan_data = run_scan_job(
                            source_dir=src_dir,
                            work_dir=wdir,
                            workers=workers,
                            ffmpeg_bin=ffmpeg_bin,
                            ffprobe_bin=ffprobe_bin,
                            progress_cb=on_prog,
                            cancel_event=ctx.scan_cancel,
                            include_images=include_images,
                        )
                        with ctx.lock:
                            ctx.load_work_dir_locked(wdir)
                            ctx.scan_state["status"] = "done"
                            loaded_items = plan_data.get("items", [])
                            v_cnt = sum(1 for x in loaded_items if x.get("media_type", "video") == "video")
                            i_cnt = sum(1 for x in loaded_items if x.get("media_type") == "image")
                            ctx.append_scan_log_locked(
                                f"스캔 완료: 총 {len(loaded_items)}개 (영상 {v_cnt} · 이미지 {i_cnt})", "ok"
                            )
                    except JobCancelled:
                        with ctx.lock:
                            ctx.scan_state["status"] = "cancelled"
                            ctx.append_scan_log_locked("스캔이 중단되었습니다.", "warn")
                    except Exception as exc:
                        with ctx.lock:
                            ctx.scan_state["status"] = "error"
                            ctx.scan_state["error"] = str(exc)
                            ctx.append_scan_log_locked(f"오류: {exc}", "err")

                threading.Thread(target=scan_worker, daemon=True).start()
                self._send_json(HTTPStatus.OK, {"ok": True, "work_dir": str(wdir)})
                return

            if parsed.path == "/api/scan/cancel":
                ctx.scan_cancel.set()
                self._send_json(HTTPStatus.OK, {"ok": True})
                return

            if parsed.path == "/api/convert/preview":
                with ctx.lock:
                    if ctx.work_dir is None:
                        self._send_json(HTTPStatus.BAD_REQUEST, {"error": "스캔된 작업이 없습니다."})
                        return
                    plan_snap = copy.deepcopy(ctx.plan)
                    wdir = ctx.work_dir
                mode = str(body.get("mode") or "output")
                if mode not in ("output", "replace-backup"):
                    mode = "output"
                fix_dates = bool(body.get("fix_missing_dates", False))
                retry_failed = bool(body.get("retry_failed", False))
                planned = plan_convert_tasks(
                    plan_snap,
                    source_dir=Path(plan_snap["source_dir"]),
                    output_dir=wdir / "converted",
                    mode=mode,
                    fix_missing_dates=fix_dates,
                    retry_failed=retry_failed,
                )
                preview_rows = []
                for t in planned["standalone"]:
                    it = t["item"]
                    rot = int(it.get("user_rotation", 0))
                    is_img = it.get("media_type", "video") == "image"
                    if t["action"] == "image_process":
                        label = f"이미지 회전 ({rot}°)" if rot else "이미지 변환"
                        if rot and it.get("needs_convert"):
                            label = f"이미지 변환 + 회전 ({rot}°)"
                    elif t["action"] == "encode":
                        label = f"재인코딩 (회전 {rot}°)" if rot else "MP4 재인코딩"
                    else:
                        label = "날짜 메타 보완 (무손실)"
                    preview_rows.append({
                        "kind": "image" if is_img else "standalone",
                        "src": it["rel_path"],
                        "dst": t["out_rel"].as_posix(),
                        "action_label": label,
                        "will_run": t["will_run"],
                        "skip_reason": t["skip_reason"],
                    })
                for m in planned["merge"]:
                    grp = m["group"]
                    g_items = m["items"]
                    rots = sorted({int(x.get("user_rotation", 0)) for x in g_items if x.get("user_rotation", 0)})
                    rot_txt = f" · 회전 {'/'.join(str(r) + '°' for r in rots)}" if rots else ""
                    preview_rows.append({
                        "kind": "merge",
                        "src": f"{grp['name']} ({len(g_items)}개 클립)",
                        "dst": m["out_rel"].as_posix(),
                        "action_label": f"병합 인코딩{rot_txt}",
                        "will_run": m["will_run"],
                        "skip_reason": m["skip_reason"],
                    })
                for d in planned["deleted"]:
                    it = d["item"]
                    preview_rows.append({
                        "kind": "delete",
                        "src": it["rel_path"],
                        "dst": f"_gphotos_work/deleted_trash/{d['trash_rel'].as_posix()}",
                        "action_label": "휴지통 이동 (안전 삭제)",
                        "will_run": d["will_run"],
                        "skip_reason": d["skip_reason"],
                    })
                self._send_json(HTTPStatus.OK, {
                    "ok": True,
                    "summary": planned["summary"],
                    "items": preview_rows,
                })
                return

            if parsed.path == "/api/convert/start":
                mode = str(body.get("mode") or "output")
                if mode not in ("output", "replace-backup"):
                    self._send_json(HTTPStatus.BAD_REQUEST, {"error": "Invalid mode"})
                    return
                crf = int(body.get("crf") or 20)
                preset = str(body.get("preset") or "medium")
                fix_dates = bool(body.get("fix_missing_dates", False))
                retry_failed = bool(body.get("retry_failed", False))

                try:
                    ffmpeg_bin, ffprobe_bin = check_ffmpeg_tools()
                except SystemExit:
                    self._send_json(HTTPStatus.BAD_REQUEST, {"error": "FFmpeg/FFprobe가 설치되어 있지 않습니다."})
                    return

                with ctx.lock:
                    if ctx.work_dir is None:
                        self._send_json(HTTPStatus.BAD_REQUEST, {"error": "먼저 영상을 스캔해 주세요."})
                        return
                    if ctx.is_busy_locked():
                        self._send_json(HTTPStatus.CONFLICT, {"error": "이미 다른 작업이 진행 중입니다."})
                        return
                    wdir = ctx.work_dir
                    src_dir = Path(ctx.plan["source_dir"])
                    out_folder = str((wdir / "converted") if mode == "output" else src_dir)
                    bak_folder = str(wdir / "original_backup") if mode == "replace-backup" else ""
                    ctx.convert_cancel.clear()
                    ctx.convert_state = {
                        "status": "running",
                        "done_units": 0,
                        "total_units": 0,
                        "current_label": "변환 준비 중...",
                        "converted_cnt": 0,
                        "merged_cnt": 0,
                        "deleted_cnt": 0,
                        "skipped_cnt": 0,
                        "failed_cnt": 0,
                        "cancelled": False,
                        "output_folder": out_folder,
                        "backup_folder": bak_folder,
                        "trash_folder": "",
                        "error": "",
                        "logs": [{"text": f"작업 시작 (모드: {mode}, CRF: {crf})", "level": "info"}],
                    }

                def convert_worker() -> None:
                    def on_prog(ev: dict[str, Any]) -> None:
                        with ctx.lock:
                            for k in ("done_units", "total_units", "current_label", "converted_cnt", "merged_cnt", "deleted_cnt", "skipped_cnt", "failed_cnt"):
                                if k in ev:
                                    ctx.convert_state[k] = ev[k]
                            if ev.get("log_line"):
                                ctx.append_convert_log_locked(ev["log_line"], ev.get("log_level", "info"))

                    try:
                        summary = run_convert_job(
                            work_dir=wdir,
                            mode=mode,
                            output_dir=None,
                            crf=crf,
                            preset=preset,
                            fix_missing_dates=fix_dates,
                            retry_failed=retry_failed,
                            ffmpeg_bin=ffmpeg_bin,
                            ffprobe_bin=ffprobe_bin,
                            progress_cb=on_prog,
                            cancel_event=ctx.convert_cancel,
                            plan_lock=ctx.lock,
                        )
                        with ctx.lock:
                            ctx.load_work_dir_locked(wdir)
                            ctx.convert_state["converted_cnt"] = summary["converted_cnt"]
                            ctx.convert_state["merged_cnt"] = summary["merged_cnt"]
                            ctx.convert_state["deleted_cnt"] = summary["deleted_cnt"]
                            ctx.convert_state["skipped_cnt"] = summary["skipped_cnt"]
                            ctx.convert_state["failed_cnt"] = summary["failed_cnt"]
                            ctx.convert_state["cancelled"] = summary["cancelled"]
                            ctx.convert_state["output_folder"] = summary["output_folder"]
                            ctx.convert_state["backup_folder"] = summary["backup_folder"]
                            ctx.convert_state["trash_folder"] = summary["trash_folder"]
                            ctx.convert_state["status"] = "cancelled" if summary["cancelled"] else "done"
                            ctx.convert_state["current_label"] = (
                                "안전하게 중단되었습니다." if summary["cancelled"] else "모든 정리/변환 작업 완료"
                            )
                    except Exception as exc:
                        with ctx.lock:
                            ctx.convert_state["status"] = "error"
                            ctx.convert_state["error"] = str(exc)
                            ctx.append_convert_log_locked(f"오류: {exc}", "err")

                threading.Thread(target=convert_worker, daemon=True).start()
                self._send_json(HTTPStatus.OK, {"ok": True})
                return

            if parsed.path == "/api/convert/cancel":
                ctx.convert_cancel.set()
                self._send_json(HTTPStatus.OK, {"ok": True})
                return

            if parsed.path == "/api/open-folder":
                target_kind = str(body.get("target") or "")
                if target_kind not in ("output", "backup", "trash", "source", "work"):
                    self._send_json(HTTPStatus.BAD_REQUEST, {"error": "Invalid target folder"})
                    return
                with ctx.lock:
                    if target_kind == "output":
                        raw = ctx.convert_state.get("output_folder") or (str(ctx.work_dir / "converted") if ctx.work_dir else "")
                    elif target_kind == "backup":
                        raw = ctx.convert_state.get("backup_folder") or (str(ctx.work_dir / "original_backup") if ctx.work_dir else "")
                    elif target_kind == "trash":
                        raw = ctx.convert_state.get("trash_folder") or (str(ctx.work_dir / "deleted_trash") if ctx.work_dir else "")
                    elif target_kind == "source":
                        raw = ctx.plan.get("source_dir", "")
                    else:
                        raw = str(ctx.work_dir) if ctx.work_dir else ""
                if not raw:
                    self._send_json(HTTPStatus.NOT_FOUND, {"error": "폴더 경로가 설정되어 있지 않습니다."})
                    return
                folder_path = Path(raw)
                if not folder_path.exists():
                    self._send_json(HTTPStatus.NOT_FOUND, {"error": f"폴더가 존재하지 않습니다: {folder_path}"})
                    return
                try:
                    open_path_in_os(folder_path)
                    self._send_json(HTTPStatus.OK, {"ok": True})
                except Exception as exc:
                    self._send_json(HTTPStatus.INTERNAL_SERVER_ERROR, {"error": str(exc)})
                return

            # Endpoints below require a loaded plan
            with ctx.lock:
                if ctx.work_dir is None:
                    self._send_json(HTTPStatus.BAD_REQUEST, {"error": "먼저 영상을 스캔해 주세요."})
                    return
                if ctx.is_busy_locked():
                    self._send_json(HTTPStatus.CONFLICT, {"error": "작업 진행 중에는 편집할 수 없습니다."})
                    return

            item_id = str(body.get("id") or "")

            def apply_item_patch_locked(item: dict[str, Any], patch: dict[str, Any]) -> None:
                if "delete" in patch:
                    del_flag = bool(patch["delete"])
                    if del_flag:
                        item["delete"] = True
                        new_groups: list[dict[str, Any]] = []
                        for grp in ctx.plan.get("merge_groups") or []:
                            rem = [x for x in (grp.get("item_ids") or []) if x != item["id"]]
                            if len(rem) >= 2:
                                if rem != grp.get("item_ids"):
                                    grp["item_ids"] = rem
                                    grp["output_rel_path"] = ""
                                    grp["status"] = "pending"
                                new_groups.append(grp)
                        ctx.plan["merge_groups"] = new_groups
                        if item.get("status") != "deleted":
                            item["status"] = "pending"
                    else:
                        item["delete"] = False
                        if item.get("status") == "deleted" and ctx.work_dir is not None:
                            trash_root = (ctx.work_dir / "deleted_trash").resolve()
                            trash_rel = item.get("trash_rel_path") or item["rel_path"]
                            src_trash = (trash_root / Path(trash_rel)).resolve()
                            source_root = Path(ctx.plan["source_dir"]).resolve()
                            dst_orig = (source_root / Path(item["rel_path"])).resolve()
                            if (
                                src_trash.is_relative_to(trash_root)
                                and dst_orig.is_relative_to(source_root)
                                and src_trash.is_file()
                                and not dst_orig.exists()
                            ):
                                dst_orig.parent.mkdir(parents=True, exist_ok=True)
                                shutil.move(str(src_trash), str(dst_orig))
                            item.pop("trash_rel_path", None)
                            item["status"] = "pending"

                if "user_rotation" in patch:
                    rot = int(patch["user_rotation"])
                    if rot not in VALID_ROTATIONS:
                        raise ValueError("Invalid rotation")
                    if rot != item.get("user_rotation"):
                        item["user_rotation"] = rot
                        if rot != 0:
                            item["convert"] = True
                        if item.get("status") == "done":
                            item["status"] = "pending"
                        for grp in ctx.plan.get("merge_groups") or []:
                            if item["id"] in (grp.get("item_ids") or []):
                                grp["status"] = "pending"

                if "convert" in patch:
                    new_conv = bool(patch["convert"])
                    if int(item.get("user_rotation", 0)) != 0:
                        new_conv = True
                    if new_conv != item.get("convert") and item.get("status") == "done":
                        item["status"] = "pending"
                    item["convert"] = new_conv

                if "reviewed" in patch:
                    item["reviewed"] = bool(patch["reviewed"])

            if parsed.path == "/api/update":
                with ctx.lock:
                    item = ctx.by_id.get(item_id)
                    if not item:
                        self._send_json(HTTPStatus.NOT_FOUND, {"error": "Item not found"})
                        return
                    try:
                        apply_item_patch_locked(item, body)
                    except ValueError as exc:
                        self._send_json(HTTPStatus.BAD_REQUEST, {"error": str(exc)})
                        return
                    ctx.save_locked()
                    self._send_json(HTTPStatus.OK, {
                        "ok": True,
                        "item": item,
                        "merge_groups": ctx.plan.get("merge_groups", []),
                    })
                return

            if parsed.path == "/api/batch-update":
                with ctx.lock:
                    raw_ids = body.get("item_ids") or []
                    patch: dict[str, Any] = {}
                    for k in ("delete", "user_rotation", "convert", "reviewed"):
                        if k in body:
                            patch[k] = body[k]
                    updated_items: list[dict[str, Any]] = []
                    seen: set[str] = set()
                    try:
                        for iid in raw_ids:
                            sid = str(iid)
                            if sid in seen:
                                continue
                            seen.add(sid)
                            it = ctx.by_id.get(sid)
                            if it:
                                apply_item_patch_locked(it, patch)
                                updated_items.append(it)
                    except ValueError as exc:
                        self._send_json(HTTPStatus.BAD_REQUEST, {"error": str(exc)})
                        return
                    ctx.save_locked()
                    self._send_json(HTTPStatus.OK, {
                        "ok": True,
                        "updated": len(updated_items),
                        "items": updated_items,
                        "merge_groups": ctx.plan.get("merge_groups", []),
                    })
                return

            if parsed.path == "/api/merge/create":
                with ctx.lock:
                    raw_ids = body.get("item_ids") or []
                    dedup_ids: list[str] = []
                    seen: set[str] = set()
                    for iid in raw_ids:
                        sid = str(iid)
                        it = ctx.by_id.get(sid)
                        if (
                            it is not None
                            and it.get("media_type", "video") == "video"
                            and not it.get("delete")
                            and sid not in seen
                        ):
                            seen.add(sid)
                            dedup_ids.append(sid)
                    if len(dedup_ids) < 2:
                        self._send_json(HTTPStatus.BAD_REQUEST, {"error": "병합은 삭제 예정이 아닌 동영상 2개 이상을 선택해야 합니다."})
                        return

                    new_groups: list[dict[str, Any]] = []
                    for grp in ctx.plan.get("merge_groups") or []:
                        rem = [x for x in (grp.get("item_ids") or []) if x not in seen]
                        if len(rem) >= 2:
                            grp["item_ids"] = rem
                            grp["output_rel_path"] = ""
                            grp["status"] = "pending"
                            new_groups.append(grp)

                    first_item = ctx.by_id[dedup_ids[0]]
                    group = {
                        "id": f"mg_{uuid.uuid4().hex[:8]}",
                        "name": sanitize_merged_filename(str(body.get("name") or ""), Path(first_item["filename"]).stem),
                        "item_ids": dedup_ids,
                        "output_rel_path": "",
                        "status": "pending",
                    }
                    new_groups.append(group)
                    ctx.plan["merge_groups"] = new_groups
                    ctx.save_locked()
                    self._send_json(HTTPStatus.OK, {"ok": True, "group": group, "merge_groups": new_groups})
                return

            if parsed.path == "/api/merge/update":
                with ctx.lock:
                    group_id = str(body.get("group_id") or "")
                    groups = ctx.plan.get("merge_groups") or []
                    target_grp = next((g for g in groups if g.get("id") == group_id), None)
                    if not target_grp:
                        self._send_json(HTTPStatus.NOT_FOUND, {"error": "Merge group not found"})
                        return

                    if "item_ids" in body:
                        dedup_ids = []
                        seen = set()
                        for iid in body["item_ids"] or []:
                            sid = str(iid)
                            if sid in ctx.by_id and sid not in seen:
                                seen.add(sid)
                                dedup_ids.append(sid)
                        if len(dedup_ids) < 2:
                            ctx.plan["merge_groups"] = [g for g in groups if g.get("id") != group_id]
                            ctx.save_locked()
                            self._send_json(HTTPStatus.OK, {"ok": True, "deleted": True, "merge_groups": ctx.plan["merge_groups"]})
                            return
                        target_grp["item_ids"] = dedup_ids
                        target_grp["output_rel_path"] = ""
                        target_grp["status"] = "pending"

                    if "name" in body:
                        first_item = ctx.by_id[target_grp["item_ids"][0]]
                        target_grp["name"] = sanitize_merged_filename(str(body["name"] or ""), Path(first_item["filename"]).stem)
                        target_grp["output_rel_path"] = ""
                        target_grp["status"] = "pending"

                    if "rotate_all" in body:
                        rot = int(body["rotate_all"])
                        if rot not in VALID_ROTATIONS:
                            self._send_json(HTTPStatus.BAD_REQUEST, {"error": "Invalid rotation"})
                            return
                        for iid in target_grp["item_ids"]:
                            it = ctx.by_id.get(iid)
                            if it:
                                it["user_rotation"] = rot
                                if rot != 0:
                                    it["convert"] = True
                        target_grp["status"] = "pending"

                    ctx.save_locked()
                    self._send_json(HTTPStatus.OK, {"ok": True, "group": target_grp, "merge_groups": ctx.plan["merge_groups"]})
                return

            if parsed.path == "/api/merge/delete":
                with ctx.lock:
                    group_id = str(body.get("group_id") or "")
                    groups = ctx.plan.get("merge_groups") or []
                    ctx.plan["merge_groups"] = [g for g in groups if g.get("id") != group_id]
                    ctx.save_locked()
                    self._send_json(HTTPStatus.OK, {"ok": True, "merge_groups": ctx.plan["merge_groups"]})
                return

            if parsed.path == "/api/merge/auto":
                with ctx.lock:
                    try:
                        max_gap = float(body.get("max_gap_sec", 15.0))
                    except (TypeError, ValueError):
                        max_gap = 15.0
                    max_gap = max(1.0, min(max_gap, 3600.0))
                    new_groups = auto_detect_merge_groups(ctx.plan, max_gap_sec=max_gap)
                    ctx.plan.setdefault("merge_groups", []).extend(new_groups)
                    ctx.save_locked()
                    self._send_json(HTTPStatus.OK, {
                        "ok": True,
                        "created": len(new_groups),
                        "merge_groups": ctx.plan["merge_groups"],
                    })
                return

            if parsed.path == "/api/rethumb":
                with ctx.lock:
                    item = ctx.by_id.get(item_id)
                    if not item or ctx.thumbs_dir is None:
                        self._send_json(HTTPStatus.NOT_FOUND, {"error": "Item not found"})
                        return
                    source_dir = Path(ctx.plan["source_dir"])
                    src_file = source_dir / Path(item["rel_path"])
                    next_idx = (int(item.get("thumb_offset_idx", 0)) + 1) % len(THUMB_OFFSETS)
                    ratio = THUMB_OFFSETS[next_idx]
                    thumb_path = ctx.thumbs_dir / f"{item_id}.jpg"
                    duration = float(item.get("duration") or 0.0)
                    ffmpeg_bin = ctx.ffmpeg_bin or "ffmpeg"
                ok = extract_thumbnail(
                    src_file,
                    thumb_path,
                    duration=duration,
                    offset_ratio=ratio,
                    ffmpeg_bin=ffmpeg_bin,
                )
                with ctx.lock:
                    if ok:
                        item["thumb_offset_idx"] = next_idx
                        item["thumb"] = f"thumbs/{item_id}.jpg"
                        ctx.save_locked()
                self._send_json(
                    HTTPStatus.OK,
                    {"ok": ok, "thumb_url": f"/thumbs/{item_id}.jpg?t={int(time.time() * 1000)}"},
                )
                return

            if parsed.path == "/api/open":
                with ctx.lock:
                    item = ctx.by_id.get(item_id)
                    if not item:
                        self._send_json(HTTPStatus.NOT_FOUND, {"error": "Item not found"})
                        return
                    src_file = (Path(ctx.plan["source_dir"]) / Path(item["rel_path"])).resolve()
                if not src_file.is_file():
                    self._send_json(HTTPStatus.NOT_FOUND, {"error": "Source file missing"})
                    return
                try:
                    open_path_in_os(src_file)
                    self._send_json(HTTPStatus.OK, {"ok": True})
                except Exception as exc:
                    self._send_json(HTTPStatus.INTERNAL_SERVER_ERROR, {"error": str(exc)})
                return

            self.send_error(HTTPStatus.NOT_FOUND)

    return AppHandler


def start_app_server(
    work_dir: Path | None = None,
    port: int = 8765,
    ffmpeg_bin: str | None = None,
    ffprobe_bin: str | None = None,
    initial_step: int = 1,
) -> tuple[ThreadingHTTPServer, int, AppContext]:
    """Start the 4-step All-in-One HTTP server on 127.0.0.1, auto-incrementing port if busy."""
    ctx = AppContext(
        work_dir=work_dir,
        ffmpeg_bin=ffmpeg_bin,
        ffprobe_bin=ffprobe_bin,
        initial_step=initial_step,
    )
    handler_cls = make_app_handler(ctx)
    last_err: Exception | None = None
    for candidate_port in range(port, port + 50):
        try:
            server = ThreadingHTTPServer(("127.0.0.1", candidate_port), handler_cls)
            return server, candidate_port, ctx
        except OSError as exc:
            last_err = exc
    raise RuntimeError(f"사용 가능한 포트를 찾지 못했습니다 ({port}~{port + 49}): {last_err}")


def start_review_server(
    work_dir: Path,
    port: int = 8765,
    ffmpeg_bin: str = "ffmpeg",
) -> tuple[ThreadingHTTPServer, int, AppContext]:
    """Backward-compatible wrapper starting the app server directly on Step 2."""
    return start_app_server(
        work_dir=work_dir,
        port=port,
        ffmpeg_bin=ffmpeg_bin,
        ffprobe_bin="ffprobe",
        initial_step=2,
    )


def _serve_web_app(
    work_dir: Path | None,
    port: int,
    no_browser: bool,
    initial_step: int,
    ffmpeg_bin: str | None = None,
    ffprobe_bin: str | None = None,
) -> int:
    try:
        server, bound_port, _ = start_app_server(
            work_dir=work_dir,
            port=port,
            ffmpeg_bin=ffmpeg_bin,
            ffprobe_bin=ffprobe_bin,
            initial_step=initial_step,
        )
    except Exception as exc:
        print(f"[오류] 웹 UI 서버 시작 실패: {exc}", file=sys.stderr)
        return 1

    url = f"http://127.0.0.1:{bound_port}"
    print(f"웹 마법사 실행 중: {url}")
    print("브라우저에서 작업을 마친 뒤 터미널에서 Ctrl+C를 눌러 종료하세요.")
    if not no_browser:
        webbrowser.open(url)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n서버를 종료합니다.")
    finally:
        server.server_close()
    return 0


def cmd_app(args: argparse.Namespace) -> int:
    """Execute the 'app' (4-step wizard) subcommand."""
    ffmpeg_bin, ffprobe_bin = locate_ffmpeg_tools()
    work_dir = Path(args.work_dir).expanduser().resolve() if args.work_dir else None
    initial_step = 2 if (work_dir is not None and (work_dir / "plan.json").exists()) else 1
    return _serve_web_app(
        work_dir=work_dir,
        port=int(args.port),
        no_browser=bool(args.no_browser),
        initial_step=initial_step,
        ffmpeg_bin=ffmpeg_bin,
        ffprobe_bin=ffprobe_bin,
    )


def cmd_review(args: argparse.Namespace) -> int:
    """Execute the 'review' subcommand (opens Step 2 directly)."""
    ffmpeg_bin, ffprobe_bin = check_ffmpeg_tools()
    work_dir = Path(args.work_dir).expanduser().resolve()
    return _serve_web_app(
        work_dir=work_dir,
        port=int(args.port),
        no_browser=bool(args.no_browser),
        initial_step=2,
        ffmpeg_bin=ffmpeg_bin,
        ffprobe_bin=ffprobe_bin,
    )


def resolve_output_rel_path(
    item: dict[str, Any],
    source_dir: Path,
    reserved_rel_outputs: set[str],
    mode: str,
) -> Path:
    """
    Compute collision-free relative output path (.mp4 for videos; .jpg/.png/.webp/.gif for images).
    """
    rel = Path(item["rel_path"])
    media_type = item.get("media_type", "video")
    if media_type == "image":
        src_ext = rel.suffix.lower()
        target_ext = src_ext if src_ext in BROWSER_NATIVE_IMAGE_EXTENSIONS else ".jpg"
        ext_tag = src_ext.lstrip(".") or "img"
    else:
        target_ext = ".mp4"
        ext_tag = rel.suffix.lstrip(".").lower() or "video"

    candidate = rel.with_suffix(target_ext)
    cand_key = candidate.as_posix().lower()

    src_same_name_conflict = (
        rel.suffix.lower() != target_ext
        and (source_dir / candidate).exists()
        and (source_dir / candidate).resolve() != (source_dir / rel).resolve()
    )

    if cand_key not in reserved_rel_outputs and not (mode == "replace-backup" and src_same_name_conflict):
        reserved_rel_outputs.add(cand_key)
        return candidate

    base_stem = rel.stem
    suffix_name = f"{base_stem}_from_{ext_tag}"
    candidate = rel.with_name(f"{suffix_name}{target_ext}")
    idx = 2
    while candidate.as_posix().lower() in reserved_rel_outputs or (
        mode == "replace-backup" and (source_dir / candidate).exists()
    ):
        candidate = rel.with_name(f"{suffix_name}_{idx}{target_ext}")
        idx += 1
    reserved_rel_outputs.add(candidate.as_posix().lower())
    return candidate


def verify_output_video(out_path: Path, ffprobe_bin: str = "ffprobe") -> None:
    """Verify converted MP4 exists, is non-empty, and has a readable video stream."""
    if not out_path.exists() or out_path.stat().st_size <= 0:
        raise RuntimeError("출력 파일이 비어 있거나 생성되지 않았습니다.")
    probe = probe_video(out_path, ffprobe_bin=ffprobe_bin)
    v_streams = [s for s in probe.get("streams") or [] if s.get("codec_type") == "video"]
    if not v_streams:
        raise RuntimeError("변환된 파일에서 비디오 스트림을 확인할 수 없습니다.")


def resolve_merge_output_rel_path(
    group: dict[str, Any],
    first_item: dict[str, Any],
    source_dir: Path,
    reserved_rel_outputs: set[str],
    mode: str,
    group_source_rel_paths: set[str],
) -> Path:
    """Compute collision-free relative .mp4 output path for a merge group."""
    first_rel = Path(first_item["rel_path"])
    parent_rel = first_rel.parent
    fname = sanitize_merged_filename(group.get("name", ""), first_rel.stem)
    candidate = parent_rel / fname

    def has_conflict(cand: Path) -> bool:
        key = cand.as_posix().lower()
        if key in reserved_rel_outputs:
            return True
        if mode == "replace-backup" and key not in group_source_rel_paths and (source_dir / cand).exists():
            return True
        return False

    if not has_conflict(candidate):
        reserved_rel_outputs.add(candidate.as_posix().lower())
        return candidate

    stem = Path(fname).stem
    idx = 2
    while True:
        alt = parent_rel / f"{stem}_{idx}.mp4"
        if not has_conflict(alt):
            reserved_rel_outputs.add(alt.as_posix().lower())
            return alt
        idx += 1


def plan_convert_tasks(
    plan: dict[str, Any],
    source_dir: Path,
    output_dir: Path,
    mode: str,
    fix_missing_dates: bool = False,
    retry_failed: bool = False,
) -> dict[str, Any]:
    """Compute standalone, merge, and trash-delete tasks without mutating plan on disk."""
    items = plan.get("items", [])
    by_id: dict[str, dict[str, Any]] = {it["id"]: it for it in items}
    valid_merge_ids = {
        it["id"] for it in items
        if it.get("media_type", "video") == "video" and not it.get("delete")
    }
    merge_groups = prune_merge_groups(plan.get("merge_groups", []), valid_merge_ids)
    merged_item_ids: set[str] = {iid for grp in merge_groups for iid in grp["item_ids"]}

    reserved_outputs: set[str] = set()
    if mode == "replace-backup":
        for item in items:
            if item.get("delete"):
                continue
            if item["id"] not in merged_item_ids and not item.get("convert") and int(item.get("user_rotation", 0)) == 0:
                if item.get("media_type", "video") == "image" or item.get("ext") == ".mp4":
                    reserved_outputs.add(item["rel_path"].lower())

    deleted_tasks: list[dict[str, Any]] = []
    standalone: list[dict[str, Any]] = []
    for item in items:
        if item.get("delete"):
            trash_rel = Path(item["rel_path"])
            src_exists = (source_dir / trash_rel).exists()
            if mode == "output":
                will_run = False
                skip_reason = "세이프 모드 (원본 유지 · 교체 모드에서 휴지통 이동)"
            elif item.get("status") == "deleted" and not src_exists:
                will_run = False
                skip_reason = "이미 휴지통 이동됨"
            else:
                will_run = True
                skip_reason = ""
            deleted_tasks.append({
                "item": item,
                "trash_rel": trash_rel,
                "will_run": will_run,
                "skip_reason": skip_reason,
            })
            continue

        if item["id"] in merged_item_ids:
            continue

        media_type = item.get("media_type", "video")
        if media_type == "image":
            should_process = bool(item.get("convert")) or int(item.get("user_rotation", 0)) != 0
            if not should_process:
                continue
            action = "image_process"
        else:
            should_encode = bool(item.get("convert")) or int(item.get("user_rotation", 0)) != 0
            should_fix_date = (
                not should_encode
                and bool(fix_missing_dates)
                and item.get("date_source") != "metadata"
                and item.get("ext") in {".mp4", ".mov", ".m4v"}
            )
            if not should_encode and not should_fix_date:
                continue
            action = "encode" if should_encode else "fix_date"

        out_rel = Path(item["output_rel_path"]) if item.get("output_rel_path") else resolve_output_rel_path(
            item, source_dir, reserved_outputs, mode
        )
        reserved_outputs.add(out_rel.as_posix().lower())
        final_dst = (output_dir / out_rel) if mode == "output" else (source_dir / out_rel)

        will_run = True
        skip_reason = ""
        if item.get("status") == "done" and final_dst.exists() and final_dst.stat().st_size > 0:
            will_run = False
            skip_reason = "이미 완료됨"
        elif item.get("status") == "failed" and not retry_failed:
            will_run = False
            skip_reason = "이전 실패 (--retry-failed 필요)"

        standalone.append({
            "item": item,
            "action": action,
            "out_rel": out_rel,
            "will_run": will_run,
            "skip_reason": skip_reason,
        })

    merge_tasks: list[dict[str, Any]] = []
    for grp in merge_groups:
        g_items = [by_id[iid] for iid in grp["item_ids"]]
        g_src_keys = {it["rel_path"].lower() for it in g_items}
        if grp.get("output_rel_path"):
            out_rel = Path(grp["output_rel_path"])
            reserved_outputs.add(out_rel.as_posix().lower())
        else:
            out_rel = resolve_merge_output_rel_path(
                grp, g_items[0], source_dir, reserved_outputs, mode, g_src_keys
            )
        final_dst = (output_dir / out_rel) if mode == "output" else (source_dir / out_rel)
        will_run = True
        skip_reason = ""
        if grp.get("status") == "done" and final_dst.exists() and final_dst.stat().st_size > 0:
            will_run = False
            skip_reason = "이미 완료됨"
        elif grp.get("status") == "failed" and not retry_failed:
            will_run = False
            skip_reason = "이전 실패 (--retry-failed 필요)"

        merge_tasks.append({
            "group": grp,
            "items": g_items,
            "out_rel": out_rel,
            "will_run": will_run,
            "skip_reason": skip_reason,
        })

    rotated_count = sum(
        1 for it in items if not it.get("delete") and int(it.get("user_rotation", 0)) != 0
    )
    standalone_run = [t for t in standalone if t["will_run"]]
    merge_run = [m for m in merge_tasks if m["will_run"]]
    deleted_run = [d for d in deleted_tasks if d["will_run"]]
    total_videos = sum(1 for it in items if it.get("media_type", "video") == "video")
    total_images = sum(1 for it in items if it.get("media_type") == "image")
    standalone_video_count = sum(1 for t in standalone_run if t["item"].get("media_type", "video") == "video")
    standalone_image_count = sum(1 for t in standalone_run if t["item"].get("media_type") == "image")

    summary = {
        "total_items": len(items),
        "total_videos": total_videos,
        "total_images": total_images,
        "standalone_count": len(standalone_run),
        "standalone_video_count": standalone_video_count,
        "standalone_image_count": standalone_image_count,
        "rotated_count": rotated_count,
        "date_fix_count": sum(1 for t in standalone_run if t["action"] == "fix_date"),
        "merge_group_count": len(merge_run),
        "merge_clip_count": sum(len(m["items"]) for m in merge_run),
        "delete_marked_count": len(deleted_tasks),
        "delete_will_run_count": len(deleted_run),
        "keep_count": max(
            0,
            len(items)
            - len(standalone)
            - sum(len(m["items"]) for m in merge_tasks)
            - len(deleted_tasks),
        ),
        "total_will_run": len(standalone_run) + len(merge_run) + len(deleted_run),
    }
    return {
        "standalone": standalone,
        "merge": merge_tasks,
        "deleted": deleted_tasks,
        "merge_groups": merge_groups,
        "summary": summary,
    }


def run_convert_job(
    work_dir: Path,
    mode: str = "output",
    output_dir: Path | None = None,
    crf: int = 20,
    preset: str = "medium",
    fix_missing_dates: bool = False,
    retry_failed: bool = False,
    ffmpeg_bin: str = "ffmpeg",
    ffprobe_bin: str = "ffprobe",
    progress_cb: Callable[[dict[str, Any]], None] | None = None,
    cancel_event: threading.Event | None = None,
    plan_lock: threading.Lock | None = None,
) -> dict[str, Any]:
    """Execute batch conversion, image rotation, merging, and safe trash moves."""
    plan_path = work_dir / "plan.json"
    plan = load_plan(work_dir)
    source_dir = Path(plan["source_dir"])
    resolved_out_dir = output_dir if output_dir is not None else (work_dir / "converted")
    backup_dir = work_dir / "original_backup"
    trash_dir = work_dir / "deleted_trash"
    temp_dir = work_dir / "tmp"
    error_log = work_dir / "convert_errors.log"

    planned = plan_convert_tasks(
        plan,
        source_dir=source_dir,
        output_dir=resolved_out_dir,
        mode=mode,
        fix_missing_dates=fix_missing_dates,
        retry_failed=retry_failed,
    )
    plan["merge_groups"] = planned["merge_groups"]

    # Persist computed output_rel_path onto live plan items/groups
    for t in planned["standalone"]:
        t["item"]["output_rel_path"] = t["out_rel"].as_posix()
    for m in planned["merge"]:
        m["group"]["output_rel_path"] = m["out_rel"].as_posix()

    temp_dir.mkdir(parents=True, exist_ok=True)
    if mode == "output":
        resolved_out_dir.mkdir(parents=True, exist_ok=True)
    else:
        backup_dir.mkdir(parents=True, exist_ok=True)

    standalone_tasks = planned["standalone"]
    merge_tasks = planned["merge"]
    deleted_tasks = planned["deleted"]
    total_units = (
        sum(1 for t in standalone_tasks if t["will_run"])
        + sum(len(m["items"]) for m in merge_tasks if m["will_run"])
        + sum(1 for d in deleted_tasks if d["will_run"])
    )

    done_units = 0
    converted_cnt = 0
    merged_cnt = 0
    deleted_cnt = 0
    skipped_cnt = 0
    failed_cnt = 0
    cancelled = False

    def save_plan() -> None:
        if plan_lock:
            with plan_lock:
                atomic_write_json(plan_path, plan)
        else:
            atomic_write_json(plan_path, plan)

    def emit(label: str, log_line: str = "", log_level: str = "info") -> None:
        if progress_cb:
            progress_cb({
                "done_units": done_units,
                "total_units": total_units,
                "current_label": label,
                "converted_cnt": converted_cnt,
                "merged_cnt": merged_cnt,
                "deleted_cnt": deleted_cnt,
                "skipped_cnt": skipped_cnt,
                "failed_cnt": failed_cnt,
                "log_line": log_line,
                "log_level": log_level,
            })

    emit("변환 및 정리 작업 준비 중...")

    for idx, task in enumerate(standalone_tasks, 1):
        if cancel_event and cancel_event.is_set():
            cancelled = True
            emit("사용자에 의해 안전하게 중단되었습니다.", "사용자 요청으로 작업을 중단했습니다.", "warn")
            break

        item = task["item"]
        action = task["action"]
        out_rel = task["out_rel"]
        src_path = source_dir / Path(item["rel_path"])
        final_dst = (resolved_out_dir / out_rel) if mode == "output" else (source_dir / out_rel)

        if not task["will_run"]:
            skipped_cnt += 1
            emit(f"건너뜀: {item['rel_path']}")
            continue

        if not src_path.exists():
            msg = f"원본 파일 없음: {src_path}"
            item["status"] = "failed"
            item["error"] = msg
            failed_cnt += 1
            done_units += 1
            emit(f"실패: {item['rel_path']}", f"[실패] {msg}", "err")
            save_plan()
            continue

        emit(f"[{idx}/{len(standalone_tasks)}] {item['rel_path']} → {out_rel.as_posix()}",
             f"처리 중: {item['rel_path']} → {out_rel.as_posix()}", "info")
        tmp_ext = out_rel.suffix.lower() if action == "image_process" else ".mp4"
        tmp_out = temp_dir / f"{item['id']}{tmp_ext}"
        if tmp_out.exists():
            tmp_out.unlink()

        try:
            if action == "image_process":
                cmd = build_ffmpeg_image_cmd(
                    src_path=src_path,
                    dst_path=tmp_out,
                    user_rotation=int(item.get("user_rotation", 0)),
                    ffmpeg_bin=ffmpeg_bin,
                )
                res = run_cmd(cmd, timeout=300)
                if res.returncode != 0:
                    raise RuntimeError(res.stderr.strip()[-600:] or f"ffmpeg exit {res.returncode}")
                verify_output_image(tmp_out)
                if tmp_ext in {".jpg", ".jpeg"}:
                    inject_jpeg_exif(src_path, tmp_out, item["best_date_utc"])
            else:
                if action == "encode":
                    cmd = build_ffmpeg_convert_cmd(
                        src_path=src_path,
                        dst_path=tmp_out,
                        user_rotation=int(item.get("user_rotation", 0)),
                        has_audio=bool(item.get("has_audio", True)),
                        creation_time_utc=item["best_date_utc"],
                        crf=int(crf),
                        preset=preset,
                        ffmpeg_bin=ffmpeg_bin,
                    )
                else:
                    cmd = build_ffmpeg_date_fix_cmd(
                        src_path=src_path,
                        dst_path=tmp_out,
                        creation_time_utc=item["best_date_utc"],
                        ffmpeg_bin=ffmpeg_bin,
                    )

                res = run_cmd(cmd, timeout=3600)
                if res.returncode != 0:
                    raise RuntimeError(res.stderr.strip()[-600:] or f"ffmpeg exit {res.returncode}")

                verify_output_video(tmp_out, ffprobe_bin=ffprobe_bin)

            ts = float(item["best_date_epoch"])
            final_dst.parent.mkdir(parents=True, exist_ok=True)
            if mode == "output":
                shutil.move(str(tmp_out), str(final_dst))
                os.utime(final_dst, (ts, ts))
            else:
                bak_path = backup_dir / Path(item["rel_path"])
                bak_path.parent.mkdir(parents=True, exist_ok=True)
                if not bak_path.exists():
                    shutil.move(str(src_path), str(bak_path))
                else:
                    src_path.unlink()
                shutil.move(str(tmp_out), str(final_dst))
                os.utime(final_dst, (ts, ts))

            item["status"] = "done"
            item.pop("error", None)
            converted_cnt += 1
            done_units += 1
            emit(f"완료: {out_rel.as_posix()}", f"✓ 완료: {out_rel.as_posix()}", "ok")
        except Exception as exc:
            if tmp_out.exists():
                try:
                    tmp_out.unlink()
                except OSError:
                    pass
            item["status"] = "failed"
            item["error"] = str(exc)
            failed_cnt += 1
            done_units += 1
            emit(f"실패: {item['rel_path']}", f"✗ 실패 ({item['rel_path']}): {exc}", "err")
            with error_log.open("a", encoding="utf-8") as ef:
                ef.write(f"[{datetime.now().isoformat()}] {item['rel_path']}\n{exc}\n\n")

        save_plan()

    if not cancelled:
        for g_idx, mtask in enumerate(merge_tasks, 1):
            if cancel_event and cancel_event.is_set():
                cancelled = True
                emit("사용자에 의해 안전하게 중단되었습니다.", "사용자 요청으로 병합 작업을 중단했습니다.", "warn")
                break

            grp = mtask["group"]
            g_items = mtask["items"]
            out_rel = mtask["out_rel"]
            final_dst = (resolved_out_dir / out_rel) if mode == "output" else (source_dir / out_rel)

            if not mtask["will_run"]:
                skipped_cnt += 1
                emit(f"병합 건너뜀: {grp['name']}")
                continue

            emit(f"[병합 {g_idx}/{len(merge_tasks)}] {grp['name']} ({len(g_items)}개 클립)",
                 f"병합 시작: {grp['name']} ({len(g_items)}개 클립) → {out_rel.as_posix()}", "info")

            grp_tmp = temp_dir / grp["id"]
            shutil.rmtree(grp_tmp, ignore_errors=True)
            grp_tmp.mkdir(parents=True, exist_ok=True)
            merged_tmp = grp_tmp / "merged.tmp.mp4"
            processed_clips = 0

            try:
                for it in g_items:
                    src_path = source_dir / Path(it["rel_path"])
                    if not src_path.exists():
                        raise RuntimeError(f"원본 파일 없음: {src_path}")

                target_w, target_h = compute_group_canvas(g_items)
                group_has_audio = any(bool(it.get("has_audio", False)) for it in g_items)

                seg_paths: list[Path] = []
                for s_idx, it in enumerate(g_items):
                    src_path = source_dir / Path(it["rel_path"])
                    seg_path = grp_tmp / f"seg_{s_idx:03d}.ts"
                    emit(f"[병합 {g_idx}/{len(merge_tasks)}] 클립 {s_idx + 1}/{len(g_items)}: {it['filename']}")
                    seg_cmd = build_ffmpeg_segment_cmd(
                        src_path=src_path,
                        dst_ts_path=seg_path,
                        user_rotation=int(it.get("user_rotation", 0)),
                        has_audio=bool(it.get("has_audio", True)),
                        group_has_audio=group_has_audio,
                        target_w=target_w,
                        target_h=target_h,
                        crf=int(crf),
                        preset=preset,
                        ffmpeg_bin=ffmpeg_bin,
                    )
                    res = run_cmd(seg_cmd, timeout=3600)
                    if res.returncode != 0:
                        raise RuntimeError(
                            f"세그먼트 실패 ({it['filename']}): "
                            + (res.stderr.strip()[-500:] or f"ffmpeg exit {res.returncode}")
                        )
                    seg_paths.append(seg_path)
                    done_units += 1
                    processed_clips += 1
                    emit(f"[병합 {g_idx}/{len(merge_tasks)}] 클립 {s_idx + 1}/{len(g_items)} 준비 완료")

                concat_list = grp_tmp / "concat_list.txt"
                with concat_list.open("w", encoding="utf-8") as cf:
                    for sp in seg_paths:
                        cf.write(f"file '{sp.name}'\n")

                concat_cmd = build_ffmpeg_concat_cmd(
                    concat_list_path=concat_list,
                    dst_mp4_path=merged_tmp,
                    group_has_audio=group_has_audio,
                    creation_time_utc=g_items[0]["best_date_utc"],
                    ffmpeg_bin=ffmpeg_bin,
                )
                res = run_cmd(concat_cmd, timeout=3600)
                if res.returncode != 0:
                    raise RuntimeError(
                        "병합 실패: " + (res.stderr.strip()[-500:] or f"ffmpeg exit {res.returncode}")
                    )

                verify_output_video(merged_tmp, ffprobe_bin=ffprobe_bin)

                ts = float(g_items[0]["best_date_epoch"])
                final_dst.parent.mkdir(parents=True, exist_ok=True)
                if mode == "output":
                    shutil.move(str(merged_tmp), str(final_dst))
                    os.utime(final_dst, (ts, ts))
                else:
                    for it in g_items:
                        src_path = source_dir / Path(it["rel_path"])
                        bak_path = backup_dir / Path(it["rel_path"])
                        bak_path.parent.mkdir(parents=True, exist_ok=True)
                        if src_path.exists():
                            if not bak_path.exists():
                                shutil.move(str(src_path), str(bak_path))
                            else:
                                src_path.unlink()
                    shutil.move(str(merged_tmp), str(final_dst))
                    os.utime(final_dst, (ts, ts))

                grp["status"] = "done"
                grp.pop("error", None)
                for it in g_items:
                    it["status"] = "merged"
                    it.pop("error", None)
                merged_cnt += 1
                emit(f"병합 완료: {out_rel.as_posix()}", f"✓ 병합 완료: {out_rel.as_posix()}", "ok")
            except Exception as exc:
                done_units += max(0, len(g_items) - processed_clips)
                grp["status"] = "failed"
                grp["error"] = str(exc)
                failed_cnt += 1
                emit(f"병합 실패: {grp['name']}", f"✗ 병합 실패 ({grp['name']}): {exc}", "err")
                with error_log.open("a", encoding="utf-8") as ef:
                    ef.write(f"[{datetime.now().isoformat()}] [MERGE {grp['name']}]\n{exc}\n\n")
            finally:
                shutil.rmtree(grp_tmp, ignore_errors=True)
                save_plan()

    if not cancelled:
        for d_idx, dtask in enumerate(deleted_tasks, 1):
            if cancel_event and cancel_event.is_set():
                cancelled = True
                emit("사용자에 의해 안전하게 중단되었습니다.", "사용자 요청으로 삭제 정리를 중단했습니다.", "warn")
                break

            item = dtask["item"]
            trash_rel = dtask["trash_rel"]
            if not dtask["will_run"]:
                skipped_cnt += 1
                emit(f"삭제 건너뜀 ({dtask['skip_reason']}): {item['rel_path']}")
                continue

            src_path = source_dir / trash_rel
            if not src_path.exists():
                skipped_cnt += 1
                done_units += 1
                emit(f"건너뜀 (원본 없음): {item['rel_path']}")
                continue

            dst_trash = trash_dir / trash_rel
            if dst_trash.exists():
                stem = dst_trash.stem
                suf = dst_trash.suffix
                c_idx = 2
                while dst_trash.exists():
                    dst_trash = dst_trash.with_name(f"{stem}_{c_idx}{suf}")
                    c_idx += 1

            try:
                dst_trash.parent.mkdir(parents=True, exist_ok=True)
                shutil.move(str(src_path), str(dst_trash))
                item["status"] = "deleted"
                item["trash_rel_path"] = dst_trash.relative_to(trash_dir).as_posix()
                item.pop("error", None)
                deleted_cnt += 1
                done_units += 1
                emit(
                    f"[삭제 {d_idx}/{len(deleted_tasks)}] 휴지통 이동: {item['rel_path']}",
                    f"🗑 휴지통 이동: {item['rel_path']} → deleted_trash/{item['trash_rel_path']}",
                    "ok",
                )
            except Exception as exc:
                item["status"] = "failed"
                item["error"] = str(exc)
                failed_cnt += 1
                done_units += 1
                emit(f"삭제 이동 실패: {item['rel_path']}", f"✗ 휴지통 이동 실패 ({item['rel_path']}): {exc}", "err")
            save_plan()

    save_plan()
    return {
        "converted_cnt": converted_cnt,
        "merged_cnt": merged_cnt,
        "deleted_cnt": deleted_cnt,
        "skipped_cnt": skipped_cnt,
        "failed_cnt": failed_cnt,
        "cancelled": cancelled,
        "output_folder": str(resolved_out_dir if mode == "output" else source_dir),
        "backup_folder": str(backup_dir) if mode == "replace-backup" else "",
        "trash_folder": str(trash_dir) if trash_dir.exists() else "",
        "error_log": str(error_log) if failed_cnt > 0 else "",
    }


def cmd_convert(args: argparse.Namespace) -> int:
    """Execute the 'convert' subcommand."""
    ffmpeg_bin, ffprobe_bin = check_ffmpeg_tools()
    work_dir = Path(args.work_dir).expanduser().resolve()
    plan = load_plan(work_dir)
    source_dir = Path(plan["source_dir"])
    mode = args.mode
    output_dir = (
        Path(args.output_dir).expanduser().resolve()
        if args.output_dir
        else (work_dir / "converted")
    )
    backup_dir = work_dir / "original_backup"

    print(f"변환 모드: {mode}")
    if mode == "output":
        print(f"출력 폴더: {output_dir} (원본 파일은 변경되지 않습니다)")
    else:
        print(f"백업 폴더: {backup_dir} (원본을 백업 후 정리된 파일로 교체합니다)")

    if args.dry_run:
        planned = plan_convert_tasks(
            plan,
            source_dir=source_dir,
            output_dir=output_dir,
            mode=mode,
            fix_missing_dates=bool(args.fix_missing_dates),
            retry_failed=bool(args.retry_failed),
        )
        tasks = planned["standalone"]
        merge_tasks = planned["merge"]
        del_tasks = planned["deleted"]
        print(f"\n[미리보기 --dry-run] 단독 변환/회전 {len(tasks)}개 / 병합 그룹 {len(merge_tasks)}개 / 휴지통 이동 {len(del_tasks)}개:")
        for t in tasks:
            item = t["item"]
            rot = int(item.get("user_rotation", 0))
            if t["action"] == "image_process":
                act_label = f"이미지변환(회전 {rot}°)"
            elif t["action"] == "encode":
                act_label = f"재인코딩(회전 {rot}°)"
            else:
                act_label = "날짜메타보완(무손실)"
            print(f"  - {item['rel_path']} -> {t['out_rel'].as_posix()} [{act_label}, 상태={item.get('status')}]")
        for m in merge_tasks:
            grp = m["group"]
            members_desc = ", ".join(
                f"{it['filename']}({it.get('user_rotation', 0)}°)" for it in m["items"]
            )
            print(f"  - [병합 그룹] {grp['name']} ({members_desc}) -> {m['out_rel'].as_posix()} [상태={grp.get('status')}]")
        for d in del_tasks:
            it = d["item"]
            print(f"  - [휴지통 이동] {it['rel_path']} ({d['skip_reason'] or '실행 예정'})")
        return 0

    def on_cli_prog(ev: dict[str, Any]) -> None:
        line = ev.get("log_line")
        if line:
            stream = sys.stderr if ev.get("log_level") == "err" else sys.stdout
            print(line, file=stream)

    summary = run_convert_job(
        work_dir=work_dir,
        mode=mode,
        output_dir=output_dir,
        crf=int(args.crf),
        preset=args.preset,
        fix_missing_dates=bool(args.fix_missing_dates),
        retry_failed=bool(args.retry_failed),
        ffmpeg_bin=ffmpeg_bin,
        ffprobe_bin=ffprobe_bin,
        progress_cb=on_cli_prog,
    )
    print("\n=== 작업 완료 ===")
    print(
        f"단독 변환/회전 성공: {summary['converted_cnt']}개 | 병합 성공: {summary['merged_cnt']}그룹 "
        f"| 휴지통 이동: {summary['deleted_cnt']}개 | 건너뜀: {summary['skipped_cnt']}개 | 실패: {summary['failed_cnt']}개"
    )
    if summary["failed_cnt"] > 0:
        print(f"실패 상세 로그: {summary['error_log']}")
        print("실패한 항목만 다시 시도하려면 --retry-failed 옵션을 추가하세요.")
    return 1 if summary["failed_cnt"] > 0 else 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Google Photos 업로드 전 사진/영상 미리보기·회전·삭제·병합 정리 도구"
    )
    sub = parser.add_subparsers(dest="command", required=True)

    p_app = sub.add_parser("app", help="4단계 웹 마법사 실행 (기본 모드)")
    p_app.add_argument("--work-dir", help="기존 _gphotos_work 폴더 경로 (선택)")
    p_app.add_argument("--port", type=int, default=8765, help="로컬 웹 서버 포트 (기본: 8765)")
    p_app.add_argument("--no-browser", action="store_true", help="브라우저 자동 열기 비활성화")
    p_app.set_defaults(func=cmd_app)

    p_scan = sub.add_parser("scan", help="폴더 스캔 및 썸네일 추출")
    p_scan.add_argument("source_dir", help="사진/영상들이 들어있는 원본 폴더 경로")
    p_scan.add_argument("--work-dir", help="작업 폴더 (기본: <원본폴더>/_gphotos_work)")
    p_scan.add_argument("--workers", type=int, default=4, help="병렬 스캔 작업 수 (기본: 4)")
    p_scan.add_argument("--no-images", action="store_true", help="이미지 파일을 제외하고 동영상만 스캔")
    p_scan.set_defaults(func=cmd_scan)

    p_rev = sub.add_parser("review", help="브라우저에서 돌아간 영상/병합 체크 (2단계 바로 열기)")
    p_rev.add_argument("--work-dir", required=True, help="_gphotos_work 폴더 경로")
    p_rev.add_argument("--port", type=int, default=8765, help="로컬 웹 서버 포트 (기본: 8765)")
    p_rev.add_argument("--no-browser", action="store_true", help="브라우저 자동 열기 비활성화")
    p_rev.set_defaults(func=cmd_review)

    p_conv = sub.add_parser("convert", help="체크된 영상 일괄 변환")
    p_conv.add_argument("--work-dir", required=True, help="_gphotos_work 폴더 경로")
    p_conv.add_argument(
        "--mode",
        choices=("output", "replace-backup"),
        default="output",
        help="output(기본, 별도 폴더에 저장) 또는 replace-backup(원본 백업 후 제자리 교체)",
    )
    p_conv.add_argument("--output-dir", help="--mode output 시 저장할 폴더 경로")
    p_conv.add_argument("--crf", type=int, default=20, help="화질 CRF (낮을수록 고화질, 기본: 20)")
    p_conv.add_argument("--preset", default="medium", help="x264 인코딩 속도 프리셋 (기본: medium)")
    p_conv.add_argument("--dry-run", action="store_true", help="실제 변환 없이 대상 목록만 출력")
    p_conv.add_argument(
        "--fix-missing-dates",
        action="store_true",
        help="회전/변환 대상이 아닌 MP4/MOV 중 촬영일 메타가 없는 파일에도 무손실로 날짜 기록",
    )
    p_conv.add_argument("--retry-failed", action="store_true", help="이전에 실패했던 항목 재시도")
    p_conv.set_defaults(func=cmd_convert)

    return parser


def main(argv: list[str] | None = None) -> int:
    raw_argv = sys.argv[1:] if argv is None else list(argv)
    if not raw_argv:
        raw_argv = ["app"]
    parser = build_parser()
    args = parser.parse_args(raw_argv)
    return int(args.func(args))


if __name__ == "__main__":
    sys.exit(main())
