"""Window capture for failure postmortems.

Pure ctypes GDI plus a minimal stdlib PNG encoder, so it works from any
thread of any process with no COM and no imaging dependency: the worker's
watcher thread captures a blocking dialog before it is reported, and the
supervisor captures the (possibly hidden) Excel main window before a
timeout kill.

PrintWindow with PW_RENDERFULLCONTENT asks the window to render even when
offscreen, but it does so by SENDING a message: on a window whose thread is
not pumping (an Excel spinning inside a VBA loop, which is precisely when a
screenshot is wanted) it blocks indefinitely. Verified live on 2026-07-25,
where it wedged the timeout path itself. Two defenses, both required:

- ``IsHungAppWindow`` selects BitBlt instead of PrintWindow for a
  non-pumping window (BitBlt reads the DC without sending anything);
- every public entry point runs the capture on a daemon thread with a hard
  join timeout, so no GDI call can ever delay a kill.

Capture failures are soft (return None) and never affect a run outcome.
"""
from __future__ import annotations

import ctypes
import ctypes.wintypes as wt
import struct
import threading
import time
import zlib
from pathlib import Path

_user32 = ctypes.windll.user32
_gdi32 = ctypes.windll.gdi32

_PW_RENDERFULLCONTENT = 0x00000002
_SRCCOPY = 0x00CC0020
_BI_RGB = 0
_DIB_RGB_COLORS = 0

_EnumWindowsProc = ctypes.WINFUNCTYPE(wt.BOOL, wt.HWND, wt.LPARAM)


class _BITMAPINFOHEADER(ctypes.Structure):
    _fields_ = [
        ("biSize", wt.DWORD),
        ("biWidth", ctypes.c_long),
        ("biHeight", ctypes.c_long),
        ("biPlanes", wt.WORD),
        ("biBitCount", wt.WORD),
        ("biCompression", wt.DWORD),
        ("biSizeImage", wt.DWORD),
        ("biXPelsPerMeter", ctypes.c_long),
        ("biYPelsPerMeter", ctypes.c_long),
        ("biClrUsed", wt.DWORD),
        ("biClrImportant", wt.DWORD),
    ]


def find_window_for_pid(pid: int, class_name: str = "XLMAIN") -> int:
    """Top-level window handle of the given class owned by the PID, or 0."""
    found = wt.HWND(0)

    def callback(hwnd: int, _lparam: int) -> bool:
        owner = wt.DWORD(0)
        _user32.GetWindowThreadProcessId(hwnd, ctypes.byref(owner))
        if owner.value != pid:
            return True
        buffer = ctypes.create_unicode_buffer(256)
        _user32.GetClassNameW(hwnd, buffer, 256)
        if buffer.value == class_name:
            found.value = hwnd
            return False
        return True

    _user32.EnumWindows(_EnumWindowsProc(callback), 0)
    return found.value or 0


def capture_window(hwnd: int, path: str | Path) -> str | None:
    """Save a PNG of the window; returns the path or None on any failure."""
    if not hwnd:
        return None
    rect = wt.RECT()
    if not _user32.GetWindowRect(hwnd, ctypes.byref(rect)):
        return None
    width = rect.right - rect.left
    height = rect.bottom - rect.top
    if width <= 0 or height <= 0 or width > 8192 or height > 8192:
        return None

    window_dc = _user32.GetWindowDC(hwnd)
    if not window_dc:
        return None
    memory_dc = _gdi32.CreateCompatibleDC(window_dc)
    bitmap = _gdi32.CreateCompatibleBitmap(window_dc, width, height)
    try:
        if not memory_dc or not bitmap:
            return None
        previous = _gdi32.SelectObject(memory_dc, bitmap)
        # PrintWindow sends a message; on a non-pumping window it never
        # returns. BitBlt only reads the device context.
        rendered = 0
        if not _user32.IsHungAppWindow(hwnd):
            rendered = _user32.PrintWindow(hwnd, memory_dc,
                                           _PW_RENDERFULLCONTENT)
        if not rendered:
            _gdi32.BitBlt(memory_dc, 0, 0, width, height, window_dc, 0, 0,
                          _SRCCOPY)
        _gdi32.SelectObject(memory_dc, previous)

        header = _BITMAPINFOHEADER()
        header.biSize = ctypes.sizeof(_BITMAPINFOHEADER)
        header.biWidth = width
        header.biHeight = -height  # top-down rows
        header.biPlanes = 1
        header.biBitCount = 32
        header.biCompression = _BI_RGB
        buffer = (ctypes.c_ubyte * (width * height * 4))()
        copied = _gdi32.GetDIBits(memory_dc, bitmap, 0, height, buffer,
                                  ctypes.byref(header), _DIB_RGB_COLORS)
        if copied != height:
            return None
        target = Path(path)
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(_encode_png(bytes(buffer), width, height))
        return str(target)
    except OSError:
        return None
    finally:
        if bitmap:
            _gdi32.DeleteObject(bitmap)
        if memory_dc:
            _gdi32.DeleteDC(memory_dc)
        _user32.ReleaseDC(hwnd, window_dc)


def _encode_png(bgra: bytes, width: int, height: int) -> bytes:
    """Minimal PNG: 8-bit RGB, filter 0 rows, one zlib IDAT."""
    rows = bytearray()
    stride = width * 4
    for y in range(height):
        rows.append(0)  # filter type None
        row = bgra[y * stride:(y + 1) * stride]
        for x in range(width):
            base = x * 4
            rows.append(row[base + 2])  # R (DIB stores BGRA)
            rows.append(row[base + 1])  # G
            rows.append(row[base + 0])  # B

    def chunk(kind: bytes, payload: bytes) -> bytes:
        body = kind + payload
        return (struct.pack(">I", len(payload)) + body
                + struct.pack(">I", zlib.crc32(body) & 0xFFFFFFFF))

    ihdr = struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0)
    return (b"\x89PNG\r\n\x1a\n"
            + chunk(b"IHDR", ihdr)
            + chunk(b"IDAT", zlib.compress(bytes(rows), 6))
            + chunk(b"IEND", b""))


def capture_window_safely(hwnd: int, path: str | Path,
                          timeout_s: float = 3.0) -> str | None:
    """capture_window on a daemon thread with a hard timeout.

    A capture must never delay killing a wedged Excel, so a GDI call that
    blocks past the timeout is simply abandoned (the thread is a daemon and
    dies with the process).
    """
    result: list[str | None] = [None]

    def work() -> None:
        try:
            result[0] = capture_window(hwnd, path)
        except Exception:  # noqa: BLE001 - capture is always best effort
            result[0] = None

    worker = threading.Thread(target=work, daemon=True,
                              name="pyvba-screenshot")
    worker.start()
    worker.join(timeout=timeout_s)
    return None if worker.is_alive() else result[0]


def capture_excel_window(pid: int, directory: str | Path,
                         label: str) -> str | None:
    """Find the Excel main window of a PID and capture it, best effort."""
    hwnd = find_window_for_pid(pid, "XLMAIN")
    if not hwnd:
        return None
    stamp = time.strftime("%Y%m%d-%H%M%S")
    return capture_window_safely(
        hwnd, Path(directory) / f"{label}-{stamp}.png")
