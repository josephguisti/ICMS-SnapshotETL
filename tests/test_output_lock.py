"""Preflight output write-lock checks."""

from __future__ import annotations

import sys
from contextlib import contextmanager
from pathlib import Path

import pytest

import icms_etl_core as core
from conftest import make_workbook, run, spec


@contextmanager
def exclusive_lock(path: Path):
    """Hold an exclusive (no-share) handle so other writers get PermissionError.

    Mimics Excel locking a workbook open on Windows. Skips on non-Windows.
    """
    if sys.platform != "win32":
        pytest.skip("exclusive lock probe is Windows-specific")

    import ctypes
    from ctypes import wintypes

    GENERIC_READ = 0x80000000
    GENERIC_WRITE = 0x40000000
    OPEN_EXISTING = 3
    FILE_ATTRIBUTE_NORMAL = 0x80
    INVALID_HANDLE_VALUE = wintypes.HANDLE(-1).value

    CreateFileW = ctypes.windll.kernel32.CreateFileW
    CreateFileW.argtypes = [
        wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD,
        wintypes.LPVOID, wintypes.DWORD, wintypes.DWORD, wintypes.HANDLE,
    ]
    CreateFileW.restype = wintypes.HANDLE
    CloseHandle = ctypes.windll.kernel32.CloseHandle

    handle = CreateFileW(
        str(path),
        GENERIC_READ | GENERIC_WRITE,
        0,  # no sharing
        None,
        OPEN_EXISTING,
        FILE_ATTRIBUTE_NORMAL,
        None,
    )
    if handle == INVALID_HANDLE_VALUE or handle is None:
        raise OSError(f"Could not exclusively lock {path}")
    try:
        yield
    finally:
        CloseHandle(handle)


def test_check_outputs_writable_empty_dir(out_dir):
    assert core.check_outputs_writable(out_dir) == []


def test_check_outputs_writable_unlocked_existing(out_dir):
    (out_dir / core.SNAPSHOT_HISTORY_FILENAME).write_text(
        "snapshot_date,CHAMP ID\n2026-01-01,1\n", encoding="utf-8"
    )
    (out_dir / core.DASHBOARD_FILENAME).write_bytes(b"PK\x03\x04placeholder")
    assert core.check_outputs_writable(out_dir) == []


def test_check_outputs_writable_reports_locked_file(out_dir):
    target = out_dir / core.DASHBOARD_FILENAME
    target.write_bytes(b"PK\x03\x04placeholder")
    with exclusive_lock(target):
        locked = core.check_outputs_writable(out_dir)
    assert locked == [target]


def test_run_snapshot_files_aborts_before_processing_when_locked(tmp_path, out_dir):
    wb = make_workbook(tmp_path / "s.xlsx", [(1, "Proj A", 10, "SJC")])
    flat = out_dir / core.DASHBOARD_FILENAME
    flat.write_bytes(b"PK\x03\x04placeholder")
    ingest = out_dir / core.INGEST_LOG_FILENAME

    with exclusive_lock(flat):
        with pytest.raises(core.OutputLockedError) as ei:
            run([spec(wb, "2026-10-19")], out_dir)
        assert flat in ei.value.locked

    assert not ingest.exists(), "locked preflight must not start processing"
    assert not (out_dir / core.SNAPSHOT_HISTORY_FILENAME).exists()


def test_ensure_outputs_writable_message_lists_filename(out_dir):
    target = out_dir / core.MASTER_FILENAME
    target.write_bytes(b"PK\x03\x04placeholder")
    with exclusive_lock(target):
        with pytest.raises(core.OutputLockedError) as ei:
            core.ensure_outputs_writable(out_dir)
    msg = str(ei.value)
    assert core.MASTER_FILENAME in msg
    assert "Excel" in msg
