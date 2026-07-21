"""Build the versioned Windows GUI executable."""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
SPEC = ROOT / "ICMS Snapshot ETL.spec"


def main() -> None:
    from version import __version__

    print(f"Building ICMS Snapshot ETL v{__version__}...")
    subprocess.check_call(
        [sys.executable, "-m", "PyInstaller", "--noconfirm", str(SPEC)],
        cwd=ROOT,
    )
    exe_path = ROOT / "dist" / f"ICMS Snapshot ETL v{__version__}.exe"
    if exe_path.exists():
        size_mb = exe_path.stat().st_size / (1024 * 1024)
        print(f"Done: {exe_path} ({size_mb:.1f} MB)")
    else:
        raise SystemExit(f"Build finished but executable not found: {exe_path}")


if __name__ == "__main__":
    main()
