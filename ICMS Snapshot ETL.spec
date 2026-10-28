# -*- mode: python ; coding: utf-8 -*-
"""One-file GUI build. Excludes unused optional stacks that PyInstaller would
otherwise follow (~80 MB of Apache Arrow alone). Do not exclude pandas.*
submodules that pandas imports at package load time (e.g. pandas.plotting)."""

import sys

sys.path.insert(0, SPECPATH)

from version import __version__

APP_NAME = f"ICMS Snapshot ETL v{__version__}"

# Safe excludes: optional third-party stacks and test tooling the app never
# imports. Keep cryptography (msoffcrypto) and Tcl/Tk (the GUI).
EXCLUDES = [
    "pyarrow",
    "pytest",
    "_pytest",
    "py",
    "pygments",
    "pluggy",
    "iniconfig",
    "matplotlib",
    "scipy",
    "PIL",
    "Pillow",
    "IPython",
    "jupyter",
    "notebook",
    "sphinx",
    "numba",
    "numexpr",
    "sklearn",
    "skimage",
    "statsmodels",
    "tables",
    "h5py",
    "lxml",
    "bs4",
    "html5lib",
    "sqlalchemy",
    "fastparquet",
    "fsspec",
    "s3fs",
    "botocore",
    "boto3",
    "zstandard",
    "lz4",
    "pandas.tests",
    "numpy.tests",
    "numpy.f2py",
    "numpy.distutils",
    "numpy.doc",
    "tkinter.test",
    "lib2to3",
    "ensurepip",
    "idlelib",
    "turtledemo",
    "turtle",
    "setuptools",
    "pkg_resources",
    "distutils",
]

a = Analysis(
    ["icms_gui.py"],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=EXCLUDES,
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name=APP_NAME,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
