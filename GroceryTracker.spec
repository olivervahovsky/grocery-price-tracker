# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path

source_dir = Path(SPECPATH)
analysis = Analysis(
    [str(source_dir / "GroceryTracker.py")],
    pathex=[str(source_dir)],
)
archive = PYZ(analysis.pure)
executable = EXE(
    archive,
    analysis.scripts,
    analysis.binaries,
    analysis.datas,
    [],
    name="GroceryTracker",
    console=False,
    upx=False,
)
