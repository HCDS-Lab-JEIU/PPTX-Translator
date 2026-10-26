"""One-folder Windows build. Never copy the complete repository into the bundle."""
import os
from pathlib import Path

folder = Path(SPECPATH).resolve()
root = folder.parent.parent
analysis = Analysis(
    [str(folder / "launcher.py")],
    pathex=[str(root)],
    binaries=[],
    datas=[(str(root / "pptx_pipeline/render_powerpoint.ps1"), "pptx_pipeline")],
    hiddenimports=['pypdfium2', 'pypdfium2_raw'],
    hookspath=[],
    runtime_hooks=[],
    excludes=["legacy", "tests", "tkinter", "PyQt5", "PyQt6", "PySide2"],
    noarchive=False,
)
archive = PYZ(analysis.pure)
executable = EXE(
    archive, analysis.scripts, [],
    exclude_binaries=True,
    name="PPTTranslator",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=os.environ.get("PPT_TRANSLATOR_BUILD_CONSOLE") == "1",
)
bundle = COLLECT(
    executable, analysis.binaries, analysis.datas,
    strip=False, upx=False, name="PPTTranslator",
)
