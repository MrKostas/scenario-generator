"""PyInstaller configuration for the Windows scenario generator."""

import os
from pathlib import Path


source_directory = Path(SPEC).resolve().parent
mode = os.environ.get("SCENARIO_BUILD_MODE", "onefile")
if mode not in {"onedir", "onefile"}:
    raise ValueError(f"Unsupported build mode: {mode}")

analysis = Analysis(
    [str(source_directory / "desktop_app.py")],
    pathex=[str(source_directory)],
    binaries=[],
    datas=[
        (str(source_directory / "scenario-generator.py"), "."),
        (str(source_directory / "scenario-generator-epi.py"), "."),
    ],
    hiddenimports=["pandas", "numpy", "openpyxl", "xlrd"],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(analysis.pure)

if mode == "onefile":
    exe = EXE(
        pyz,
        analysis.scripts,
        analysis.binaries,
        analysis.datas,
        [],
        name="ScenarioGenerator",
        console=False,
        version=str(source_directory / "version_info.txt"),
        upx=False,
    )
else:
    exe = EXE(
        pyz,
        analysis.scripts,
        [],
        name="ScenarioGenerator",
        console=False,
        version=str(source_directory / "version_info.txt"),
        exclude_binaries=True,
        upx=False,
    )
    coll = COLLECT(
        exe,
        analysis.binaries,
        analysis.datas,
        name="ScenarioGenerator",
        upx=False,
    )
