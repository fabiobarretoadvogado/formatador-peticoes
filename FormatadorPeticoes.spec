# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path
import sys


ROOT = Path(SPECPATH)
PYTHON_ROOT = Path(sys.base_prefix)
TCL_ROOT = PYTHON_ROOT / "tcl"
PYTHON_DLLS = PYTHON_ROOT / "DLLs"


def _tk_runtime():
    datas = [
        (TCL_ROOT / "tcl8.6", "_tcl_data"),
        (TCL_ROOT / "tk8.6", "_tk_data"),
        (PYTHON_ROOT / "Lib" / "tkinter", "tkinter"),
    ]
    binaries = [
        (PYTHON_DLLS / "tcl86t.dll", "."),
        (PYTHON_DLLS / "tk86t.dll", "."),
    ]
    ausentes = [
        str(path)
        for path, _ in [*datas, *binaries]
        if not path.exists()
    ]
    if ausentes:
        raise SystemExit("Runtime Tcl/Tk incompleto: " + ", ".join(ausentes))
    return (
        [(str(path), destino) for path, destino in datas],
        [(str(path), destino) for path, destino in binaries],
    )


TK_DATAS, TK_BINARIES = _tk_runtime()


def _icon_is_readable(path: Path) -> bool:
    try:
        with path.open("rb"):
            return True
    except OSError:
        # OneDrive can leave the icon as an online-only or locked file.  The
        # application can still be built with the default executable icon.
        return False


a = Analysis(
    [str(ROOT / "aplicativo.py")],
    pathex=[str(ROOT)],
    binaries=TK_BINARIES,
    datas=[
        (str(ROOT / "modelos" / "modelo_oficial.docx"), "modelos"),
        (str(ROOT / "modelos" / "estilos.json"), "modelos"),
    ] + TK_DATAS,
    hiddenimports=[
        "scripts.numeracao_word",
        "_tkinter",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="FormatadorPeticoesApp",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=(
        str(ROOT / "assets" / "aplicativo.ico")
        if _icon_is_readable(ROOT / "assets" / "aplicativo.ico")
        else None
    ),
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="FormatadorPeticoes",
)
