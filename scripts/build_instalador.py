"""Gera o portátil validado e o instalador Windows a partir de arquivos limpos."""
from __future__ import annotations

import os
import subprocess
import sys
import tempfile
import tomllib
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from app_version import VERSION
from scripts.build_portatil import main as build_portatil


def main() -> None:
    metadata = tomllib.loads((ROOT / "pyproject.toml").read_text(encoding="utf-8"))
    if metadata["project"]["version"] != VERSION:
        raise SystemExit("As versões de app_version.py e pyproject.toml devem ser iguais.")
    compiler = Path(os.environ.get("ISCC_PATH", r"C:\Program Files (x86)\Inno Setup 6\ISCC.exe"))
    if not compiler.is_file():
        raise SystemExit("Instale Inno Setup 6 ou defina ISCC_PATH com o caminho de ISCC.exe.")
    build_portatil()
    with tempfile.TemporaryDirectory(prefix="formatador_instalador_") as temp:
        with zipfile.ZipFile(ROOT / "dist" / f"FormatadorPeticoes-portatil-{VERSION}.zip") as pacote:
            pacote.extractall(temp)
        source = Path(temp) / "FormatadorPeticoes"
        # Nenhum arquivo de entrada, resultado ou execução integra o instalador.
        for proibida in ("entrada", "saida", "trabalho", "execucoes", "tmp"):
            if list(source.rglob(proibida)):
                raise RuntimeError(f"Pasta de documentos encontrada no pacote: {proibida}")
        subprocess.run([
            str(compiler), f"/DAppVersion={VERSION}", f"/DSourceDir={source}",
            f"/DOutputDir={ROOT / 'dist'}", str(ROOT / "installer" / "FormatadorPeticoes.iss"),
        ], check=True, cwd=ROOT)
    print(f"Instalador criado: {ROOT / 'dist' / 'FormatadorPeticoes-Setup.exe'}")


if __name__ == "__main__":
    main()
