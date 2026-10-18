from __future__ import annotations

import importlib.util
import json
import os
import shutil
import subprocess
import sys
import tempfile
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from app_version import VERSION
from scripts.gerar_icone import main as gerar_icone


ROOT = Path(__file__).resolve().parents[1]
DIST = ROOT / "dist"


def main() -> None:
    gerar_icone()
    if str(ROOT) not in sys.path:
        sys.path.insert(0, str(ROOT))
    dependencias_ausentes = [
        modulo
        for modulo in ("PyInstaller", "markdown_it", "scripts.numeracao_word")
        if importlib.util.find_spec(modulo) is None
    ]
    if dependencias_ausentes:
        raise SystemExit(
            "Dependências do pacote portátil não estão disponíveis neste Python: "
            + ", ".join(dependencias_ausentes)
            + ". Execute:\n"
            f'  "{sys.executable}" -m pip install -e ".[build]"'
        )

    with tempfile.TemporaryDirectory(prefix="formatador_portatil_") as temporario:
        temporario_raiz = Path(temporario)
        dist_temporario = temporario_raiz / "dist"
        build_temporario = temporario_raiz / "build"
        pacote = dist_temporario / "FormatadorPeticoes"

        subprocess.run(
            [
                sys.executable,
                "-m",
                "PyInstaller",
                "--noconfirm",
                "--clean",
                "--distpath",
                str(dist_temporario),
                "--workpath",
                str(build_temporario),
                str(ROOT / "FormatadorPeticoes.spec"),
            ],
            cwd=ROOT,
            check=True,
        )

        modelos_destino = pacote / "modelos"
        modelos_destino.mkdir(parents=True, exist_ok=True)
        for nome in ("modelo_oficial.docx", "estilos.json"):
            shutil.copy2(ROOT / "modelos" / nome, modelos_destino / nome)

        saida_teste = temporario_raiz / "teste_resultados"
        entrada_markdown = temporario_raiz / "peticao_markdown.md"
        entrada_markdown.write_text(
            "## DOS FATOS\n\nTexto de teste do pacote portátil.\n\n"
            "> Citação de teste.\n\n1. Item numerado.\n\n"
            "| Campo | Valor |\n| --- | --- |\n| Origem | Markdown |\n",
            encoding="utf-8",
        )
        ambiente = os.environ.copy()
        ambiente["LOCALAPPDATA"] = str(temporario_raiz / "dados_locais")
        try:
            teste = subprocess.run(
                [
                    str(pacote / "FormatadorPeticoes.exe"),
                    str(ROOT / "modelos" / "modelo_oficial.docx"),
                    str(entrada_markdown),
                    "--saida",
                    str(saida_teste),
                    "--sem-pdf",
                    "--silencioso",
                ],
                cwd=pacote,
                env=ambiente,
                check=False,
                timeout=90,
            )
        except subprocess.TimeoutExpired as exc:
            raise RuntimeError(
                "O executável foi criado, mas o teste de DOCX e Markdown não terminou em 90 segundos."
            ) from exc
        saidas_docx = list(saida_teste.rglob("*_formatada.docx"))
        if (
            teste.returncode != 0
            or len(saidas_docx) < 2
            or not any(saida_teste.rglob("entrada_intermediaria.docx"))
            or not any(saida_teste.rglob("manifesto_entrada.json"))
        ):
            raise RuntimeError(
                "O executável foi criado, mas não passou no teste automático de DOCX e Markdown."
            )

        diagnostico_interface = temporario_raiz / "diagnostico_interface.json"
        try:
            teste_interface = subprocess.run(
                [
                    str(pacote / "FormatadorPeticoes.exe"),
                    "--validar-interface",
                    str(diagnostico_interface),
                ],
                cwd=pacote,
                env=ambiente,
                check=False,
                timeout=30,
            )
        except subprocess.TimeoutExpired as exc:
            raise RuntimeError(
                "O executável processa documentos, mas a validação da interface não terminou em 30 segundos."
            ) from exc
        if teste_interface.returncode != 0 or not diagnostico_interface.is_file():
            detalhe = (
                diagnostico_interface.read_text(encoding="utf-8")
                if diagnostico_interface.is_file()
                else "O executável não produziu o diagnóstico esperado."
            )
            raise RuntimeError(
                "O executável processa documentos, mas o frontend gráfico não iniciou:\n" + detalhe
            )

        diagnostico = json.loads(diagnostico_interface.read_text(encoding="utf-8"))
        if not diagnostico.get("tk_importado") or not diagnostico.get("funcional"):
            raise RuntimeError(
                "O pacote seria publicado sem o frontend completo:\n"
                + diagnostico_interface.read_text(encoding="utf-8")
            )
        print("Frontend Tcl/Tk completo validado.")

        zip_base = temporario_raiz / "FormatadorPeticoes-portatil"
        zip_criado = Path(
            shutil.make_archive(
                str(zip_base), "zip", root_dir=dist_temporario, base_dir=pacote.name
            )
        )

        DIST.mkdir(parents=True, exist_ok=True)
        destino = DIST / f"FormatadorPeticoes-portatil-{VERSION}.zip"
        shutil.copy2(zip_criado, destino)
        print(f"Pacote criado e testado: {destino}")


if __name__ == "__main__":
    main()
