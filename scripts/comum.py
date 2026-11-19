from __future__ import annotations

import hashlib
import json
import os
import re
import shutil
import tempfile
from datetime import datetime, timezone
from itertools import count
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
TIPO_PECA_SUPORTADO = "peticao_inicial"


def pasta_dados_locais() -> Path:
    base = os.environ.get("LOCALAPPDATA")
    return Path(base) / "FormatadorPeticoes" if base else Path.home() / ".formatador_peticoes"


def carregar_json(caminho: Path) -> dict[str, Any]:
    return json.loads(caminho.read_text(encoding="utf-8"))


def salvar_json(caminho: Path, dados: Any) -> None:
    caminho.parent.mkdir(parents=True, exist_ok=True)
    caminho.write_text(json.dumps(dados, ensure_ascii=False, indent=2), encoding="utf-8")


def hash_texto(texto: str) -> str:
    normalizado = " ".join(texto.split()).casefold()
    return hashlib.sha256(normalizado.encode("utf-8")).hexdigest()


def agora_iso() -> str:
    return datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds")


def garantir_diretorios(root: Path = ROOT) -> None:
    for nome in ("modelos", "entrada", "trabalho", "saida"):
        (root / nome).mkdir(parents=True, exist_ok=True)


def backup_entrada(origem: Path, root: Path = ROOT) -> Path:
    pasta = root / "trabalho" / "backups"
    pasta.mkdir(parents=True, exist_ok=True)
    destino = pasta / f"{origem.stem}_{datetime.now():%Y%m%d_%H%M%S}{origem.suffix}"
    shutil.copy2(origem, destino)
    return destino


def nome_seguro(nome: str, limite: int = 80) -> str:
    base = nome.strip()
    for extensao in (".docx", ".md"):
        if base.casefold().endswith(extensao):
            base = base[:-len(extensao)]
            break
    base = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", base)
    base = re.sub(r"\s+", " ", base).strip(" .")
    return (base or "peticao")[:limite]


def publicar_docx(origem: Path, pasta: Path, nome: str, *, substituir: Path | None = None) -> Path:
    """Entrega o Word no destino e reserva um novo nome quando já existe outro arquivo."""
    pasta.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(prefix=".formatador-", suffix=".tmp", dir=pasta, delete=False) as temporario:
        preparado = Path(temporario.name)
    try:
        shutil.copy2(origem, preparado)
        if substituir is not None:
            preparado.replace(substituir)
            return substituir
        for numero in count(1):
            sufixo = "" if numero == 1 else f"_{numero}"
            destino = pasta / f"{nome}_formatada{sufixo}.docx"
            try:
                with destino.open("xb"):
                    pass
            except FileExistsError:
                continue
            try:
                preparado.replace(destino)
            except OSError:
                destino.unlink(missing_ok=True)
                raise
            return destino
    finally:
        preparado.unlink(missing_ok=True)
