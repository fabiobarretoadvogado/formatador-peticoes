from __future__ import annotations

import hashlib
import json
import shutil
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
TIPO_PECA_SUPORTADO = "peticao_inicial"


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
