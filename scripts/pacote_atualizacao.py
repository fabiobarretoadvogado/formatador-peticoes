"""Cria um snapshot e um ZIP apenas com os arquivos novos ou alterados."""
from __future__ import annotations

import json
import shutil
import tempfile
import zipfile
from pathlib import Path

from scripts.atualizacoes import buscar_release, interpretar_asset, baixar_verificado
from scripts.protocolo_atualizacao import (
    MANIFESTO, MAX_MANIFESTO, REPOSITORIO, caminho_relativo, ler_manifesto,
    nome_pacote, sha256, validar_manifesto,
)


def manifesto_publicado() -> dict | None:
    release = buscar_release()
    if not any(a.get("name") == MANIFESTO for a in release.get("assets", [])):
        return None
    asset = interpretar_asset(release, MANIFESTO, MAX_MANIFESTO)
    with tempfile.TemporaryDirectory(prefix="base_atualizacao_") as temp:
        arquivo = baixar_verificado(asset, Path(temp))
        return ler_manifesto(arquivo, asset.versao)


def gerar_pacote(pasta: Path, destino: Path, numero: str, anterior: dict | None = None) -> dict:
    if anterior is not None:
        validar_manifesto(anterior)
    anteriores = {a["caminho"]: a for a in (anterior or {}).get("arquivos", [])}
    arquivos, alterados = [], []
    for caminho in sorted(pasta.rglob("*")):
        if caminho.is_symlink() or getattr(caminho, "is_junction", lambda: False)():
            raise ValueError("O pacote não pode conter links ou junções.")
        if not caminho.is_file() or caminho.name == MANIFESTO:
            continue
        relativo = caminho_relativo(caminho.relative_to(pasta).as_posix())
        item = {"caminho": relativo, "tamanho": caminho.stat().st_size,
                "sha256": sha256(caminho), "pacote": numero}
        antigo = anteriores.get(relativo)
        if antigo and antigo["sha256"] == item["sha256"] and antigo["tamanho"] == item["tamanho"]:
            item["pacote"] = antigo["pacote"]
        else:
            alterados.append(caminho)
        arquivos.append(item)
    dados = validar_manifesto({"protocolo": 1, "repositorio": REPOSITORIO,
                               "versao": numero, "arquivos": arquivos}, numero)
    destino.mkdir(parents=True, exist_ok=True)
    pacote = destino / nome_pacote(numero)
    with zipfile.ZipFile(pacote, "w", compression=zipfile.ZIP_DEFLATED) as zipado:
        for caminho in alterados:
            zipado.write(caminho, caminho.relative_to(pasta).as_posix())
    manifesto = destino / MANIFESTO
    manifesto.write_text(json.dumps(dados, ensure_ascii=False, indent=2), encoding="utf-8")
    shutil.copy2(manifesto, pasta / MANIFESTO)
    print(f"Atualização incremental: {len(alterados)} de {len(arquivos)} arquivo(s), {pacote.stat().st_size} bytes.")
    return dados
