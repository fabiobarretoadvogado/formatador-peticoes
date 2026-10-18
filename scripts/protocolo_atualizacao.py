"""Formato compartilhado entre compilação, download e inicialização de versões."""
from __future__ import annotations

import hashlib
import json
import os
import re
from pathlib import Path

REPOSITORIO = "fabiobarretoadvogado/formatador-peticoes"
MANIFESTO = "manifesto_atualizacao.json"
EXECUTAVEL = "FormatadorPeticoesApp.exe"
INICIADOR = "FormatadorPeticoes.exe"
MAX_MANIFESTO = 4 * 1024 * 1024
MAX_ARQUIVO = 300 * 1024 * 1024
MAX_TOTAL = 1024 * 1024 * 1024


def versao(valor: str) -> tuple[int, int, int]:
    if not isinstance(valor, str) or not re.fullmatch(r"\d+\.\d+\.\d+", valor):
        raise ValueError("Versão inválida.")
    return tuple(int(x) for x in valor.split("."))


def nome_pacote(numero: str) -> str:
    versao(numero)
    return f"FormatadorPeticoes-atualizacao-{numero}.zip"


def caminho_relativo(valor: str) -> str:
    if not isinstance(valor, str) or not valor or len(valor) > 240:
        raise ValueError("Caminho inválido na atualização.")
    partes = valor.split("/")
    for parte in partes:
        if (parte in {"", ".", ".."} or parte.endswith((".", " "))
                or re.search(r'[<>:"\\|?*\x00-\x1f]', parte)
                or re.fullmatch(r"(?:con|prn|aux|nul|com[1-9]|lpt[1-9])(?:\..*)?", parte, re.I)):
            raise ValueError("Caminho inseguro na atualização.")
    permitido = (
        valor == EXECUTAVEL
        or (len(partes) > 1 and partes[0] == "_internal")
        or valor in {"modelos/modelo_oficial.docx", "modelos/estilos.json"}
    )
    if not permitido:
        raise ValueError(f"Arquivo fora do programa: {valor}")
    return valor


def dentro(raiz: Path, relativo: str) -> Path:
    raiz = raiz.absolute()
    alvo = raiz.joinpath(*relativo.split("/"))
    if not alvo.resolve().is_relative_to(raiz.resolve()):
        raise ValueError("O caminho ultrapassa a pasta do programa.")
    for parte in (raiz, *alvo.parents, alvo):
        if parte == raiz or parte.is_relative_to(raiz):
            if parte.is_symlink() or getattr(parte, "is_junction", lambda: False)():
                raise ValueError("Links e junções não são permitidos na atualização.")
    return alvo


def sha256(arquivo: Path) -> str:
    digest = hashlib.sha256()
    with arquivo.open("rb") as stream:
        for bloco in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(bloco)
    return digest.hexdigest()


def conferir_arquivo(arquivo: Path, item: dict) -> bool:
    return (arquivo.is_file() and not arquivo.is_symlink()
            and arquivo.stat().st_size == item["tamanho"] and sha256(arquivo) == item["sha256"])


def validar_manifesto(dados: dict, numero: str | None = None) -> dict:
    if not isinstance(dados, dict) or dados.get("protocolo") != 1 or dados.get("repositorio") != REPOSITORIO:
        raise ValueError("Manifesto de atualização incompatível.")
    atual = dados.get("versao")
    versao(atual)
    if numero is not None and atual != numero:
        raise ValueError("A versão do manifesto difere da publicação.")
    arquivos = dados.get("arquivos")
    if not isinstance(arquivos, list) or not 1 <= len(arquivos) <= 10000:
        raise ValueError("Lista de arquivos inválida.")
    vistos, total = set(), 0
    for item in arquivos:
        if not isinstance(item, dict):
            raise ValueError("Arquivo inválido no manifesto.")
        nome = caminho_relativo(item.get("caminho"))
        if nome.casefold() in vistos:
            raise ValueError("Arquivos duplicados no manifesto.")
        vistos.add(nome.casefold())
        tamanho = item.get("tamanho")
        if type(tamanho) is not int or not 0 <= tamanho <= MAX_ARQUIVO:
            raise ValueError("Tamanho de arquivo inválido.")
        total += tamanho
        if not isinstance(item.get("sha256"), str) or not re.fullmatch(r"[a-f0-9]{64}", item["sha256"]):
            raise ValueError("Hash inválido no manifesto.")
        origem = item.get("pacote")
        if not isinstance(origem, str) or versao(origem) > versao(atual):
            raise ValueError("Origem de arquivo inválida.")
    if total > MAX_TOTAL or EXECUTAVEL.casefold() not in vistos:
        raise ValueError("Pacote incompleto ou excessivamente grande.")
    return dados


def ler_manifesto(caminho: Path, numero: str | None = None) -> dict:
    if caminho.stat().st_size > MAX_MANIFESTO:
        raise ValueError("Manifesto excessivamente grande.")
    return validar_manifesto(json.loads(caminho.read_text(encoding="utf-8")), numero)


def verificar_pasta(pasta: Path, dados: dict) -> None:
    validar_manifesto(dados, pasta.name)
    for item in dados["arquivos"]:
        if not conferir_arquivo(dentro(pasta, item["caminho"]), item):
            raise ValueError(f"Arquivo ausente ou alterado: {item['caminho']}")


def gravar_json_atomico(caminho: Path, dados: dict) -> None:
    temporario = caminho.with_name(caminho.name + ".tmp")
    if caminho.is_symlink() or temporario.is_symlink():
        raise ValueError("Link inválido no estado da atualização.")
    with temporario.open("w", encoding="utf-8") as stream:
        json.dump(dados, stream, ensure_ascii=False, indent=2)
        stream.flush()
        os.fsync(stream.fileno())
    temporario.replace(caminho)
