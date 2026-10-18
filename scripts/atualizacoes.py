from __future__ import annotations

import hashlib
import json
import re
from dataclasses import dataclass
from pathlib import Path
from urllib.error import HTTPError
from urllib.parse import urlparse
from urllib.request import Request, urlopen

from app_version import INSTALLER_NAME, REPOSITORY, VERSION

MAX_INSTALLER_BYTES = 300 * 1024 * 1024


def versao_numerica(valor: str) -> tuple[int, int, int]:
    if not re.fullmatch(r"v?\d+\.\d+\.\d+", valor):
        raise ValueError("Versão inválida; use MAIOR.MENOR.CORREÇÃO.")
    return tuple(int(p) for p in valor.removeprefix("v").split("."))


@dataclass(frozen=True)
class Atualizacao:
    versao: str
    url: str
    sha256: str
    tamanho: int


def interpretar_release(release: dict, atual: str = VERSION) -> Atualizacao | None:
    if release.get("draft") or release.get("prerelease"):
        return None
    tag = release.get("tag_name", "")
    if versao_numerica(tag) <= versao_numerica(atual):
        return None
    for asset in release.get("assets", []):
        if asset.get("name") != INSTALLER_NAME:
            continue
        url = asset.get("browser_download_url", "")
        parsed = urlparse(url)
        esperado = f"/{REPOSITORY}/releases/download/{tag}/{INSTALLER_NAME}"
        digest = asset.get("digest") or ""
        tamanho = asset.get("size", 0)
        if (parsed.scheme != "https" or parsed.netloc != "github.com"
                or parsed.path != esperado or parsed.query or parsed.fragment):
            raise ValueError("O instalador não pertence ao repositório oficial.")
        if not re.fullmatch(r"sha256:[a-fA-F0-9]{64}", digest):
            raise ValueError("A versão não contém uma verificação SHA-256 válida do GitHub.")
        if not isinstance(tamanho, int) or not 0 < tamanho <= MAX_INSTALLER_BYTES:
            raise ValueError("Tamanho do instalador inválido.")
        return Atualizacao(tag.removeprefix("v"), url, digest[7:].lower(), tamanho)
    raise ValueError("A nova versão ainda não possui instalador para Windows.")


def consultar_atualizacao() -> Atualizacao | None:
    request = Request(
        f"https://api.github.com/repos/{REPOSITORY}/releases/latest",
        headers={"Accept": "application/vnd.github+json", "User-Agent": f"FormatadorPeticoes/{VERSION}"},
    )
    try:
        with urlopen(request, timeout=20) as response:
            release = json.load(response)
    except HTTPError as exc:
        if exc.code == 404:
            raise RuntimeError("Nenhuma versão pública encontrada. Verifique se o repositório e a release estão acessíveis.") from exc
        raise
    return interpretar_release(release)


def baixar_instalador(atualizacao: Atualizacao, destino: Path) -> Path:
    destino.mkdir(parents=True, exist_ok=True)
    arquivo = destino / INSTALLER_NAME
    parcial = arquivo.with_suffix(".download")
    digest = hashlib.sha256()
    total = 0
    try:
        request = Request(atualizacao.url, headers={"User-Agent": f"FormatadorPeticoes/{VERSION}"})
        with urlopen(request, timeout=60) as response, parcial.open("wb") as stream:
            while bloco := response.read(1024 * 1024):
                total += len(bloco)
                if total > atualizacao.tamanho or total > MAX_INSTALLER_BYTES:
                    raise ValueError("O download excedeu o tamanho publicado.")
                stream.write(bloco)
                digest.update(bloco)
        if total != atualizacao.tamanho or digest.hexdigest() != atualizacao.sha256:
            raise ValueError("A verificação do instalador falhou. Nenhuma atualização foi executada.")
        parcial.replace(arquivo)
        return arquivo
    finally:
        parcial.unlink(missing_ok=True)
