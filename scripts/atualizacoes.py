from __future__ import annotations

import hashlib
import json
import os
import re
import shutil
import stat
import subprocess
import tempfile
import zipfile
from dataclasses import dataclass
from pathlib import Path
from urllib.error import HTTPError
from urllib.parse import urlparse
from urllib.request import Request, urlopen

from app_version import INSTALLER_NAME, REPOSITORY, VERSION
from scripts.protocolo_atualizacao import (
    MANIFESTO, MAX_MANIFESTO, MAX_TOTAL, INICIADOR,
    caminho_relativo, conferir_arquivo, dentro, ler_manifesto,
    nome_pacote, sha256, validar_manifesto, verificar_pasta,
)

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
    tipo: str = "instalador"
    nome: str = INSTALLER_NAME


@dataclass(frozen=True)
class AtualizacaoPreparada:
    raiz: Path
    versao: str
    manifesto_sha256: str
    bytes_baixados: int


def interpretar_asset(release: dict, nome: str, limite: int = MAX_INSTALLER_BYTES) -> Atualizacao:
    tag = release["tag_name"]
    versao_numerica(tag)
    encontrados = [a for a in release.get("assets", []) if a.get("name") == nome]
    if len(encontrados) != 1:
        raise ValueError(f"Arquivo de atualização ausente ou duplicado: {nome}")
    asset = encontrados[0]
    url = asset.get("browser_download_url", "")
    parsed = urlparse(url)
    esperado = f"/{REPOSITORY}/releases/download/{tag}/{nome}"
    digest, tamanho = asset.get("digest") or "", asset.get("size", 0)
    if (parsed.scheme != "https" or parsed.netloc != "github.com"
            or parsed.path != esperado or parsed.query or parsed.fragment):
        raise ValueError("O arquivo não pertence ao repositório oficial.")
    if not re.fullmatch(r"sha256:[a-fA-F0-9]{64}", digest):
        raise ValueError("A publicação não contém uma verificação SHA-256 válida do GitHub.")
    if type(tamanho) is not int or not 0 < tamanho <= limite:
        raise ValueError("Tamanho do arquivo de atualização inválido.")
    return Atualizacao(tag.removeprefix("v"), url, digest[7:].lower(), tamanho,
                       "incremental" if nome == MANIFESTO else "instalador", nome)


def interpretar_release(release: dict, atual: str = VERSION) -> Atualizacao | None:
    if release.get("draft") or release.get("prerelease"):
        return None
    tag = release.get("tag_name", "")
    if versao_numerica(tag) <= versao_numerica(atual):
        return None
    if any(a.get("name") == MANIFESTO for a in release.get("assets", [])):
        return interpretar_asset(release, MANIFESTO, MAX_MANIFESTO)
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
    return baixar_verificado(atualizacao, destino)


def baixar_verificado(atualizacao: Atualizacao, destino: Path) -> Path:
    destino.mkdir(parents=True, exist_ok=True)
    if Path(atualizacao.nome).name != atualizacao.nome or "/" in atualizacao.nome or "\\" in atualizacao.nome:
        raise ValueError("Nome do download inválido.")
    arquivo = destino / atualizacao.nome
    parcial = arquivo.with_suffix(".download")
    if arquivo.is_symlink() or parcial.is_symlink():
        raise ValueError("Link inválido na pasta de atualização.")
    if arquivo.is_file() and arquivo.stat().st_size == atualizacao.tamanho and sha256(arquivo) == atualizacao.sha256:
        return arquivo
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


def buscar_release(numero: str | None = None) -> dict:
    if numero is not None:
        versao_numerica(numero)
    endpoint = f"tags/v{numero}" if numero else "latest"
    request = Request(
        f"https://api.github.com/repos/{REPOSITORY}/releases/{endpoint}",
        headers={"Accept": "application/vnd.github+json", "User-Agent": f"FormatadorPeticoes/{VERSION}"},
    )
    with urlopen(request, timeout=20) as response:
        dados = response.read(MAX_MANIFESTO + 1)
    if len(dados) > MAX_MANIFESTO:
        raise ValueError("Resposta do GitHub excessivamente grande.")
    release = json.loads(dados)
    if release.get("draft") or release.get("prerelease"):
        raise ValueError("A versão ainda não foi publicada como estável.")
    if numero is not None and release.get("tag_name") != f"v{numero}":
        raise ValueError("A publicação não corresponde à versão solicitada.")
    return release


def _extrair_necessarios(pacote: Path, itens: list[dict], destino: Path) -> None:
    with zipfile.ZipFile(pacote) as arquivo:
        membros, vistos, total = {}, set(), 0
        for info in arquivo.infolist():
            nome = caminho_relativo(info.filename)
            modo = stat.S_IFMT(info.external_attr >> 16)
            if nome.casefold() in vistos or modo not in {0, stat.S_IFREG}:
                raise ValueError("Pacote contém arquivo duplicado, diretório ou link.")
            vistos.add(nome.casefold())
            total += info.file_size
            if total > MAX_TOTAL:
                raise ValueError("Pacote descompactado excessivamente grande.")
            membros[nome] = info
        for item in itens:
            info = membros.get(item["caminho"])
            if info is None or info.file_size != item["tamanho"]:
                raise ValueError(f"Arquivo incorreto no pacote: {item['caminho']}")
            alvo = dentro(destino, item["caminho"])
            alvo.parent.mkdir(parents=True, exist_ok=True)
            total, digest = 0, hashlib.sha256()
            with arquivo.open(info) as origem, alvo.open("xb") as stream:
                for bloco in iter(lambda: origem.read(1024 * 1024), b""):
                    total += len(bloco)
                    if total > item["tamanho"]:
                        raise ValueError("Arquivo descompactado excedeu o tamanho esperado.")
                    digest.update(bloco)
                    stream.write(bloco)
            if total != item["tamanho"] or digest.hexdigest() != item["sha256"]:
                raise ValueError(f"Arquivo adulterado: {item['caminho']}")


def preparar_incremental(atualizacao: Atualizacao, pasta_atual: Path, cache: Path) -> AtualizacaoPreparada:
    if atualizacao.tipo != "incremental":
        raise ValueError("A publicação não oferece atualização pelo próprio programa.")
    pasta_atual = pasta_atual.absolute()
    if pasta_atual.parent.name != "versoes" or not (pasta_atual.parent.parent / INICIADOR).is_file():
        raise ValueError("Instale uma vez a versão com o novo atualizador para habilitar as próximas atualizações.")
    raiz = pasta_atual.parent.parent
    dentro(raiz, "versoes/" + pasta_atual.name)
    manifesto = baixar_verificado(atualizacao, cache / atualizacao.versao)
    dados = ler_manifesto(manifesto, atualizacao.versao)
    if versao_numerica(dados["versao"]) <= versao_numerica(pasta_atual.name):
        raise ValueError("Atualização anterior ou igual à instalação atual.")
    final = dentro(raiz, "versoes/" + atualizacao.versao)
    if final.exists():
        if sha256(final / MANIFESTO) != atualizacao.sha256:
            raise ValueError("Já existe uma pasta diferente para essa versão; a instalação atual foi preservada.")
        verificar_pasta(final, dados)
        return AtualizacaoPreparada(raiz, atualizacao.versao, atualizacao.sha256, 0)
    preparo = Path(tempfile.mkdtemp(prefix=".preparando-", dir=raiz / "versoes"))
    baixados = atualizacao.tamanho
    try:
        grupos = {}
        for item in dados["arquivos"]:
            origem = dentro(pasta_atual, item["caminho"])
            alvo = dentro(preparo, item["caminho"])
            if conferir_arquivo(origem, item):
                alvo.parent.mkdir(parents=True, exist_ok=True)
                # A cópia evita compartilhar arquivos editáveis entre versões.
                shutil.copy2(origem, alvo)
            else:
                grupos.setdefault(item["pacote"], []).append(item)
        for numero, itens in grupos.items():
            release = buscar_release(numero)
            asset = interpretar_asset(release, nome_pacote(numero))
            pacote = baixar_verificado(asset, cache / numero)
            baixados += asset.tamanho
            _extrair_necessarios(pacote, itens, preparo)
        shutil.copy2(manifesto, preparo / MANIFESTO)
        # A pasta temporária tem outro nome; a versão é conferida ao ativar.
        for item in dados["arquivos"]:
            if not conferir_arquivo(dentro(preparo, item["caminho"]), item):
                raise ValueError("A atualização ficou incompleta; a versão atual foi preservada.")
        preparo.rename(final)
        return AtualizacaoPreparada(raiz, atualizacao.versao, atualizacao.sha256, baixados)
    finally:
        if preparo.exists():
            dentro(raiz, "versoes/" + preparo.name)
            shutil.rmtree(preparo)


def iniciar_atualizacao(preparada: AtualizacaoPreparada) -> None:
    iniciador = dentro(preparada.raiz, INICIADOR)
    subprocess.Popen([
        str(iniciador), "--ativar", preparada.versao,
        "--manifesto-sha256", preparada.manifesto_sha256,
        "--aguardar-pid", str(os.getpid()),
    ], close_fds=True, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
