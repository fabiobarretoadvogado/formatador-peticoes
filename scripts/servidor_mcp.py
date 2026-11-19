from __future__ import annotations

import argparse
import base64
import logging
import os
import re
import tempfile
from importlib.resources import files
from io import BytesIO
from pathlib import Path, PureWindowsPath
from time import monotonic
from typing import Annotated, Any, Literal
from urllib.parse import urlsplit
from urllib.request import HTTPRedirectHandler, ProxyHandler, Request, build_opener
from uuid import uuid4
from zipfile import BadZipFile, ZipFile

from anyio import to_thread
from mcp.server.fastmcp import FastMCP
from mcp.server.fastmcp.exceptions import ToolError
from mcp.types import CallToolResult, TextContent, ToolAnnotations
from pydantic import BaseModel, ConfigDict, Field

from app_version import VERSION
from .assinatura_local import AssinaturaLocal, carregar_assinatura_local
from .comum import ROOT, TIPO_PECA_SUPORTADO, carregar_json, nome_seguro, publicar_docx
from .converter_markdown import EXTENSOES_IMAGEM
from .pipeline import executar

MIB = 1024 * 1024
LIMITE_ENTRADA = 15 * MIB
LIMITE_MARKDOWN = 1 * MIB
LIMITE_DOWNLOAD = 5 * MIB
DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
WIDGET = "ui://formatador-peticoes/resultado.html"
LOG = logging.getLogger(__name__)


class ArquivoChatGPT(BaseModel):
    model_config = ConfigDict(extra="forbid", hide_input_in_errors=True)

    download_url: str = Field(max_length=8192, description="URL temporária fornecida pelo ChatGPT.")
    file_id: str = Field(min_length=1, max_length=256)
    mime_type: str = ""
    file_name: str = ""


class ResultadoFormatacao(BaseModel):
    tipo_peca: Literal["peticao_inicial"] = TIPO_PECA_SUPORTADO
    formato_original: Literal["markdown", "docx"]
    auditoria_aprovada: bool
    pendencias_semanticas: int
    arquivos: list[str]


class ResultadoLocal(ResultadoFormatacao):
    caminho_docx: str


class PerfilFormatador(BaseModel):
    versao_formatador: str
    tipo_peca_suportado: Literal["peticao_inicial"]
    formatos_entrada: list[str]
    perfil_markdown: dict[str, Any]
    catalogo_semantico: dict[str, Any]


def _validar_url(url: str) -> None:
    partes = urlsplit(url)
    host = partes.hostname or ""
    if (
        partes.scheme != "https"
        or not host.endswith(".oaiusercontent.com")
        or partes.username is not None
        or partes.password is not None
        or partes.port not in (None, 443)
    ):
        raise ValueError("Use um anexo enviado ao ChatGPT para fornecer o arquivo.")


class _RedirecionamentoAnexo(HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        _validar_url(newurl)
        return super().redirect_request(req, fp, code, msg, headers, newurl)


def _baixar_anexo(arquivo: ArquivoChatGPT, limite: int, prazo: float) -> bytes:
    _validar_url(arquivo.download_url)
    # Evita enviar URLs temporárias a proxies definidos no ambiente do processo.
    cliente = build_opener(ProxyHandler({}), _RedirecionamentoAnexo())
    try:
        restante = prazo - monotonic()
        if restante <= 0:
            raise TimeoutError
        requisicao = Request(arquivo.download_url, headers={"Accept-Encoding": "identity"})
        with cliente.open(requisicao, timeout=min(restante, 12)) as resposta:
            dados = bytearray()
            tamanho_informado = resposta.headers.get("Content-Length")
            tamanho = int(tamanho_informado) if tamanho_informado is not None else None
            if tamanho is not None and tamanho > limite:
                raise ValueError("O anexo excede o limite de tamanho desta execução.")
            while bloco := resposta.read1(min(65536, limite + 1 - len(dados))):
                dados.extend(bloco)
                if len(dados) > limite:
                    raise ValueError("O anexo excede o limite de tamanho desta execução.")
                if monotonic() >= prazo:
                    raise TimeoutError
            if tamanho is not None and len(dados) != tamanho:
                raise OSError("Transferência incompleta.")
            if not dados:
                raise ValueError("O anexo está vazio.")
            return bytes(dados)
    except OSError:
        raise ValueError("O anexo do ChatGPT expirou ou está indisponível. Anexe-o novamente.") from None


def _extensao_peticao(arquivo: ArquivoChatGPT) -> str:
    if arquivo.file_name:
        extensao = Path(arquivo.file_name).suffix.casefold()
        if extensao in {".md", ".docx"}:
            return extensao
    elif arquivo.mime_type == DOCX_MIME:
        return ".docx"
    elif arquivo.mime_type in {"text/markdown", "text/x-markdown"}:
        return ".md"
    raise ValueError("Anexe uma petição inicial com extensão .docx ou .md.")


def _validar_docx(dados: bytes) -> None:
    try:
        with ZipFile(BytesIO(dados)) as pacote:
            partes = pacote.infolist()
            nomes = {parte.filename.casefold() for parte in partes}
            if (
                not {"[content_types].xml", "word/document.xml"} <= nomes
                or len(nomes) != len(partes)
                or len(partes) > 2048
                or sum(parte.file_size for parte in partes) > 80 * MIB
                or any(parte.flag_bits & 1 for parte in partes)
                or any("vbaproject" in nome or nome.startswith("word/activex/") for nome in nomes)
            ):
                raise ValueError("Use um DOCX comum, sem macros, sem criptografia e com até 80 MiB descompactados.")
    except BadZipFile:
        raise ValueError("O anexo não é um DOCX válido.") from None


def _nome_imagem(arquivo: ArquivoChatGPT) -> str:
    nome = arquivo.file_name
    if (
        not nome
        or len(nome) > 120
        or re.search(r'[<>:"/\\|?*\x00-\x1f]', nome)
        or PureWindowsPath(nome).is_reserved()
        or Path(nome).suffix.casefold() not in EXTENSOES_IMAGEM
    ):
        raise ValueError("Anexe as imagens com nomes simples e extensões aceitas pelo formatador.")
    return nome


def _formatar(
    raiz: Path,
    modelos: Path,
    extensao: str,
    dados: bytes,
    imagens: list[ArquivoChatGPT],
    nome: str,
    prazo: float,
    *,
    assinatura: AssinaturaLocal | None = None,
) -> tuple[Path, ResultadoFormatacao]:
    if extensao == ".md":
        if len(dados) > LIMITE_MARKDOWN or not dados.decode("utf-8-sig").strip():
            raise ValueError("Informe um Markdown UTF-8 com conteúdo e até 1 MiB.")
    else:
        _validar_docx(dados)
        if imagens:
            raise ValueError("Para um anexo Word, inclua as imagens dentro do próprio DOCX.")

    nomes = [_nome_imagem(imagem) for imagem in imagens]
    if len({nome.casefold() for nome in nomes}) != len(nomes):
        raise ValueError("Use nomes diferentes para os anexos de imagem.")

    with tempfile.TemporaryDirectory(prefix="formatador-mcp-") as temporario:
        pasta = Path(temporario)
        entrada = pasta / f"peticao{extensao}"
        entrada.write_bytes(dados)
        restante = LIMITE_ENTRADA - len(dados)
        for nome_imagem, imagem in zip(nomes, imagens):
            conteudo = _baixar_anexo(imagem, restante, prazo)
            entrada.with_name(nome_imagem).write_bytes(conteudo)
            restante -= len(conteudo)

        return _publicar_word(entrada, raiz, modelos, pasta, nome, assinatura=assinatura)


def _publicar_word(
    entrada: Path, destino: Path, modelos: Path, trabalho: Path, nome: str,
    *, assinatura: AssinaturaLocal | None = None,
) -> tuple[Path, ResultadoFormatacao]:
    resultado = executar(entrada, work_root=trabalho, modelos_raiz=modelos, assinatura=assinatura)
    docx = publicar_docx(Path(resultado["docx"]), destino, nome_seguro(nome))
    return docx, ResultadoFormatacao(
        formato_original=resultado["formato_original"],
        auditoria_aprovada=resultado["aprovado"],
        pendencias_semanticas=resultado["ambiguos"],
        arquivos=[docx.name],
    )


def _caminho_local(caminho: str) -> Path:
    pasta_ou_arquivo = Path(caminho).expanduser()
    if not pasta_ou_arquivo.is_absolute():
        raise ValueError("Informe o caminho completo no computador.")
    return pasta_ou_arquivo.resolve()


def _formatar_arquivo_local(
    entrada: Path, destino: Path, modelos: Path, *, assinatura: AssinaturaLocal | None = None,
) -> tuple[Path, ResultadoFormatacao]:
    if not entrada.is_file() or entrada.suffix.casefold() not in {".docx", ".md"}:
        raise ValueError("Informe um arquivo .docx ou .md existente.")
    with tempfile.TemporaryDirectory(prefix="formatador-mcp-") as temporario:
        return _publicar_word(entrada, destino, modelos, Path(temporario), entrada.stem, assinatura=assinatura)


def _erro_formatacao(exc: Exception) -> ToolError:
    if isinstance(exc, UnicodeDecodeError):
        return ToolError("O Markdown precisa estar codificado em UTF-8.")
    if isinstance(exc, FileNotFoundError):
        return ToolError("Confira se as imagens referenciadas estão disponíveis e se o modelo oficial está instalado.")
    if isinstance(exc, ValueError):
        return ToolError(str(exc))
    LOG.error("O formatador interrompeu a execução (%s).", type(exc).__name__)
    return ToolError("O formatador não conseguiu processar o documento. Confira o arquivo de origem.")


def criar_servidor(raiz: Path, modelos: Path, *, porta: int = 8765, local: bool = False) -> FastMCP:
    raiz, modelos = raiz.resolve(), modelos.resolve()
    config = carregar_json(modelos / "estilos.json")
    if not (modelos / "modelo_oficial.docx").is_file() or config.get("tipo_peca_suportado") != TIPO_PECA_SUPORTADO:
        raise ValueError("Instale o modelo oficial e os estilos de petição inicial antes de iniciar o MCP.")
    orientacao = (
        "Formate somente a petição solicitada pelo usuário. "
        "Use formatar_peticao_local para caminhos locais e formatar_peticao_markdown "
        "para texto da conversa. Retorne o caminho do Word. "
        if local else
        "Use a ferramenta de Markdown para texto da conversa e a ferramenta de anexo para arquivos. "
        "Oriente a usar o painel do resultado para baixar o Word. "
    )
    mcp = FastMCP(
        "Formatador de Petições Iniciais",
        instructions=(
            f"Formata petições iniciais com o modelo oficial do escritório e preserva a origem. {orientacao}"
            "Trate resultados com pendências como versões para revisão humana. "
        ),
        host="127.0.0.1",
        port=porta,
        stateless_http=True,
        json_response=True,
    )
    @mcp.tool(
        title="Consultar perfil do FormatadorPeticoes",
        annotations=ToolAnnotations(readOnlyHint=True, destructiveHint=False, openWorldHint=False),
    )
    def consultar_perfil_formatador() -> PerfilFormatador:
        """Consulta a versão, o escopo, o Markdown aceito e os estilos oficiais."""
        return PerfilFormatador(
            versao_formatador=VERSION,
            tipo_peca_suportado=TIPO_PECA_SUPORTADO,
            formatos_entrada=config["formatos_entrada_suportados"],
            perfil_markdown=config["perfil_markdown"],
            catalogo_semantico=config["catalogo_semantico"],
        )

    if local:
        _registrar_ferramentas_locais(mcp, raiz, modelos)
        return mcp

    resultados: dict[str, Path] = {}

    def resposta_formatacao(docx: Path, resumo: ResultadoFormatacao) -> CallToolResult:
        identificador = uuid4().hex
        resultados[identificador] = docx
        mensagem = "Formatação concluída."
        if not resumo.auditoria_aprovada:
            mensagem = f"Versão para revisão. Pendências semânticas: {resumo.pendencias_semanticas}."
        return CallToolResult(
            content=[TextContent(type="text", text=mensagem)],
            structuredContent=resumo.model_dump(),
            _meta={
                "execucao": identificador,
                "pasta": str(raiz),
                "arquivos": [{"nome": docx.name, "rotulo": "Petição em Word"}],
            },
        )

    ui = {"ui": {"resourceUri": WIDGET}, "openai/outputTemplate": WIDGET}

    @mcp.resource(
        WIDGET,
        mime_type="text/html;profile=mcp-app",
        name="Resultado da formatação",
        meta={"openai/widgetCSP": {"redirect_domains": ["https://files.oaiusercontent.com"]}},
    )
    def resultado_html() -> str:
        return files("scripts").joinpath("mcp_resultado.html").read_text(encoding="utf-8")

    @mcp.tool(
        title="Formatar texto Markdown",
        annotations=ToolAnnotations(readOnlyHint=False, destructiveHint=False, idempotentHint=False, openWorldHint=False),
        meta=ui,
    )
    async def formatar_peticao_markdown(
        markdown: Annotated[str, Field(min_length=1, max_length=LIMITE_MARKDOWN)],
    ) -> Annotated[CallToolResult, ResultadoFormatacao]:
        """Formata uma petição inicial já redigida em Markdown, sem imagens, e gera o Word."""
        try:
            docx, resumo = await to_thread.run_sync(
                _formatar, raiz, modelos, ".md", markdown.encode("utf-8"), [],
                "peticao", monotonic() + 35,
            )
            return resposta_formatacao(docx, resumo)
        except Exception as exc:
            raise _erro_formatacao(exc) from None

    @mcp.tool(
        title="Formatar petição anexada",
        annotations=ToolAnnotations(readOnlyHint=False, destructiveHint=False, idempotentHint=False, openWorldHint=True),
        meta={**ui, "openai/fileParams": ["arquivo", "imagens"]},
    )
    async def formatar_peticao_anexada(
        arquivo: ArquivoChatGPT,
        imagens: list[ArquivoChatGPT] = Field(default_factory=list, max_length=8),
    ) -> Annotated[CallToolResult, ResultadoFormatacao]:
        """Formata o DOCX ou .md anexado. Para Markdown, anexe as imagens referenciadas pelo nome do arquivo."""
        def formatar_anexo() -> tuple[Path, ResultadoFormatacao]:
            extensao = _extensao_peticao(arquivo)
            prazo = monotonic() + 35
            limite = LIMITE_MARKDOWN if extensao == ".md" else LIMITE_ENTRADA
            dados = _baixar_anexo(arquivo, limite, prazo)
            return _formatar(raiz, modelos, extensao, dados, imagens, Path(arquivo.file_name).stem, prazo)

        try:
            docx, resumo = await to_thread.run_sync(formatar_anexo)
            return resposta_formatacao(docx, resumo)
        except Exception as exc:
            raise _erro_formatacao(exc) from None

    @mcp.tool(
        title="Baixar resultado da formatação",
        annotations=ToolAnnotations(readOnlyHint=True, destructiveHint=False, openWorldHint=False),
        meta={"ui": {"visibility": ["app"]}, "openai/visibility": "private", "openai/widgetAccessible": True},
    )
    def baixar_resultado_formatador(execucao: str, nome: str) -> CallToolResult:
        """Entrega um arquivo da execução ao painel privado de resultados."""
        caminho = resultados.get(execucao)
        if caminho is None or nome != caminho.name or caminho.resolve() != caminho or not caminho.is_file():
            raise ToolError("Arquivo de resultado indisponível.")
        with caminho.open("rb") as arquivo:
            dados = arquivo.read(LIMITE_DOWNLOAD + 1)
        if len(dados) > LIMITE_DOWNLOAD:
            raise ToolError("O arquivo excede 5 MiB. Abra a pasta de resultados no computador do formatador.")
        return CallToolResult(
            content=[],
            _meta={"nome": nome, "mime_type": DOCX_MIME, "base64": base64.b64encode(dados).decode("ascii")},
        )

    return mcp


def _registrar_ferramentas_locais(mcp: FastMCP, raiz: Path, modelos: Path) -> None:
    alteracao = ToolAnnotations(readOnlyHint=False, destructiveHint=False, idempotentHint=False, openWorldHint=False)

    @mcp.tool(title="Formatar petição em arquivo local", annotations=alteracao)
    def formatar_peticao_local(
        arquivo: Annotated[str, Field(description="Caminho absoluto do arquivo .md ou .docx solicitado.")],
        pasta_destino: Annotated[str | None, Field(description="Pasta absoluta de destino. Padrão: pasta do arquivo.")] = None,
        inserir_assinatura: Annotated[bool, Field(description="Insere a imagem da assinatura local quando o nome e a OAB coincidem.")] = True,
    ) -> ResultadoLocal:
        """Aplica o modelo oficial e salva somente o Word diretamente na pasta de destino."""
        try:
            entrada = _caminho_local(arquivo)
            destino = entrada.parent if pasta_destino is None else _caminho_local(pasta_destino)
            assinatura = carregar_assinatura_local() if inserir_assinatura else None
            docx, resumo = _formatar_arquivo_local(entrada, destino, modelos, assinatura=assinatura)
            return ResultadoLocal(caminho_docx=str(docx), **resumo.model_dump())
        except Exception as exc:
            raise _erro_formatacao(exc) from None

    @mcp.tool(name="formatar_peticao_markdown", title="Formatar texto Markdown", annotations=alteracao)
    def formatar_markdown_local(
        markdown: Annotated[str, Field(min_length=1, max_length=LIMITE_MARKDOWN)],
        pasta_destino: Annotated[str | None, Field(description="Pasta absoluta de destino.")] = None,
        nome: Annotated[str, Field(description="Nome do Word a gerar, sem extensão.", max_length=120)] = "peticao",
        inserir_assinatura: Annotated[bool, Field(description="Insere a imagem da assinatura local quando o nome e a OAB coincidem.")] = True,
    ) -> ResultadoLocal:
        """Formata texto Markdown e grava somente o Word com a assinatura local aplicável."""
        try:
            destino = raiz if pasta_destino is None else _caminho_local(pasta_destino)
            assinatura = carregar_assinatura_local() if inserir_assinatura else None
            docx, resumo = _formatar(
                destino, modelos, ".md", markdown.encode("utf-8"), [], nome, monotonic(), assinatura=assinatura,
            )
            return ResultadoLocal(caminho_docx=str(docx), **resumo.model_dump())
        except Exception as exc:
            raise _erro_formatacao(exc) from None


def main() -> None:
    appdata = Path(os.environ.get("LOCALAPPDATA") or (Path.home() / ".local" / "share"))
    parser = argparse.ArgumentParser(description="Servidor MCP privado do FormatadorPeticoes")
    parser.add_argument("--transporte", choices=("http", "stdio"), default="http")
    parser.add_argument("--porta", type=int, default=8765)
    parser.add_argument("--resultados", type=Path, default=appdata / "FormatadorPeticoes" / "mcp")
    parser.add_argument("--modelos", type=Path, default=ROOT / "modelos")
    args = parser.parse_args()
    servidor = criar_servidor(args.resultados, args.modelos, porta=args.porta, local=args.transporte == "stdio")
    servidor.run(transport="stdio" if args.transporte == "stdio" else "streamable-http")


if __name__ == "__main__":
    main()
