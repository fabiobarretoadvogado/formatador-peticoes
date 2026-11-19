import asyncio
import base64
from io import BytesIO
from pathlib import Path

import pytest
from docx import Document
from docx.oxml.ns import qn
from PIL import Image

pytest.importorskip("mcp")

from mcp.shared.memory import create_connected_server_and_client_session

from scripts import servidor_mcp


MODELOS = Path(__file__).resolve().parents[1] / "modelos"
MARKDOWN = (
    "AO JUÍZO DA VARA CÍVEL DA COMARCA DE TESTE\n\n"
    "Ação de cobrança\n\n"
    "## Dos fatos\n\n"
    "A parte autora apresentou os comprovantes referentes à obrigação discutida.\n\n"
    "## Dos pedidos\n\n"
    "Ante o exposto, requer:\n\n"
    "1. A procedência do pedido.\n\n"
    "Termos em que, pede deferimento.\n"
)


def test_fluxo_markdown_no_mcp(tmp_path):
    servidor = servidor_mcp.criar_servidor(tmp_path, MODELOS)

    async def verificar():
        async with create_connected_server_and_client_session(servidor) as sessao:
            ferramentas = {tool.name: tool for tool in (await sessao.list_tools()).tools}
            anexos = ferramentas["formatar_peticao_anexada"]
            assert anexos.meta["openai/fileParams"] == ["arquivo", "imagens"]
            esquema_arquivo = anexos.inputSchema["$defs"]["ArquivoChatGPT"]
            assert set(esquema_arquivo["properties"]) == {"download_url", "file_id", "mime_type", "file_name"}
            assert set(esquema_arquivo["required"]) == {"download_url", "file_id"}
            assert anexos.inputSchema["properties"]["imagens"]["type"] == "array"
            assert anexos.outputSchema

            painel = await sessao.read_resource(anexos.meta["ui"]["resourceUri"])
            assert painel.contents[0].mimeType == "text/html;profile=mcp-app"
            assert "baixar_resultado_formatador" in painel.contents[0].text

            perfil = await sessao.call_tool("consultar_perfil_formatador", {})
            assert not perfil.isError
            assert perfil.structuredContent["perfil_markdown"]["titulos"] == ["##", "###"]

            formatado = await sessao.call_tool("formatar_peticao_markdown", {"markdown": MARKDOWN})
            assert not formatado.isError, formatado.content
            assert formatado.structuredContent["formato_original"] == "markdown"
            assert "gerar_pdf" not in anexos.inputSchema["properties"]
            assert "execucao" not in formatado.structuredContent
            nome, = formatado.structuredContent["arquivos"]
            caminho = tmp_path / nome
            assert list(tmp_path.iterdir()) == [caminho]

            args = {"execucao": formatado.meta["execucao"], "nome": nome}
            word = await sessao.call_tool("baixar_resultado_formatador", args)
            assert not word.isError
            assert not word.content
            dados = base64.b64decode(word.meta["base64"])
            assert dados == caminho.read_bytes()
            documento = Document(BytesIO(dados))
            pedido = next(p for p in documento.paragraphs if p.text == "A procedência do pedido.")
            assert pedido.style.name == "FB - Pedido"
            numeracao = documento.part.numbering_part.element
            definicoes = list(numeracao)
            primeira_lista = definicoes.index(numeracao.find(qn("w:num")))
            assert all(definicoes.index(no) < primeira_lista for no in numeracao.findall(qn("w:abstractNum")))
            for argumentos in ({**args, "execucao": "../"}, {**args, "nome": "../entrada/peticao.md"}):
                assert (await sessao.call_tool("baixar_resultado_formatador", argumentos)).isError

    asyncio.run(verificar())


def test_rejeita_markdown_incompativel_e_url_externa(tmp_path):
    servidor = servidor_mcp.criar_servidor(tmp_path, MODELOS)

    async def verificar():
        async with create_connected_server_and_client_session(servidor) as sessao:
            markdown = await sessao.call_tool("formatar_peticao_markdown", {"markdown": "# Inválido"})
            assert markdown.isError
            assert "use apenas ## e ###" in markdown.content[0].text
            for url in (
                "http://files.oaiusercontent.com/peticao",
                "https://files.oaiusercontent.com@invalid.example/peticao",
                "https://invalid.example/peticao",
            ):
                anexo = await sessao.call_tool(
                    "formatar_peticao_anexada",
                    {"arquivo": {"download_url": url, "file_id": "file_teste", "file_name": "peticao.md"}},
                )
                assert anexo.isError
                assert "anexo enviado ao ChatGPT" in anexo.content[0].text
            assert not list(tmp_path.iterdir())

    asyncio.run(verificar())


def test_anexo_docx_sai_diretamente_na_pasta_selecionada(tmp_path, monkeypatch):
    origem = Document()
    origem.add_paragraph("AO JUÍZO DA VARA CÍVEL DA COMARCA DE TESTE")
    origem.add_paragraph("A parte autora apresentou os comprovantes referentes à obrigação discutida.")
    buffer = BytesIO()
    origem.save(buffer)
    dados = buffer.getvalue()

    def baixar_anexo(arquivo, limite, prazo):
        assert arquivo.file_id == "file_teste"
        assert len(dados) <= limite
        return dados

    monkeypatch.setattr(servidor_mcp, "_baixar_anexo", baixar_anexo)
    servidor = servidor_mcp.criar_servidor(tmp_path, MODELOS)

    async def verificar():
        async with create_connected_server_and_client_session(servidor) as sessao:
            anexo = await sessao.call_tool(
                "formatar_peticao_anexada",
                {
                    "arquivo": {
                        "download_url": "https://files.oaiusercontent.com/peticao",
                        "file_id": "file_teste",
                        "mime_type": servidor_mcp.DOCX_MIME,
                    },
                },
            )
            assert not anexo.isError, anexo.content
            assert anexo.structuredContent["formato_original"] == "docx"
            assert anexo.structuredContent["arquivos"] == ["peticao_formatada.docx"]
            assert [caminho.name for caminho in tmp_path.iterdir()] == ["peticao_formatada.docx"]

    asyncio.run(verificar())


def test_anexo_markdown_incorpora_imagem_local(tmp_path, monkeypatch):
    imagem = BytesIO()
    Image.new("RGB", (32, 24), "white").save(imagem, format="PNG")
    arquivos = {
        "file_texto": (MARKDOWN + "\n![Comprovante](comprovante.png)\n").encode(),
        "file_comprovante": imagem.getvalue(),
    }

    def anexo(file_id, nome):
        return {"download_url": f"https://files.oaiusercontent.com/{file_id}", "file_id": file_id, "file_name": nome}

    monkeypatch.setattr(servidor_mcp, "_baixar_anexo", lambda arquivo, limite, prazo: arquivos[arquivo.file_id])
    servidor = servidor_mcp.criar_servidor(tmp_path, MODELOS)

    async def verificar():
        async with create_connected_server_and_client_session(servidor) as sessao:
            resposta = await sessao.call_tool(
                "formatar_peticao_anexada",
                {"arquivo": anexo("file_texto", "peticao.md"), "imagens": [anexo("file_comprovante", "comprovante.png")]},
            )
            assert not resposta.isError, resposta.content
            caminho = tmp_path / "peticao_formatada.docx"
            documento = Document(caminho)
            assert len(documento.inline_shapes) == 1
            assert list(tmp_path.iterdir()) == [caminho]

    asyncio.run(verificar())


def test_mcp_local_publica_apenas_word_na_pasta_solicitada(tmp_path):
    origem = tmp_path / "modelo de teste.md"
    origem.write_text(MARKDOWN, encoding="utf-8")
    destino = tmp_path / "destino"
    destino.mkdir()
    servidor = servidor_mcp.criar_servidor(tmp_path, MODELOS, local=True)

    async def verificar():
        async with create_connected_server_and_client_session(servidor) as sessao:
            assert {tool.name for tool in (await sessao.list_tools()).tools} == {
                "consultar_perfil_formatador", "formatar_peticao_local", "formatar_peticao_markdown",
            }
            formato = await sessao.call_tool(
                "formatar_peticao_local", {"arquivo": str(origem), "pasta_destino": str(destino)},
            )
            assert not formato.isError, formato.content
            caminho = destino / "modelo de teste_formatada.docx"
            assert formato.structuredContent["caminho_docx"] == str(caminho)
            assert list(destino.iterdir()) == [caminho]
            assert origem.read_text(encoding="utf-8") == MARKDOWN
            assert any(p.style.name == "FB - Pedido" for p in Document(caminho).paragraphs)

    asyncio.run(verificar())
