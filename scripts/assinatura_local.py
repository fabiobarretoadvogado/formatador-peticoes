from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass
from io import BytesIO

from docx.shared import Cm
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.oxml.text.paragraph import CT_P
from PIL import Image

if __package__:
    from .comum import carregar_json, pasta_dados_locais
else:
    from comum import carregar_json, pasta_dados_locais


ESTILO_ASSINATURA = "FB - Assinatura"
MARCADOR_NUMERO = r"\s*(?:n\s*\.?\s*[º°o]?\s*|[º°]\s*)?"
NUMERO_OAB = r"\d{2,3}\.\d{3}|\d{1,6}"
RE_OAB = re.compile(rf"OAB\s*/\s*([A-Z]{{2}}){MARCADOR_NUMERO}({NUMERO_OAB})", re.I)
REGISTRO_LINHA = rf"OAB(?:\s*/\s*[A-Z]{{2}})?{MARCADOR_NUMERO}(?:{NUMERO_OAB})"
LINHA_OAB = re.compile(
    rf"(?:advogad[oa]\s*[-–—,:]?\s*)?{REGISTRO_LINHA}"
    rf"(?:\s*(?:[-–—/,;]|e)\s*{REGISTRO_LINHA})*\s*\.?",
    re.I,
)
RE_ROTULO = re.compile(r"^(?:rol\s+de|lista\s+de|anexos?\b|documentos?\b|advogad[oa]\b|procurador[a]?\b)", re.I)


def parece_linha_oab(texto: str) -> bool:
    return bool(LINHA_OAB.fullmatch(" ".join(texto.split())))


def parece_nome_signatario(texto: str) -> bool:
    palavras = texto.split()
    return bool(
        2 <= len(palavras) <= 10
        and not RE_ROTULO.search(texto)
        and all(
            palavra.casefold() in {"de", "da", "do", "das", "dos", "e"}
            or (palavra[0].isupper() and palavra.replace("-", "").replace("'", "").replace("’", "").isalpha())
            for palavra in palavras
        )
    )


def _normalizar_nome(texto: str) -> str:
    decomposto = unicodedata.normalize("NFKD", " ".join(texto.split()).casefold())
    return "".join(c for c in decomposto if not unicodedata.combining(c))


@dataclass(frozen=True)
class AssinaturaLocal:
    nome: str
    registros: frozenset[str]
    imagem: bytes
    largura_cm: float

    def corresponde_oab(self, texto: str) -> bool:
        texto = " ".join(texto.split())
        if not parece_linha_oab(texto):
            return False
        registros = RE_OAB.findall(texto)
        return (
            bool(registros) and len(registros) == len(re.findall(r"\bOAB\b", texto, re.I))
            and {uf.upper() + numero.replace(".", "") for uf, numero in registros} <= self.registros
        )


def carregar_assinatura_local() -> AssinaturaLocal | None:
    pasta = pasta_dados_locais()
    caminho = pasta / "assinatura.json"
    if not caminho.is_file():
        return None
    try:
        config = carregar_json(caminho)
        nome = config["nome"].strip()
        registros = frozenset(registro.upper() for registro in config["registros"])
        largura_cm = float(config["largura_cm"])
        if not nome or not registros or any(not re.fullmatch(r"[A-Z]{2}\d{1,6}", r) for r in registros):
            raise ValueError("O nome e os registros da OAB devem estar completos.")
        if not 1 <= largura_cm <= 10:
            raise ValueError("A largura da imagem deve estar entre 1 e 10 cm.")
        imagem = (pasta / "assinatura.png").read_bytes()
        with Image.open(BytesIO(imagem), formats=["PNG"]) as png:
            png.verify()
    except (OSError, ValueError, TypeError, KeyError, AttributeError) as exc:
        raise ValueError("Não foi possível carregar a imagem ou a configuração da assinatura local.") from exc
    return AssinaturaLocal(nome, registros, imagem, largura_cm)


def localizar_bloco_assinatura(doc, assinatura: AssinaturaLocal) -> tuple[Paragraph, ...]:
    """Encontra o último bloco com o nome e a inscrição configurados."""
    nome = _normalizar_nome(assinatura.nome)
    bloco: tuple[Paragraph, ...] = ()
    anterior: Paragraph | None = None
    for elemento in _elementos_assinatura(doc):
        if elemento is None:
            anterior = None
            continue
        texto = elemento.text.strip()
        if not texto:
            continue
        linha_nome, quebra, restante = texto.partition("\n")
        if quebra and _normalizar_nome(linha_nome) == nome and assinatura.corresponde_oab(restante):
            bloco = (elemento,)
        elif anterior is not None and _normalizar_nome(anterior.text) == nome and assinatura.corresponde_oab(texto):
            bloco = anterior, elemento
        anterior = elemento
    return bloco


def _elementos_assinatura(doc):
    for elemento in doc.iter_inner_content():
        if isinstance(elemento, Paragraph):
            yield elemento
        else:
            yield None
            if isinstance(elemento, Table) and len(elemento.rows) == len(elemento.columns) == 1:
                yield from elemento.cell(0, 0).paragraphs
                yield None


def inserir_imagem_assinatura(bloco: tuple[Paragraph, ...], assinatura: AssinaturaLocal) -> bool:
    """Insere a imagem local apenas quando o bloco ainda não tem uma imagem."""
    primeiro = bloco[0]
    ultimo = bloco[-1]
    atual = primeiro._p
    ja_tem_imagem = False
    while atual is not ultimo._p:
        ja_tem_imagem |= bool(atual.xpath(".//w:drawing | .//w:pict"))
        atual = atual.getnext()
    if ja_tem_imagem or ultimo._p.xpath(".//w:drawing | .//w:pict"):
        return False
    for vizinho in (primeiro._p.itersiblings(preceding=True), ultimo._p.itersiblings()):
        for elemento in vizinho:
            if not isinstance(elemento, CT_P) or elemento.text.strip():
                break
            if elemento.xpath(".//w:drawing | .//w:pict"):
                return False
    novo = primeiro.insert_paragraph_before(style=ESTILO_ASSINATURA)
    novo.add_run().add_picture(BytesIO(assinatura.imagem), width=Cm(assinatura.largura_cm))
    return True


def manter_blocos_assinatura(doc) -> None:
    """Encadeia os parágrafos de assinatura e os espaços existentes entre eles."""
    anterior = None
    vazios = []
    for paragrafo in _elementos_assinatura(doc):
        if paragrafo is None:
            anterior, vazios = None, []
            continue
        if paragrafo.style.name == ESTILO_ASSINATURA:
            formato = paragrafo.paragraph_format
            formato.keep_together = True
            formato.keep_with_next = False
            if anterior is not None:
                for vinculado in (anterior, *vazios):
                    vinculado.paragraph_format.keep_with_next = True
            anterior, vazios = paragrafo, []
        elif anterior is not None and not paragrafo.text.strip() and not paragrafo._p.xpath(".//w:drawing | .//w:pict"):
            vazios.append(paragrafo)
        else:
            anterior, vazios = None, []
