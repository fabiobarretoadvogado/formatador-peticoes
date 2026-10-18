from __future__ import annotations

from docx.oxml import OxmlElement
from docx.oxml.ns import qn


ESTILOS_TITULO = {
    "FB - Título Principal": 0,
    "FB - Título Secundário": 1,
}


def _proximo_id(elemento, nome: str, atributo: str) -> str:
    usados = [
        int(no.get(qn(atributo)))
        for no in elemento.findall(qn(nome))
        if no.get(qn(atributo), "").isdigit()
    ]
    return str(max(usados, default=0) + 1)


def _filho(nome: str, valor: str | None = None):
    no = OxmlElement(f"w:{nome}")
    if valor is not None:
        no.set(qn("w:val"), valor)
    return no


def _nivel(indice: int, texto: str, estilo_id: str, esquerda: int):
    nivel = OxmlElement("w:lvl")
    nivel.set(qn("w:ilvl"), str(indice))
    nivel.append(_filho("start", "1"))
    nivel.append(_filho("numFmt", "decimal"))
    nivel.append(_filho("pStyle", estilo_id))
    nivel.append(_filho("lvlText", texto))
    nivel.append(_filho("lvlJc", "left"))

    ppr = OxmlElement("w:pPr")
    ind = OxmlElement("w:ind")
    ind.set(qn("w:left"), str(esquerda))
    ind.set(qn("w:hanging"), str(esquerda))
    ppr.append(ind)
    nivel.append(ppr)

    rpr = OxmlElement("w:rPr")
    fontes = OxmlElement("w:rFonts")
    for atributo in ("ascii", "hAnsi", "eastAsia", "cs"):
        fontes.set(qn(f"w:{atributo}"), "Segoe UI Variable Text")
    rpr.append(fontes)
    rpr.append(OxmlElement("w:b"))
    rpr.append(OxmlElement("w:bCs"))
    rpr.append(_filho("sz", "24"))
    rpr.append(_filho("szCs", "24"))
    nivel.append(rpr)
    return nivel


def criar_lista_multinivel_titulos(doc, *, vincular_estilos: bool = False) -> str:
    """Cria a lista decimal 1 / 1.1 usada pelos dois estilos de título."""
    numeracao = doc.part.numbering_part.element
    abstract_id = _proximo_id(numeracao, "w:abstractNum", "w:abstractNumId")
    num_id = _proximo_id(numeracao, "w:num", "w:numId")

    abstrato = OxmlElement("w:abstractNum")
    abstrato.set(qn("w:abstractNumId"), abstract_id)
    abstrato.append(_filho("nsid", "46425449"))
    abstrato.append(_filho("multiLevelType", "multilevel"))
    abstrato.append(_filho("tmpl", "46425449"))
    estilos = [
        doc.styles["FB - Título Principal"].style_id,
        doc.styles["FB - Título Secundário"].style_id,
        *(doc.styles[f"Heading {indice}"].style_id for indice in range(3, 10)),
    ]
    recuos = [283, 567, 720, 864, 1008, 1152, 1296, 1440, 1584]
    for indice, (estilo_id, esquerda) in enumerate(zip(estilos, recuos)):
        texto = ".".join(f"%{nivel}" for nivel in range(1, indice + 2))
        abstrato.append(_nivel(indice, texto, estilo_id, esquerda))
    numeracao.append(abstrato)

    lista = OxmlElement("w:num")
    lista.set(qn("w:numId"), num_id)
    lista.append(_filho("abstractNumId", abstract_id))
    numeracao.append(lista)

    if vincular_estilos:
        for nome, indice in ESTILOS_TITULO.items():
            ppr = doc.styles[nome].element.get_or_add_pPr()
            anterior = ppr.find(qn("w:numPr"))
            if anterior is not None:
                ppr.remove(anterior)
            ppr.append(criar_numpr(num_id, indice))

    return num_id


def criar_numpr(num_id: str, nivel: int):
    numpr = OxmlElement("w:numPr")
    numpr.append(_filho("ilvl", str(nivel)))
    numpr.append(_filho("numId", num_id))
    return numpr

