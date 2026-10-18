from __future__ import annotations

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm


ESTILOS_TITULO = {
    "FB - Título Principal": 0,
    "FB - Título Secundário": 1,
}


def garantir_estilos_titulos(doc) -> None:
    """Mantém a regra de negrito e primeira linha também no modelo."""
    for nome in ESTILOS_TITULO:
        estilo = doc.styles[nome]
        estilo.font.bold = True
        estilo.font.cs_bold = True
    secundario = doc.styles["FB - Título Secundário"]
    secundario.paragraph_format.left_indent = Cm(1)
    secundario.paragraph_format.first_line_indent = Cm(1)
    ind = secundario.element.pPr.find(qn("w:ind"))
    for atributo in ("hanging", "hangingChars", "firstLineChars", "leftChars", "startChars"):
        ind.attrib.pop(qn(f"w:{atributo}"), None)

    # Corrige também as definições já vinculadas aos estilos no modelo.
    ids = {doc.styles[nome].style_id: nivel for nome, nivel in ESTILOS_TITULO.items()}
    numeracao = doc.part.numbering_part.element
    for nome, indice in ESTILOS_TITULO.items():
        estilo = doc.styles[nome]
        numpr = estilo.element.pPr.find(qn("w:numPr"))
        if numpr is None or numpr.numId is None:
            continue
        lista = next((n for n in numeracao.findall(qn("w:num"))
                      if n.get(qn("w:numId")) == str(numpr.numId.val)), None)
        if lista is None:
            continue
        abstract_id = lista.find(qn("w:abstractNumId")).get(qn("w:val"))
        abstrato = next((n for n in numeracao.findall(qn("w:abstractNum"))
                         if n.get(qn("w:abstractNumId")) == abstract_id), None)
        if abstrato is None:
            continue
        nivel = next((n for n in abstrato.findall(qn("w:lvl"))
                      if n.get(qn("w:ilvl")) == str(indice)), None)
        if nivel is None:
            continue
        vinculo = nivel.find(qn("w:pStyle"))
        if vinculo is None:
            vinculo = OxmlElement("w:pStyle")
            formato = nivel.find(qn("w:numFmt"))
            nivel.insert(list(nivel).index(formato) + 1 if formato is not None else 0, vinculo)
        vinculo.set(qn("w:val"), estilo.style_id)
    for nivel in numeracao.findall(".//" + qn("w:lvl")):
        estilo = nivel.find(qn("w:pStyle"))
        if estilo is None or estilo.get(qn("w:val")) not in ids:
            continue
        rpr = nivel.find(qn("w:rPr"))
        if rpr is None:
            rpr = OxmlElement("w:rPr")
            nivel.append(rpr)
        for nome in ("b", "bCs"):
            for existente in list(rpr.findall(qn(f"w:{nome}"))):
                rpr.remove(existente)
            rpr.append(OxmlElement(f"w:{nome}"))
        if ids[estilo.get(qn("w:val"))] == 1:
            # Com primeira linha positiva, uma tabulação após o número pode
            # apontar para trás e juntar a numeração ao texto no Word.
            sufixo = nivel.find(qn("w:suff"))
            if sufixo is None:
                sufixo = OxmlElement("w:suff")
                texto = nivel.find(qn("w:lvlText"))
                nivel.insert(list(nivel).index(texto) if texto is not None else 0, sufixo)
            sufixo.set(qn("w:val"), "space")
            ppr = nivel.find(qn("w:pPr"))
            if ppr is None:
                ppr = OxmlElement("w:pPr")
                nivel.insert(list(nivel).index(rpr), ppr)
            ind = ppr.find(qn("w:ind"))
            if ind is None:
                ind = OxmlElement("w:ind")
                ppr.append(ind)
            for atributo in ("hanging", "hangingChars", "firstLineChars", "leftChars", "startChars"):
                ind.attrib.pop(qn(f"w:{atributo}"), None)
            ind.set(qn("w:left"), "567")
            ind.set(qn("w:firstLine"), "567")


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
    if indice == 1:
        nivel.append(_filho("suff", "space"))
    nivel.append(_filho("lvlText", texto))
    nivel.append(_filho("lvlJc", "left"))

    ppr = OxmlElement("w:pPr")
    ind = OxmlElement("w:ind")
    ind.set(qn("w:left"), str(esquerda))
    ind.set(qn("w:firstLine" if indice == 1 else "w:hanging"), str(esquerda))
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
    garantir_estilos_titulos(doc)
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

