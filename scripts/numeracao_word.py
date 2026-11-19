from __future__ import annotations

from copy import deepcopy
from uuid import uuid4
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm


ESTILOS_TITULO = {
    "FB - Título Principal": 0,
    "FB - Título Secundário": 1,
}


def garantir_estilos_titulos(doc) -> None:
    """Mantém negrito, caixa alta e recuo de primeira linha também no modelo."""
    for nome in ESTILOS_TITULO:
        estilo = doc.styles[nome]
        estilo.font.bold = True
        estilo.font.cs_bold = True
        estilo.font.all_caps = True
        estilo.font.small_caps = False
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


def _nivel(indice: int, texto: str, estilo_id: str | None, esquerda: int, *, formato="decimal", pendente=None):
    nivel = OxmlElement("w:lvl")
    nivel.set(qn("w:ilvl"), str(indice))
    nivel.append(_filho("start", "1"))
    nivel.append(_filho("numFmt", formato))
    if estilo_id is not None:
        nivel.append(_filho("pStyle", estilo_id))
    if indice == 1:
        nivel.append(_filho("suff", "space"))
    nivel.append(_filho("lvlText", texto))
    nivel.append(_filho("lvlJc", "left"))

    ppr = OxmlElement("w:pPr")
    ind = OxmlElement("w:ind")
    ind.set(qn("w:left"), str(esquerda))
    ind.set(qn("w:firstLine" if indice == 1 else "w:hanging"), str(esquerda if pendente is None else pendente))
    ppr.append(ind)
    nivel.append(ppr)

    rpr = OxmlElement("w:rPr")
    fontes = OxmlElement("w:rFonts")
    for atributo in ("ascii", "hAnsi", "eastAsia", "cs"):
        fontes.set(qn(f"w:{atributo}"), "Segoe UI Variable Text")
    rpr.append(fontes)
    if estilo_id is not None:
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
    numeracao.insert_element_before(abstrato, "w:num", "w:numIdMacAtCleanup")

    lista = OxmlElement("w:num")
    lista.set(qn("w:numId"), num_id)
    lista.append(_filho("abstractNumId", abstract_id))
    numeracao.insert_element_before(lista, "w:numIdMacAtCleanup")

    if vincular_estilos:
        for nome, indice in ESTILOS_TITULO.items():
            ppr = doc.styles[nome].element.get_or_add_pPr()
            anterior = ppr.find(qn("w:numPr"))
            if anterior is not None:
                ppr.remove(anterior)
            ppr.append(criar_numpr(num_id, indice))

    return num_id


def isolar_lista_pedidos(doc, num_id: int) -> str | None:
    """Copia a lista de pedidos e mantém os níveis subordinados da origem."""
    numeracao = doc.part.numbering_part.element
    try:
        origem = numeracao.num_having_numId(num_id)
    except KeyError:
        return None
    abstract_origem = str(origem.abstractNumId.val)
    modelo = numeracao.find(f"w:abstractNum[@w:abstractNumId='{abstract_origem}']", namespaces=numeracao.nsmap)
    if modelo is None:
        return None
    abstrato = deepcopy(modelo)
    abstract_id = _proximo_id(numeracao, "w:abstractNum", "w:abstractNumId")
    novo_id = _proximo_id(numeracao, "w:num", "w:numId")
    abstrato.set(qn("w:abstractNumId"), abstract_id)
    # O Word pode reunir listas que mantenham a mesma identidade do original.
    for tag in ("w:nsid", "w:tmpl"):
        identidade = abstrato.find(qn(tag))
        if identidade is not None:
            abstrato.remove(identidade)
    abstrato.insert(0, _filho("nsid", uuid4().hex[:8].upper()))
    nivel = abstrato.find("w:lvl[@w:ilvl='0']", namespaces=abstrato.nsmap)
    romano = _nivel(0, "%1.", None, 1701, formato="upperRoman", pendente=425)
    if nivel is not None:
        abstrato.replace(nivel, romano)
    else:
        abstrato.insert_element_before(romano, "w:lvl")
    numeracao.insert_element_before(abstrato, "w:num", "w:numIdMacAtCleanup")

    lista = deepcopy(origem)
    lista.set(qn("w:numId"), novo_id)
    lista.abstractNumId.val = int(abstract_id)
    for variante in lista.findall("w:lvlOverride[@w:ilvl='0']", namespaces=lista.nsmap):
        lista.remove(variante)
    numeracao.insert_element_before(lista, "w:numIdMacAtCleanup")
    return novo_id


def numeracao_do_paragrafo(paragrafo):
    """Obtém o vínculo direto ou herdado de uma lista no Word."""
    ppr = paragrafo._p.pPr
    estilo = paragrafo.style
    while True:
        if ppr is not None and ppr.numPr is not None and ppr.numPr.numId is not None:
            return ppr.numPr if ppr.numPr.numId.val != 0 else None
        if estilo is None:
            return None
        ppr = estilo.element.pPr
        estilo = estilo.base_style


def formatos_das_listas(doc) -> dict[tuple[str, str], str]:
    try:
        raiz = doc.part.numbering_part.element
    except (KeyError, NotImplementedError):
        return {}

    def niveis(elemento):
        return {
            nivel.get(qn("w:ilvl")): nivel.find(qn("w:numFmt")).get(qn("w:val"))
            for nivel in elemento.findall(qn("w:lvl")) if nivel.find(qn("w:numFmt")) is not None
        }

    abstratos = {n.get(qn("w:abstractNumId")): niveis(n) for n in raiz.findall(qn("w:abstractNum"))}
    formatos = {}
    for lista in raiz.findall(qn("w:num")):
        formato = abstratos.get(str(lista.abstractNumId.val), {}).copy()
        for substituicao in lista.findall(qn("w:lvlOverride")):
            formato.update(niveis(substituicao))
        formatos.update({(str(lista.numId), nivel): valor for nivel, valor in formato.items()})
    return formatos


def criar_numpr(num_id: str, nivel: int):
    numpr = OxmlElement("w:numPr")
    numpr.append(_filho("ilvl", str(nivel)))
    numpr.append(_filho("numId", num_id))
    return numpr

