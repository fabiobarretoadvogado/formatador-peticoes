from __future__ import annotations

import argparse
import re
from pathlib import Path

try:
    from .assinatura_local import parece_linha_oab, parece_nome_signatario
    from .comum import ROOT, TIPO_PECA_SUPORTADO, carregar_json, salvar_json
except ImportError:
    from assinatura_local import parece_linha_oab, parece_nome_signatario
    from comum import ROOT, TIPO_PECA_SUPORTADO, carregar_json, salvar_json

RE_ENDERECAMENTO = re.compile(r"^\s*(excelent[ií]ssim[oa]|merit[ií]ssim[oa]|ao\s+ju[ií]zo|[àa]o\s+excelent[ií]ssim)", re.I)
RE_DESTINATARIO = re.compile(
    r"^(?:ao|à|a)\s+(?:(?:_+|\d+\s*[ºªo])\s+)?"
    r"(?:ju[ií]zado|vara|tribunal|turma|c[aâ]mara|(?:ju[ií]z[ao]|presidente|relator[a]?)\s+d[ao])\b",
    re.I,
)
RE_VOCATIVO_ORGAO = re.compile(
    r"^(?:(?:ao|[àa])\s+)?(?:egr[eé]gi[oa]|colend[oa]|em[eé]rit[oa])\s+(?P<orgao>.+?)\s*[.:!]?\s*$",
    re.I,
)
RE_ORGAO_JUDICIAL = re.compile(r"\b(?:tribunal|turma|c[aâ]mara|ju[ií]zo)\b", re.I)
RE_ORGAO_EM_NARRATIVA = re.compile(
    r"\b(?:decidiu|decide|entendeu|entende|reconheceu|reconhece|assentou|firmou|julgou|"
    r"julga|determinou|determina|compete|conforme|segundo|ac[oó]rd[aã]o|foi|foram|remeteu)\b", re.I,
)
NUMERO_CNJ = r"[0-9]{7}-[0-9]{2}\.[0-9]{4}\.[0-9]\.[0-9]{2}\.[0-9]{4}"
NUMERO_PROCESSO = rf"(?:{NUMERO_CNJ}|[0-9]{{5,20}})"
RE_PROCESSO = re.compile(
    rf"(?:processo(?:\s+de\s+origem|\s+origin[aá]rio)?|autos|proc\.?)\s*(?:n\s*\.?\s*[º°o]?\s*)?[:–-]?\s*"
    rf"{NUMERO_PROCESSO}(?:\s*[/–-]\s*{NUMERO_PROCESSO})?\s*\.?",
    re.I,
)
RE_NUMERO_UNICO = re.compile(rf"n[uú]mero\s+[uú]nico\s*:\s*{NUMERO_CNJ}", re.I)
RE_PEDIDO = re.compile(r"^\s*(ante\s+o\s+exposto|diante\s+do\s+exposto|pelo\s+exposto|requer(?:-se)?)\b", re.I)
RE_FECHAMENTO = re.compile(r"\b(termos em que|pede deferimento|nestes termos)\b", re.I)
DATA_ENCERRAMENTO = (
    r"(?:\d{1,2}(?:º)?\s+de\s+"
    r"(?:janeiro|fevereiro|março|abril|maio|junho|julho|agosto|setembro|outubro|novembro|dezembro)"
    r"\s+de\s+\d{4}|\d{1,2}[/.-]\d{1,2}[/.-]\d{4}|"
    r"(?:na\s+)?data\s+(?:do\s+protocolo(?:\s+eletr[oô]nico)?|"
    r"da\s+assinatura(?:\s+(?:eletr[oô]nica|digital))?))"
)
RE_DATA_ENCERRAMENTO = re.compile(rf"^{DATA_ENCERRAMENTO}\.?$", re.I)
RE_LOCAL_DATA = re.compile(rf"^(?P<local>[^,\n]{{1,100}}),\s*{DATA_ENCERRAMENTO}\.?$", re.I)
RE_UF_FINAL = re.compile(r"\s*(?:[/–-]\s*[A-Z]{2}|\(\s*[A-Z]{2}\s*\))$", re.I)
RE_LISTA = re.compile(r"^\s*(?:\d+[.)]|[a-z][.)]|[ivxlcdm]+[.)]|[-–—•])\s+", re.I)
RE_LISTA_NUMERADA = re.compile(r"^\s*(?:\d+[.)]|[a-z][.)]|[ivxlcdm]+[.)])\s+", re.I)
RE_ASSINATURA = re.compile(
    r"\b(?:OAB(?:/\w+)?\s*(?:n\s*\.?\s*[º°o]?|[º°])?\s*\d+|advogad[oa]|procurador[a]?)\b",
    re.I,
)
RE_RESUMO = re.compile(r"^\s*resumo\s*:", re.I)
RE_QUALIFICACAO = re.compile(r"\b(já\s+qualificad[oa]|qualificad[oa]\s+nos\s+autos|vem,?\s+respeitosamente|em\s+face\s+de)\b", re.I)
RE_DOCUMENTO_PARTE = re.compile(r"\b(?:CPF|CNPJ|RG)\b", re.I)
RE_DADOS_PARTE = re.compile(
    r"\b(?:residente(?:\s+e\s+domiciliad[oa])?|domiciliad[oa]|com\s+sede|representad[oa]\s+por|"
    r"nacionalidade|estado\s+civil|profiss[aã]o|maior\s+e\s+capaz)\b",
    re.I,
)
RE_ACAO = re.compile(r"\b(ação|mandado\s+de\s+segurança|cumprimento\s+de\s+sentença|execução|procedimento)\b", re.I)
RE_DENOMINACAO_PECA = re.compile(
    r"contrarraz[oõ]es\s+ao\s+recurso\s+(?:especial|extraordin[aá]rio)", re.I,
)
RE_MARKDOWN_TITULO_PRINCIPAL = re.compile(r"^\s*##(?!#)\s+")
RE_MARKDOWN_TITULO_SECUNDARIO = re.compile(r"^\s*###\s+")
RE_MARKDOWN_TITULO_TERCIARIO = re.compile(r"^\s*####\s+")
RE_MARKDOWN_CITACAO = re.compile(r"^\s*>\s?")
RE_MARKDOWN_TABELA = re.compile(r"^\s*\|.*\|\s*$")
RE_BLOCO_PEDIDOS = re.compile(
    r"(?:#{2,4}\s+)?(?:(?:\d+(?:\.\d+)*|[ivxlcdm]+)[.)]?\s*(?:[-–—:]\s*)?\s+)?"
    r"(?:(?:dos?\s+)?pedidos?(?:\s+e\s+(?:dos?\s+)?requerimentos?)?|"
    r"(?:dos?\s+)?requerimentos?)(?:\s+finais?)?\s*:?",
    re.I,
)
RE_TITULO_ENUMERADO = re.compile(
    r"^(?P<numero>(?:\d{1,3}|[IVXLCDM]+)(?:\.\d{1,3})*)(?:[.)]?\s*[-–—:]\s+|[.)]?\s+)(?P<titulo>\S.*)$", re.I,
)
RE_ROMANO_VALIDO = re.compile(r"M{0,3}(?:CM|CD|D?C{0,3})(?:XC|XL|L?X{0,3})(?:IX|IV|V?I{0,3})", re.I)
RE_NOME_SECAO = re.compile(
    r"^(?:(?:da|do|das|dos)\s+)?(?:breve\s+)?(?:s[ií]ntese|fatos?|direito|fundamentos?|"
    r"pedidos?|requerimentos?|provas?|documentos?|preliminares?|m[eé]rito|conclus[aã]o|"
    r"tutela|responsabilidade|compet[eê]ncia|gratuidade|justi[cç]a\s+gratuita|"
    r"valor\s+da\s+causa|requisitos?)\b", re.I,
)
RE_TITULO_EM_NARRATIVA = re.compile(
    r"\b(?:requer(?:-se)?|seja|sejam|consta|constam|conclui|resulta|dever[aá]|"
    r"dever[aã]o|foi|foram|comprova|comprovam|determina|pede|decidiu)\b", re.I,
)
RE_VERBO_PROSA = re.compile(
    r"\b(?:é|s[aã]o|foi|foram|h[aá]|houve|possui|possuem|re[uú]ne|re[uú]nem|"
    r"demonstra|demonstram|evidencia|evidenciam|comprova|comprovam|revela|revelam|"
    r"configura|configuram|constitui|constituem|apresenta|apresentam|permanece|permanecem|"
    r"decorre|decorrem|sustenta|sustentam|alega|alegam|afirma|afirmam|informa|informam|"
    r"preenche|preenchem|verifica-se|observa-se|trata-se)\b", re.I,
)
RE_ROTULO_INFORMATIVO = re.compile(r"^(?:descri[cç][aã]o|fonte|nota)\s*:", re.I)
ESTILOS_CITACAO = {"quote", "citação", "citacao", "fb - citação"}
TITULOS_ORIGEM = {
    "heading 1": "FB - Título Principal",
    "título 1": "FB - Título Principal",
    "titulo 1": "FB - Título Principal",
    "fb - título principal": "FB - Título Principal",
    "heading 2": "FB - Título Secundário",
    "título 2": "FB - Título Secundário",
    "titulo 2": "FB - Título Secundário",
    "fb - título secundário": "FB - Título Secundário",
}
TITULOS = frozenset(TITULOS_ORIGEM.values())


def _caixa_alta(texto: str) -> bool:
    letras = [c for c in texto if c.isalpha()]
    return bool(letras) and sum(c.isupper() for c in letras) / len(letras) >= 0.9


def parece_vocativo_orgao(texto: str) -> bool:
    encontrado = RE_VOCATIVO_ORGAO.fullmatch(texto)
    if not encontrado:
        return False
    orgao = encontrado["orgao"].rstrip(".:! ")
    return bool(
        len(orgao.split()) <= 24
        and RE_ORGAO_JUDICIAL.search(orgao)
        and not re.search(r"[,;:!?]", orgao)
        and not RE_ORGAO_EM_NARRATIVA.search(orgao)
    )


def parece_enderecamento(texto: str) -> bool:
    destinatario = texto.rstrip(" .,:")
    return parece_vocativo_orgao(texto) or bool(
        len(destinatario.split()) <= 32
        and (RE_ENDERECAMENTO.search(destinatario) or RE_DESTINATARIO.search(destinatario))
        and not re.search(r"[,;:!?]", destinatario)
        and not RE_ORGAO_EM_NARRATIVA.search(destinatario)
    )


def nivel_titulo_enumerado(texto: str) -> int | None:
    """Separa rótulos de seção enumerados de itens narrativos e pedidos."""
    encontrado = RE_TITULO_ENUMERADO.fullmatch(texto)
    if not encontrado:
        return None
    numero, titulo = encontrado["numero"], encontrado["titulo"]
    raiz = numero.split(".", 1)[0]
    if raiz[0].isalpha() and not RE_ROMANO_VALIDO.fullmatch(raiz):
        return None
    if "." in numero:
        # A hierarquia explícita admite subtítulos longos e com mais de uma
        # frase. Itens corridos com pontuação de encerramento continuam na
        # classificação conservadora; estilos Heading 1/2 prevalecem antes.
        if titulo[0].isupper() and not titulo.endswith((".", ";", "!", "?")):
            return 2
        return None
    if (len(titulo.split()) > 16 or re.search(r"[,;.!?]", titulo)
            or not RE_NOME_SECAO.search(titulo) or RE_TITULO_EM_NARRATIVA.search(titulo)):
        return None
    return 1


def parece_qualificacao(texto: str) -> bool:
    """Reconhece a apresentação das partes mesmo quando veio com estilo Quote."""
    return bool(
        RE_QUALIFICACAO.search(texto)
        or (RE_DOCUMENTO_PARTE.search(texto) and RE_DADOS_PARTE.search(texto))
    )


def parece_prosa_discursiva(texto: str) -> bool:
    """Usa sinais de frase completa, em vez de depender só do tamanho."""
    if (len(texto.split()) < 10 or not texto.endswith((".", "!", "?"))
            or RE_ROTULO_INFORMATIVO.search(texto) or not RE_VERBO_PROSA.search(texto)):
        return False
    letras = [caractere for caractere in texto if caractere.isalpha()]
    return bool(letras) and sum(caractere.islower() for caractere in letras) / len(letras) >= 0.70


def parece_fechamento_local_data(p: dict, texto: str, anterior: str = "", seguinte: str = "") -> bool:
    """Reconhece a linha de local/data, sem confundi-la com fatos datados."""
    texto = " ".join(texto.split())
    local_data = RE_LOCAL_DATA.fullmatch(texto)
    if local_data:
        cidade = RE_UF_FINAL.sub("", local_data["local"]).strip()
        palavras = cidade.replace("-", " ").split()
        # Exige formato de nome de localidade; frases como "O contrato foi
        # assinado em Aracaju, ..." continuam sendo conteúdo do corpo.
        if 1 <= len(palavras) <= 8 and all(
            palavra.casefold() in {"de", "da", "do", "das", "dos", "e"}
            or (palavra[0].isupper() and palavra.replace("'", "").replace("’", "").isalpha())
            or bool(re.fullmatch(r"d['’][A-ZÀ-ÖØ-Ý][a-zA-ZÀ-ÖØ-öø-ÿ]+", palavra))
            for palavra in palavras
        ):
            return True
    if not RE_DATA_ENCERRAMENTO.fullmatch(texto):
        return False
    return bool(
        p.get("posicao_relativa", 0) >= 0.70
        or RE_FECHAMENTO.search(anterior)
        or RE_ASSINATURA.search(seguinte)
    )


def parece_nome_assinatura(p: dict, texto: str, anterior: str = "", seguinte: str = "") -> bool:
    nome, quebra, oab = p.get("texto", texto).strip().partition("\n")
    return bool(
        parece_nome_signatario(nome)
        and not RE_NOME_SECAO.search(nome)
        and parece_linha_oab(oab if quebra else seguinte)
        and (
            p.get("posicao_relativa", 0) >= 0.50
            or (p.get("estilo_original") or "").casefold() == "fb - assinatura"
            or parece_fechamento_local_data(p, anterior)
            or RE_FECHAMENTO.search(anterior)
        )
    )


def classificar_um(p: dict, anterior: str = "", seguinte: str = "", *, em_abertura: bool | None = None) -> dict:
    texto = " ".join(p.get("texto", "").split())
    palavras = texto.split()
    estilo_original = (p.get("estilo_original") or "").casefold()
    citacao = estilo_original in ESTILOS_CITACAO
    if em_abertura is None:
        em_abertura = p.get("posicao_relativa", 1) < 0.30
    estilo, conf, motivo = "FB - Corpo", 0.70, "Regra padrão; requer contexto semântico"
    if not texto:
        estilo, conf, motivo = "FB - Corpo", 1.0, "Parágrafo vazio preservado"
    elif RE_MARKDOWN_TITULO_SECUNDARIO.search(texto):
        estilo, conf, motivo = "FB - Título Secundário", 0.99, "Subtítulo estrutural marcado no texto"
    elif RE_MARKDOWN_TITULO_TERCIARIO.search(texto):
        estilo, conf, motivo = "FB - Título Secundário", 0.98, "Subtítulo de terceiro nível mapeado para secundário"
    elif RE_MARKDOWN_TITULO_PRINCIPAL.search(texto):
        estilo, conf, motivo = "FB - Título Principal", 0.99, "Título estrutural marcado no texto"
    elif RE_MARKDOWN_CITACAO.search(texto):
        estilo, conf, motivo = "FB - Citação", 0.99, "Citação estrutural marcada no texto"
    elif RE_MARKDOWN_TABELA.search(texto):
        estilo, conf, motivo = "FB - Tabela", 0.99, "Linha de tabela estrutural marcada no texto"
    elif RE_RESUMO.search(texto):
        estilo, conf, motivo = "FB - Resumo", 0.98, "Resumo inicial identificado por rótulo"
    elif not citacao and (parece_enderecamento(texto) or em_abertura and estilo_original == "fb - endereçamento"):
        estilo, conf, motivo = "FB - Endereçamento", 0.99, "Fórmula de endereçamento judicial"
    elif em_abertura and not citacao and (RE_NUMERO_UNICO.fullmatch(texto) or RE_PROCESSO.fullmatch(texto)):
        estilo, conf, motivo = "FB - Identificação do Processo", 0.99, "Identificação dos autos no bloco de abertura"
    elif parece_fechamento_local_data(p, texto, anterior, seguinte):
        estilo, conf, motivo = "FB - Fechamento", 0.99, "Linha de encerramento com local e data ou referência ao protocolo/assinatura"
    elif not citacao and parece_nome_assinatura(p, texto, anterior, seguinte):
        estilo, conf, motivo = "FB - Assinatura", 0.98, "Nome do subscritor junto ao registro profissional"
    elif not citacao and parece_linha_oab(texto):
        estilo, conf, motivo = "FB - Assinatura", 0.98, "Registro profissional em linha própria"
    elif not citacao and RE_DENOMINACAO_PECA.fullmatch(texto) and not estilo_original.startswith("fb - título"):
        estilo, conf, motivo = "FB - Identificação da Ação", 0.99, "Denominação isolada da peça; não é seção argumentativa"
    elif estilo_original in TITULOS_ORIGEM:
        estilo, conf, motivo = TITULOS_ORIGEM[estilo_original], 0.99, "Nível do título no estilo estrutural original"
    elif not citacao and RE_BLOCO_PEDIDOS.fullmatch(texto):
        estilo, conf, motivo = "FB - Título Principal", 0.99, "Título do bloco de pedidos ou requerimentos"
    elif p.get("posicao_relativa", 1) < 0.30 and parece_qualificacao(texto):
        estilo, conf, motivo = "FB - Qualificação das Partes", 0.98, "Dados pessoais e representação no bloco inicial das partes"
    elif (nivel := nivel_titulo_enumerado(texto)) and not citacao:
        estilo = "FB - Título Principal" if nivel == 1 else "FB - Título Secundário"
        conf, motivo = 0.98, "Rótulo de seção com numeração explícita; não é item de lista"
    elif not citacao and (_caixa_alta(texto) or p.get("negrito")) and RE_ACAO.search(texto) and len(palavras) <= 18:
        estilo, conf, motivo = "FB - Identificação da Ação", 0.94, "Denominação curta da ação em destaque"
    elif citacao:
        estilo, conf, motivo = "FB - Citação", 0.99, "Estilo de citação identificado no documento original"
    elif RE_FECHAMENTO.search(texto):
        estilo, conf, motivo = "FB - Fechamento", 0.97, "Fórmula convencional de fechamento"
    elif estilo_original == "fb - assinatura":
        estilo, conf, motivo = "FB - Assinatura", 0.99, "Bloco de assinatura no estilo original do escritório"
    elif RE_PEDIDO.search(texto):
        estilo, conf, motivo = "FB - Corpo", 0.97, "Parágrafo introdutório do bloco de pedidos"
    elif p.get("possui_numeracao") or RE_LISTA.search(texto) or estilo_original in {"fb - lista", "fb - pedido"}:
        estilo, conf, motivo = "FB - Lista", 0.93, "Numeração estrutural ou marcador no início"
    elif _caixa_alta(texto) and len(palavras) <= 14:
        estilo = "FB - Título Principal" if p.get("posicao_relativa", 1) < 0.2 else "FB - Título Secundário"
        conf, motivo = 0.85, "Bloco curto em caixa alta, sem evidência estrutural conclusiva"
    elif p.get("negrito") and len(palavras) <= 14:
        estilo = "FB - Título Principal" if p.get("posicao_relativa", 1) < 0.2 else "FB - Título Secundário"
        conf, motivo = 0.75, "Bloco curto em negrito, sem evidência estrutural conclusiva"
    elif (p.get("recuo_esquerdo_cm") or 0) >= 3.0 or (texto.startswith(('"', '“')) and texto.endswith(('"', '”'))):
        estilo, conf, motivo = "FB - Citação", 0.91, "Recuo elevado ou aspas delimitadoras"
    elif parece_prosa_discursiva(texto):
        estilo, conf, motivo = "FB - Corpo", 0.96, "Prosa discursiva com oração verbal e pontuação de encerramento"
    elif len(texto) > 180:
        estilo, conf, motivo = "FB - Corpo", 0.96, "Parágrafo argumentativo longo"
    elif len(palavras) >= 12:
        estilo, conf, motivo = "FB - Corpo", 0.90, "Parágrafo discursivo"
    return {
        "id": p["id"], "hash": p.get("hash"), "texto": texto,
        "categoria_provavel": estilo, "estilo_destino": estilo,
        "confianca": conf, "motivo": motivo,
        "necessita_analise_semantica": conf < 0.95 and bool(texto),
    }


def _classificar_pedidos_finais(ps, itens) -> None:
    limites = TITULOS | {"FB - Endereçamento", "FB - Fechamento", "FB - Assinatura"}
    aberturas = [
        i for i, item in enumerate(itens)
        if item["estilo_destino"] in TITULOS and RE_BLOCO_PEDIDOS.fullmatch(item["texto"])
    ]
    if not aberturas:
        for i, item in enumerate(itens):
            if item["estilo_destino"] != "FB - Corpo" or not RE_PEDIDO.search(item["texto"]):
                continue
            proximos = (p for p in itens[i + 1:i + 12] if p["estilo_destino"] in limites)
            if next(proximos, {}).get("estilo_destino") == "FB - Fechamento":
                aberturas.append(i)
    if not aberturas:
        return
    for i in range(aberturas[-1] + 1, len(itens)):
        item = itens[i]
        if item["estilo_destino"] in limites:
            break
        p = ps[i]
        lista_numerada = p.get("possui_numeracao") and p.get("formato_lista") not in {"bullet", "none"}
        lista = item["estilo_destino"] == "FB - Lista" and (lista_numerada or RE_LISTA_NUMERADA.search(item["texto"]))
        # Há pedidos originalmente recuados com o estilo genérico Quote.
        pedido_recuado = (
            item["estilo_destino"] == "FB - Citação" and lista_numerada
            and (p.get("estilo_original") or "").casefold() == "quote"
        )
        if lista or pedido_recuado:
            item.update({
                "categoria_provavel": "FB - Pedido",
                "estilo_destino": "FB - Pedido",
                "confianca": 0.99,
                "motivo": "Item numerado do bloco final de pedidos",
                "necessita_analise_semantica": False,
            })


def classificar(estrutura: dict) -> dict:
    ps = estrutura["paragrafos"]
    itens = []
    em_resumo = False
    linhas_abertura = 0
    em_abertura = True
    for i, p in enumerate(ps):
        texto = " ".join(p.get("texto", "").split())
        if not texto:
            em_resumo = False
        elif RE_RESUMO.search(texto):
            em_resumo = True

        anterior = next(
            (ps[j].get("texto", "") for j in range(i - 1, -1, -1) if ps[j].get("texto", "").strip()),
            "",
        )
        seguinte = next(
            (ps[j].get("texto", "") for j in range(i + 1, len(ps)) if ps[j].get("texto", "").strip()),
            "",
        )
        item = classificar_um(p, anterior, seguinte, em_abertura=em_abertura and linhas_abertura < 12)
        papel = item["estilo_destino"]
        if papel == "FB - Endereçamento":
            em_abertura, linhas_abertura, em_resumo = True, 0, False
        elif texto:
            linhas_abertura += 1
            if (
                papel in {"FB - Qualificação das Partes", "FB - Identificação da Ação", "FB - Fechamento"}
                or papel in TITULOS and linhas_abertura > 1 and not item["necessita_analise_semantica"]
                or papel != "FB - Identificação do Processo" and len(texto.split()) >= 20
            ):
                em_abertura = False
        if (em_resumo and texto and p.get("posicao_relativa", 1) < 0.15
                and papel not in TITULOS | {"FB - Endereçamento", "FB - Identificação do Processo"}):
            item.update({
                "categoria_provavel": "FB - Resumo",
                "estilo_destino": "FB - Resumo",
                "confianca": 0.98,
                "motivo": "Parágrafo integrante do bloco de resumo inicial",
                "necessita_analise_semantica": False,
            })
        itens.append(item)
    _classificar_pedidos_finais(ps, itens)
    ambiguos = []
    for i, item in enumerate(itens):
        if item["necessita_analise_semantica"]:
            ambiguos.append({
                "id": item["id"], "hash": item["hash"],
                "paragrafo_anterior": ps[i - 1]["texto"] if i else "",
                "paragrafo_atual": ps[i]["texto"],
                "paragrafo_seguinte": ps[i + 1]["texto"] if i + 1 < len(ps) else "",
                "pre_classificacao": item["estilo_destino"], "confianca_regra": item["confianca"],
            })
    return {
        "documento": estrutura["documento"],
        "tipo_peca": TIPO_PECA_SUPORTADO,
        "classificacoes": itens,
        "ambiguos": ambiguos,
    }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("estrutura", type=Path)
    ap.add_argument("--saida", type=Path, default=ROOT / "trabalho" / "classificacao.json")
    args = ap.parse_args()
    salvar_json(args.saida, classificar(carregar_json(args.estrutura)))


if __name__ == "__main__":
    main()
