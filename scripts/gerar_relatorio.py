from __future__ import annotations

from pathlib import Path


def gerar_relatorio(
    estrutura: dict,
    classificacao: dict,
    auditoria: dict,
    destino: Path,
    *,
    metadados_entrada: dict | None = None,
) -> Path:
    aplicacao = auditoria.get("aplicacao", {})
    papel_timbrado = aplicacao.get("papel_timbrado", {})
    metadados_entrada = metadados_entrada or {}
    formato = metadados_entrada.get("formato_original", "docx")
    rotulo_formato = "Markdown (.md)" if formato == "markdown" else "Word (.docx)"
    origem = metadados_entrada.get("origem", estrutura["documento"])
    intermediario = metadados_entrada.get("docx_intermediario")
    auditoria_ingestao = metadados_entrada.get("auditoria_conteudo", {})
    linhas = [
        "# Relatório de formatação", "",
        f"Documento de origem: `{origem}`", "",
        f"Formato original: **{rotulo_formato}**", "",
    ]
    if intermediario:
        linhas.extend([
            f"DOCX intermediário utilizado: `{intermediario}`", "",
            "Auditoria da conversão: "
            f"**{'APROVADA' if auditoria_ingestao.get('aprovada') else 'REQUER REVISÃO'}**", "",
        ])
    linhas.extend([
        "Escopo aplicado: **petição inicial**", "",
        f"Total de parágrafos: **{auditoria['total_paragrafos']}**", "",
        "## Alterações", "",
    ])
    for estilo, total in sorted(aplicacao.get("contagem_estilos", auditoria.get("estilos", {})).items()):
        linhas.append(f"- {total} parágrafo(s) classificados como **{estilo}**.")
    linhas.extend([
        f"- {aplicacao.get('formatacoes_diretas_removidas', 0)} propriedade(s) de formatação direta incompatível removida(s).",
        f"- {aplicacao.get('cabecalhos_negrito_caixa_alta', 0)} parágrafo(s) de endereçamento ou identificação do processo receberam negrito integral e caixa alta visual.",
        f"- {len(aplicacao.get('margens_modelo', []))} seção(ões) receberam as margens e as distâncias de cabeçalho/rodapé do modelo oficial.",
        f"- {papel_timbrado.get('cabecalhos_adicionados', 0)} cabeçalho(s) vazio(s) receberam o papel timbrado do modelo oficial.",
        f"- {papel_timbrado.get('rodapes_adicionados', 0)} rodapé(s) vazio(s) receberam a paginação do modelo oficial; conteúdo existente foi preservado.",
        f"- {aplicacao.get('negritos_qualificacao_ajustados', 0)} parágrafo(s) de qualificação tiveram o negrito restringido a nomes, entidades, ações e fases processuais.",
        f"- {aplicacao.get('imagens_assinatura_centralizadas', 0)} imagem(ns) isolada(s) de assinatura foram centralizadas.",
        f"- {aplicacao.get('imagens_gerais_centralizadas', 0)} imagem(ns) geral(is) em bloco foram centralizadas.",
        f"- {aplicacao.get('negritos_corpo_removidos', 0)} run(s) de negrito direto foram removidos dos estilos regulares de texto.",
        f"- {aplicacao.get('normalizacoes_markdown', 0)} parágrafo(s) tiveram marcações estruturais convertidas em formatação Word.",
        f"- {aplicacao.get('tabelas_markdown_convertidas', 0)} tabela(s) textual(is) foram convertida(s) em tabela(s) Word.",
        f"- {aplicacao.get('titulos_literais_ajustados', 0)} título(s) com número digitado foram vinculados à lista multinível automática 1 / 1.1.",
        f"- {aplicacao.get('pedidos_literais_ajustados', 0)} pedido(s) com numeração literal foram normalizados em algarismos romanos.",
        f"- {aplicacao.get('assinaturas_em_tabela', 0)} parágrafo(s) de assinatura em tabela receberam o estilo de assinatura sem espaçamento.",
        f"- {aplicacao.get('vazios_antes_assinatura_removidos', 0)} parágrafo(s) vazio(s) entre data e assinatura foram removidos.",
        "- Tópicos, subtópicos e identificação da ação receberam 24 pt de espaço anterior, equivalente a uma linha de separação.",
        f"- {len(classificacao.get('ambiguos', []))} parágrafo(s) encaminhado(s) à conferência semântica.", "",
        "## Auditoria", "",
        f"- Erros: **{auditoria['resumo']['erros']}**",
        f"- Avisos: **{auditoria['resumo']['avisos']}**",
        f"- Resultado: **{'APROVADO' if auditoria['resumo']['aprovado'] else 'REQUER REVISÃO'}**", "",
        "## Alertas", "",
    ])
    alertas = aplicacao.get("alertas", []) + auditoria.get("problemas", [])
    if not alertas:
        linhas.append("- Nenhum alerta registrado.")
    else:
        for a in alertas:
            local = a.get("paragrafo", a.get("secao", "—"))
            detalhe = a.get("alerta", a.get("codigo", "alerta"))
            linhas.append(f"- Item {local}: {detalhe}.")
    avisos_ingestao = metadados_entrada.get("avisos", [])
    if avisos_ingestao:
        linhas.extend(["", "## Avisos da entrada", ""])
        linhas.extend(f"- {aviso}" for aviso in avisos_ingestao)
    destino.parent.mkdir(parents=True, exist_ok=True)
    destino.write_text("\n".join(linhas) + "\n", encoding="utf-8")
    return destino
