from __future__ import annotations

import argparse
import sys
import tempfile
from pathlib import Path

if __package__ in (None, ""):
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
    from scripts.aplicar_estilos import aplicar
    from scripts.classificar_regras import classificar
    from scripts.comum import ROOT, TIPO_PECA_SUPORTADO, backup_entrada, carregar_json, garantir_diretorios, salvar_json
    from scripts.configurar_visibilidade_estilos import ajustar_arquivo
    from scripts.converter_markdown import normalizar_entrada
    from scripts.extrair_documento import extrair
    from scripts.gerar_pdf import gerar_pdf
    from scripts.gerar_relatorio import gerar_relatorio
    from scripts.validar_documento import validar
else:
    from .aplicar_estilos import aplicar
    from .classificar_regras import classificar
    from .comum import ROOT, TIPO_PECA_SUPORTADO, backup_entrada, carregar_json, garantir_diretorios, salvar_json
    from .configurar_visibilidade_estilos import ajustar_arquivo
    from .converter_markdown import normalizar_entrada
    from .extrair_documento import extrair
    from .gerar_pdf import gerar_pdf
    from .gerar_relatorio import gerar_relatorio
    from .validar_documento import validar


def _consolidar(classificacao: dict, semantica: Path | None, config: dict) -> int:
    if not semantica or not semantica.exists():
        return 0
    permitidos = set(config["estilos_permitidos"])
    respostas = carregar_json(semantica)
    if isinstance(respostas, dict):
        respostas = respostas.get("classificacoes", respostas.get("respostas", []))
    if not isinstance(respostas, list):
        return 0
    mapa = {x.get("id"): x for x in respostas if isinstance(x, dict) and x.get("id") is not None}
    ids_ambiguos = {x.get("id") for x in classificacao.get("ambiguos", [])}
    resolvidos = set()
    for item in classificacao["classificacoes"]:
        if item["id"] not in ids_ambiguos:
            continue
        s = mapa.get(item["id"])
        if not s or s.get("estilo_destino") not in permitidos:
            continue
        if not s.get("hash") or s.get("hash") != item.get("hash"):
            continue
        if float(s.get("confianca", 0)) > 0.90:
            item.update({
                "estilo_destino": s["estilo_destino"],
                "confianca": float(s["confianca"]),
                "motivo": s.get("justificativa", "Classificação confirmada em revisão humana"),
                "origem": "revisao_humana",
                "necessita_analise_semantica": False,
            })
            resolvidos.add(item["id"])
    classificacao["ambiguos"] = [
        item for item in classificacao.get("ambiguos", []) if item.get("id") not in resolvidos
    ]
    return len(resolvidos)


def _carregar_artefatos(root: Path) -> tuple[Path, dict]:
    garantir_diretorios(root)
    modelo = root / "modelos" / "modelo_oficial.docx"
    config_path = root / "modelos" / "estilos.json"
    ausentes = [caminho for caminho in (modelo, config_path) if not caminho.is_file()]
    if ausentes:
        lista = ", ".join(str(caminho) for caminho in ausentes)
        raise FileNotFoundError(
            "A execução foi interrompida porque o modelo oficial ou suas regras não foram encontrados: "
            f"{lista}. Restaure os arquivos aprovados na pasta 'modelos'; o sistema não cria nem substitui "
            "esses artefatos automaticamente."
        )
    config = carregar_json(config_path)
    if config.get("tipo_peca_suportado") != TIPO_PECA_SUPORTADO:
        raise ValueError(
            "As regras instaladas não correspondem ao escopo atual de petição inicial. "
            "Restaure o arquivo modelos/estilos.json aprovado para esta versão."
        )
    return modelo, config


def executar_docx(
    entrada: Path,
    semantica: Path | None = None,
    skip_pdf: bool = False,
    work_root: Path | None = None,
    *,
    metadados_entrada: dict | None = None,
    manifesto_entrada: Path | None = None,
    backup_original: Path | None = None,
) -> dict:
    """Executa o núcleo estável do formatador sobre uma entrada DOCX."""
    root = (work_root or ROOT).resolve()
    modelo, config = _carregar_artefatos(root)
    entrada = entrada.resolve()
    if not entrada.exists() or entrada.suffix.lower() != ".docx": raise ValueError("Informe um arquivo DOCX existente")
    metadados_entrada = metadados_entrada or {
        "formato_original": "docx",
        "origem": str(entrada),
        "entrada_pipeline": str(entrada),
        "docx_intermediario": None,
        "avisos": [],
        "auditoria_conteudo": {"aprovada": True, "modo": "entrada DOCX nativa"},
    }
    estrutura = extrair(entrada)
    backup = backup_original or backup_entrada(entrada, root)
    salvar_json(root / "trabalho" / "estrutura_documento.json", estrutura)
    classificacao = classificar(estrutura); _consolidar(classificacao, semantica, config)
    salvar_json(root / "trabalho" / "classificacao.json", classificacao)
    salvar_json(root / "trabalho" / "solicitacoes_semanticas.json", {
        "instrucoes": (
            "Use o catalogo_semantico para decidir a funcao do paragrafo. "
            "O escopo atual e exclusivamente peticao inicial; nao aplique estruturas proprias de outras pecas. "
            "Nao escolha apenas pela aparencia visual do modelo. "
            "Responda com id, estilo_destino, confianca e justificativa; "
            "use apenas estilos permitidos."
        ),
        "estilos_permitidos": config["estilos_permitidos"],
        "tipo_peca": TIPO_PECA_SUPORTADO,
        "regras_contextuais": config.get("regras_contextuais", {}),
        "formatacao_estilos": config.get("formatacao_estilos", {}),
        "catalogo_semantico": config.get("catalogo_semantico", {}),
        "itens": classificacao["ambiguos"],
    })
    saida_docx = root / "saida" / "peticao_formatada.docx"
    aplicacao = aplicar(entrada, modelo, classificacao, config, saida_docx)
    # The formatter imports the authored styles into the output document. Apply
    # the same visibility policy to the final package so Word does not expose
    # its built-in styles in the user-facing style pane.
    with tempfile.TemporaryDirectory(dir=root / "trabalho") as temporario_dir:
        temporario = Path(temporario_dir) / saida_docx.name
        ajustar_arquivo(saida_docx, temporario, config["estilos_permitidos"])
        temporario.replace(saida_docx)
    auditoria = validar(saida_docx, config, aplicacao)
    pendencias_semanticas = len(classificacao["ambiguos"])
    auditoria["resumo"]["pendencias_semanticas"] = pendencias_semanticas
    auditoria["resumo"]["aprovado"] = auditoria["resumo"]["aprovado"] and pendencias_semanticas == 0
    salvar_json(root / "trabalho" / "auditoria.json", auditoria)
    relatorio = gerar_relatorio(
        estrutura,
        classificacao,
        auditoria,
        root / "saida" / "relatorio_formatacao.md",
        metadados_entrada=metadados_entrada,
    )
    pdf, erro_pdf, pdf_gerado = root / "saida" / "peticao_formatada.pdf", None, False
    if not skip_pdf:
        try:
            gerar_pdf(saida_docx, pdf)
            pdf_gerado = True
        except Exception as exc: erro_pdf = str(exc)
    return {
        "docx": str(saida_docx),
        "pdf": str(pdf) if pdf_gerado else None,
        "relatorio": str(relatorio),
        "backup": str(backup),
        "tipo_peca": TIPO_PECA_SUPORTADO,
        "ambiguos": len(classificacao["ambiguos"]),
        "erro_pdf": erro_pdf,
        "aprovado": auditoria["resumo"]["aprovado"],
        "entrada_original": metadados_entrada.get("origem", str(entrada)),
        "formato_original": metadados_entrada.get("formato_original", "docx"),
        "entrada_pipeline": str(entrada),
        "docx_intermediario": metadados_entrada.get("docx_intermediario"),
        "manifesto_ingestao": str(manifesto_entrada) if manifesto_entrada else None,
        "metadados_entrada": metadados_entrada,
    }


def executar(entrada: Path, semantica: Path | None = None, skip_pdf: bool = False, work_root: Path | None = None) -> dict:
    """Normaliza DOCX/Markdown e encaminha um DOCX auditado ao núcleo."""
    root = (work_root or ROOT).resolve()
    modelo, config = _carregar_artefatos(root)
    entrada = entrada.expanduser().resolve()
    if not entrada.is_file() or entrada.suffix.casefold() not in {".docx", ".md"}:
        raise ValueError("Informe uma petição inicial existente em DOCX ou Markdown (.md)")
    formato = "markdown" if entrada.suffix.casefold() == ".md" else "docx"
    if formato not in set(config.get("formatos_entrada_suportados", ["docx"])):
        raise ValueError(f"O formato de entrada {formato} não está autorizado em modelos/estilos.json")
    backup = backup_entrada(entrada, root)
    ingestao = normalizar_entrada(
        entrada,
        modelo,
        root,
        perfil_markdown=config.get("perfil_markdown"),
    )
    return executar_docx(
        ingestao.entrada_docx,
        semantica=semantica,
        skip_pdf=skip_pdf,
        work_root=root,
        metadados_entrada=ingestao.dados,
        manifesto_entrada=ingestao.manifesto,
        backup_original=backup,
    )


def main():
    ap = argparse.ArgumentParser(description="Padroniza uma petição DOCX ou Markdown sem alterar o original")
    ap.add_argument("entrada", type=Path)
    ap.add_argument("--semantica", type=Path)
    ap.add_argument("--skip-pdf", action="store_true")
    ap.add_argument("--work-root", type=Path, help="Pasta isolada para esta execução")
    args = ap.parse_args()
    print(executar(args.entrada, args.semantica, args.skip_pdf, args.work_root))


if __name__ == "__main__": main()
