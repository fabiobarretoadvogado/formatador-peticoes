from __future__ import annotations

import argparse
import sys
import tempfile
from pathlib import Path

if __package__ in (None, ""):
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
    from scripts.aplicar_estilos import aplicar
    from scripts.assinatura_local import AssinaturaLocal, carregar_assinatura_local
    from scripts.classificar_regras import classificar
    from scripts.comum import ROOT, TIPO_PECA_SUPORTADO, backup_entrada, carregar_json, nome_seguro, publicar_docx
    from scripts.configurar_visibilidade_estilos import IDIOMA_REVISAO, ajustar_arquivo
    from scripts.converter_markdown import normalizar_entrada
    from scripts.extrair_documento import extrair
    from scripts.validar_documento import validar
else:
    from .aplicar_estilos import aplicar
    from .assinatura_local import AssinaturaLocal, carregar_assinatura_local
    from .classificar_regras import classificar
    from .comum import ROOT, TIPO_PECA_SUPORTADO, backup_entrada, carregar_json, nome_seguro, publicar_docx
    from .configurar_visibilidade_estilos import IDIOMA_REVISAO, ajustar_arquivo
    from .converter_markdown import normalizar_entrada
    from .extrair_documento import extrair
    from .validar_documento import validar


def _consolidar(classificacao: dict, semantica: Path | dict | None, config: dict) -> int:
    if semantica is None or isinstance(semantica, Path) and not semantica.exists():
        return 0
    permitidos = set(config["estilos_permitidos"])
    respostas = carregar_json(semantica) if isinstance(semantica, Path) else semantica
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


def _carregar_artefatos(modelos: Path) -> tuple[Path, dict]:
    modelo = modelos / "modelo_oficial.docx"
    config_path = modelos / "estilos.json"
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
    semantica: Path | dict | None = None,
    work_root: Path | None = None,
    *,
    modelos_raiz: Path | None = None,
    metadados_entrada: dict | None = None,
    backup_original: Path | None = None,
    assinatura: AssinaturaLocal | None = None,
) -> dict:
    """Executa o núcleo estável do formatador sobre uma entrada DOCX."""
    root = (work_root or ROOT).resolve()
    modelo, config = _carregar_artefatos(modelos_raiz or root / "modelos")
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
    classificacao = classificar(estrutura)
    _consolidar(classificacao, semantica, config)
    saida_docx = root / "saida" / "peticao_formatada.docx"
    aplicacao = aplicar(entrada, modelo, classificacao, config, saida_docx, assinatura=assinatura)
    paragrafos_assinatura = set(aplicacao["paragrafos_assinatura_local"])
    pendencias = [item for item in classificacao["ambiguos"] if item["id"] not in paragrafos_assinatura]
    solicitacoes = {
        "estilos_permitidos": config["estilos_permitidos"],
        "tipo_peca": TIPO_PECA_SUPORTADO,
        "regras_contextuais": config.get("regras_contextuais", {}),
        "formatacao_estilos": config.get("formatacao_estilos", {}),
        "catalogo_semantico": config.get("catalogo_semantico", {}),
        "itens": pendencias,
    }
    # The formatter imports the authored styles into the output document. Apply
    # the same visibility policy to the final package so Word does not expose
    # its built-in styles in the user-facing style pane.
    trabalho = root / "trabalho"
    trabalho.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(dir=trabalho) as temporario_dir:
        temporario = Path(temporario_dir) / saida_docx.name
        ajustar_arquivo(saida_docx, temporario, config["estilos_permitidos"], idioma_revisao=IDIOMA_REVISAO)
        temporario.replace(saida_docx)
    auditoria = validar(saida_docx, config, aplicacao)
    pendencias_semanticas = len(pendencias)
    auditoria["resumo"]["pendencias_semanticas"] = pendencias_semanticas
    auditoria["resumo"]["aprovado"] = auditoria["resumo"]["aprovado"] and pendencias_semanticas == 0
    return {
        "docx": str(saida_docx),
        "backup": str(backup),
        "tipo_peca": TIPO_PECA_SUPORTADO,
        "ambiguos": pendencias_semanticas,
        "aprovado": auditoria["resumo"]["aprovado"],
        "entrada_original": metadados_entrada.get("origem", str(entrada)),
        "formato_original": metadados_entrada.get("formato_original", "docx"),
        "entrada_pipeline": str(entrada),
        "docx_intermediario": metadados_entrada.get("docx_intermediario"),
        "metadados_entrada": metadados_entrada,
        "solicitacoes_semanticas": solicitacoes,
        "auditoria": auditoria,
    }


def executar(
    entrada: Path,
    semantica: Path | dict | None = None,
    work_root: Path | None = None,
    *,
    modelos_raiz: Path | None = None,
    assinatura: AssinaturaLocal | None = None,
) -> dict:
    """Normaliza DOCX/Markdown e encaminha um DOCX auditado ao núcleo."""
    root = (work_root or ROOT).resolve()
    modelo, config = _carregar_artefatos(modelos_raiz or root / "modelos")
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
        backup if formato == "docx" else ingestao.entrada_docx,
        semantica=semantica,
        work_root=root,
        modelos_raiz=modelos_raiz,
        metadados_entrada=ingestao.dados,
        backup_original=backup,
        assinatura=assinatura,
    )


def main():
    ap = argparse.ArgumentParser(description="Padroniza uma petição DOCX ou Markdown sem alterar o original")
    ap.add_argument("entrada", type=Path)
    ap.add_argument("--semantica", type=Path)
    ap.add_argument("--skip-pdf", action="store_true", help=argparse.SUPPRESS)
    ap.add_argument("--work-root", type=Path, help="Pasta que contém os modelos oficiais")
    ap.add_argument("--saida", type=Path, help="Pasta de destino do Word formatado")
    args = ap.parse_args()
    entrada = args.entrada.expanduser().resolve()
    modelos = (args.work_root or ROOT) / "modelos"
    destino = (args.saida or entrada.parent).expanduser().resolve()
    with tempfile.TemporaryDirectory(prefix="formatador-") as temporario:
        resultado = executar(
            entrada, args.semantica, Path(temporario), modelos_raiz=modelos, assinatura=carregar_assinatura_local(),
        )
        print(publicar_docx(Path(resultado["docx"]), destino, nome_seguro(entrada.stem)))


if __name__ == "__main__": main()
