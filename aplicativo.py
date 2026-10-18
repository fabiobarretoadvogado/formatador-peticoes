from __future__ import annotations

import argparse
import base64
import ctypes
import json
import os
import queue
import re
import shutil
import subprocess
import sys
import threading
import traceback
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path


ERRO_PREPARACAO_TK = ""


def _preparar_tcl_windows() -> None:
    """Inicializa o Tcl antes do Tk em instalações Windows incompletas.

    Algumas distribuições do Python incluem `_tkinter`, as DLLs e os scripts
    Tcl/Tk, mas o `Tcl_Init` não registra a biblioteca antes de procurar
    `init.tcl`.  Registrar os caminhos dentro de um interpretador Tcl
    temporário torna tanto a execução local quanto o pacote PyInstaller
    determinísticos, sem depender de variáveis permanentes do usuário.
    """

    if os.name != "nt":
        return

    # O runtime hook do PyInstaller define estas variáveis antes de carregar
    # o Tcl. Nesta distribuição do Windows isso faz o sistema de arquivos do
    # Tcl considerar até caminhos existentes como ausentes. A biblioteca é
    # registrada de forma segura no interpretador temporário mais abaixo.
    os.environ.pop("TCL_LIBRARY", None)
    os.environ.pop("TK_LIBRARY", None)

    bases: list[Path] = []
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        bases.append(Path(meipass))
    for base in (Path(sys.base_prefix), Path(sys.prefix)):
        if base not in bases:
            bases.append(base)

    def localizar_arquivo(*relativos: str) -> Path | None:
        for base in bases:
            for relativo in relativos:
                candidato = base / relativo
                if candidato.is_file():
                    return candidato
        return None

    def localizar_pasta(marcador: str, *relativos: str) -> Path | None:
        for base in bases:
            for relativo in relativos:
                candidato = base / relativo
                if (candidato / marcador).is_file():
                    return candidato
        return None

    tcl_dll = localizar_arquivo("tcl86t.dll", "DLLs/tcl86t.dll")
    tcl_library = localizar_pasta("init.tcl", "_tcl_data", "tcl/tcl8.6", "tcl8.6")
    tk_library = localizar_pasta("tk.tcl", "_tk_data", "tcl/tk8.6", "tk8.6")
    if not (tcl_dll and tcl_library and tk_library):
        raise RuntimeError("O runtime Tcl/Tk necessário ao frontend não foi encontrado.")

    tcl = ctypes.WinDLL(str(tcl_dll))
    tcl.Tcl_FindExecutable.argtypes = [ctypes.c_char_p]
    tcl.Tcl_CreateInterp.restype = ctypes.c_void_p
    tcl.Tcl_Eval.argtypes = [ctypes.c_void_p, ctypes.c_char_p]
    tcl.Tcl_Eval.restype = ctypes.c_int
    tcl.Tcl_Init.argtypes = [ctypes.c_void_p]
    tcl.Tcl_Init.restype = ctypes.c_int
    tcl.Tcl_GetStringResult.argtypes = [ctypes.c_void_p]
    tcl.Tcl_GetStringResult.restype = ctypes.c_char_p
    tcl.Tcl_DeleteInterp.argtypes = [ctypes.c_void_p]
    tcl.Tcl_FindExecutable(os.fsencode(sys.executable))

    interpretador = tcl.Tcl_CreateInterp()
    if not interpretador:
        raise RuntimeError("Não foi possível criar o interpretador Tcl do frontend.")
    try:
        biblioteca_tcl = tcl_library.as_posix().replace("}", "\\}")
        biblioteca_tk = tk_library.as_posix().replace("}", "\\}")
        comando = (
            f"set env(TCL_LIBRARY) {{{biblioteca_tcl}}}; "
            f"set env(TK_LIBRARY) {{{biblioteca_tk}}}"
        ).encode("utf-8")
        if tcl.Tcl_Eval(interpretador, comando) != 0 or tcl.Tcl_Init(interpretador) != 0:
            detalhe = tcl.Tcl_GetStringResult(interpretador)
            mensagem = detalhe.decode("utf-8", "replace") if detalhe else "falha desconhecida"
            raise RuntimeError(f"Falha ao inicializar Tcl/Tk: {mensagem}")
    finally:
        tcl.Tcl_DeleteInterp(interpretador)


try:
    _preparar_tcl_windows()
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    TK_IMPORTADO = True
except Exception:
    TK_IMPORTADO = False
    ERRO_PREPARACAO_TK = traceback.format_exc()

    class _TkAusente:
        class Tk:
            pass

    tk = _TkAusente()  # type: ignore[assignment]
    filedialog = messagebox = ttk = None  # type: ignore[assignment]

from scripts.pipeline import executar, executar_docx
from app_version import VERSION
from scripts.atualizacoes import consultar_atualizacao, preparar_incremental, iniciar_atualizacao
from scripts.iniciar_aplicativo import confirmar_inicio


APP_NAME = "Formatador de Petições Iniciais"
ENTRADA_FILTER = [
    ("Petições iniciais em Word ou Markdown", "*.docx *.md"),
    ("Documentos Word", "*.docx"),
    ("Markdown UTF-8", "*.md"),
]
ETAPAS = ("Selecionar", "Formatar", "Revisar", "Resultados")
CORES = {
    "fundo": "#F8FAFC",
    "cartao": "#FFFFFF",
    "texto": "#0F172A",
    "texto_secundario": "#475569",
    "primaria": "#2563EB",
    "primaria_escura": "#1D4ED8",
    "primaria_clara": "#DBEAFE",
    "borda": "#E2E8F0",
    "sucesso": "#15803D",
    "sucesso_claro": "#DCFCE7",
    "aviso": "#A16207",
    "aviso_claro": "#FEF3C7",
    "erro": "#DC2626",
}
ERRO_INTERFACE_TK = ERRO_PREPARACAO_TK


@dataclass(frozen=True)
class ResultadoPortatil:
    origem: Path
    entrada_pipeline: Path
    formato_original: str
    docx_intermediario: Path | None
    manifesto_ingestao: Path | None
    backup_original: Path | None
    pasta_resultado: Path
    docx: Path
    pdf: Path | None
    relatorio: Path
    ambiguos: int
    aprovado: bool
    erro_pdf: str | None
    execucao_local: Path
    gerar_pdf: bool


def nome_seguro(nome: str, limite: int = 80) -> str:
    base = str(nome).strip()
    for extensao in (".docx", ".md"):
        if base.lower().endswith(extensao):
            base = base[:-len(extensao)]
            break
    base = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", base)
    base = re.sub(r"\s+", " ", base).strip(" .")
    return (base or "peticao")[:limite]


def pasta_do_aplicativo() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def pasta_do_pacote() -> Path:
    return Path(getattr(sys, "_MEIPASS", pasta_do_aplicativo())).resolve()


def localizar_modelos() -> Path:
    candidatos = [
        pasta_do_aplicativo() / "modelos",
        pasta_do_pacote() / "modelos",
        Path(__file__).resolve().parent / "modelos",
    ]
    for pasta in candidatos:
        if (pasta / "modelo_oficial.docx").is_file() and (pasta / "estilos.json").is_file():
            return pasta
    caminhos = "\n".join(f"- {p}" for p in candidatos)
    raise FileNotFoundError(
        "O modelo oficial ou o arquivo de estilos não foi encontrado. "
        f"Pastas verificadas:\n{caminhos}"
    )


def raiz_execucoes_locais() -> Path:
    base = os.environ.get("LOCALAPPDATA")
    if base:
        return Path(base) / "FormatadorPeticoes" / "execucoes"
    return Path.home() / ".formatador_peticoes" / "execucoes"


def criar_pasta_unica(raiz: Path, stem: str) -> Path:
    raiz.mkdir(parents=True, exist_ok=True)
    prefixo = f"{nome_seguro(stem)}_{datetime.now():%Y%m%d_%H%M%S}"
    candidata = raiz / prefixo
    indice = 1
    while candidata.exists():
        candidata = raiz / f"{prefixo}_{indice:02d}"
        indice += 1
    candidata.mkdir(parents=True)
    return candidata


def validar_entrada(caminho: Path) -> Path:
    arquivo = caminho.expanduser().resolve()
    if not arquivo.is_file():
        raise FileNotFoundError(f"Arquivo não encontrado: {arquivo}")
    if arquivo.suffix.lower() not in {".docx", ".md"} or arquivo.name.startswith("~$"):
        raise ValueError(f"Somente petições iniciais em DOCX ou Markdown (.md) são aceitas: {arquivo.name}")
    try:
        with arquivo.open("rb") as stream:
            stream.read(1)
    except PermissionError as exc:
        complemento = " e fechado no Word" if arquivo.suffix.lower() == ".docx" else ""
        raise PermissionError(
            f"O Windows ou o OneDrive negou acesso a {arquivo.name}. "
            f"Confirme que o arquivo está disponível neste dispositivo{complemento}."
        ) from exc
    return arquivo


def processar_arquivo(
    origem: Path,
    destino_raiz: Path,
    gerar_pdf: bool = True,
    *,
    modelos_raiz: Path | None = None,
    execucoes_raiz: Path | None = None,
) -> ResultadoPortatil:
    origem = validar_entrada(origem)
    modelos = (modelos_raiz or localizar_modelos()).resolve()
    destino_raiz = destino_raiz.expanduser().resolve()
    execucoes_raiz = (execucoes_raiz or raiz_execucoes_locais()).resolve()

    execucao = criar_pasta_unica(execucoes_raiz, origem.stem)
    modelos_execucao = execucao / "modelos"
    modelos_execucao.mkdir(parents=True)
    shutil.copy2(modelos / "modelo_oficial.docx", modelos_execucao / "modelo_oficial.docx")
    shutil.copy2(modelos / "estilos.json", modelos_execucao / "estilos.json")

    resultado = executar(origem, skip_pdf=not gerar_pdf, work_root=execucao)
    pasta_resultado = criar_pasta_unica(destino_raiz, origem.stem)
    return _publicar_resultado(origem, resultado, execucao, pasta_resultado, gerar_pdf)


def _publicar_resultado(
    origem: Path,
    resultado: dict,
    execucao: Path,
    pasta_resultado: Path,
    gerar_pdf: bool,
) -> ResultadoPortatil:
    pasta_resultado.mkdir(parents=True, exist_ok=True)
    stem = nome_seguro(origem.stem)

    docx_origem = Path(resultado["docx"])
    docx_destino = pasta_resultado / f"{stem}_formatada.docx"
    shutil.copy2(docx_origem, docx_destino)

    relatorio_origem = Path(resultado["relatorio"])
    relatorio_destino = pasta_resultado / "relatorio_formatacao.md"
    shutil.copy2(relatorio_origem, relatorio_destino)

    manifesto_origem = Path(resultado["manifesto_ingestao"]) if resultado.get("manifesto_ingestao") else None
    manifesto_destino: Path | None = None
    if manifesto_origem and manifesto_origem.is_file():
        manifesto_destino = pasta_resultado / "manifesto_entrada.json"
        shutil.copy2(manifesto_origem, manifesto_destino)

    intermediario_origem = Path(resultado["docx_intermediario"]) if resultado.get("docx_intermediario") else None
    intermediario_publicado: Path | None = None
    if intermediario_origem and intermediario_origem.is_file():
        intermediario_publicado = pasta_resultado / "entrada_intermediaria.docx"
        shutil.copy2(intermediario_origem, intermediario_publicado)

    caminho_pdf_publicado = pasta_resultado / f"{stem}_formatada.pdf"
    pdf_destino: Path | None = None
    if resultado.get("pdf"):
        pdf_destino = caminho_pdf_publicado
        shutil.copy2(Path(resultado["pdf"]), pdf_destino)
    elif caminho_pdf_publicado.exists():
        caminho_pdf_publicado.unlink()

    auditoria = execucao / "trabalho" / "auditoria.json"
    if auditoria.is_file():
        shutil.copy2(auditoria, pasta_resultado / "auditoria.json")

    ambiguos = int(resultado.get("ambiguos", 0))
    solicitacoes = execucao / "trabalho" / "solicitacoes_semanticas.json"
    pendencias_publicadas = pasta_resultado / "pendencias_semanticas.json"
    if ambiguos > 0 and solicitacoes.is_file():
        shutil.copy2(solicitacoes, pendencias_publicadas)
    elif pendencias_publicadas.exists():
        pendencias_publicadas.unlink()

    resumo = {
        "origem": str(origem),
        "formato_original": resultado.get("formato_original", origem.suffix.lower().lstrip(".")),
        "entrada_pipeline": resultado.get("entrada_pipeline"),
        "docx_intermediario": str(intermediario_publicado) if intermediario_publicado else None,
        "manifesto_entrada": str(manifesto_destino) if manifesto_destino else None,
        "resultado": str(docx_destino),
        "pdf": str(pdf_destino) if pdf_destino else None,
        "tipo_peca": resultado.get("tipo_peca", "peticao_inicial"),
        "aprovado": bool(resultado.get("aprovado", False)),
        "pendencias_semanticas": ambiguos,
        "erro_pdf": resultado.get("erro_pdf"),
        "execucao_local": str(execucao),
        "data": datetime.now().astimezone().isoformat(timespec="seconds"),
    }
    (pasta_resultado / "resultado_execucao.json").write_text(
        json.dumps(resumo, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    return ResultadoPortatil(
        origem=origem,
        entrada_pipeline=Path(resultado.get("entrada_pipeline") or origem),
        formato_original=str(resultado.get("formato_original", "docx")),
        docx_intermediario=intermediario_origem,
        manifesto_ingestao=manifesto_origem,
        backup_original=Path(resultado["backup"]) if resultado.get("backup") else None,
        pasta_resultado=pasta_resultado,
        docx=docx_destino,
        pdf=pdf_destino,
        relatorio=relatorio_destino,
        ambiguos=ambiguos,
        aprovado=bool(resultado.get("aprovado", False)),
        erro_pdf=resultado.get("erro_pdf"),
        execucao_local=execucao,
        gerar_pdf=gerar_pdf,
    )


def carregar_dados_revisao(resultado: ResultadoPortatil) -> dict:
    caminho = resultado.execucao_local / "trabalho" / "solicitacoes_semanticas.json"
    if not caminho.is_file():
        raise FileNotFoundError("As pendências semânticas desta execução não foram encontradas.")
    return json.loads(caminho.read_text(encoding="utf-8"))


def reprocessar_resultado(
    resultado: ResultadoPortatil,
    decisoes: list[dict],
) -> ResultadoPortatil:
    dados = carregar_dados_revisao(resultado)
    itens = {int(item["id"]): item for item in dados.get("itens", [])}
    permitidos = set(dados.get("estilos_permitidos", []))
    normalizadas = []
    ids_vistos = set()
    for decisao in decisoes:
        identificador = int(decisao.get("id", 0))
        item = itens.get(identificador)
        estilo = decisao.get("estilo_destino")
        if not item:
            raise ValueError(f"A pendência {identificador} não pertence a esta execução.")
        if identificador in ids_vistos:
            raise ValueError(f"A pendência {identificador} foi respondida mais de uma vez.")
        if estilo not in permitidos:
            raise ValueError(f"O estilo escolhido para a pendência {identificador} não é autorizado.")
        if decisao.get("hash") != item.get("hash"):
            raise ValueError(
                f"O texto da pendência {identificador} mudou. Reabra a revisão antes de continuar."
            )
        ids_vistos.add(identificador)
        normalizadas.append({
            "id": identificador,
            "hash": item["hash"],
            "estilo_destino": estilo,
            "confianca": 1.0,
            "justificativa": decisao.get("justificativa") or "Estilo confirmado na interface de revisão.",
        })

    faltantes = sorted(set(itens) - ids_vistos)
    if faltantes:
        lista = ", ".join(str(identificador) for identificador in faltantes)
        raise ValueError(f"Confirme todas as pendências antes de continuar. Faltam: {lista}.")

    respostas = resultado.execucao_local / "trabalho" / "respostas_semanticas.json"
    respostas.write_text(
        json.dumps({"tipo_peca": "peticao_inicial", "classificacoes": normalizadas}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    metadados_entrada = None
    if resultado.manifesto_ingestao and resultado.manifesto_ingestao.is_file():
        metadados_entrada = json.loads(resultado.manifesto_ingestao.read_text(encoding="utf-8"))
    novo_resultado = executar_docx(
        resultado.entrada_pipeline,
        semantica=respostas,
        skip_pdf=not resultado.gerar_pdf,
        work_root=resultado.execucao_local,
        metadados_entrada=metadados_entrada,
        manifesto_entrada=resultado.manifesto_ingestao,
        backup_original=resultado.backup_original,
    )
    return _publicar_resultado(
        resultado.origem,
        novo_resultado,
        resultado.execucao_local,
        resultado.pasta_resultado,
        resultado.gerar_pdf,
    )


def pasta_saida_padrao(arquivos: list[Path]) -> Path:
    if arquivos:
        return arquivos[0].expanduser().resolve().parent / "Formatados"
    documentos = Path.home() / "Documents"
    return documentos / "Formatador de Petições Iniciais" / "Resultados"


def interface_tk_funcional() -> bool:
    global ERRO_INTERFACE_TK
    if not TK_IMPORTADO:
        ERRO_INTERFACE_TK = ERRO_PREPARACAO_TK or "O módulo tkinter não foi incluído nesta instalação."
        return False
    try:
        teste = tk.Tk()
        teste.withdraw()
        teste.update_idletasks()
        teste.destroy()
        ERRO_INTERFACE_TK = ""
        return True
    except Exception:
        ERRO_INTERFACE_TK = traceback.format_exc()
        return False


def _executar_powershell_dialogo(script: str) -> list[str]:
    flags = 0x08000000 if os.name == "nt" else 0
    processo = subprocess.run(
        ["powershell.exe", "-NoProfile", "-STA", "-Command", script],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        creationflags=flags,
        check=False,
    )
    if processo.returncode != 0:
        raise RuntimeError(processo.stderr.strip() or "Não foi possível abrir o seletor do Windows.")
    caminhos: list[str] = []
    for linha in processo.stdout.splitlines():
        linha = linha.strip()
        if not linha:
            continue
        try:
            caminhos.append(base64.b64decode(linha).decode("utf-8"))
        except (ValueError, UnicodeDecodeError):
            continue
    return caminhos


def selecionar_arquivos_nativo() -> list[Path]:
    script = r"""
Add-Type -AssemblyName System.Windows.Forms
$d = New-Object System.Windows.Forms.OpenFileDialog
  $d.Title = 'Escolha as petições iniciais para formatar'
  $d.Filter = 'Petições iniciais (*.docx;*.md)|*.docx;*.md|Documentos Word (*.docx)|*.docx|Markdown UTF-8 (*.md)|*.md'
$d.Multiselect = $true
if ($d.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
  foreach ($p in $d.FileNames) {
    [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($p))
  }
}
"""
    return [Path(p) for p in _executar_powershell_dialogo(script)]


def selecionar_pasta_nativo(inicial: Path) -> Path | None:
    inicial_b64 = base64.b64encode(str(inicial).encode("utf-8")).decode("ascii")
    script = rf"""
Add-Type -AssemblyName System.Windows.Forms
$inicial = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('{inicial_b64}'))
$d = New-Object System.Windows.Forms.FolderBrowserDialog
$d.Description = 'Escolha a pasta onde os resultados serão salvos'
if (Test-Path -LiteralPath $inicial) {{ $d.SelectedPath = $inicial }}
if ($d.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {{
  [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($d.SelectedPath))
}}
"""
    resultados = _executar_powershell_dialogo(script)
    return Path(resultados[0]) if resultados else None


def mensagem_nativa(texto: str, tipo: str = "info") -> None:
    if os.name != "nt":
        print(texto)
        return
    icones = {"info": 0x40, "aviso": 0x30, "erro": 0x10}
    ctypes.windll.user32.MessageBoxW(None, texto, APP_NAME, icones.get(tipo, 0x40))


def executar_modo_nativo(args: argparse.Namespace) -> int:
    arquivos = [Path(p) for p in args.arquivos]
    vieram_por_argumento = bool(arquivos)
    if not arquivos:
        try:
            arquivos = selecionar_arquivos_nativo()
        except Exception as exc:
            if not args.silencioso:
                mensagem_nativa(str(exc), "erro")
            return 1
    if not arquivos:
        return 0

    destino = args.saida or pasta_saida_padrao(arquivos)
    # In the native fallback, use the deterministic default directly.  The
    # portable build may not contain Tcl/Tk, so opening a second folder dialog
    # after the file picker makes the application look stalled.  A custom
    # destination remains available through --saida.

    if not args.silencioso:
        mensagem_nativa(
            f"O processamento de {len(arquivos)} documento(s) será iniciado.\n\n"
            "Aguarde a mensagem de conclusão.",
            "info",
        )

    resultados: list[ResultadoPortatil] = []
    falhas: list[tuple[Path, str]] = []
    for arquivo in arquivos:
        try:
            resultados.append(processar_arquivo(arquivo, Path(destino), not args.sem_pdf))
        except Exception as exc:
            falhas.append((arquivo, str(exc)))

    pendencias = sum(r.ambiguos for r in resultados)
    erros_pdf = sum(1 for r in resultados if r.erro_pdf)
    linhas = [f"{len(resultados)} documento(s) concluído(s)."]
    if pendencias:
        linhas.append(f"{pendencias} pendência(s) semântica(s).")
    if erros_pdf:
        linhas.append(f"PDF não gerado em {erros_pdf} documento(s); o DOCX foi preservado.")
    if falhas:
        linhas.append(f"{len(falhas)} falha(s):")
        linhas.extend(f"• {arquivo.name}: {erro}" for arquivo, erro in falhas)

    if not args.silencioso:
        mensagem_nativa("\n".join(linhas), "aviso" if falhas or pendencias or erros_pdf else "info")
        if resultados and os.name == "nt":
            Path(destino).mkdir(parents=True, exist_ok=True)
            os.startfile(Path(destino))  # type: ignore[attr-defined]
    return 1 if falhas else 0


class Aplicativo(tk.Tk):
    def __init__(
        self,
        arquivos_iniciais: list[Path] | None = None,
        saida_inicial: Path | None = None,
        gerar_pdf: bool = True,
        auto_iniciar: bool = False,
    ) -> None:
        super().__init__()
        self.title(f"{APP_NAME} — {VERSION}")
        self.geometry("1100x780")
        self.minsize(940, 680)
        self.configure(background=CORES["fundo"])

        self.arquivos: list[Path] = []
        self.resultados: list[ResultadoPortatil] = []
        self.falhas: list[tuple[Path, str]] = []
        self.pendencias_revisao: list[dict] = []
        self.decisoes: dict[tuple[str, int], dict] = {}
        self.indice_revisao = 0
        self.etapa_atual = 1
        self.fila: queue.Queue[tuple[str, object]] = queue.Queue()
        self.cancelar_evento = threading.Event()
        self.em_execucao = False
        self.atualizando = False

        self.saida_var = tk.StringVar()
        self.pdf_var = tk.BooleanVar(value=gerar_pdf)
        self.status_var = tk.StringVar(value="Pronto para selecionar as petições iniciais.")
        self.progresso_var = tk.DoubleVar(value=0)
        self.estilo_var = tk.StringVar()
        self.descricao_estilo_var = tk.StringVar()
        self.contador_revisao_var = tk.StringVar()

        self._configurar_estilos()
        self._montar_estrutura()
        self._mostrar_selecao()
        self.adicionar_arquivos(arquivos_iniciais or [])
        self.saida_var.set(str(saida_inicial or pasta_saida_padrao(self.arquivos)))

        self.bind_all("<Control-o>", lambda _e: self._selecionar_arquivos())
        self.bind_all("<Control-Return>", self._atalho_principal)
        self.protocol("WM_DELETE_WINDOW", self._fechar)
        self.after(100, self._ler_fila)

        if auto_iniciar and self.arquivos:
            self.after(450, self.iniciar_processamento)

    def _configurar_estilos(self) -> None:
        estilo = ttk.Style(self)
        try:
            estilo.theme_use("clam")
        except tk.TclError:
            pass
        estilo.configure(".", font=("Segoe UI", 10), foreground=CORES["texto"])
        estilo.configure("App.TFrame", background=CORES["fundo"])
        estilo.configure("Card.TFrame", background=CORES["cartao"], relief="solid", borderwidth=1)
        estilo.configure("CardBody.TFrame", background=CORES["cartao"], relief="flat", borderwidth=0)
        estilo.configure("TLabel", background=CORES["fundo"], foreground=CORES["texto"])
        estilo.configure("Card.TLabel", background=CORES["cartao"], foreground=CORES["texto"])
        estilo.configure(
            "Title.TLabel", background=CORES["fundo"], foreground=CORES["texto"],
            font=("Segoe UI Semibold", 21),
        )
        estilo.configure(
            "Section.TLabel", background=CORES["fundo"], foreground=CORES["texto"],
            font=("Segoe UI Semibold", 16),
        )
        estilo.configure(
            "CardTitle.TLabel", background=CORES["cartao"], foreground=CORES["texto"],
            font=("Segoe UI Semibold", 12),
        )
        estilo.configure(
            "Muted.TLabel", background=CORES["fundo"], foreground=CORES["texto_secundario"]
        )
        estilo.configure(
            "CardMuted.TLabel", background=CORES["cartao"], foreground=CORES["texto_secundario"]
        )
        estilo.configure(
            "Badge.TLabel", background=CORES["primaria_clara"], foreground=CORES["primaria_escura"],
            font=("Segoe UI Semibold", 9), padding=(9, 4),
        )
        estilo.configure(
            "Stepper.Active.TLabel", background=CORES["primaria"], foreground="#FFFFFF",
            font=("Segoe UI Semibold", 9), padding=(12, 7), anchor="center",
        )
        estilo.configure(
            "Stepper.Done.TLabel", background=CORES["sucesso_claro"], foreground=CORES["sucesso"],
            font=("Segoe UI Semibold", 9), padding=(12, 7), anchor="center",
        )
        estilo.configure(
            "Stepper.Pending.TLabel", background="#E2E8F0", foreground=CORES["texto_secundario"],
            font=("Segoe UI", 9), padding=(12, 7), anchor="center",
        )
        estilo.configure(
            "Primary.TButton", background=CORES["primaria"], foreground="#FFFFFF",
            font=("Segoe UI Semibold", 10), padding=(18, 10), borderwidth=0,
        )
        estilo.map(
            "Primary.TButton",
            background=[("active", CORES["primaria_escura"]), ("disabled", "#94A3B8")],
            foreground=[("disabled", "#F8FAFC")],
        )
        estilo.configure(
            "Secondary.TButton", background="#FFFFFF", foreground=CORES["texto"],
            padding=(14, 9), borderwidth=1,
        )
        estilo.map("Secondary.TButton", background=[("active", "#F1F5F9")])
        estilo.configure(
            "Danger.TButton", background="#FFFFFF", foreground=CORES["erro"],
            padding=(14, 9), borderwidth=1,
        )
        estilo.map("Danger.TButton", background=[("active", "#FEF2F2")])
        estilo.configure("Treeview", rowheight=30, background="#FFFFFF", fieldbackground="#FFFFFF")
        estilo.configure("Treeview.Heading", font=("Segoe UI Semibold", 9), padding=(8, 7))
        estilo.map("Treeview", background=[("selected", CORES["primaria_clara"])], foreground=[("selected", CORES["texto"])])
        estilo.configure("TEntry", padding=7)
        estilo.configure("TCombobox", padding=7)
        estilo.configure("Horizontal.TProgressbar", troughcolor="#E2E8F0", background=CORES["primaria"])

    def _montar_estrutura(self) -> None:
        principal = ttk.Frame(self, style="App.TFrame", padding=(30, 22, 30, 16))
        principal.pack(fill="both", expand=True)
        principal.columnconfigure(0, weight=1)
        principal.rowconfigure(2, weight=1)

        cabecalho = ttk.Frame(principal, style="App.TFrame")
        cabecalho.grid(row=0, column=0, sticky="ew")
        cabecalho.columnconfigure(0, weight=1)
        ttk.Label(cabecalho, text=APP_NAME, style="Title.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            cabecalho,
            text="Fluxo seguro para padronizar, revisar e abrir os documentos finais.",
            style="Muted.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(3, 0))
        ttk.Label(cabecalho, text="ESCOPO: PETIÇÃO INICIAL", style="Badge.TLabel").grid(
            row=0, column=1, sticky="e"
        )
        self.botao_atualizacao = ttk.Button(
            cabecalho, text=f"Versão {VERSION} · Verificar atualizações",
            command=self._verificar_atualizacoes,
        )
        self.botao_atualizacao.grid(row=1, column=1, sticky="e", padx=(12, 0))

        self.barra_etapas = ttk.Frame(principal, style="App.TFrame")
        self.barra_etapas.grid(row=1, column=0, sticky="ew", pady=(22, 18))
        for coluna in range(len(ETAPAS)):
            self.barra_etapas.columnconfigure(coluna, weight=1, uniform="etapas")
        self.rotulos_etapas: list[ttk.Label] = []
        for indice, nome in enumerate(ETAPAS, start=1):
            rotulo = ttk.Label(self.barra_etapas, text=f"{indice}. {nome}", anchor="center")
            rotulo.grid(row=0, column=indice - 1, sticky="ew", padx=(0 if indice == 1 else 4, 0))
            self.rotulos_etapas.append(rotulo)

        self.conteudo = ttk.Frame(principal, style="App.TFrame")
        self.conteudo.grid(row=2, column=0, sticky="nsew")
        self.conteudo.columnconfigure(0, weight=1)
        self.conteudo.rowconfigure(0, weight=1)

        rodape = ttk.Frame(principal, style="App.TFrame")
        rodape.grid(row=3, column=0, sticky="ew", pady=(12, 0))
        rodape.columnconfigure(0, weight=1)
        ttk.Separator(rodape).grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 9))
        ttk.Label(rodape, textvariable=self.status_var, style="Muted.TLabel").grid(row=1, column=0, sticky="w")
        ttk.Label(rodape, text="Ctrl+O: adicionar  |  Ctrl+Enter: ação principal", style="Muted.TLabel").grid(
            row=1, column=1, sticky="e"
        )

    def _limpar_conteudo(self) -> None:
        for filho in self.conteudo.winfo_children():
            filho.destroy()

    def _atualizar_etapas(self, etapa: int) -> None:
        self.etapa_atual = etapa
        for indice, rotulo in enumerate(self.rotulos_etapas, start=1):
            if indice < etapa:
                rotulo.configure(style="Stepper.Done.TLabel")
            elif indice == etapa:
                rotulo.configure(style="Stepper.Active.TLabel")
            else:
                rotulo.configure(style="Stepper.Pending.TLabel")

    def _mostrar_selecao(self) -> None:
        self._limpar_conteudo()
        self._atualizar_etapas(1)
        self.status_var.set("Selecione uma ou mais petições iniciais em DOCX ou Markdown.")
        painel = ttk.Frame(self.conteudo, style="App.TFrame")
        painel.grid(row=0, column=0, sticky="nsew")
        painel.columnconfigure(0, weight=1)
        painel.rowconfigure(2, weight=1)

        topo = ttk.Frame(painel, style="App.TFrame")
        topo.grid(row=0, column=0, sticky="ew")
        topo.columnconfigure(0, weight=1)
        ttk.Label(topo, text="Escolha os documentos", style="Section.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            topo,
            text="O original nunca é alterado. Cada resultado será salvo em uma pasta separada.",
            style="Muted.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(3, 0))
        ttk.Button(
            topo, text="Adicionar DOCX ou MD", style="Secondary.TButton", command=self._selecionar_arquivos
        ).grid(row=0, column=1, rowspan=2, sticky="e")

        barra = ttk.Frame(painel, style="App.TFrame")
        barra.grid(row=1, column=0, sticky="ew", pady=(15, 7))
        ttk.Label(barra, text="Petições iniciais selecionadas", style="Muted.TLabel").pack(side="left")

        self.quadro_lista = ttk.Frame(painel, style="Card.TFrame", padding=1)
        self.quadro_lista.grid(row=2, column=0, sticky="nsew")
        self.quadro_lista.columnconfigure(0, weight=1)
        self.quadro_lista.rowconfigure(0, weight=1)
        self.lista = ttk.Treeview(
            self.quadro_lista, columns=("documento", "pasta", "situacao"), show="headings", selectmode="extended"
        )
        self.lista.heading("documento", text="Documento")
        self.lista.heading("pasta", text="Local de origem")
        self.lista.heading("situacao", text="Situação")
        self.lista.column("documento", width=330, minwidth=180)
        self.lista.column("pasta", width=470, minwidth=220)
        self.lista.column("situacao", width=120, minwidth=100, anchor="center")
        self.lista.grid(row=0, column=0, sticky="nsew")
        rolagem = ttk.Scrollbar(self.quadro_lista, orient="vertical", command=self.lista.yview)
        rolagem.grid(row=0, column=1, sticky="ns")
        self.lista.configure(yscrollcommand=rolagem.set)

        self.acoes_lista = ttk.Frame(painel, style="App.TFrame")
        self.acoes_lista.grid(row=3, column=0, sticky="e", pady=(8, 0))
        self.botao_remover_selecionados = ttk.Button(
            self.acoes_lista, text="Remover selecionadas", command=self._remover_selecionado
        )
        self.botao_remover_selecionados.pack(side="right")
        self.botao_limpar_lista = ttk.Button(self.acoes_lista, text="Limpar lista", command=self._limpar)
        self.botao_limpar_lista.pack(side="right", padx=(0, 8))

        opcoes = ttk.Frame(painel, style="Card.TFrame", padding=15)
        opcoes.grid(row=4, column=0, sticky="ew", pady=(14, 0))
        opcoes.columnconfigure(0, weight=1)
        ttk.Label(opcoes, text="Pasta onde os resultados serão salvos", style="CardTitle.TLabel").grid(
            row=0, column=0, columnspan=2, sticky="w"
        )
        ttk.Entry(opcoes, textvariable=self.saida_var).grid(row=1, column=0, sticky="ew", pady=(7, 5))
        ttk.Button(opcoes, text="Escolher pasta", style="Secondary.TButton", command=self._selecionar_saida).grid(
            row=1, column=1, padx=(8, 0)
        )
        ttk.Label(
            opcoes, text="Uma subpasta será criada para cada documento.", style="CardMuted.TLabel"
        ).grid(row=2, column=0, sticky="w")
        ttk.Checkbutton(
            opcoes, text="Gerar também uma cópia em PDF usando o Microsoft Word", variable=self.pdf_var
        ).grid(row=3, column=0, columnspan=2, sticky="w", pady=(10, 0))
        try:
            localizar_modelos()
            texto_modelo, cor_modelo = "Modelo oficial e regras: disponíveis", CORES["sucesso"]
        except FileNotFoundError:
            texto_modelo, cor_modelo = "Modelo oficial ou regras: não encontrados", CORES["erro"]
        ttk.Label(
            opcoes, text=texto_modelo, style="CardMuted.TLabel", foreground=cor_modelo
        ).grid(row=4, column=0, columnspan=2, sticky="w", pady=(8, 0))

        acoes = ttk.Frame(painel, style="App.TFrame")
        acoes.grid(row=5, column=0, sticky="e", pady=(16, 0))
        self.botao_processar = ttk.Button(
            acoes, text="Formatar petições iniciais", style="Primary.TButton", command=self.iniciar_processamento
        )
        self.botao_processar.pack(side="right")
        self._atualizar_lista_selecao()

    def _atualizar_lista_selecao(self) -> None:
        if not hasattr(self, "lista") or not self.lista.winfo_exists():
            return
        for item in self.lista.get_children():
            self.lista.delete(item)
        for indice, arquivo in enumerate(self.arquivos):
            situacao = "Pronto" if arquivo.is_file() else "Não encontrado"
            self.lista.insert("", "end", iid=str(indice), values=(arquivo.name, str(arquivo.parent), situacao))

    def adicionar_arquivos(self, arquivos: list[Path]) -> None:
        existentes = {p.resolve() for p in self.arquivos}
        for caminho in arquivos:
            try:
                arquivo = caminho.expanduser().resolve()
            except OSError:
                continue
            if arquivo.suffix.lower() not in {".docx", ".md"} or arquivo.name.startswith("~$"):
                continue
            if arquivo not in existentes:
                self.arquivos.append(arquivo)
                existentes.add(arquivo)
        self._atualizar_lista_selecao()

    def _selecionar_arquivos(self) -> None:
        if self.em_execucao:
            return
        escolhidos = filedialog.askopenfilenames(
            title="Escolha as petições iniciais", filetypes=ENTRADA_FILTER
        )
        if escolhidos:
            estava_vazio = not self.arquivos
            if self.etapa_atual != 1:
                self._mostrar_selecao()
            self.adicionar_arquivos([Path(p) for p in escolhidos])
            if estava_vazio and self.arquivos:
                self.saida_var.set(str(pasta_saida_padrao(self.arquivos)))

    def _remover_selecionado(self) -> None:
        indices = sorted((int(item) for item in self.lista.selection()), reverse=True)
        for indice in indices:
            del self.arquivos[indice]
        self._atualizar_lista_selecao()

    def _limpar(self) -> None:
        self.arquivos.clear()
        self._atualizar_lista_selecao()

    def _selecionar_saida(self) -> None:
        escolhida = filedialog.askdirectory(
            title="Escolha a pasta dos resultados", initialdir=self.saida_var.get() or None
        )
        if escolhida:
            self.saida_var.set(escolhida)

    def iniciar_processamento(self) -> None:
        if self.atualizando:
            messagebox.showinfo("Atualizações", "Aguarde a consulta ou o download terminar.")
            return
        if self.em_execucao:
            return
        if not self.arquivos:
            messagebox.showwarning(APP_NAME, "Adicione pelo menos uma petição inicial em DOCX ou Markdown (.md).")
            return
        if not self.saida_var.get().strip():
            messagebox.showwarning(APP_NAME, "Escolha a pasta onde os resultados serão salvos.")
            return
        ausentes = [arquivo.name for arquivo in self.arquivos if not arquivo.is_file()]
        if ausentes:
            messagebox.showwarning(APP_NAME, "Estes arquivos não foram encontrados:\n\n" + "\n".join(ausentes))
            return

        self.em_execucao = True
        self.cancelar_evento.clear()
        self.resultados = []
        self.falhas = []
        self.decisoes.clear()
        self.progresso_var.set(0)
        self.status_var.set("Preparando a formatação...")
        self._mostrar_processamento(revisao=False)
        threading.Thread(
            target=self._executar_lote,
            args=(list(self.arquivos), Path(self.saida_var.get()), self.pdf_var.get()),
            daemon=True,
        ).start()

    def _mostrar_processamento(self, revisao: bool) -> None:
        self._limpar_conteudo()
        self._atualizar_etapas(3 if revisao else 2)
        painel = ttk.Frame(self.conteudo, style="App.TFrame")
        painel.grid(row=0, column=0, sticky="nsew")
        painel.columnconfigure(0, weight=1)
        painel.rowconfigure(2, weight=1)
        titulo = "Aplicando as decisões da revisão" if revisao else "Formatando os documentos"
        apoio = (
            "As escolhas confirmadas estão sendo aplicadas e novamente verificadas."
            if revisao else
            "O tempo varia conforme o tamanho dos arquivos e a geração do PDF."
        )
        ttk.Label(painel, text=titulo, style="Section.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(painel, text=apoio, style="Muted.TLabel").grid(row=1, column=0, sticky="w", pady=(3, 14))

        lista_frame = ttk.Frame(painel, style="Card.TFrame", padding=1)
        lista_frame.grid(row=2, column=0, sticky="nsew")
        lista_frame.columnconfigure(0, weight=1)
        lista_frame.rowconfigure(0, weight=1)
        self.lista_processamento = ttk.Treeview(
            lista_frame, columns=("documento", "situacao"), show="headings", selectmode="none"
        )
        self.lista_processamento.heading("documento", text="Documento")
        self.lista_processamento.heading("situacao", text="Situação")
        self.lista_processamento.column("documento", width=720)
        self.lista_processamento.column("situacao", width=210, anchor="center")
        self.lista_processamento.grid(row=0, column=0, sticky="nsew")
        documentos = [r.origem for r in self.resultados] if revisao else self.arquivos
        for indice, arquivo in enumerate(documentos):
            self.lista_processamento.insert("", "end", iid=str(indice), values=(arquivo.name, "Aguardando"))

        self.progresso = ttk.Progressbar(painel, maximum=100, variable=self.progresso_var)
        self.progresso.grid(row=3, column=0, sticky="ew", pady=(16, 8))
        linha = ttk.Frame(painel, style="App.TFrame")
        linha.grid(row=4, column=0, sticky="ew")
        linha.columnconfigure(0, weight=1)
        ttk.Label(linha, textvariable=self.status_var, style="Muted.TLabel").grid(row=0, column=0, sticky="w")
        if not revisao:
            self.botao_cancelar = ttk.Button(
                linha, text="Cancelar próximos documentos", style="Danger.TButton", command=self._cancelar_processamento
            )
            self.botao_cancelar.grid(row=0, column=1, sticky="e")

    def _cancelar_processamento(self) -> None:
        if not self.em_execucao:
            return
        self.cancelar_evento.set()
        if hasattr(self, "botao_cancelar") and self.botao_cancelar.winfo_exists():
            self.botao_cancelar.configure(state="disabled")
        self.status_var.set("Cancelamento solicitado. O documento atual será concluído com segurança.")

    def _executar_lote(self, arquivos: list[Path], destino: Path, gerar_pdf: bool) -> None:
        inicializou_com = False
        try:
            try:
                import pythoncom  # type: ignore

                pythoncom.CoInitialize()
                inicializou_com = True
            except ImportError:
                pass

            total = len(arquivos)
            resultados: list[ResultadoPortatil] = []
            falhas: list[tuple[Path, str]] = []
            for indice, arquivo in enumerate(arquivos):
                if self.cancelar_evento.is_set():
                    for restante in range(indice, total):
                        self.fila.put(("arquivo_status", (restante, "Cancelado")))
                    break
                self.fila.put(("arquivo_status", (indice, "Formatando")))
                self.fila.put(("status", f"Formatando {arquivo.name} ({indice + 1}/{total})..."))
                try:
                    resultado = processar_arquivo(arquivo, destino, gerar_pdf)
                    resultados.append(resultado)
                    situacao = "Revisão necessária" if resultado.ambiguos else "Formatado"
                    self.fila.put(("arquivo_status", (indice, situacao)))
                except Exception as exc:
                    falhas.append((arquivo, str(exc)))
                    self.fila.put(("arquivo_status", (indice, "Falha")))
                    traceback.print_exc()
                self.fila.put(("progresso", (indice + 1) / total * 100))
            self.fila.put(("concluido", (resultados, falhas)))
        except Exception as exc:
            self.fila.put(("fatal", str(exc)))
        finally:
            if inicializou_com:
                try:
                    import pythoncom  # type: ignore

                    pythoncom.CoUninitialize()
                except Exception:
                    pass

    def _preparar_revisao(self) -> None:
        self.pendencias_revisao = []
        for resultado in self.resultados:
            if resultado.ambiguos <= 0:
                continue
            dados = carregar_dados_revisao(resultado)
            for item in dados.get("itens", []):
                self.pendencias_revisao.append({"resultado": resultado, "dados": dados, "item": item})
        self.indice_revisao = min(self.indice_revisao, max(0, len(self.pendencias_revisao) - 1))

    def _mostrar_revisao(self) -> None:
        self._limpar_conteudo()
        self._atualizar_etapas(3)
        painel = ttk.Frame(self.conteudo, style="App.TFrame")
        painel.grid(row=0, column=0, sticky="nsew")
        painel.columnconfigure(1, weight=1)
        painel.rowconfigure(1, weight=1)

        cabecalho = ttk.Frame(painel, style="App.TFrame")
        cabecalho.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 12))
        cabecalho.columnconfigure(0, weight=1)
        ttk.Label(cabecalho, text="Revise os trechos duvidosos", style="Section.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            cabecalho,
            text="Confirme a função de cada trecho. O texto jurídico não será reescrito.",
            style="Muted.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(3, 0))
        ttk.Button(
            cabecalho, text="Aceitar todas as sugestões", style="Secondary.TButton",
            command=self._aceitar_todas_sugestoes,
        ).grid(row=0, column=1, rowspan=2, sticky="e")

        lateral = ttk.Frame(painel, style="Card.TFrame", padding=12)
        lateral.grid(row=1, column=0, sticky="nsew", padx=(0, 12))
        lateral.columnconfigure(0, weight=1)
        lateral.rowconfigure(2, weight=1)
        ttk.Label(lateral, text="Pendências", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(lateral, textvariable=self.contador_revisao_var, style="CardMuted.TLabel").grid(
            row=1, column=0, sticky="w", pady=(2, 8)
        )
        self.lista_revisao = ttk.Treeview(lateral, columns=("item", "status"), show="headings", height=15)
        self.lista_revisao.heading("item", text="Documento / trecho")
        self.lista_revisao.heading("status", text="Estado")
        self.lista_revisao.column("item", width=225, minwidth=180)
        self.lista_revisao.column("status", width=90, anchor="center")
        self.lista_revisao.grid(row=2, column=0, sticky="nsew")
        self.lista_revisao.bind("<<TreeviewSelect>>", self._selecionar_item_revisao)

        detalhe = ttk.Frame(painel, style="Card.TFrame", padding=16)
        detalhe.grid(row=1, column=1, sticky="nsew")
        detalhe.columnconfigure(0, weight=1)
        detalhe.rowconfigure(3, weight=1)
        self.revisao_documento_var = tk.StringVar()
        ttk.Label(detalhe, textvariable=self.revisao_documento_var, style="CardTitle.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        self.contexto_anterior = self._criar_caixa_contexto(detalhe, "Parágrafo anterior", 1, 1, "#F8FAFC")
        self.contexto_atual = self._criar_caixa_contexto(detalhe, "Trecho a classificar", 2, 2, "#EFF6FF")
        self.contexto_seguinte = self._criar_caixa_contexto(detalhe, "Parágrafo seguinte", 3, 1, "#F8FAFC")

        escolha = ttk.Frame(detalhe, style="CardBody.TFrame")
        escolha.grid(row=4, column=0, sticky="ew", pady=(13, 0))
        escolha.columnconfigure(1, weight=1)
        ttk.Label(escolha, text="Sugestão do programa", style="CardMuted.TLabel").grid(row=0, column=0, sticky="w")
        self.sugestao_var = tk.StringVar()
        ttk.Label(escolha, textvariable=self.sugestao_var, style="CardTitle.TLabel").grid(
            row=0, column=1, sticky="w", padx=(12, 0)
        )
        ttk.Label(escolha, text="Função confirmada", style="CardMuted.TLabel").grid(row=1, column=0, sticky="w", pady=(10, 0))
        self.combo_estilos = ttk.Combobox(escolha, textvariable=self.estilo_var, state="readonly")
        self.combo_estilos.grid(row=1, column=1, sticky="ew", padx=(12, 0), pady=(10, 0))
        self.combo_estilos.bind("<<ComboboxSelected>>", lambda _e: self._atualizar_descricao_estilo())
        ttk.Label(
            escolha, textvariable=self.descricao_estilo_var, style="CardMuted.TLabel", wraplength=610,
        ).grid(row=2, column=1, sticky="w", padx=(12, 0), pady=(5, 0))

        navegacao = ttk.Frame(detalhe, style="CardBody.TFrame")
        navegacao.grid(row=5, column=0, sticky="ew", pady=(14, 0))
        navegacao.columnconfigure(1, weight=1)
        self.botao_anterior = ttk.Button(
            navegacao, text="Anterior", style="Secondary.TButton", command=lambda: self._mover_revisao(-1)
        )
        self.botao_anterior.grid(row=0, column=0, sticky="w")
        self.botao_confirmar = ttk.Button(
            navegacao, text="Confirmar e próximo", style="Primary.TButton", command=self._confirmar_e_avancar
        )
        self.botao_confirmar.grid(row=0, column=2, sticky="e")

        acoes = ttk.Frame(painel, style="App.TFrame")
        acoes.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(14, 0))
        acoes.columnconfigure(0, weight=1)
        ttk.Button(acoes, text="Ver resultados parciais", style="Secondary.TButton", command=self._mostrar_resultados).grid(
            row=0, column=0, sticky="w"
        )
        self.botao_aplicar_revisao = ttk.Button(
            acoes, text="Aplicar revisão e concluir", style="Primary.TButton", command=self._aplicar_revisao
        )
        self.botao_aplicar_revisao.grid(row=0, column=1, sticky="e")
        self._preencher_lista_revisao()
        self._atualizar_item_revisao()

    def _criar_caixa_contexto(
        self, pai: ttk.Frame, titulo: str, linha: int, altura: int, fundo: str
    ) -> tk.Text:
        quadro = ttk.Frame(pai, style="CardBody.TFrame")
        quadro.grid(row=linha, column=0, sticky="nsew", pady=(10 if linha == 1 else 7, 0))
        quadro.columnconfigure(0, weight=1)
        ttk.Label(quadro, text=titulo, style="CardMuted.TLabel").grid(row=0, column=0, sticky="w")
        texto = tk.Text(
            quadro, height=altura, wrap="word", relief="solid", borderwidth=1,
            background=fundo, foreground=CORES["texto"], font=("Segoe UI", 10),
            padx=10, pady=8, highlightthickness=0,
        )
        texto.grid(row=1, column=0, sticky="nsew", pady=(4, 0))
        texto.configure(state="disabled")
        return texto

    @staticmethod
    def _definir_texto(caixa: tk.Text, valor: str) -> None:
        caixa.configure(state="normal")
        caixa.delete("1.0", "end")
        caixa.insert("1.0", valor or "—")
        caixa.configure(state="disabled")

    def _chave_pendencia(self, entrada: dict) -> tuple[str, int]:
        return (str(entrada["resultado"].execucao_local), int(entrada["item"]["id"]))

    def _preencher_lista_revisao(self) -> None:
        for item in self.lista_revisao.get_children():
            self.lista_revisao.delete(item)
        for indice, entrada in enumerate(self.pendencias_revisao):
            documento = entrada["resultado"].origem.stem
            identificador = entrada["item"]["id"]
            estado = "Confirmado" if self._chave_pendencia(entrada) in self.decisoes else "Pendente"
            self.lista_revisao.insert("", "end", iid=str(indice), values=(f"{documento} · trecho {identificador}", estado))
        confirmados = sum(self._chave_pendencia(entrada) in self.decisoes for entrada in self.pendencias_revisao)
        self.contador_revisao_var.set(f"{confirmados} de {len(self.pendencias_revisao)} confirmados")
        estado = "normal" if self.pendencias_revisao and confirmados == len(self.pendencias_revisao) else "disabled"
        self.botao_aplicar_revisao.configure(state=estado)

    def _selecionar_item_revisao(self, _evento: object = None) -> None:
        selecao = self.lista_revisao.selection()
        if not selecao:
            return
        self.indice_revisao = int(selecao[0])
        self._atualizar_item_revisao(sincronizar_lista=False)

    def _atualizar_item_revisao(self, sincronizar_lista: bool = True) -> None:
        if not self.pendencias_revisao:
            return
        entrada = self.pendencias_revisao[self.indice_revisao]
        item = entrada["item"]
        dados = entrada["dados"]
        resultado = entrada["resultado"]
        self.revisao_documento_var.set(f"{resultado.origem.name}  ·  trecho {item['id']}")
        self._definir_texto(self.contexto_anterior, item.get("paragrafo_anterior", ""))
        self._definir_texto(self.contexto_atual, item.get("paragrafo_atual", ""))
        self._definir_texto(self.contexto_seguinte, item.get("paragrafo_seguinte", ""))
        confianca = float(item.get("confianca_regra", 0)) * 100
        self.sugestao_var.set(f"{item.get('pre_classificacao', '—')}  ({confianca:.0f}% de confiança)")
        permitidos = list(dados.get("estilos_permitidos", []))
        self.combo_estilos.configure(values=permitidos)
        decisao = self.decisoes.get(self._chave_pendencia(entrada))
        self.estilo_var.set((decisao or {}).get("estilo_destino", item.get("pre_classificacao", "")))
        self._atualizar_descricao_estilo()
        self.botao_anterior.configure(state="normal" if self.indice_revisao > 0 else "disabled")
        ultimo = self.indice_revisao == len(self.pendencias_revisao) - 1
        self.botao_confirmar.configure(text="Confirmar" if ultimo else "Confirmar e próximo")
        if sincronizar_lista:
            self.lista_revisao.selection_set(str(self.indice_revisao))
            self.lista_revisao.see(str(self.indice_revisao))

    def _atualizar_descricao_estilo(self) -> None:
        if not self.pendencias_revisao:
            self.descricao_estilo_var.set("")
            return
        dados = self.pendencias_revisao[self.indice_revisao]["dados"]
        catalogo = dados.get("catalogo_semantico", {})
        descricao = catalogo.get(self.estilo_var.get(), {}).get("descricao", "Estilo autorizado pelo modelo oficial.")
        self.descricao_estilo_var.set(descricao)

    def _mover_revisao(self, deslocamento: int) -> None:
        novo = max(0, min(len(self.pendencias_revisao) - 1, self.indice_revisao + deslocamento))
        self.indice_revisao = novo
        self._atualizar_item_revisao()

    def _confirmar_e_avancar(self) -> None:
        if not self.estilo_var.get() or not self.pendencias_revisao:
            messagebox.showwarning(APP_NAME, "Escolha a função do trecho antes de confirmar.")
            return
        entrada = self.pendencias_revisao[self.indice_revisao]
        item = entrada["item"]
        self.decisoes[self._chave_pendencia(entrada)] = {
            "id": int(item["id"]),
            "hash": item["hash"],
            "estilo_destino": self.estilo_var.get(),
            "justificativa": "Estilo confirmado pela pessoa usuária na interface de revisão.",
        }
        if self.indice_revisao < len(self.pendencias_revisao) - 1:
            self.indice_revisao += 1
        self._preencher_lista_revisao()
        self._atualizar_item_revisao()

    def _aceitar_todas_sugestoes(self) -> None:
        if not messagebox.askyesno(
            APP_NAME,
            "Confirmar todas as sugestões automáticas?\n\nVocê ainda poderá revisar cada escolha antes de aplicar.",
        ):
            return
        for entrada in self.pendencias_revisao:
            item = entrada["item"]
            self.decisoes[self._chave_pendencia(entrada)] = {
                "id": int(item["id"]),
                "hash": item["hash"],
                "estilo_destino": item["pre_classificacao"],
                "justificativa": "Sugestão automática confirmada em lote pela pessoa usuária.",
            }
        self._preencher_lista_revisao()
        self._atualizar_item_revisao()

    def _aplicar_revisao(self) -> None:
        faltantes = [
            indice for indice, entrada in enumerate(self.pendencias_revisao)
            if self._chave_pendencia(entrada) not in self.decisoes
        ]
        if faltantes:
            self.indice_revisao = faltantes[0]
            self._atualizar_item_revisao()
            messagebox.showwarning(APP_NAME, "Confirme todas as pendências antes de concluir.")
            return
        self.em_execucao = True
        self.progresso_var.set(0)
        self.status_var.set("Aplicando as decisões confirmadas...")
        self._mostrar_processamento(revisao=True)
        threading.Thread(target=self._executar_revisao, daemon=True).start()

    def _executar_revisao(self) -> None:
        inicializou_com = False
        try:
            try:
                import pythoncom  # type: ignore

                pythoncom.CoInitialize()
                inicializou_com = True
            except ImportError:
                pass
            novos: list[ResultadoPortatil] = []
            falhas: list[tuple[Path, str]] = []
            total = len(self.resultados)
            for indice, resultado in enumerate(self.resultados):
                if resultado.ambiguos <= 0:
                    novos.append(resultado)
                    self.fila.put(("arquivo_status", (indice, "Já revisado")))
                else:
                    self.fila.put(("arquivo_status", (indice, "Aplicando revisão")))
                    self.fila.put(("status", f"Revisando {resultado.origem.name} ({indice + 1}/{total})..."))
                    prefixo = str(resultado.execucao_local)
                    decisoes = [valor for (execucao, _), valor in self.decisoes.items() if execucao == prefixo]
                    try:
                        novo = reprocessar_resultado(resultado, decisoes)
                        novos.append(novo)
                        self.fila.put(("arquivo_status", (indice, "Concluído" if not novo.ambiguos else "Ainda pendente")))
                    except Exception as exc:
                        novos.append(resultado)
                        falhas.append((resultado.origem, str(exc)))
                        self.fila.put(("arquivo_status", (indice, "Falha")))
                        traceback.print_exc()
                self.fila.put(("progresso", (indice + 1) / max(total, 1) * 100))
            self.fila.put(("revisao_concluida", (novos, falhas)))
        except Exception as exc:
            self.fila.put(("fatal", str(exc)))
        finally:
            if inicializou_com:
                try:
                    import pythoncom  # type: ignore

                    pythoncom.CoUninitialize()
                except Exception:
                    pass

    def _ler_fila(self) -> None:
        try:
            while True:
                evento, valor = self.fila.get_nowait()
                if evento == "status":
                    self.status_var.set(str(valor))
                elif evento == "progresso":
                    self.progresso_var.set(float(valor))
                elif evento == "arquivo_status":
                    indice, situacao = valor  # type: ignore[misc]
                    if hasattr(self, "lista_processamento") and self.lista_processamento.winfo_exists():
                        iid = str(indice)
                        if self.lista_processamento.exists(iid):
                            valores = list(self.lista_processamento.item(iid, "values"))
                            if len(valores) >= 2:
                                valores[1] = situacao
                                self.lista_processamento.item(iid, values=valores)
                elif evento == "concluido":
                    resultados, falhas = valor  # type: ignore[misc]
                    self._finalizar(resultados, falhas)
                elif evento == "revisao_concluida":
                    resultados, falhas = valor  # type: ignore[misc]
                    self._finalizar_revisao(resultados, falhas)
                elif evento == "fatal":
                    self._falha_fatal(str(valor))
                elif evento == "atualizacao_consultada":
                    self._oferecer_atualizacao(valor)
                elif evento == "atualizacao_erro":
                    self._encerrar_atualizacao()
                    messagebox.showerror("Atualizações", str(valor))
                elif evento == "atualizacao_preparada":
                    self._reiniciar_atualizado(valor)
        except queue.Empty:
            pass
        self.after(100, self._ler_fila)

    def _finalizar(
        self,
        resultados: list[ResultadoPortatil],
        falhas: list[tuple[Path, str]],
    ) -> None:
        self.em_execucao = False
        self.resultados = resultados
        self.falhas = falhas
        pendencias = sum(resultado.ambiguos for resultado in resultados)
        if pendencias:
            self.status_var.set(f"Formatação concluída. {pendencias} trecho(s) precisam de confirmação.")
            try:
                self._preparar_revisao()
                self._mostrar_revisao()
            except Exception as exc:
                self.falhas.append((Path("revisão"), str(exc)))
                self._mostrar_resultados()
                messagebox.showerror(APP_NAME, f"Não foi possível abrir a revisão:\n\n{exc}")
        else:
            self.status_var.set("Formatação e verificação concluídas.")
            self._mostrar_resultados()

    def _finalizar_revisao(
        self,
        resultados: list[ResultadoPortatil],
        falhas: list[tuple[Path, str]],
    ) -> None:
        self.em_execucao = False
        self.resultados = resultados
        if falhas:
            self.falhas.extend(falhas)
            self._preparar_revisao()
            self._mostrar_revisao()
            detalhes = "\n".join(f"• {arquivo.name}: {erro}" for arquivo, erro in falhas)
            messagebox.showerror(APP_NAME, "A revisão não pôde ser aplicada a todos os documentos:\n\n" + detalhes)
            return
        self.status_var.set("Revisão aplicada. Os documentos finais estão prontos para conferência no Word.")
        self._mostrar_resultados()

    def _situacao_resultado(self, resultado: ResultadoPortatil) -> str:
        if resultado.ambiguos:
            return "Revisão necessária"
        if not resultado.aprovado:
            return "Verificar relatório"
        return "Formatado e verificado"

    def _mostrar_resultados(self) -> None:
        self._limpar_conteudo()
        self._atualizar_etapas(4)
        painel = ttk.Frame(self.conteudo, style="App.TFrame")
        painel.grid(row=0, column=0, sticky="nsew")
        painel.columnconfigure(0, weight=1)
        painel.rowconfigure(2, weight=1)
        concluidos = sum(not resultado.ambiguos and resultado.aprovado for resultado in self.resultados)
        pendencias = sum(resultado.ambiguos for resultado in self.resultados)
        ttk.Label(painel, text="Resultados", style="Section.TLabel").grid(row=0, column=0, sticky="w")
        resumo = f"{len(self.resultados)} documento(s) processado(s) · {concluidos} verificado(s)"
        if pendencias:
            resumo += f" · {pendencias} trecho(s) ainda pendente(s)"
        if self.falhas:
            resumo += f" · {len(self.falhas)} falha(s)"
        ttk.Label(painel, text=resumo, style="Muted.TLabel").grid(row=1, column=0, sticky="w", pady=(3, 14))

        lista_frame = ttk.Frame(painel, style="Card.TFrame", padding=1)
        lista_frame.grid(row=2, column=0, sticky="nsew")
        lista_frame.columnconfigure(0, weight=1)
        lista_frame.rowconfigure(0, weight=1)
        self.lista_resultados = ttk.Treeview(
            lista_frame, columns=("documento", "situacao", "docx", "pdf"), show="headings", selectmode="browse"
        )
        for coluna, titulo in (("documento", "Documento"), ("situacao", "Situação"), ("docx", "DOCX"), ("pdf", "PDF")):
            self.lista_resultados.heading(coluna, text=titulo)
        self.lista_resultados.column("documento", width=480)
        self.lista_resultados.column("situacao", width=220, anchor="center")
        self.lista_resultados.column("docx", width=95, anchor="center")
        self.lista_resultados.column("pdf", width=95, anchor="center")
        self.lista_resultados.grid(row=0, column=0, sticky="nsew")
        self.lista_resultados.bind("<<TreeviewSelect>>", lambda _e: self._atualizar_botoes_resultado())
        for indice, resultado in enumerate(self.resultados):
            pdf = "Disponível" if resultado.pdf else ("Falhou" if resultado.erro_pdf else "Não solicitado")
            self.lista_resultados.insert(
                "", "end", iid=str(indice),
                values=(resultado.origem.name, self._situacao_resultado(resultado), "Disponível", pdf),
            )
        if self.resultados:
            self.lista_resultados.selection_set("0")

        if self.falhas:
            texto_falhas = " | ".join(f"{arquivo.name}: {erro}" for arquivo, erro in self.falhas)
            ttk.Label(
                painel, text="Atenção: " + texto_falhas, foreground=CORES["erro"], wraplength=980,
            ).grid(row=3, column=0, sticky="w", pady=(9, 0))

        abertura = ttk.Frame(painel, style="App.TFrame")
        abertura.grid(row=4, column=0, sticky="ew", pady=(14, 0))
        self.botao_abrir_docx = ttk.Button(
            abertura, text="Abrir DOCX", style="Secondary.TButton", command=lambda: self._abrir_resultado("docx")
        )
        self.botao_abrir_docx.pack(side="left")
        self.botao_abrir_pdf = ttk.Button(
            abertura, text="Abrir PDF", style="Secondary.TButton", command=lambda: self._abrir_resultado("pdf")
        )
        self.botao_abrir_pdf.pack(side="left", padx=(8, 0))
        self.botao_abrir_relatorio = ttk.Button(
            abertura, text="Abrir relatório", style="Secondary.TButton", command=lambda: self._abrir_resultado("relatorio")
        )
        self.botao_abrir_relatorio.pack(side="left", padx=(8, 0))
        self.botao_abrir_pasta = ttk.Button(
            abertura, text="Abrir pasta", style="Secondary.TButton", command=lambda: self._abrir_resultado("pasta_resultado")
        )
        self.botao_abrir_pasta.pack(side="left", padx=(8, 0))

        acoes = ttk.Frame(painel, style="App.TFrame")
        acoes.grid(row=5, column=0, sticky="ew", pady=(14, 0))
        acoes.columnconfigure(0, weight=1)
        if pendencias:
            ttk.Button(
                acoes, text="Voltar à revisão", style="Secondary.TButton", command=self._voltar_revisao
            ).grid(row=0, column=0, sticky="w")
        ttk.Button(
            acoes, text="Nova formatação", style="Primary.TButton", command=self._novo_lote
        ).grid(row=0, column=1, sticky="e")
        self._atualizar_botoes_resultado()

    def _resultado_selecionado(self) -> ResultadoPortatil | None:
        if not hasattr(self, "lista_resultados") or not self.lista_resultados.winfo_exists():
            return None
        selecao = self.lista_resultados.selection()
        return self.resultados[int(selecao[0])] if selecao else None

    def _atualizar_botoes_resultado(self) -> None:
        resultado = self._resultado_selecionado()
        if not resultado:
            for nome in ("botao_abrir_docx", "botao_abrir_pdf", "botao_abrir_relatorio", "botao_abrir_pasta"):
                if hasattr(self, nome):
                    getattr(self, nome).configure(state="disabled")
            return
        self.botao_abrir_docx.configure(state="normal" if resultado.docx.is_file() else "disabled")
        self.botao_abrir_pdf.configure(state="normal" if resultado.pdf and resultado.pdf.is_file() else "disabled")
        self.botao_abrir_relatorio.configure(state="normal" if resultado.relatorio.is_file() else "disabled")
        self.botao_abrir_pasta.configure(state="normal" if resultado.pasta_resultado.is_dir() else "disabled")

    def _abrir_resultado(self, atributo: str) -> None:
        resultado = self._resultado_selecionado()
        if not resultado:
            return
        caminho = getattr(resultado, atributo, None)
        if not caminho or not Path(caminho).exists():
            messagebox.showwarning(APP_NAME, "O arquivo selecionado não está disponível.")
            return
        if os.name == "nt":
            os.startfile(Path(caminho))  # type: ignore[attr-defined]

    def _voltar_revisao(self) -> None:
        self._preparar_revisao()
        self._mostrar_revisao()

    def _novo_lote(self) -> None:
        self.arquivos.clear()
        self.resultados.clear()
        self.falhas.clear()
        self.pendencias_revisao.clear()
        self.decisoes.clear()
        self.indice_revisao = 0
        self.progresso_var.set(0)
        self._mostrar_selecao()

    def _falha_fatal(self, erro: str) -> None:
        self.em_execucao = False
        self.status_var.set("A execução foi interrompida.")
        if self.resultados:
            self._mostrar_resultados()
        else:
            self._mostrar_selecao()
        messagebox.showerror(APP_NAME, erro)

    def _atalho_principal(self, _evento: object = None) -> None:
        if self.em_execucao:
            return
        if self.etapa_atual == 1:
            self.iniciar_processamento()
        elif self.etapa_atual == 3:
            if len(self.decisoes) == len(self.pendencias_revisao):
                self._aplicar_revisao()
            else:
                self._confirmar_e_avancar()
        elif self.etapa_atual == 4:
            self._novo_lote()

    def _fechar(self) -> None:
        if self.atualizando:
            messagebox.showinfo("Atualizações", "Aguarde a consulta ou o download da atualização terminar.")
            return
        if self.em_execucao:
            if messagebox.askyesno(
                APP_NAME,
                "Há um documento em processamento.\n\nDeseja cancelar os próximos e aguardar a conclusão segura do atual?",
            ):
                self._cancelar_processamento()
            return
        self.destroy()

    def _encerrar_atualizacao(self) -> None:
        self.atualizando = False
        self.botao_atualizacao.configure(state="normal")
        self.status_var.set("Pronto.")

    def _verificar_atualizacoes(self) -> None:
        if self.em_execucao or self.atualizando:
            messagebox.showinfo("Atualizações", "Conclua a operação atual antes de atualizar.")
            return
        if not getattr(sys, "frozen", False) or pasta_do_aplicativo().parent.name != "versoes":
            messagebox.showinfo("Atualizações", "O novo atualizador está disponível na versão instalada do programa.")
            return
        self.atualizando = True
        self.botao_atualizacao.configure(state="disabled")
        self.status_var.set("Consultando novas versões no GitHub…")

        def consultar() -> None:
            try:
                self.fila.put(("atualizacao_consultada", consultar_atualizacao()))
            except Exception as exc:
                self.fila.put(("atualizacao_erro", f"Não foi possível consultar as atualizações.\n\n{exc}"))

        threading.Thread(target=consultar, daemon=True).start()

    def _oferecer_atualizacao(self, atualizacao) -> None:
        if atualizacao is None:
            self._encerrar_atualizacao()
            messagebox.showinfo("Atualizações", f"Você já utiliza a versão mais recente ({VERSION}).")
            return
        if self.em_execucao or any(resultado.ambiguos for resultado in self.resultados):
            self._encerrar_atualizacao()
            messagebox.showinfo("Atualizações", "Há uma nova versão. Conclua a formatação e a revisão antes de instalá-la.")
            return
        if atualizacao.tipo != "incremental":
            self._encerrar_atualizacao()
            messagebox.showinfo("Atualizações", "A nova versão ainda não oferece atualização automática. Tente novamente quando a publicação estiver completa.")
            return
        if not messagebox.askyesno(
            "Atualizações",
            f"A versão {atualizacao.versao} está disponível.\n\n"
            "Atualizar e reiniciar? O programa prepara a atualização e reabre automaticamente. "
            "Seus documentos e resultados serão preservados.",
        ):
            self._encerrar_atualizacao()
            return
        self.status_var.set("Preparando a atualização e conferindo os arquivos…")

        def baixar() -> None:
            try:
                destino = raiz_execucoes_locais().parent / "atualizacoes" / atualizacao.versao
                preparada = preparar_incremental(atualizacao, pasta_do_aplicativo(), destino.parent)
                self.fila.put(("atualizacao_preparada", preparada))
            except Exception as exc:
                self.fila.put(("atualizacao_erro", f"Não foi possível baixar a atualização.\n\n{exc}"))

        threading.Thread(target=baixar, daemon=True).start()

    def _reiniciar_atualizado(self, preparada) -> None:
        if self.em_execucao or any(resultado.ambiguos for resultado in self.resultados):
            self._encerrar_atualizacao()
            messagebox.showinfo("Atualizações", "Conclua a revisão e consulte novamente para aplicar a atualização.")
            return
        try:
            iniciar_atualizacao(preparada)
        except (OSError, ValueError) as exc:
            self._encerrar_atualizacao()
            messagebox.showerror("Atualizações", f"Não foi possível iniciar a atualização.\n\n{exc}")
            return
        self.destroy()


def analisar_argumentos(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("arquivos", nargs="*")
    parser.add_argument("--saida", type=Path)
    parser.add_argument("--sem-pdf", action="store_true")
    parser.add_argument("--nao-iniciar", action="store_true")
    parser.add_argument("--silencioso", action="store_true")
    parser.add_argument("--validar-interface", type=Path)
    parser.add_argument("--confirmar-inicio", type=Path)
    parser.add_argument("--token-inicio", default="")
    parser.add_argument("--versao", action="version", version=VERSION)
    return parser.parse_args(argv)


def main() -> None:
    global ERRO_INTERFACE_TK
    args = analisar_argumentos()
    if args.validar_interface:
        funcional = interface_tk_funcional()
        if funcional:
            try:
                teste_aplicativo = Aplicativo(auto_iniciar=False)
                teste_aplicativo.withdraw()
                teste_aplicativo.update_idletasks()
                teste_aplicativo.destroy()
            except Exception:
                funcional = False
                ERRO_INTERFACE_TK = traceback.format_exc()
        args.validar_interface.parent.mkdir(parents=True, exist_ok=True)
        args.validar_interface.write_text(
            json.dumps(
                {"funcional": funcional, "tk_importado": TK_IMPORTADO, "erro": ERRO_INTERFACE_TK},
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )
        if funcional:
            confirmar_inicio(args.confirmar_inicio, args.token_inicio, pasta_do_aplicativo(), VERSION)
        raise SystemExit(0 if funcional else 2)
    if args.silencioso or not interface_tk_funcional():
        raise SystemExit(executar_modo_nativo(args))
    arquivos = [Path(p) for p in args.arquivos]
    app = Aplicativo(
        arquivos_iniciais=arquivos,
        saida_inicial=args.saida,
        gerar_pdf=not args.sem_pdf,
        auto_iniciar=bool(arquivos) and not args.nao_iniciar,
    )
    app.update_idletasks()
    app.after(150, lambda: confirmar_inicio(args.confirmar_inicio, args.token_inicio, pasta_do_aplicativo(), VERSION))
    erro_atualizacao = pasta_do_aplicativo().parent.parent / "erro_atualizacao.txt"
    if pasta_do_aplicativo().parent.name == "versoes" and erro_atualizacao.is_file():
        erro = erro_atualizacao.read_text(encoding="utf-8")[:2000]
        erro_atualizacao.unlink()
        app.after(300, lambda: messagebox.showwarning("Atualizações", "A atualização não pôde ser concluída. A versão anterior foi mantida.\n\n" + erro))
    app.mainloop()


if __name__ == "__main__":
    main()
