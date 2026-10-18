"""Iniciador estável: ativa versões completas e recupera falhas na abertura."""
from __future__ import annotations

import argparse
import ctypes
import json
import os
import re
import shutil
import subprocess
import sys
import time
import uuid
from contextlib import contextmanager
from pathlib import Path

from scripts.protocolo_atualizacao import (
    EXECUTAVEL, MANIFESTO, dentro, gravar_json_atomico, ler_manifesto,
    sha256, verificar_pasta, versao,
)


def _ler(caminho: Path) -> dict:
    if caminho.is_symlink() or caminho.stat().st_size > 16384:
        raise ValueError("Estado de atualização inválido.")
    return json.loads(caminho.read_text(encoding="utf-8"))


@contextmanager
def trava(raiz: Path):
    arquivo = dentro(raiz, "atualizacao.lock").open("a+b")
    try:
        arquivo.seek(0)
        arquivo.write(b"0")
        arquivo.flush()
        arquivo.seek(0)
        if os.name == "nt":
            import msvcrt
            msvcrt.locking(arquivo.fileno(), msvcrt.LK_NBLCK, 1)
        else:
            import fcntl
            fcntl.flock(arquivo.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
        try:
            yield
        finally:
            arquivo.seek(0)
            if os.name == "nt":
                msvcrt.locking(arquivo.fileno(), msvcrt.LK_UNLCK, 1)
            else:
                fcntl.flock(arquivo.fileno(), fcntl.LOCK_UN)
    finally:
        arquivo.close()


def pasta_versao(raiz: Path, numero: str) -> Path:
    versao(numero)
    return dentro(raiz, "versoes/" + numero)


def versao_ativa(raiz: Path) -> str:
    opcoes = []
    for nome in ("versao_inicial.json", "versao_ativa.json"):
        caminho = dentro(raiz, nome)
        if caminho.is_file():
            dados = _ler(caminho)
            numero = dados["versao"]
            pasta = pasta_versao(raiz, numero)
            ler_manifesto(dentro(pasta, MANIFESTO), numero)
            if dentro(pasta, EXECUTAVEL).is_file():
                opcoes.append(numero)
    if not opcoes:
        raise ValueError("Não há uma versão completa instalada. Execute o instalador de recuperação.")
    return max(opcoes, key=versao)


def aguardar_processo(pid: int, segundos: int = 90) -> None:
    if pid <= 0:
        raise ValueError("Processo de origem inválido.")
    if os.name != "nt":
        raise OSError("A ativação com espera de processo requer Windows.")
    kernel = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel.OpenProcess.argtypes = [ctypes.c_uint32, ctypes.c_int, ctypes.c_uint32]
    kernel.OpenProcess.restype = ctypes.c_void_p
    kernel.WaitForSingleObject.argtypes = [ctypes.c_void_p, ctypes.c_uint32]
    kernel.CloseHandle.argtypes = [ctypes.c_void_p]
    handle = kernel.OpenProcess(0x00100000, False, pid)
    if not handle:
        if ctypes.get_last_error() == 87:
            return  # O programa já terminou.
        raise OSError("Não foi possível aguardar o encerramento do programa.")
    try:
        if kernel.WaitForSingleObject(handle, segundos * 1000) != 0:
            raise TimeoutError("O programa ainda está aberto; a versão anterior foi preservada.")
    finally:
        kernel.CloseHandle(handle)


def abrir(raiz: Path, numero: str, argumentos: list[str], confirmacao: Path | None = None, token: str = ""):
    pasta = pasta_versao(raiz, numero)
    comando = [str(dentro(pasta, EXECUTAVEL)), *argumentos]
    if confirmacao:
        comando.extend(["--confirmar-inicio", str(confirmacao), "--token-inicio", token])
    return subprocess.Popen(comando, cwd=pasta, close_fds=True)


def confirmar_inicio(caminho: Path | None, token: str, pasta: Path, numero: str) -> None:
    if caminho is None:
        return
    if pasta.parent.name != "versoes" or not re.fullmatch(r"[a-f0-9]{32}", token):
        raise ValueError("Confirmação de inicialização inválida.")
    raiz = pasta.parent.parent
    esperado = dentro(raiz, ".confirmacoes/" + token + ".json")
    if caminho.absolute() != esperado.absolute():
        raise ValueError("Destino inválido da confirmação.")
    gravar_json_atomico(esperado, {"versao": numero, "token": token, "pid": os.getpid()})


def _confirmada(caminho: Path, numero: str, token: str) -> bool:
    if not caminho.is_file():
        return False
    try:
        dados = _ler(caminho)
        return dados.get("versao") == numero and dados.get("token") == token
    except (ValueError, OSError):
        return False


def recuperar_interrupcao(raiz: Path) -> None:
    journal = dentro(raiz, "transacao_atualizacao.json")
    if not journal.is_file():
        return
    dados = _ler(journal)
    nova, anterior, token = dados["nova"], dados["anterior"], dados["token"]
    versao(nova)
    versao(anterior)
    if not isinstance(token, str) or not re.fullmatch(r"[a-f0-9]{32}", token):
        raise ValueError("Transação de atualização inválida.")
    confirmacao = dentro(raiz, ".confirmacoes/" + token + ".json")
    if _confirmada(confirmacao, nova, token):
        pasta = pasta_versao(raiz, nova)
        try:
            verificar_pasta(pasta, ler_manifesto(pasta / MANIFESTO, nova))
            gravar_json_atomico(raiz / "versao_ativa.json", {"versao": nova, "anterior": anterior})
        except (ValueError, OSError):
            gravar_json_atomico(raiz / "versao_ativa.json", {"versao": anterior})
    else:
        gravar_json_atomico(raiz / "versao_ativa.json", {"versao": anterior})
    journal.unlink()
    confirmacao.unlink(missing_ok=True)


def limpar_versoes_antigas(raiz: Path, atual: str, anterior: str) -> None:
    manter = {atual, anterior, _ler(raiz / "versao_inicial.json")["versao"]}
    versoes = dentro(raiz, "versoes")
    for candidata in versoes.iterdir():
        if candidata.name in manter or not re.fullmatch(r"\d+\.\d+\.\d+", candidata.name):
            continue
        try:
            alvo = pasta_versao(raiz, candidata.name)
            dados = ler_manifesto(dentro(alvo, MANIFESTO), candidata.name)
            esperados = {a["caminho"] for a in dados["arquivos"]} | {MANIFESTO}
            encontrados = set()
            for arquivo in alvo.rglob("*"):
                relativo = arquivo.relative_to(alvo).as_posix()
                dentro(alvo, relativo)
                if arquivo.is_file():
                    encontrados.add(relativo)
            # Pastas com arquivos desconhecidos não são removidas.
            if encontrados == esperados and alvo.parent == versoes:
                shutil.rmtree(alvo)
        except (OSError, ValueError):
            continue


def ativar(raiz: Path, numero: str, digest: str, *, iniciar=abrir, timeout: float = 40,
           argumentos: list[str] | None = None) -> bool:
    """A referência ativa só muda após o aplicativo confirmar que abriu."""
    with trava(raiz):
        recuperar_interrupcao(raiz)
        anterior = versao_ativa(raiz)
        processo, confirmacao = None, None
        journal = raiz / "transacao_atualizacao.json"
        try:
            if versao(numero) <= versao(anterior):
                raise ValueError("A atualização precisa ser posterior à versão instalada.")
            pasta = pasta_versao(raiz, numero)
            manifesto = dentro(pasta, MANIFESTO)
            if sha256(manifesto) != digest:
                raise ValueError("O manifesto foi alterado após o download.")
            verificar_pasta(pasta, ler_manifesto(manifesto, numero))
            token = uuid.uuid4().hex
            pasta_confirmacao = dentro(raiz, ".confirmacoes")
            pasta_confirmacao.mkdir(exist_ok=True)
            confirmacao = dentro(raiz, ".confirmacoes/" + token + ".json")
            gravar_json_atomico(journal, {"nova": numero, "anterior": anterior, "token": token})
            processo = iniciar(raiz, numero, argumentos or [], confirmacao, token)
            limite = time.monotonic() + timeout
            while not _confirmada(confirmacao, numero, token):
                if processo.poll() is not None or time.monotonic() >= limite:
                    raise RuntimeError("A nova versão não confirmou a abertura. A versão anterior foi restaurada.")
                time.sleep(0.1)
            gravar_json_atomico(raiz / "versao_ativa.json", {"versao": numero, "anterior": anterior})
            journal.unlink()
            confirmacao.unlink(missing_ok=True)
            (raiz / "erro_atualizacao.txt").unlink(missing_ok=True)
            try:
                limpar_versoes_antigas(raiz, numero, anterior)
            except (OSError, ValueError, KeyError):
                pass
            return True
        except Exception as exc:
            if processo is not None and processo.poll() is None:
                processo.terminate()
                try:
                    processo.wait(timeout=10)
                except subprocess.TimeoutExpired:
                    processo.kill()
                    processo.wait(timeout=10)
            gravar_json_atomico(raiz / "versao_ativa.json", {"versao": anterior})
            journal.unlink(missing_ok=True)
            if confirmacao:
                confirmacao.unlink(missing_ok=True)
            dentro(raiz, "erro_atualizacao.txt").write_text(str(exc), encoding="utf-8")
            iniciar(raiz, anterior, [])
            return False


def main() -> None:
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--ativar")
    parser.add_argument("--manifesto-sha256", default="")
    parser.add_argument("--aguardar-pid", type=int)
    parser.add_argument("--validar-ativacao", type=Path)
    args, restantes = parser.parse_known_args()
    raiz = Path(sys.executable if getattr(sys, "frozen", False) else __file__).resolve().parent
    try:
        if args.ativar:
            if args.aguardar_pid is not None:
                aguardar_processo(args.aguardar_pid)
            argumentos = ["--validar-interface", str(args.validar_ativacao)] if args.validar_ativacao else []
            raise SystemExit(0 if ativar(raiz, args.ativar, args.manifesto_sha256, argumentos=argumentos) else 1)
        with trava(raiz):
            recuperar_interrupcao(raiz)
            numero = versao_ativa(raiz)
        processo = abrir(raiz, numero, restantes)
        raise SystemExit(processo.wait())
    except Exception as exc:
        if os.name == "nt":
            ctypes.windll.user32.MessageBoxW(None, str(exc), "Atualização do Formatador", 0x10)
        else:
            print(exc, file=sys.stderr)
        raise SystemExit(1)


if __name__ == "__main__":
    main()
