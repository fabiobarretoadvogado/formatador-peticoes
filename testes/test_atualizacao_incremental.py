import copy
import hashlib
import io
import json
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch

from scripts.atualizacoes import (
    _extrair_necessarios, interpretar_release, preparar_incremental,
)
from scripts.iniciar_aplicativo import ativar, recuperar_interrupcao, versao_ativa, confirmar_inicio
from scripts.pacote_atualizacao import gerar_pacote
from scripts.protocolo_atualizacao import (
    EXECUTAVEL, INICIADOR, MANIFESTO, REPOSITORIO, gravar_json_atomico,
    nome_pacote, sha256, validar_manifesto,
)


def release(numero, arquivos):
    return {"tag_name": "v" + numero, "assets": [
        {"name": nome, "browser_download_url": f"https://github.com/{REPOSITORIO}/releases/download/v{numero}/{nome}",
         "size": len(conteudo), "digest": "sha256:" + hashlib.sha256(conteudo).hexdigest()}
        for nome, conteudo in arquivos.items()
    ]}


class Processo:
    def __init__(self, codigo=None):
        self.codigo = codigo
        self.encerrado = False

    def poll(self):
        return self.codigo

    def terminate(self):
        self.encerrado = True
        self.codigo = 1

    def wait(self, timeout=10):
        return self.codigo


class IncrementalTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.base = Path(self.temp.name)
        self.app = self.base / "instalado"
        self.app.mkdir()
        (self.app / INICIADOR).write_bytes(b"iniciador estavel")
        self.atual = self.app / "versoes/1.0.0"
        self.atual.mkdir(parents=True)
        self.escrever(self.atual, EXECUTAVEL, b"programa antigo")
        self.escrever(self.atual, "_internal/runtime.dll", b"biblioteca estavel" * 1000)
        self.escrever(self.atual, "_internal/obsoleto.dll", b"somente na antiga")
        self.escrever(self.atual, "modelos/estilos.json", b"regras antigas")
        self.old = gerar_pacote(self.atual, self.base / "dist0", "1.0.0")
        gravar_json_atomico(self.app / "versao_inicial.json", {"versao": "1.0.0"})
        self.nova = self.base / "compilado"
        self.nova.mkdir()
        self.escrever(self.nova, EXECUTAVEL, b"programa novo")
        self.escrever(self.nova, "_internal/runtime.dll", b"biblioteca estavel" * 1000)
        self.escrever(self.nova, "modelos/estilos.json", b"regras novas")
        self.new = gerar_pacote(self.nova, self.base / "dist1", "1.1.0", self.old)
        self.assets = {n.name: n.read_bytes() for n in (self.base / "dist1").iterdir()}
        self.rel = release("1.1.0", self.assets)
        self.atualizacao = interpretar_release(self.rel, "1.0.0")
        self.urls = []

    def escrever(self, pasta, nome, conteudo):
        alvo = pasta / nome
        alvo.parent.mkdir(parents=True, exist_ok=True)
        alvo.write_bytes(conteudo)

    def urlopen(self, request, timeout=0):
        url = request.full_url
        self.urls.append(url)
        return io.BytesIO(self.assets[url.rsplit("/", 1)[-1]])

    def preparar(self):
        with patch("scripts.atualizacoes.urlopen", side_effect=self.urlopen), patch("scripts.atualizacoes.buscar_release", return_value=self.rel):
            return preparar_incremental(self.atualizacao, self.atual, self.base / "cache")

    def test_pacote_contem_somente_alterados_e_reutiliza_bibliotecas(self):
        with zipfile.ZipFile(self.base / "dist1" / nome_pacote("1.1.0")) as zipado:
            self.assertEqual(set(zipado.namelist()), {EXECUTAVEL, "modelos/estilos.json"})
        runtime = next(a for a in self.new["arquivos"] if a["caminho"] == "_internal/runtime.dll")
        self.assertEqual(runtime["pacote"], "1.0.0")

    def test_prepara_versao_completa_sem_alterar_a_atual(self):
        original = (self.atual / EXECUTAVEL).read_bytes()
        preparada = self.preparar()
        alvo = self.app / "versoes/1.1.0"
        self.assertEqual((alvo / EXECUTAVEL).read_bytes(), b"programa novo")
        self.assertEqual((self.atual / EXECUTAVEL).read_bytes(), original)
        self.assertEqual((alvo / "_internal/runtime.dll").read_bytes(), (self.atual / "_internal/runtime.dll").read_bytes())
        self.assertFalse((alvo / "_internal/obsoleto.dll").exists())
        self.assertEqual(versao_ativa(self.app), "1.0.0")
        self.assertEqual(len(self.urls), 2)
        self.assertTrue(all("Setup.exe" not in u for u in self.urls))
        self.assertEqual(preparada.manifesto_sha256, sha256(alvo / MANIFESTO))
        self.assertEqual(preparada.bytes_baixados, sum(len(x) for x in self.assets.values()))

    def test_cache_e_versao_preparada_sao_reutilizados_sem_download(self):
        self.preparar()
        self.urls.clear()
        self.assertEqual(self.preparar().bytes_baixados, 0)
        self.assertEqual(self.urls, [])

    def test_arquivo_local_corrompido_e_baixado_da_origem_correta(self):
        self.escrever(self.atual, "_internal/runtime.dll", b"corrompido")
        origem = self.base / "dist0" / nome_pacote("1.0.0")
        assets0 = {origem.name: origem.read_bytes()}
        self.assets.update(assets0)
        with patch("scripts.atualizacoes.urlopen", side_effect=self.urlopen), patch("scripts.atualizacoes.buscar_release", side_effect=lambda n: release("1.0.0", assets0) if n == "1.0.0" else self.rel):
            preparar_incremental(self.atualizacao, self.atual, self.base / "cache")
        self.assertEqual((self.app / "versoes/1.1.0/_internal/runtime.dll").read_bytes(), b"biblioteca estavel" * 1000)
        self.assertEqual((self.atual / "_internal/runtime.dll").read_bytes(), b"corrompido")

    def test_download_adulterado_e_interrompido_nao_toca_na_instalacao(self):
        self.assets[nome_pacote("1.1.0")] = b"pacote adulterado"
        with self.assertRaises(ValueError):
            self.preparar()
        self.assertFalse((self.app / "versoes/1.1.0").exists())
        self.assertFalse(list((self.app / "versoes").glob(".preparando-*")))
        self.assertEqual((self.atual / EXECUTAVEL).read_bytes(), b"programa antigo")
        self.assertEqual(versao_ativa(self.app), "1.0.0")

    def test_falha_de_rede_preserva_versao_atual(self):
        with patch("scripts.atualizacoes.urlopen", side_effect=OSError("rede indisponivel")):
            with self.assertRaises(OSError):
                preparar_incremental(self.atualizacao, self.atual, self.base / "cache")
        self.assertEqual(versao_ativa(self.app), "1.0.0")

    def test_ativa_somente_apos_confirmar_abertura(self):
        preparada = self.preparar()
        def iniciar(raiz, numero, argumentos, confirmacao=None, token=""):
            self.assertEqual(versao_ativa(raiz), "1.0.0")
            confirmar_inicio(confirmacao, token, raiz / "versoes" / numero, numero)
            return Processo()
        self.assertTrue(ativar(self.app, "1.1.0", preparada.manifesto_sha256, iniciar=iniciar))
        self.assertEqual(versao_ativa(self.app), "1.1.0")
        self.assertFalse((self.app / "transacao_atualizacao.json").exists())
        self.assertTrue(self.atual.is_dir())

    def test_falha_de_inicializacao_reabre_anterior(self):
        preparada = self.preparar()
        chamadas = []
        def iniciar(raiz, numero, argumentos, confirmacao=None, token=""):
            chamadas.append(numero)
            return Processo(1)
        self.assertFalse(ativar(self.app, "1.1.0", preparada.manifesto_sha256, iniciar=iniciar))
        self.assertEqual(chamadas, ["1.1.0", "1.0.0"])
        self.assertEqual(versao_ativa(self.app), "1.0.0")
        self.assertTrue((self.app / "erro_atualizacao.txt").is_file())

    def test_timeout_encerra_apenas_candidato_e_reabre_anterior(self):
        preparada = self.preparar()
        candidato = Processo()
        with patch("scripts.iniciar_aplicativo.time.sleep"):
            self.assertFalse(ativar(self.app, "1.1.0", preparada.manifesto_sha256,
                                   iniciar=lambda *args: candidato, timeout=0))
        self.assertTrue(candidato.encerrado)
        self.assertEqual(versao_ativa(self.app), "1.0.0")

    def test_alteracao_apos_download_impede_ativacao(self):
        preparada = self.preparar()
        (self.app / "versoes/1.1.0" / EXECUTAVEL).write_bytes(b"adulterado")
        chamadas = []
        self.assertFalse(ativar(self.app, "1.1.0", preparada.manifesto_sha256,
                               iniciar=lambda *args: chamadas.append(args[1]) or Processo(0)))
        self.assertEqual(chamadas, ["1.0.0"])
        self.assertEqual(versao_ativa(self.app), "1.0.0")

    def test_interrupcao_antes_de_confirmar_restaura_anterior(self):
        self.preparar()
        gravar_json_atomico(self.app / "transacao_atualizacao.json",
                           {"nova": "1.1.0", "anterior": "1.0.0", "token": "a" * 32})
        recuperar_interrupcao(self.app)
        self.assertEqual(versao_ativa(self.app), "1.0.0")

    def test_interrupcao_apos_confirmar_completa_ativacao(self):
        self.preparar()
        (self.app / ".confirmacoes").mkdir()
        confirmar_inicio(self.app / ".confirmacoes" / ("a" * 32 + ".json"), "a" * 32,
                         self.app / "versoes/1.1.0", "1.1.0")
        gravar_json_atomico(self.app / "transacao_atualizacao.json",
                           {"nova": "1.1.0", "anterior": "1.0.0", "token": "a" * 32})
        recuperar_interrupcao(self.app)
        self.assertEqual(versao_ativa(self.app), "1.1.0")

    def test_manifesto_rejeita_caminhos_colisoes_tamanhos_e_origens(self):
        for nome in ("../fora.exe", "/fora.exe", "C:/fora.exe", "_internal/../fora.exe", "_internal/a:stream", "_internal/CON.txt", "_internal/a ", "documentos/cliente.docx"):
            with self.subTest(nome=nome):
                dados = copy.deepcopy(self.new)
                dados["arquivos"][0]["caminho"] = nome
                with self.assertRaises(ValueError):
                    validar_manifesto(dados)
        dados = copy.deepcopy(self.new)
        dados["arquivos"].append(copy.deepcopy(dados["arquivos"][0]))
        with self.assertRaises(ValueError):
            validar_manifesto(dados)
        for campo, valor in (("tamanho", -1), ("tamanho", True), ("sha256", "invalido"), ("pacote", "99.0.0")):
            dados = copy.deepcopy(self.new)
            dados["arquivos"][0][campo] = valor
            with self.assertRaises(ValueError):
                validar_manifesto(dados)

    def test_zip_rejeita_escape_e_duplicacao_sem_escrever_fora(self):
        pacote = self.base / "inseguro.zip"
        with zipfile.ZipFile(pacote, "w") as z:
            z.writestr("../fora.exe", b"invalido")
        with self.assertRaises(ValueError):
            _extrair_necessarios(pacote, [], self.base / "preparo")
        self.assertFalse((self.base / "fora.exe").exists())

    def test_manifesto_tem_preferencia_sobre_instalador_e_exige_digest(self):
        self.assertEqual(self.atualizacao.tipo, "incremental")
        invalida = copy.deepcopy(self.rel)
        next(a for a in invalida["assets"] if a["name"] == MANIFESTO)["digest"] = None
        with self.assertRaises(ValueError):
            interpretar_release(invalida, "1.0.0")

    def test_documentos_fora_do_programa_nao_sao_alterados(self):
        documentos = self.base / "documentos"
        documentos.mkdir()
        original = documentos / "peticao.docx"
        original.write_bytes(b"documento confidencial")
        self.preparar()
        self.assertEqual(original.read_bytes(), b"documento confidencial")

    def test_modelo_e_aplicativo_podem_vir_de_versoes_distintas(self):
        terceira = self.base / "compilado2"
        terceira.mkdir()
        self.escrever(terceira, EXECUTAVEL, b"programa novo")
        self.escrever(terceira, "_internal/runtime.dll", b"biblioteca estavel" * 1000)
        self.escrever(terceira, "modelos/estilos.json", b"regras mais novas")
        gerar_pacote(terceira, self.base / "dist2", "1.2.0", self.new)
        assets2 = {n.name: n.read_bytes() for n in (self.base / "dist2").iterdir()}
        rel2 = release("1.2.0", assets2)
        self.assets.update(assets2)
        with patch("scripts.atualizacoes.urlopen", side_effect=self.urlopen), patch("scripts.atualizacoes.buscar_release", side_effect=lambda n: self.rel if n == "1.1.0" else rel2):
            preparar_incremental(interpretar_release(rel2, "1.0.0"), self.atual, self.base / "cache")
        self.assertEqual((self.app / "versoes/1.2.0" / EXECUTAVEL).read_bytes(), b"programa novo")
        self.assertEqual((self.app / "versoes/1.2.0/modelos/estilos.json").read_bytes(), b"regras mais novas")
        self.assertFalse(any('/v1.0.0/' in u for u in self.urls))
