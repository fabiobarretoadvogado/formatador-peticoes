import hashlib
import io
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from app_version import INSTALLER_NAME, REPOSITORY
from scripts.atualizacoes import baixar_instalador, interpretar_release


class AtualizacoesTests(unittest.TestCase):
    def release(self, conteudo=b"instalador de teste"):
        return {
            "tag_name": "v1.1.0",
            "assets": [{
                "name": INSTALLER_NAME,
                "browser_download_url": f"https://github.com/{REPOSITORY}/releases/download/v1.1.0/{INSTALLER_NAME}",
                "digest": "sha256:" + hashlib.sha256(conteudo).hexdigest(),
                "size": len(conteudo),
            }],
        }

    def test_compara_versoes_numericamente_sem_downgrade(self):
        release = self.release()
        self.assertIsNone(interpretar_release(release, "1.2.0"))
        self.assertIsNone(interpretar_release(release, "1.1.0"))
        release["tag_name"] = "v1.10.0"
        release["assets"][0]["browser_download_url"] = release["assets"][0]["browser_download_url"].replace("v1.1.0", "v1.10.0")
        self.assertEqual(interpretar_release(release, "1.2.0").versao, "1.10.0")

    def test_ignora_rascunhos_e_pre_lancamentos(self):
        for campo in ("draft", "prerelease"):
            release = self.release()
            release[campo] = True
            self.assertIsNone(interpretar_release(release, "1.0.0"))

    def test_rejeita_origem_estranha_e_digest_ausente(self):
        for alteracao in (
            {"browser_download_url": "https://outro.example/instalador.exe"},
            {"browser_download_url": f"https://github.com/outro/projeto/releases/download/v1.1.0/{INSTALLER_NAME}"},
            {"digest": None},
            {"size": 0},
        ):
            with self.subTest(alteracao=alteracao):
                release = self.release()
                release["assets"][0].update(alteracao)
                with self.assertRaises(ValueError):
                    interpretar_release(release, "1.0.0")

    def test_download_confere_hash_e_remove_parcial(self):
        conteudo = b"instalador de teste"
        atualizacao = interpretar_release(self.release(conteudo), "1.0.0")
        with tempfile.TemporaryDirectory() as temp:
            with patch("scripts.atualizacoes.urlopen", return_value=io.BytesIO(conteudo)):
                arquivo = baixar_instalador(atualizacao, Path(temp))
            self.assertEqual(arquivo.read_bytes(), conteudo)
            arquivo.unlink()
            with patch("scripts.atualizacoes.urlopen", return_value=io.BytesIO(b"arquivo adulterado")):
                with self.assertRaises(ValueError):
                    baixar_instalador(atualizacao, Path(temp))
            self.assertFalse(arquivo.exists())
            self.assertFalse(list(Path(temp).glob("*.download")))


if __name__ == "__main__":
    unittest.main()
