import unittest

from scripts.classificar_regras import classificar, classificar_um


class FechamentoTests(unittest.TestCase):
    def paragrafo(self, texto, **extras):
        return {"id": 1, "texto": texto, "posicao_relativa": 0.9, **extras}

    def test_local_e_data_do_protocolo_sao_fechamento_sem_revisao(self):
        textos = (
            "Aracaju/SE, data do protocolo.",
            "ARACAJU/SE, DATA DO PROTOCOLO.",
            "São Cristóvão/SE, data do protocolo eletrônico.",
            "Nossa Senhora do Socorro/SE, data da assinatura eletrônica.",
            "Brasília - DF, na data da assinatura digital.",
            "São João d'Aliança/GO, data do protocolo.",
            "Aracaju (SE), 17 de outubro de 2026.",
            "Aracaju, 17/10/2026.",
            "Aracaju/SE, 17.10.2026.",
            "Aracaju/SE, 17-10-2026.",
        )
        for texto in textos:
            with self.subTest(texto=texto):
                resultado = classificar_um(self.paragrafo(texto), seguinte="FÁBIO AUGUSTO MENDONÇA BARRETO")
                self.assertEqual(resultado["estilo_destino"], "FB - Fechamento")
                self.assertFalse(resultado["necessita_analise_semantica"])
                self.assertEqual(resultado["texto"], texto)

    def test_fechamento_prevalece_sobre_negrito_titulo_e_assinatura(self):
        resultado = classificar_um(
            self.paragrafo("Aracaju/SE, data do protocolo.", negrito=True, estilo_original="Heading 2", alinhamento="CENTER (1)"),
            seguinte="OAB/SE nº 12.563",
        )
        self.assertEqual(resultado["estilo_destino"], "FB - Fechamento")

    def test_referencias_a_datas_nos_fatos_continuam_corpo(self):
        for texto in (
            "O contrato foi assinado em 17 de outubro de 2026.",
            "A data do protocolo comprova a tempestividade do pedido.",
            "O contrato foi assinado em Aracaju, 17 de outubro de 2026.",
            "O prazo começa na data da assinatura eletrônica.",
        ):
            with self.subTest(texto=texto):
                resultado = classificar_um(self.paragrafo(texto))
                self.assertEqual(resultado["estilo_destino"], "FB - Corpo")

    def test_data_isolada_exige_contexto_de_encerramento(self):
        p = self.paragrafo("Data do protocolo.", posicao_relativa=0.2)
        self.assertEqual(classificar_um(p)["estilo_destino"], "FB - Corpo")
        self.assertEqual(classificar_um(p, anterior="Termos em que, pede deferimento.")["estilo_destino"], "FB - Fechamento")
        self.assertEqual(classificar_um(self.paragrafo("17/10/2026."))["estilo_destino"], "FB - Fechamento")

    def test_local_data_fecha_bloco_de_pedidos(self):
        estrutura = {"documento": "amostra.docx", "paragrafos": [
            self.paragrafo("Diante do exposto, requer:", posicao_relativa=0.6),
            self.paragrafo("I. A procedência dos pedidos;", id=2, posicao_relativa=0.7),
            self.paragrafo("Aracaju/SE, data do protocolo.", id=3, posicao_relativa=0.85),
            self.paragrafo("FÁBIO AUGUSTO MENDONÇA BARRETO", id=4, posicao_relativa=0.9),
            self.paragrafo("OAB/SE nº 12.563", id=5, posicao_relativa=0.95),
        ]}
        resultado = classificar(estrutura)
        self.assertEqual([item["estilo_destino"] for item in resultado["classificacoes"]], [
            "FB - Corpo", "FB - Pedido", "FB - Fechamento", "FB - Assinatura", "FB - Assinatura",
        ])
