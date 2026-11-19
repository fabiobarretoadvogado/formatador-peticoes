from io import BytesIO
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from PIL import Image

from scripts.assinatura_local import AssinaturaLocal, ESTILO_ASSINATURA
from scripts.pipeline import executar


MODELOS = Path(__file__).resolve().parents[1] / "modelos"
NOME = "Ana Fictícia da Silva"
OAB = "Advogada – OAB/SE nº 99.999"


def _assinatura():
    png = BytesIO()
    Image.new("RGBA", (516, 150)).save(png, "PNG")
    return AssinaturaLocal(NOME, frozenset({"SE99999"}), png.getvalue(), 5.16)


def test_markdown_insere_no_ultimo_bloco_sem_alterar_o_texto(tmp_path):
    origem = tmp_path / "peticao.md"
    blocos = ["Aracaju/SE, 18 de novembro de 2026.", NOME, OAB, "## Fundamentos", NOME, OAB]
    markdown = "\n\n".join(blocos)
    origem.write_text(markdown, encoding="utf-8")

    assinatura = _assinatura()
    resultado = executar(origem, work_root=tmp_path / "work", modelos_raiz=MODELOS, assinatura=assinatura)
    word = Document(resultado["docx"])
    imagem, nome, oab = word.paragraphs[-3:]
    assert [p.text for p in word.paragraphs if p.text.strip()] == [p.removeprefix("## ") for p in blocos]
    assert [p.style.name for p in (imagem, nome, oab)] == [ESTILO_ASSINATURA] * 3
    assert imagem.paragraph_format.keep_with_next and nome.paragraph_format.keep_with_next
    assert len(word.inline_shapes) == 1
    assert abs(word.inline_shapes[0].width.cm - 5.16) < 0.01
    id_imagem = imagem._p.xpath(".//a:blip")[0].get(qn("r:embed"))
    assert word.part.related_parts[id_imagem].blob == assinatura.imagem
    assert origem.read_text(encoding="utf-8") == markdown


def test_word_preserva_imagem_existente_e_exige_registro_correto(tmp_path):
    assinatura = _assinatura()
    for numero, registro in enumerate((OAB, "OAB/SE nº 88.888")):
        doc = Document()
        doc.add_paragraph("Aracaju/SE, 18 de novembro de 2026.")
        if numero == 0:
            doc.add_paragraph().add_run().add_picture(BytesIO(assinatura.imagem))
        doc.add_paragraph(NOME)
        doc.add_paragraph(registro)
        origem = tmp_path / f"peticao{numero}.docx"
        doc.save(origem)
        resultado = executar(
            origem, work_root=tmp_path / f"work{numero}", modelos_raiz=MODELOS, assinatura=assinatura,
        )
        final = Document(resultado["docx"])
        assert len(final.inline_shapes) == (1 if numero == 0 else 0)
        assert not resultado["auditoria"]["aplicacao"]["assinatura_local_inserida"]


def test_reconhece_e_mantem_assinatura_existente_sem_perfil_local(tmp_path):
    doc = Document()
    doc.add_paragraph("Aracaju/SE, 18 de novembro de 2026.")
    doc.add_paragraph().add_run().add_picture(BytesIO(_assinatura().imagem))
    doc.add_paragraph(NOME)
    doc.add_paragraph()
    doc.add_paragraph(OAB)
    doc.add_paragraph("ROL DE TESTEMUNHAS", style="Heading 1")
    origem = tmp_path / "assinada.docx"
    doc.save(origem)
    resultado = executar(origem, work_root=tmp_path / "work", modelos_raiz=MODELOS)
    final = Document(resultado["docx"])
    imagem, nome, vazio, oab = final.paragraphs[1:5]
    assert [p.style.name for p in (imagem, nome, oab)] == [ESTILO_ASSINATURA] * 3
    assert all(p.paragraph_format.keep_with_next for p in (imagem, nome, vazio))
    assert oab.paragraph_format.keep_together
    assert not oab.paragraph_format.keep_with_next
    assert len(final.inline_shapes) == 1
    assert final.paragraphs[-1].style.name == "FB - Título Principal"


def test_insere_assinatura_local_em_peca_com_tabela(tmp_path):
    doc = Document()
    tabela = doc.add_table(rows=1, cols=1)
    tabela.cell(0, 0).text = NOME + "\n" + OAB
    origem = tmp_path / "tabela.docx"
    doc.save(origem)
    resultado = executar(origem, work_root=tmp_path / "work", modelos_raiz=MODELOS, assinatura=_assinatura())
    final = Document(resultado["docx"])
    imagem, texto = final.tables[0].cell(0, 0).paragraphs
    assert len(final.inline_shapes) == 1
    assert imagem.style.name == texto.style.name == ESTILO_ASSINATURA
    assert imagem.paragraph_format.keep_with_next and texto.paragraph_format.keep_together
