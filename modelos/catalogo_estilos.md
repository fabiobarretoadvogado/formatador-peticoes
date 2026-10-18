# Catálogo semântico dos estilos

Este arquivo explica a função de cada estilo para revisão humana. A versão
estruturada, usada pelo pipeline e pela classificação semântica, está em
`estilos.json`, no campo `catalogo_semantico`. O `modelo_oficial.docx` define a
aparência visual.

## Formatos de entrada

O pipeline aceita DOCX nativo e Markdown UTF-8. O Markdown é convertido uma
única vez em DOCX intermediário estrutural antes da classificação. O perfil
vigente está em `estilos.json`, na chave `perfil_markdown`, e cobre `##`/`###`,
parágrafos, citações simples, listas de um nível, tabelas GFM simples,
negrito/itálico e imagens locais relativas isoladas em bloco. Recursos fora
desse perfil interrompem a conversão em vez de serem descartados.

| Estilo | Aplicar quando | Formatação no modelo oficial | Não confundir com |
|---|---|---|---|
| `FB - Endereçamento` | Destinatário formal: juízo, tribunal ou autoridade. | Segoe UI Variable Text 12 pt, negrito, esquerda, 12 pt antes/depois, entrelinhas 1,25, manter com o próximo. | Identificação do processo ou da ação. |
| `FB - Identificação do Processo` | Número ou identificação formal dos autos, normalmente no bloco inicial. | Segoe UI Variable Text 12 pt, negrito, esquerda, 12/12 pt, entrelinhas 1,25, sem recuo. | Menções ao processo dentro da argumentação. |
| `FB - Resumo` | Ementa ou síntese inicial, depois do processo e antes da qualificação. | Segoe UI Variable Text 10 pt, justificado, recuo esquerdo de 9 cm, 12/12 pt, entrelinhas 1,25; linha inferior CCB595 apenas no último parágrafo. | Corpo argumentativo comum. |
| `FB - Qualificação das Partes` | Apresentação de autor, réu, inventariante, espólio ou representante. | Segoe UI Variable Text 12 pt, regular, justificado, primeira linha de 1 cm, 12/12 pt, entrelinhas 1,25; negrito restrito a nomes de pessoas, entidades, ações e fases processuais. | Fundamentação ou narrativa dos fatos. |
| `FB - Identificação da Ação` | Nome da ação, do rito ou do procedimento, como `RITO: ARROLAMENTO COMUM`. | Segoe UI Variable Text 12 pt, negrito, centralizado, 12/12 pt, entrelinhas 1,25, manter com o próximo. | Título de seção argumentativa. |
| `FB - Título Principal` | Seção de primeiro nível, especialmente quando a origem usa Heading 1/Título 1. | Nível 1 da lista multinível decimal (`1`, `2`, `3`); número e título em Segoe UI Variable Text 12 pt, negrito, esquerda, recuo esquerdo de 0,5 cm e primeira linha de -0,5 cm, 36/12 pt, entrelinhas 1,25, manter com o próximo. | Nome da ação ou título secundário. |
| `FB - Título Secundário` | Subseção de segundo nível, especialmente quando a origem usa Heading 2/Título 2. | Nível 2 da mesma lista (`1.1`, `1.2`, `2.1`); número e título em Segoe UI Variable Text 12 pt, negrito, esquerda, recuo esquerdo de 1 cm e primeira linha de -0,5 cm, 36/12 pt, entrelinhas 1,25, manter com o próximo. | Destaque em negrito dentro do corpo. |
| `FB - Corpo` | Prosa jurídica, fatos, fundamentos, transições e frase introdutória dos pedidos. | Segoe UI Variable Text 12 pt, regular, justificado, primeira linha de 1 cm, 12/12 pt, entrelinhas 1,25. | Item de pedido, citação, assinatura ou título. |
| `FB - Corpo com Recuo` | Bloco informativo destacado, como dados bancários, que não é citação, lista ou pedido. | Segoe UI Variable Text 12 pt, regular, justificado, recuo esquerdo de 1,25 cm e primeira linha de 1 cm, 12/12 pt, entrelinhas 1,25. | Citação literal ou enumeração. |
| `FB - Tabela` | Parágrafo dentro de uma célula de tabela. | Segoe UI Variable Text 12 pt, regular, esquerda, sem recuos, 0/0 pt, entrelinhas simples. | Texto corrido com vários campos. |
| `FB - Descrição` | Legenda ou descrição textual imediatamente abaixo de uma imagem ou documento reproduzido. | Segoe UI Variable Text 8 pt, regular, esquerda, recuo esquerdo e direito de 2 cm, 6/6 pt, entrelinhas simples. | Texto corrido, citação ou assinatura. |
| `FB - Citação` | Transcrição literal de lei, decisão, doutrina, documento ou fala de terceiro. | Segoe UI Variable Text 12 pt, regular, justificado, recuo esquerdo de 3 cm, sem recuo de primeira linha, 12/12 pt, entrelinhas 1,25. | Prosa própria apenas recuada. |
| `FB - Lista` | Enumeração de requisitos, documentos, fatos ou alternativas fora dos pedidos finais. | Segoe UI Variable Text 12 pt, regular, justificado, recuo esquerdo de 3 cm, sem recuo de primeira linha, 12/12 pt, entrelinhas 1,25. | Itens numerados do bloco final de pedidos. |
| `FB - Pedido` | Item numerado que formula uma providência no bloco `DOS PEDIDOS`. | Segoe UI Variable Text 12 pt, regular, justificado, recuo esquerdo de 3 cm, sem recuo de primeira linha, 12/12 pt, entrelinhas 1,25. | Frase introdutória dos pedidos e listas gerais. |
| `FB - Fechamento` | Fórmula de encerramento, local e data. | Segoe UI Variable Text 12 pt, regular, centralizado, 0/0 pt, entrelinhas simples; sem espaçamento antes ou depois. | Nome e identificação profissional do signatário. |
| `FB - Assinatura` | Nome do signatário e identificação profissional, como OAB; imagem isolada de assinatura. | Segoe UI Variable Text 12 pt, negrito, centralizado, 0/0 pt, entrelinhas simples; imagem de assinatura centralizada, sem espaçamento antes ou depois. | Fechamento, local ou data. |
| `FB - Nota de Rodapé` | Texto da nota vinculada a uma chamada no corpo. | Segoe UI Variable Text 10 pt, regular, justificado, sem recuos, 0/0 pt, entrelinhas simples. | Conteúdo fixo do rodapé da página. |
| `FB - Rodapé` | Conteúdo da área de rodapé, como identificação ou paginação. | Segoe UI Variable Text 10 pt, regular, centralizado, sem recuos, 0/0 pt, entrelinhas simples. | Nota de rodapé vinculada a uma chamada. |

## Regras de contexto

Nos estilos regulares de texto, o conteúdo permanece sem negrito direto. As exceções ficam restritas aos papéis estruturais descritos no catálogo, como qualificação das partes, títulos e assinatura.

- O escopo de desenvolvimento e validação atual é exclusivamente petição
  inicial. Outras peças não são bloqueadas, mas usam o mesmo perfil de regras.
- Estruturas próprias de recursos, contestações, réplicas, manifestações e
  outras peças ainda não possuem regras nesta versão.
- Marcações `##`, `###`, `>` e tabelas delimitadas por `|`, quando presentes no
  texto de origem, são convertidas em estrutura Word e não permanecem visíveis.
- Tabelas de uma célula que contenham nome e OAB são tratadas como bloco de
  assinatura, centralizado e sem espaçamento.

- A informação estrutural de origem prevalece: tabela, nota, rodapé, Quote,
  Heading 1 e Heading 2 não devem ser reclassificados apenas por palavras-chave.
- `FB - Corpo` é usado para “Diante do exposto, requer-se...” e expressões
  equivalentes; os itens numerados seguintes usam `FB - Pedido`.
- `FB - Lista` só vira `FB - Pedido` quando estiver no bloco final de pedidos e
  formular uma providência ao juízo.
- Recuo, negrito, caixa alta e numeração são sinais auxiliares, não provas
  isoladas da função semântica.
