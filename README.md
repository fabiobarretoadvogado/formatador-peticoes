# Formatador de petições iniciais

Pipeline local e auditável para aplicar exclusivamente os estilos oficiais do escritório a petições iniciais recebidas em Word ou Markdown, sem sobrescrever o original.

## Instalar no Windows

Baixe `FormatadorPeticoes-Setup.exe` na seção **Releases** do repositório e execute
o instalador. O aplicativo instala para o usuário atual, cria um atalho no menu
Iniciar e pode ser removido pelas configurações do Windows. Não é necessário
instalar Python. O Microsoft Word é necessário para exportar PDF.

O instalador e o pacote portátil incluem apenas o programa, bibliotecas, o modelo
oficial vazio e as regras de estilos. Documentos de clientes, execuções anteriores,
resultados e arquivos temporários não fazem parte da distribuição.

## Atualizações pelo GitHub

No aplicativo, clique em **Verificar atualizações** e confirme **Atualizar e reiniciar**.
O programa baixa os pacotes que contêm os arquivos alterados, confere tamanho e
SHA-256 publicados pelo GitHub e prepara a nova versão em uma pasta separada.
Bibliotecas que continuam iguais são reaproveitadas da instalação atual, sem
novo download. O programa fecha e reabre automaticamente, sem assistente de instalação.

A versão ativa só muda depois de o aplicativo novo confirmar que abriu a interface.
Se o download, a verificação ou a abertura falhar, a versão anterior permanece
disponível. Interrupções durante a troca são recuperadas na próxima abertura.
Documentos, resultados e execuções ficam fora das pastas de versões do programa.
A instalação inicial e a versão anterior são mantidas para recuperação; versões
intermediárias antigas são removidas somente se contiverem apenas arquivos conhecidos do programa.

A migração das versões até 1.0.3 exige uma última instalação completa. A partir
da 1.0.4, as atualizações normais são feitas pelo próprio programa. O computador
precisa de acesso à internet e as releases e seus pacotes devem permanecer públicos.

Para publicar uma alteração:

1. Altere o código e incremente a versão em `app_version.py` e `pyproject.toml`.
2. Envie as alterações para `main`.
3. O GitHub Actions executa os testes, compila e verifica a interface, o
   processamento DOCX/Markdown e a ativação de uma versão. Publica o instalador,
   o ZIP portátil, o manifesto e o pacote incremental em uma nova release.

O manifesto associa cada arquivo ao pacote da release em que ele foi alterado.
Isso permite atualizar diretamente de versões antigas sem aplicar todas as
versões intermediárias. Os pacotes antigos referenciados não devem ser apagados.

Uma alteração com a mesma versão executa os testes, mas não substitui uma release
existente. Falhas de testes ou compilação interrompem a publicação. Pull requests
executam apenas a validação, sem publicar. O canal de atualização é definido em
`app_version.py`; ao mudar de repositório, ajuste `REPOSITORY` e os links do instalador.

Para compilar localmente, use Python 3.12, instale Inno Setup 6 e execute:

```powershell
python -m pip install -e ".[build]" pytest
python -m pytest -q
python scripts/build_instalador.py
```

O resultado é `dist/FormatadorPeticoes-Setup.exe`. O processo também gera um ZIP
portátil limpo e versionado. O instalador não possui assinatura digital de editor.

Esta versão suporta somente petições iniciais. Recursos, contestações,
réplicas, manifestações e outras peças serão adicionadas apenas depois que
o fluxo de petição inicial estiver validado. O sistema não bloqueia outras peças:
ao selecionar um arquivo, o usuário decide usar o perfil atual, ciente de que as
regras foram desenhadas e validadas somente para petições iniciais.

## Início rápido

1. Abra `modelos/modelo_oficial.docx` no Microsoft Word e ajuste/valide os estilos `FB - ...` conforme o manual do escritório.
2. Coloque a petição inicial em `entrada/`.
3. Execute, a partir desta pasta:

```powershell
python scripts/pipeline.py entrada/peticao_original.docx
```

Uma petição Markdown UTF-8 usa o mesmo comando:

```powershell
python scripts/pipeline.py entrada/peticao_original.md
```

Para testar sem exportar o PDF:

```powershell
python scripts/pipeline.py entrada/peticao_original.docx --skip-pdf
```

O resultado fica em:

- `saida/peticao_formatada.docx`
- `saida/peticao_formatada.pdf`
- `saida/relatorio_formatacao.md`

Os JSONs de trabalho e a cópia de segurança ficam em `trabalho/`.
Para entradas Markdown, essa pasta também recebe `entrada_intermediaria.docx` e
`manifesto_entrada.json`, com os hashes da origem e do intermediário, versão do
parser, imagens incorporadas e resultado da auditoria estrutural da conversão.

O pipeline exige `modelos/modelo_oficial.docx` e `modelos/estilos.json` na pasta
da execução. Se qualquer um deles estiver ausente, o processamento é interrompido
com uma mensagem de recuperação. O sistema nunca cria nem substitui silenciosamente
o modelo oficial ou suas regras.

## Aplicativo portátil para a equipe

O launcher recomendado para distribuição é `FormatadorPeticoes.exe`. Ele inclui
o Python e as bibliotecas necessárias, aceita arquivos arrastados sobre o
executável e também oferece uma janela para selecionar um ou mais DOCX ou Markdown.

Na versão portátil:

- a equipe não precisa instalar Python;
- a interface conduz o trabalho em quatro etapas: selecionar, formatar, revisar
  e abrir os resultados;
- trechos de classificação duvidosa são mostrados com o parágrafo anterior e
  o seguinte, para confirmação sem editar arquivos técnicos;
- cada execução intermediária fica em `%LOCALAPPDATA%\FormatadorPeticoes`;
- o DOCX, o PDF, o relatório, a auditoria e eventuais pendências são copiados
  para a pasta de resultados; entradas Markdown incluem ainda o manifesto e o
  DOCX intermediário utilizados;
- o modelo e as regras ficam em `versoes/<versão>/modelos`, junto da versão correspondente;
- o Word continua sendo necessário somente para a geração do PDF.

Na tela de revisão, a sugestão automática, a confiança e a descrição da função
do estilo aparecem juntas. A conclusão só é liberada depois que todas as
pendências forem confirmadas. A aplicação rejeita respostas antigas se o texto
do parágrafo tiver mudado. Ao final, a interface permite abrir diretamente o
DOCX, o PDF, o relatório técnico ou a pasta do documento.

O pacote oficial inclui o runtime Tcl/Tk e, ao abrir o EXE sem argumentos, deve
exibir sempre a interface completa. O seletor nativo do Windows permanece apenas
como contingência interna e não é aceito pela validação do pacote de distribuição.
Arquivos arrastados sobre o EXE também passam pela interface, com processamento
automático; não é necessário abrir ou acompanhar um CMD.

Para gerar o pacote de distribuição em uma máquina de desenvolvimento:

```powershell
python -m pip install -e ".[build]"
python scripts/build_portatil.py
```

O arquivo pronto para a equipe será criado em `dist/`, com a versão no nome,
por exemplo `FormatadorPeticoes-portatil-1.0.0.zip`. A pasta deve ser extraída por completo;
não se deve copiar apenas o EXE. O build executa testes de DOCX, Markdown e da
janela completa e é interrompido se somente o seletor nativo estiver disponível.

## Catálogo semântico dos estilos

O arquivo `modelos/estilos.json` possui o `catalogo_semantico`, com a finalidade
de cada estilo, e `formatacao_estilos`, com fonte, tamanho, alinhamento, recuos,
espaçamentos e observações visuais. Esses campos orientam a classificação
semântica; o `modelos/modelo_oficial.docx` continua sendo a fonte de verdade da
aparência no Word.

O catálogo semântico para consulta está em `modelos/estilos.json`.

Algumas distinções importantes:

- `FB - Corpo` é usado na frase introdutória dos pedidos.
- `FB - Pedido` é reservado aos itens numerados do bloco final de pedidos.
- `FB - Lista` é usado para enumerações que não formulam pedidos finais.
- `FB - Identificação da Ação` nomeia a ação, o rito ou o procedimento; não é
  um título de seção argumentativa.
- `FB - Citação` exige transcrição ou função de fonte; o recuo sozinho não é
  suficiente.
- `FB - Descrição` é reservado ao texto iniciado por “Descrição:” que aparece
  imediatamente abaixo de uma imagem ou documento reproduzido.

## Escopo atual: petição inicial

O pipeline usa um único perfil estrutural: endereçamento ao juízo de origem,
identificação do processo quando existente, qualificação das partes,
identificação da ação, fatos, fundamentos, pedidos, fechamento e assinatura.
Não existem nesta versão regras especiais para segundo endereçamento,
interposição, recorrente, recorrido ou razões dirigidas a tribunal.

Quando o DOCX contiver marcações textuais de edição, `##`/`###` são convertidos
em títulos, `>` em citação, `**` é removido em favor das regras de negrito dos
estilos e quadros delimitados por `|` são convertidos em tabelas Word.

## Entrada Markdown

O Markdown é uma modalidade de entrada. O resultado oficial continua sendo o
DOCX formatado pelo fluxo existente. Antes da extração, o sistema usa
`markdown-it-py` para criar uma árvore sintática e gerar um DOCX intermediário
baseado no `modelo_oficial.docx`, mas com estilos Word estruturais neutros. A
classificação e a aplicação dos estilos `FB - ...` continuam ocorrendo somente
no pipeline.

Perfil suportado nesta versão:

- arquivos UTF-8 ou UTF-8 com BOM;
- `##` como título principal e `###` como título secundário;
- parágrafos separados por linha vazia, com quebras Markdown suaves e forçadas;
- citações `>` simples, sem blocos aninhados;
- listas numéricas, alfabéticas (`a)`/`a.`) e com marcadores, em um nível;
- tabelas simples no formato GFM, sem células mescladas;
- `**negrito**` e `*itálico*` no DOCX intermediário; a saída final segue as
  regras de negrito de `modelos/estilos.json`;
- imagens locais isoladas em bloco, com caminho relativo à pasta do Markdown e
  redução automática à largura útil da página.

São rejeitados com mensagem explícita: títulos `#` ou `####` e superiores,
listas aninhadas, links, HTML, blocos de código, notas de rodapé, imagens remotas
ou fora da pasta da origem e imagens misturadas com texto no mesmo parágrafo.
Nenhum desses recursos é descartado silenciosamente.

## Resumo na primeira página

Use `FB - Resumo` para a ementa ou síntese inicial. Todos os estilos usam Segoe UI Variable Text; o resumo usa tamanho 10 pt e recuo esquerdo de 9 cm. Ele deve vir após a identificação do processo e antes da qualificação das partes. Somente o último parágrafo do resumo recebe a linha inferior #CCB595, sem linhas sob o processo ou entre os parágrafos internos do bloco.

Os estilos `FB - Citação`, `FB - Lista` e `FB - Pedido` usam recuo esquerdo de 3 cm. Lista e pedido não acrescentam recuo especial de primeira linha.

Nos estilos regulares (`FB - Corpo`, `FB - Corpo com Recuo`, `FB - Citação`, `FB - Lista` e `FB - Pedido`), o texto permanece regular; o pipeline remove negrito direto, preservando apenas as exceções estruturais previstas no catálogo.

Tópicos e subtópicos recebem 36 pt de espaço anterior: os 12 pt normais mais 24 pt equivalentes a uma linha de parágrafo separando-os do bloco anterior. A identificação da ação permanece com 12 pt antes e 12 pt depois.

O endereçamento e a identificação do processo ficam sempre em negrito e caixa alta, inclusive quando a origem contém trechos em minúsculas ou com negrito desativado. A caixa alta usa o efeito nativo do Word, preservando o texto e os campos originais.

Os títulos usam uma única lista multinível nativa do Word: `FB - Título Principal` gera `1`, `2`, `3`; `FB - Título Secundário` gera `1.1`, `1.2`, `2.1`. Número e texto ficam em 12 pt e em negrito. Prefixos digitados, como `2.` ou `II.`, são retirados ao vincular o título à lista, evitando combinações duplicadas como `II. 2.`.

Os subtítulos mantêm recuo esquerdo de 1 cm e especial **Primeira linha: 1 cm**.
O estilo, o parágrafo e a numeração usam a mesma definição, sem deslocamento.
Títulos e subtítulos ficam integralmente em negrito, inclusive números e trechos
que tenham vindo com negrito desativado na origem.

No bloco final de pedidos, a frase introdutória (por exemplo, “Diante do exposto, requer-se...”) usa `FB - Corpo`. Somente os itens numerados em algarismos romanos (`I`, `II`, `III`...) usam `FB - Pedido` e o recuo de 3 cm.

Os estilos estruturais de origem `Heading 1`/`Título 1` e `Heading 2`/`Título 2` prevalecem sobre heurísticas textuais. Assim, a numeração principal permanece sequencial (`1`, `2`, `3`...) e a secundária se comunica com ela (`3.1`, `3.2`...), sempre na fonte oficial.

## Qualificação, ação e notas de rodapé

- `FB - Qualificação das Partes`: texto inicial que identifica autor, réu e representação; deixa em negrito apenas nomes de pessoas, entidades, ações e fases processuais.
- `FB - Identificação da Ação`: denominação centralizada da demanda, entre a qualificação da parte autora e a apresentação da parte ré.
- `FB - Assinatura`: centraliza também a imagem isolada que antecede o bloco de assinatura.
- O bloco final de data, assinatura, nome e OAB usa `FB - Fechamento`/`FB - Assinatura` sem espaçamento antes ou depois dos parágrafos.
- `FB - Nota de Rodapé`: aplica a fonte oficial em 10 pt às notas; no Word, use esse estilo no texto da nota, não no marcador numérico.

## Revisão semântica assistida

O classificador por regras envia todo item abaixo de 95% de confiança para `trabalho/solicitacoes_semanticas.json`. Uma resposta aceita tem este formato:

```json
{
  "classificacoes": [
    {
      "id": 12,
      "hash": "hash SHA-256 informado na solicitação",
      "estilo_destino": "FB - Título Secundário",
      "confianca": 0.98,
      "justificativa": "Introduz nova subseção argumentativa"
    }
  ]
}
```

Reexecute:

```powershell
python scripts/pipeline.py entrada/peticao_original.docx --semantica trabalho/respostas_semanticas.json
```

Na interface, a revisão de uma origem Markdown reutiliza exatamente o mesmo
DOCX intermediário e o mesmo manifesto, preservando os hashes da classificação.

Somente estilos autorizados são aceitos. Respostas semânticas com confiança de até 0,90 não substituem a regra. O `hash` deve ser exatamente o mesmo recebido em `solicitacoes_semanticas.json`; respostas antigas são ignoradas caso o texto tenha mudado. No aplicativo, essa revisão é feita pela tela, sem editar JSON manualmente.

## Limites da versão inicial

- A paginação registrada na extração é aproximada até o Word repaginar o arquivo.
- Em entradas DOCX, notas de rodapé, imagens, tabelas, campos e seções são preservados porque o original é editado por cópia; a classificação automática cobre parágrafos do corpo principal.
- Cabeçalhos e rodapés vazios recebem o papel timbrado e a paginação do modelo oficial. Conteúdo já existente nessas partes é preservado. A regra também cobre a primeira página, páginas pares e seções vinculadas.
- Todas as seções recebem as margens, medianiz e distâncias de cabeçalho/rodapé do modelo oficial. No modelo incluído: superior 3 cm, inferior 2,75 cm e laterais 2 cm. O tamanho da página, a orientação e as quebras de seção são preservados.
- Em entradas Markdown, aplicam-se estritamente os recursos declarados no perfil acima; notas de rodapé, HTML e posicionamento flutuante não são inferidos.
- O modelo incluído é um ponto de partida técnico. A fonte de verdade deve ser validada e, se necessário, substituída pelo modelo aprovado no Word.
- O PDF requer Microsoft Word e `pywin32`. Falhas de exportação são registradas sem apagar o DOCX e o relatório.
- A revisão visual final continua obrigatória antes do uso jurídico.

## Testes

```powershell
python -m unittest discover -s testes -v
```
