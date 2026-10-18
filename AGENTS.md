# Objetivo

Padronizar petições iniciais conforme os estilos oficiais do escritório.

# Regras obrigatórias

- Nunca sobrescrever o documento original.
- Nunca criar estilos novos sem autorização.
- Aplicar somente estilos listados em `modelos/estilos.json`.
- Não alterar o conteúdo jurídico durante a formatação.
- Não corrigir fatos, datas, valores ou pedidos sem solicitação expressa.
- Não excluir parágrafos nem reorganizar blocos sem autorização.
- Classificar semanticamente antes de aplicar estilos.
- Utilizar Python para alterações mecânicas no DOCX.
- Produzir relatório completo das alterações.
- Encaminhar casos ambíguos para revisão humana.
- Usar `modelo_oficial.docx` como fonte de verdade; alterações visuais definitivas devem ser validadas no Microsoft Word.
- Gerar o PDF pelo Microsoft Word quando a fidelidade visual for relevante.
- Tratar `peticao_inicial` como o único tipo de peça suportado nesta versão.
- Não aplicar regras específicas de recursos, contestações, réplicas ou outras peças.
- Para entrada Markdown, respeitar estritamente `perfil_markdown` de
  `modelos/estilos.json` e interromper a conversão diante de recurso não suportado.
- Preservar o Markdown original, o manifesto de entrada e o DOCX intermediário;
  revisões semânticas devem reutilizar exatamente o mesmo intermediário.

# Operação

1. Leia `modelos/estilos.json`, incluindo `catalogo_semantico` e
   `formatacao_estilos` de cada estilo.
2. Execute o pipeline sobre uma cópia/entrada, nunca sobre o arquivo de saída.
3. Analise `trabalho/solicitacoes_semanticas.json` quando houver ambiguidades.
4. Grave as decisões em `trabalho/respostas_semanticas.json` e reexecute com `--semantica`.
5. Confira `trabalho/auditoria.json` e `saida/relatorio_formatacao.md`.
6. Renderize e revise todas as páginas do DOCX final antes de aprová-lo.
7. Em entradas Markdown, confira `trabalho/manifesto_entrada.json` e a auditoria
   estrutural da conversão antes de revisar o DOCX final.
