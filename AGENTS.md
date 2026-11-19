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
- Entregar apenas o DOCX formatado, diretamente na pasta selecionada.
- Manter auditoria, manifesto e decisões de revisão em memória. Usar arquivos
  temporários para o processamento e a revisão do DOCX.
- Encaminhar casos ambíguos para revisão humana.
- Usar `modelo_oficial.docx` como fonte de verdade; alterações visuais definitivas devem ser validadas no Microsoft Word.
- Tratar `peticao_inicial` como o único tipo de peça suportado nesta versão.
- Não aplicar regras específicas de recursos, contestações, réplicas ou outras peças.
- Para entrada Markdown, respeitar estritamente `perfil_markdown` de
  `modelos/estilos.json` e interromper a conversão diante de recurso não suportado.
- Preservar o Markdown original. Revisões semânticas devem reutilizar
  exatamente o mesmo DOCX intermediário durante a sessão.

# Operação

1. Leia `modelos/estilos.json`, incluindo `catalogo_semantico` e
   `formatacao_estilos` de cada estilo.
2. Execute o pipeline sobre uma cópia/entrada, nunca sobre o arquivo de saída.
3. Confira as pendências e confirme as funções dos parágrafos na tela de revisão.
4. Confira a validação estrutural antes de revisar o DOCX final.
5. Renderize e revise todas as páginas do DOCX final antes de aprová-lo.
