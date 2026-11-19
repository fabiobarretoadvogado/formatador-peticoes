# FormatadorPeticoes no ChatGPT

O servidor usa o mesmo modelo oficial, catálogo de estilos e pipeline do aplicativo.
Formata petições iniciais já redigidas, preserva o arquivo de origem e salva
somente o DOCX formatado na pasta de destino. Quando houver ambiguidades, o
painel indica que o Word requer revisão.

## Usar como plugin local no Codex

O plugin **FormatadorPeticoes** chama o servidor diretamente no computador. A
ferramenta recebe o caminho da petição em DOCX ou Markdown e a pasta de destino.
Se você omitir a pasta, o Word fica junto do original. Para texto Markdown da
conversa, indique a pasta ou use o destino padrão
`%LOCALAPPDATA%\FormatadorPeticoes\mcp`.

Quando há uma assinatura configurada neste computador, o plugin insere a imagem
no último bloco cujo nome e OAB correspondem à configuração. Para gerar a peça
sem acrescentar a imagem, indique isso no pedido (`inserir_assinatura=false`).

Exemplo de pedido no chat do Codex:

> Use o plugin FormatadorPeticoes para formatar a petição neste arquivo e
> salvar o Word diretamente na pasta que indiquei.

O arquivo de configuração do plugin inicia `scripts.servidor_mcp` com
`--transporte stdio`. Essa conexão local dispensa túnel, chave de API e abertura
do aplicativo gráfico. As instruções a seguir são para a conexão pelo ChatGPT web.

## Iniciar no Windows

Na pasta `formatador_peticoes`, use Python 3.12 e execute no PowerShell:

```powershell
py -3.12 -m venv .venv-mcp
.\.venv-mcp\Scripts\python.exe -m pip install -e ".[mcp]"
.\.venv-mcp\Scripts\python.exe -m scripts.servidor_mcp
```

O endereço local é `http://127.0.0.1:8765/mcp`. Mantenha essa janela aberta.
A opção `--resultados` permite indicar outra pasta de resultados. A opção
`--modelos` permite indicar a pasta que contém `modelo_oficial.docx` e
`estilos.json`. O transporte `--transporte stdio` também está disponível para
clientes MCP locais.

## Conectar ao ChatGPT

Use um workspace com suporte a ferramentas MCP de escrita e permissão para
adicionar um plugin personalizado. A disponibilidade e as permissões variam
conforme o plano. [Ajuda do ChatGPT](https://help.openai.com/en/articles/12584461-connecting-custom-mcp-servers-in-chatgpt)

1. Crie ou selecione um túnel nos [ajustes da OpenAI Platform](https://platform.openai.com/settings/organization/tunnels) e associe-o ao workspace de destino do ChatGPT.
2. Baixe o [cliente oficial do túnel para Windows](https://github.com/openai/tunnel-client/releases/latest) e execute os comandos a seguir em uma segunda janela do PowerShell. Use o identificador do túnel e uma chave de execução com permissão **Tunnels Read + Use**.
3. No [ChatGPT web](https://chatgpt.com/plugins), abra **Plugins → + → Add custom MCP server**. Informe **FormatadorPeticoes** como nome, escolha **Tunnel** em **Connection** e selecione o túnel. Use **No authentication** para o servidor local, que é acessível pelo túnel privado associado ao workspace. Restrinja o plugin aos usuários que precisam usá-lo.
4. Selecione **Create as a plugin**, confira as ferramentas detectadas e teste com uma petição fictícia. Para manter o serviço acessível, deixe abertos o servidor e o cliente do túnel.

```powershell
$env:CONTROL_PLANE_API_KEY = [System.Net.NetworkCredential]::new(
    '', (Read-Host 'Chave de execução do túnel' -AsSecureString)
).Password
$tunnelId = Read-Host 'Identificador do túnel'
tunnel-client init --sample sample_mcp_remote_no_auth --profile formatador-peticoes --tunnel-id $tunnelId --mcp-server-url http://127.0.0.1:8765/mcp
tunnel-client doctor --profile formatador-peticoes --explain
tunnel-client run --profile formatador-peticoes
```

O acesso externo passa pelo túnel privado da OpenAI. O servidor HTTP escuta
apenas no próprio computador. [Documentação do Secure MCP Tunnel](https://developers.openai.com/api/docs/guides/secure-mcp-tunnels)

## Usar na conversa

- Para texto já redigido na conversa, peça “Formate esta petição inicial no FormatadorPeticoes e devolva o Word”.
- Para arquivos, anexe o DOCX ou o Markdown UTF-8 e peça “Formate a petição inicial anexada usando o FormatadorPeticoes”. Se o Markdown tiver imagens, anexe-as junto com a petição e referencie cada uma pelo nome do arquivo, incluindo a extensão. O MCP recebe até oito imagens na mesma pasta da petição.
- Ao preparar o download, o painel envia o DOCX ao ChatGPT para gerar um link temporário, quando o plugin tem acesso aos recursos de arquivo. Por padrão, o Word fica diretamente em `%LOCALAPPDATA%\FormatadorPeticoes\mcp`. Cada download pelo painel fica disponível durante a execução do servidor que gerou o arquivo.

O MCP aceita Markdown de até 1 MiB e até 15 MiB de entrada por execução, somando
a petição e as imagens. O download no ChatGPT é de até 5 MiB por arquivo.
Os arquivos maiores gerados ficam disponíveis na pasta local.
