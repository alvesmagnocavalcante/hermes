# HERMES — Documentação Técnica

Versão documental: 24 de agosto de 2026  
Versão do projeto: `0.1.0`

## 1. Objetivo

O HERMES é uma aplicação Python para conferência, conciliação e transformação de relatórios contábeis, financeiros, fiscais e operacionais. A mesma base de código pode ser executada como aplicação web ou desktop.

O sistema recebe arquivos exportados de sistemas como CMFlex, CAP, Opera, Simphony, SEFAZ, Portal Nacional, prefeituras e inventário. Cada automação identifica os arquivos, normaliza os dados, executa as regras de negócio e disponibiliza o resultado na interface e em arquivos Excel, PDF ou CSV.

## 2. Escopo funcional

O catálogo atual contém 12 automações:

| Chave | Automação | Entradas | Saídas | Hotel selecionável | Abas-base no Excel |
|---|---|---:|---|---|---|
| `receita` | Conciliação de Receita | 2 | Excel e PDF | Sim | Sim |
| `diarias` | Conciliação da Receita de Diárias | 2 | Excel e PDF | Sim | Não |
| `folha` | Lançamento da Folha de Pagamento | 6 ou 7 | Excel, CSV e PDF | Não | Não |
| `cupons_hospede` | Cupons Emitidos x Conta do Hóspede | 3 | Excel e PDF | Não; identificação automática | Não |
| `rps` | RPS de Serviços Prestados | 3 | Excel e PDF | Não; perfil automático | Não |
| `debito` | Relatório de Notas de Débito | 1 ou mais | Excel e PDF | Não | Não |
| `entrada` | Notas Fiscais de Entrada em Atraso | 2 | Excel e PDF | Não | Não |
| `cupons` | Conferência dos Cupons | 3 | Excel e PDF | Sim | Sim |
| `servicos` | Notas de Serviços Tomados | 3 ou mais | Excel e PDF | Não; identificação pelo CAP | Não |
| `receber` | Conferência do Contas a Receber | 6 | Excel e PDF | Sim | Sim |
| `pagar` | Conferência do Contas a Pagar | 8 | Excel e PDF | Sim | Sim |
| `custos` | Custos da Mercadoria Vendida — CMV | 4 | Excel e PDF | Não | Sim |

Os hotéis disponíveis no seletor são `Cumbuco`, `Magna`, `Taiba` e `Charme`. Relatórios anteriormente associados ao Wind devem usar `Cumbuco`.

## 3. Tecnologias

| Componente | Uso |
|---|---|
| Python 3.12+ | Linguagem e runtime |
| Flet 0.86.x | Interface web e desktop |
| OpenPyXL 3.1.5+ | Leitura e geração de arquivos Excel modernos |
| xlrd 2.0.2+ | Leitura de arquivos `.xls` binários |
| ReportLab 4.4.2+ | Geração dos relatórios PDF |
| Logfire 4.41+ | Eventos operacionais, falhas, duração e métricas básicas do processo |
| `xml.etree.ElementTree` | Leitura dos XMLs do Opera |
| `asyncio` | Coordenação da interface e limitação de trabalhos simultâneos |
| Docker Compose | Empacotamento e execução web em contêiner |
| `unittest` | Testes automatizados |

As dependências e faixas de versões são declaradas em `pyproject.toml`; `uv.lock` fixa as versões efetivamente resolvidas.

## 4. Arquitetura

### 4.1 Fluxo principal

```text
Usuário
  -> interface Flet
  -> validação de quantidade e tamanho do upload
  -> AutomationSpec do catálogo
  -> analisador da automação
  -> resultado tipado em memória
  -> tabela, indicadores e resumo
  -> exportador Excel/PDF/CSV
  -> download ou gravação escolhida pelo usuário
```

### 4.2 Camadas

- `main.py`: configura logging, assets e inicia o Flet.
- `hermes_ui/app.py`: navegação, seleção de hotel, upload, filtros, paginação, resumo, exportação e mensagens de erro.
- `hermes_ui/registry.py`: catálogo declarativo das automações, colunas da interface, formatos, extensões, adaptação dos resultados e nomes de saída.
- `hermes_ui/theme.py`: tokens semânticos do tema Material e alternância entre os modos claro e escuro.
- `hermes_ui/presentation.py`: normalização e classificação visual dos status retornados pelas automações.
- `hermes_ui/runtime.py`: validação de quantidade/tamanho e semáforo de concorrência.
- `hermes_ui/telemetry.py`: configuração opcional do Logfire e seleção dos atributos operacionais enviados.
- `automations/*.py`: reconhecimento de arquivos, leitura, regras de negócio e exportadores.
- `automations/common.py`: normalização, conversão monetária, parsing de datas, leitura comum e inclusão das planilhas-base no Excel.
- `automations/excel_reader.py`: compatibilidade com `.xlsx`, `.xlsm`, `.xltx`, `.xltm`, `.xls` binário, HTML com extensão `.xls` e arquivos modernos renomeados como `.xls`.
- `assets/folha/modelo_folha.xlsm`: parametrização padrão da folha de pagamento.
- `tests/`: testes unitários e de regressão.

### 4.3 Contrato das automações

Cada item do catálogo é um `AutomationSpec` com:

- `key`: identificador usado na interface e no arquivo exportado;
- `name` e `description`: textos apresentados ao usuário;
- `module`: módulo Python responsável pela regra;
- `analyzer`: função de análise;
- `rows_attribute`: coleção exibida na tabela, quando aplicável;
- `columns`: campos, rótulos e indicação de conteúdo numérico;
- `extensions`: extensões permitidas no seletor;
- `formats`: formatos de exportação;
- `hotel_option`: habilita o seletor e o hotel no nome da saída.

O analisador recebe `list[Path]`. Quando `hotel_option=True`, recebe também o hotel selecionado. Os exportadores recebem o resultado da análise e o caminho de saída.

### 4.4 Concorrência

Análises e exportações pesadas são executadas por `asyncio.to_thread()`, evitando bloquear a interface. O semáforo `JOB_LIMITER` limita a quantidade global de operações simultâneas. Operações excedentes aguardam capacidade; não são descartadas.

## 5. Ciclo de vida dos arquivos

### 5.1 Execução web

1. O navegador envia o conteúdo para o servidor HERMES.
2. A aplicação cria `/tmp/hermes_upload_*` dentro do contêiner.
3. A automação analisa os arquivos.
4. O diretório temporário é removido ao terminar a análise, inclusive quando ocorre exceção tratada.
5. Para as cinco automações que anexam planilhas-base, o último conjunto selecionado permanece somente na memória da sessão enquanto o resultado estiver aberto.
6. Na exportação, a aplicação cria `/tmp/hermes_export_*`, gera o resultado e envia os bytes ao navegador.
7. O diretório de exportação é removido após o envio.

Os dados em memória são liberados ao usar **Limpar**, selecionar novos arquivos, mudar de automação ou encerrar/expirar a sessão. O timeout padrão de uma sessão desconectada é 3.600 segundos.

### 5.2 Execução desktop

Os arquivos são lidos diretamente dos caminhos escolhidos. O HERMES não duplica os arquivos de entrada. Nas automações com abas-base, os arquivos originais precisam continuar acessíveis até a exportação.

### 5.3 Persistência

- Não existe banco de dados.
- O `compose.yaml` não declara volume persistente.
- O sistema não mantém histórico de análises.
- Somente o arquivo salvo pelo usuário permanece como resultado.
- Logs operacionais são enviados para `stdout`/`stderr` e ficam sob responsabilidade do mecanismo de logs do Docker.

## 6. Normalização e critérios comuns

- Textos usados como chave são convertidos para maiúsculas, sem acentos e sem pontuação.
- CNPJ e chaves fiscais são comparados sem caracteres de formatação.
- Valores aceitam números do Excel e textos nos padrões brasileiro e internacional.
- Datas aceitam os formatos previstos em cada relatório e valores de data nativos do Excel.
- A tolerância monetária padrão das conciliações é `R$ 0,01`.
- Arquivos são identificados principalmente por cabeçalhos e conteúdo. Algumas rotinas usam partes controladas do nome quando o layout não distingue relatórios semelhantes.
- A ordem de seleção normalmente não importa; quantidade, tipo e período continuam obrigatórios.

## 7. Regras por automação

### 7.1 Conciliação de Receita

Módulo: `automations/conciliacao_receita.py`.

- Compara `Movimento` da Contabilidade com `CASHIER_DEBIT` do Journal do Opera por data.
- Preenche todo o mês predominante, inclusive dias sem movimento.
- Para divergências, tenta identificar combinações de até três descrições do Journal que expliquem o valor.
- Em `Magna`, ignora somente `Ajuste Taxa de Turismo` e `Taxa de Turismo`.
- Exporta `Resumo` e `Conciliação`; o hotel aparece no conteúdo e no nome do arquivo.

### 7.2 Receita de Diárias

Módulo: `automations/conciliacao_receita_diarias.py`.

- Lê a aba do hotel na planilha de códigos de transação.
- Classifica cada `TRX_CODE` como diária e/ou diária média.
- Totaliza quantidade e `CASHIER_DEBIT` por código.
- Código sem lançamento recebe `Sem movimento`.
- O código `1011` é exclusivo do `Taíba`; em `Cumbuco`, `Magna` e `Charme` ele é desconsiderado mesmo que esteja marcado na planilha.
- Exporta `Resumo` e `Detalhamento`.

### 7.3 Folha de Pagamento

Módulo: `automations/lancamento_folha_pagamento.py`.

- Sem férias exige seis relatórios: resumo mensal, INSS, FGTS, IRRF, provisão de férias e provisão de 13º.
- Com férias exige sete relatórios: resumo mensal, INSS, FGTS, recibo de férias, líquido de férias, provisão de férias e provisão de 13º.
- Recibo e líquido de férias são obrigatórios em conjunto.
- Usa `assets/folha/modelo_folha.xlsm`, salvo quando o usuário envia um modelo junto aos relatórios.
- Gera lançamentos por centro de custo, elimina totalizadores e evita duplicidade dos eventos detalhados em relatórios específicos.
- Não gera `REF PLANO ODONTOLÓGICO` nem `REF PLANO DE SAÚDE`.
- Eventos adicionais parametrizados: `271`, `311` e `359`.
- Excel: `Resumo`, `Detalhamento`, `Folha_Mensal`, `Ferias`, `Provisao_Ferias`, `Provisao_13` e `Rateios_Mensais`, conforme existência de movimentos.
- CSV: layout CMFlex com 13 campos, `;` como separador, vírgula decimal e sem cabeçalho.

### 7.4 Cupons Emitidos x Conta do Hóspede

Módulo: `automations/conciliacao_cupons_hospedes.py`.

- Identifica hotel e mapeamento pela melhor relação entre contas do BI/PDV e `CHECK#` do Journal.
- Compara conta, data e valor.
- Pode retornar `Conciliado`, `Conciliado - data diferente`, `Não cobrado`, `Valor divergente`, `Lançado em outra data`, `Ausente na conta` ou `Journal não cobre a data`.
- A situação de período incompleto deve ser corrigida com novo Journal antes de ser tratada como divergência definitiva.

### 7.5 RPS de Serviços Prestados

Módulo: `automations/conferencia_rps_servicos_prestados.py`.

- Compara Opera, Fiscal CMFlex e Prefeitura por RPS, situação e valor.
- Detecta automaticamente o perfil de códigos de serviço com melhor correspondência: padrão, Taíba, Cumbuco ou Charme.
- Também reconhece descrições adicionais de diárias, ajustes, early/late check, serviços do hotel, SPA, shooting, itens e multas. A descrição é normalizada para ignorar acentos, espaços e pontuação.
- Diferencia ausência, divergência de valor, cancelamento, irregularidade e registro fora do período Fiscal.
- Exporta `Resumo` e `Conferencia_RPS`.

### 7.6 Relatório de Notas de Débito

Módulo: `automations/relatorio_notas_debito.py`.

- Aceita uma ou mais planilhas.
- Consolida hotel, comprador, nota, emissão, item e valor.
- Reconhece variações estruturais do relatório e ignora linhas sem conteúdo útil.
- Gera relatório único em Excel e resumo em PDF.

### 7.7 Notas Fiscais de Entrada em Atraso

Módulo: `automations/notas_entrada_atrasadas.py`.

- Relaciona Manifesto e Detalhe por empresa e chave da nota.
- Usa a primeira data de entrada encontrada.
- Arredonda o tempo transcorrido para o dia inteiro mais próximo.
- Ceará: `Em dia` até 5 dias, `Alerta` de 6 a 10 e `Em atraso` a partir de 11.
- Outros estados: `Em dia` até 19 dias, `Alerta` de 20 a 30 e `Em atraso` acima de 30 dias.
- Nota sem entrada é calculada até a data atual.
- Exporta `Resumo` e `Análise`.

### 7.8 Conferência dos Cupons

Módulo: `automations/conferencia_cupons.py`.

- Compara Simphony, Fiscal e SEFAZ por chave, tipo, data e valor.
- Exibe `Status Simphony`: aprovado, cancelado ou ausente.
- Documento cancelado no Simphony concilia quando Fiscal e SEFAZ estão vazios ou zerados.
- Cancelamento com valor diferente de zero no Fiscal ou na SEFAZ é divergente.
- Excel: `Resumo`, `Conciliação`, `Cupons_NFCe`, `Notas_NFe` e abas-base.

### 7.9 Notas de Serviços Tomados

Módulo: `automations/conferencia_notas_servicos_tomados.py`.

- Exige um CAP, um Alterador ISS e pelo menos uma fonte externa.
- As fontes externas podem incluir Portal Nacional, Prefeitura e relatórios municipais em Excel, XLS, HTML/XLS ou CSV.
- Consolida a mesma nota entre fontes externas e relaciona CAP por CNPJ/número; quando seguro, usa número e valor como contingência.
- Verifica valor bruto, existência no CAP, aprovação BPM, hotel, existência na Prefeitura e ISS retido.
- Portal sem fonte municipal continua como `Ausente na Prefeitura`.
- ISS somente é comparado quando aplicável/retido; relatórios de São Paulo são tratados conforme sua regra específica.
- Exporta `Resumo`, `Pendências`, `Conciliadas` e `Base completa`.

### 7.10 Contas a Receber

Módulo: `automations/conferencia_contas_receber.py`.

- Clientes: Balancete por subconta x Posição por cliente.
- Notas a faturar: Borderô x débito do Razão a faturar.
- Comissões: Agregados x movimento do Razão de comissões.
- Consolida nomes relacionados, incluindo todas as variações de CVC, sem misturar BRT e BWT.
- Exporta `Resumo`, `Clientes` e abas-base.

### 7.11 Contas a Pagar

Módulo: `automations/conferencia_contas_pagar.py`.

- Fornecedores: Balancete x Posição por fornecedor.
- Adiantamentos: Balancete de adiantamentos x adiantamentos em aberto.
- Impostos: agregados IRRF/CSRF/ISS x movimentos credores do Razão de impostos.
- Identifica os alteradores `IRRF`, `CSRF` e `ISS` pelo conteúdo de `Descricao`/`Historico`, usando o nome apenas como compatibilidade para relatórios antigos ou vazios.
- Distingue o balancete de adiantamentos do balancete de fornecedores pelo conteúdo de `DescricaoConta`; os arquivos podem manter os nomes originais do CMFlex.
- Consolida grupos comerciais, incluindo CVC; BRT e BWT permanecem separados.
- Exporta `Resumo`, `Fornecedores e Adiantamentos` e abas-base.

### 7.12 Custos da Mercadoria Vendida

Módulo: `automations/conferencia_custos_mercadoria.py`.

- Executa duas análises independentes.
- `Entradas`: documentos CAP x débitos da Contabilidade cujo histórico contém simultaneamente `Nota Fiscal`, `Mercadoria` e `Terceiros`.
- `Saldo final`: grupos do Inventário x saldo atual das contas contábeis de estoque.
- Transferências, requisições, ajustes e integrações de custo não entram em `Entradas`.
- Contas sem movimento continuam usando o saldo final encontrado nas linhas subsequentes do Razão.
- Exporta `Entradas`, `Saldo final` e abas-base.

## 8. Exportações

### 8.1 Nomes

Automações com seletor usam `<chave>_<hotel>_resultado.<extensão>`. Exemplos:

- `receita_charme_resultado.xlsx`;
- `diarias_magna_resultado.pdf`;
- `cupons_cumbuco_resultado.xlsx`;
- `receber_taiba_resultado.xlsx`;
- `pagar_charme_resultado.pdf`.

As demais usam `<chave>_resultado.<extensão>`.

### 8.2 Abas-base

As automações `receita`, `cupons`, `receber`, `pagar` e `custos` acrescentam ao Excel uma cópia consultável de todas as abas das planilhas analisadas. As abas do resultado ficam primeiro. Os nomes são saneados, limitados a 31 caracteres e recebem sufixos quando houver colisão.

A cópia preserva os valores exibidos para consulta; não existe garantia de reprodução integral de macros, estilos, objetos ou fórmulas da origem.

## 9. Configuração e execução

### 9.1 Desenvolvimento com `uv`

```powershell
$env:UV_CACHE_DIR = "$PWD\.uv-cache"
uv sync
uv run python main.py
```

Para iniciar explicitamente em modo web:

```powershell
$env:FLET_FORCE_WEB_SERVER = "true"
$env:FLET_SERVER_IP = "0.0.0.0"
$env:FLET_SERVER_PORT = "8000"
uv run python main.py
```

### 9.2 Docker

```powershell
docker compose build
docker compose up -d
docker compose ps
```

Parada:

```powershell
docker compose down
```

Logs:

```powershell
docker compose logs -f hermes
```

### 9.3 Variáveis de ambiente

| Variável | Padrão | Finalidade |
|---|---:|---|
| `NGINX_BIND_ADDRESS` | `0.0.0.0` | Interface em que o proxy reverso atende |
| `NGINX_PORT` | `80` | Porta HTTP publicada pelo Nginx |
| `FLET_MAX_UPLOAD_SIZE` | `104857600` | Limite individual aceito pelo Flet, em bytes |
| `FLET_SESSION_TIMEOUT` | `30` | Limite de retenção da sessão desconectada, em segundos |
| `HERMES_MAX_CONCURRENT_JOBS` | `2` | Análises/exportações pesadas simultâneas |
| `HERMES_MAX_FILES_PER_JOB` | `20` | Arquivos por operação |
| `HERMES_MAX_TOTAL_UPLOAD_SIZE` | `209715200` | Soma máxima por operação, em bytes |
| `HERMES_CPU_LIMIT` | `2.0` | CPUs do contêiner |
| `HERMES_MEMORY_LIMIT` | `2g` | Memória do contêiner |
| `HERMES_LOG_LEVEL` | `INFO` | Nível dos logs |
| `HERMES_TELEMETRY_ENABLED` | `true` | Permite desativar integralmente a integração Logfire |
| `HERMES_LOGFIRE_SYSTEM_METRICS` | `true` | Coleta métricas básicas de CPU e memória |
| `LOGFIRE_ENVIRONMENT` | `production` | Ambiente exibido nos filtros do Logfire |
| `LOGFIRE_TOKEN` | vazio | Write Token opcional; a credencial do CLI é usada por padrão |

Exemplo de `.env` para rede local:

```dotenv
NGINX_BIND_ADDRESS=0.0.0.0
NGINX_PORT=80
HERMES_MEMORY_LIMIT=2g
HERMES_CPU_LIMIT=2.0
HERMES_TELEMETRY_ENABLED=true
HERMES_LOGFIRE_SYSTEM_METRICS=true
LOGFIRE_ENVIRONMENT=production
LOGFIRE_TOKEN=
```

### 9.4 Telemetria Logfire

O SDK é configurado uma única vez em `main.py`. A credencial é gerada por
`uv run logfire auth` e `uv run logfire projects use --org
'alvesmagnocavalcante' 'carmel'` na pasta `.logfire`. No Docker, essa pasta é
montada em `/app/.logfire` somente para leitura. Ela está excluída do Git e da
imagem Docker. `LOGFIRE_TOKEN` permanece disponível como configuração opcional.

São enviados apenas metadados operacionais:

- resultado de autenticação: autorizado, recusado, bloqueado ou erro de configuração;
- chave da automação, hotel, sucesso, quantidade e volume total dos arquivos;
- quantidade de registros processados, duração e tipo da exceção em falhas;
- totais conciliados, pendentes e informativos, percentual de qualidade e modo
  de execução web/desktop;
- mensagem operacional da falha limitada a 500 caracteres, com caminhos,
  segredos e identificadores fiscais extensos removidos;
- formato e duração das exportações;
- métricas básicas de CPU e memória do processo/servidor.

Não são enviados usuário, senha, endereço IP, nome ou conteúdo dos arquivos,
chaves fiscais, CNPJ, valores financeiros ou mensagens de exceção. Falhas no
Logfire são isoladas e não interrompem a execução das automações.

O firewall do Windows deve permitir a porta 80 somente para a rede administrativa
`10.197.0.0/22`. Em um PowerShell executado **como administrador**:

```powershell
New-NetFirewallRule -DisplayName "HERMES Nginx - Rede Administrativa" `
  -Direction Inbound -Protocol TCP -LocalPort 80 `
  -RemoteAddress 10.197.0.0/22 -Action Allow -Profile Domain,Private
```

O Nginx preserva WebSocket e encaminha as requisições para `hermes:8000`;
essa porta fica disponível apenas na rede interna do Compose.

O HERMES apresenta uma tela de login e valida a credencial em
`nginx/.htpasswd`, usando hash bcrypt com custo 12. O arquivo está excluído do
Git e é montado no contêiner `hermes` somente para leitura. O estado autenticado
permanece apenas na memória da sessão Flet e é removido nos eventos de
desconexão ou fechamento. Para trocar a senha, gere novamente o arquivo e
recrie o serviço `hermes`. Enquanto HTTPS não estiver configurado, a
autenticação não criptografa o tráfego da rede.

## 10. Segurança operacional

- O contêiner executa com usuário não administrativo (`UID 10001`).
- Todas as capabilities Linux são removidas e `no-new-privileges` é habilitado.
- O contêiner possui limite de CPU, memória e processos.
- O healthcheck consulta a aplicação localmente a cada 30 segundos.
- O HERMES exige login e limita cinco tentativas incorretas por cliente em uma janela de cinco minutos.
- O Nginx atua como proxy reverso e preserva as conexões WebSocket do Flet.
- A porta da aplicação não é publicada diretamente; somente o Nginx atende na porta 80.
- Em rede não confiável ou internet, adicione HTTPS e autenticação ao proxy.
- Controle acesso aos relatórios exportados, pois eles podem conter dados pessoais, fiscais, financeiros e trabalhistas.

## 11. Testes e validação

Executar a suíte:

```powershell
$env:UV_CACHE_DIR = "$PWD\.uv-cache"
uv run python -m unittest discover -s tests -v
```

Validar sintaxe/importação:

```powershell
uv run python -m compileall -q automations hermes_ui tests
```

Validar o contêiner:

```powershell
docker compose build
docker compose up -d
docker compose ps
```

O estado esperado é `healthy`.

## 12. Manutenção

### 12.1 Alterar uma regra existente

1. Localize o módulo em `automations/`.
2. Preserve leitura, regra e exportação no mesmo módulo.
3. Adicione um teste que reproduza o caso informado pelo usuário.
4. Execute toda a suíte, não somente o teste novo.
5. Atualize esta documentação e o manual quando a mudança for visível ao usuário.
6. Reconstrua o Docker quando a execução publicada usar contêiner.

### 12.2 Adicionar uma automação

1. Crie `automations/nova_automacao.py`.
2. Implemente um analisador que receba caminhos e devolva objetos estáveis.
3. Implemente `save_excel` e `save_pdf`; inclua `save_csv` somente quando houver layout definido.
4. Registre um `AutomationSpec` em `hermes_ui/registry.py`.
5. Adicione a orientação de arquivos em `AutomationView._file_guidance()`.
6. Crie testes de reconhecimento, regra, erro e exportação.
7. Documente entradas, resultados e interpretação.

### 12.3 Modelo da folha

O arquivo `assets/folha/modelo_folha.xlsm` é binário e contém VBA. Alterações devem preservar `vbaProject.bin`, tabelas e intervalos. Após editar:

- abra o modelo com `keep_vba=True`;
- valide os eventos por `read_mappings()`;
- teste a integridade ZIP do XLSM;
- confirme que as macros continuam presentes;
- execute os testes da folha e a suíte completa.

## 13. Limitações conhecidas

- O sistema depende dos layouts e cabeçalhos previstos. Mudanças no relatório de origem podem exigir atualização do leitor.
- Não há histórico, banco de dados ou reprocessamento automático.
- O PDF é um resumo; o Excel é a fonte detalhada para auditoria.
- Abas-base são cópias de valores para consulta, não réplicas visuais completas.
- A seleção incorreta de hotel pode aplicar regras e nomes de saída incorretos; por isso a troca de hotel limpa o resultado anterior.
- O acesso em rede local depende da porta publicada, do firewall e da disponibilidade do computador servidor.

## 14. Documentos relacionados

- [README do projeto](../README.md)
- [Manual do Usuário](GUIA_USUARIO_HERMES.md)
