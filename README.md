# HERMES — Painel de Automação de Planilhas

Aplicação desktop modular em Flet para conferência, conciliação e geração de relatórios a partir de planilhas Excel (`.xlsx`, `.xlsm`, `.xls`, `.xltx` e `.xltm`), CSV e XML.

O painel utiliza tema escuro, menu lateral recolhível, componentes responsivos, tabelas com filtros e paginação, indicadores visuais, processamento assíncrono e exportação para Excel, PDF ou CSV conforme a automação.

## O que o sistema faz

O HERMES centraliza conferências financeiras, fiscais e contábeis que antes precisavam ser realizadas manualmente em várias planilhas.

Em cada automação, o sistema:

1. Solicita os relatórios necessários para a conferência escolhida.
2. Identifica automaticamente cada arquivo pelas colunas e pelo formato esperado.
3. Lê e normaliza chaves, documentos, CNPJ, datas, nomes e valores monetários.
4. Cruza as informações entre sistemas como CAP, Contabilidade, Opera, Simphony, Fiscal, SEFAZ, prefeituras e inventário.
5. Classifica cada registro como conciliado, divergente, ausente, atrasado ou não escriturado.
6. Exibe totais, indicadores e gráfico de distribuição.
7. Permite pesquisar, filtrar, selecionar linhas e navegar pelos resultados.
8. Gera Excel detalhado e PDF de resumo sem alterar os arquivos originais.

O processamento ocorre em segundo plano, mantendo a interface disponível enquanto as planilhas são analisadas. Erros de arquivo, colunas ausentes ou formatos incompatíveis são apresentados ao usuário por mensagens visuais.

O objetivo é reduzir o tempo de conferência, identificar integrações ausentes e tornar as divergências rastreáveis até o documento de origem.

### Formatos de entrada

Todas as automações que utilizam planilhas Excel aceitam `.xlsx`, `.xlsm`, `.xls`, `.xltx` e `.xltm`. O leitor também reconhece relatórios antigos `.xls` em formato binário, relatórios HTML entregues com extensão `.xls` e arquivos modernos que tenham sido fornecidos com essa extensão.

## Guia do usuário

- [Guia em Word](docs/GUIA_USUARIO_HERMES.docx) — versão pronta para distribuição aos usuários.
- [Guia em Markdown](docs/GUIA_USUARIO_HERMES.md) — versão para consulta no repositório.

## Requisitos

- Windows
- Python 3.12 ou superior
- [uv](https://docs.astral.sh/uv/) — recomendado

## Instalação e execução

Com `uv`:

```powershell
uv sync
uv run python main.py
```

Alternativa com `pip`:

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install -e .
python main.py
```

Execute sempre o `main.py` pela raiz do projeto. Os módulos dentro de `automations/` são carregados automaticamente e não devem ser executados diretamente.

## Automações disponíveis

| Automação | Arquivos | Finalidade |
|---|---:|---|
| Conciliação de Receita | 2 | Compara Movimento da Contabilidade com `CASHIER_DEBIT` do Opera. |
| Conciliação da Receita de Diárias | 2 | Classifica e totaliza o `CASHIER_DEBIT` pelos `TRX_CODE` marcados como diária e diária média para cada hotel. |
| Lançamento da Folha de Pagamento | 7 | Reconhece os relatórios pelo conteúdo e gera as importações da folha mensal, férias, provisões de férias e 13º, além dos rateios de INSS, FGTS e planos por centro de custo. Todos os totalizadores são excluídos para impedir valores duplicados. |
| Cupons Emitidos x Conta do Hóspede | 3 | Confere se os cupons do BI/PDV constam no `CHECK#` do Journal e se o valor foi efetivamente cobrado na conta do hóspede. |
| RPS de Serviços Prestados | 3 | Confere se os RPS encerrados no Opera integraram no Fiscal do CMFlex, foram emitidos na Prefeitura e possuem o mesmo valor nas três fontes. |
| Relatório de Notas de Débito | 1 ou mais | Consolida hotel, comprador, nota, emissão, item e valor. |
| Notas Fiscais de Entrada em Atraso | 2 | Compara Manifesto e Detalhe, aplicando prazo de 11 dias para o Ceará e 30 dias para outros estados. |
| Conferência dos Cupons | 3 | Compara Simphony, Fiscal e SEFAZ por chave e valor, distinguindo NFC-e e NF-e/Unknown. |
| Notas de Serviços Tomados | 5 | Compara Portal Nacional, prefeituras, CAP, BPM, hotel e ISS retido. |
| Contas a Receber | 6 | Confere clientes, notas a faturar e comissões contra o financeiro. |
| Contas a Pagar | 8 | Confere fornecedores, adiantamentos e impostos contra o financeiro. |
| Custos da Mercadoria Vendida | 4 | Compara entradas do CAP e saldos do inventário com a contabilidade. |

As planilhas usadas para validação ficam em `PROCESSOS AUTOMAÇÃO/`, organizadas por atividade.

### Atividade 2 — Folha de pagamento

Selecione simultaneamente os sete relatórios: resumo mensal, relação de INSS, relação de FGTS, recibo de férias, líquido de férias, provisão de férias e provisão de 13º. Eles podem estar em CSV ou Excel (`.xlsx`, `.xlsm`, `.xls`, `.xltx` ou `.xltm`). Os nomes dos arquivos não são usados na identificação; o sistema reconhece cada relatório pelo conteúdo.

A planilha `Projeto 2 - DP CARMEL PADRÃO...xlsm` contém os de/para e as regras adotadas pelo departamento. Ela é carregada automaticamente da pasta de exemplo. Se uma versão atualizada da planilha for selecionada junto aos CSVs, essa versão passa a ser usada na análise.

O Excel exportado separa as saídas em `Folha_Mensal`, `Ferias`, `Provisao_Ferias`, `Provisao_13` e `Rateios_Mensais`, além das abas de resumo e detalhamento. A opção CSV segue o modelo do CMFlex: campos separados por ponto e vírgula, valores com vírgula decimal, sem cabeçalho e com 13 colunas. A data é preenchida automaticamente com o fim da competência. O filtro permite exportar todas as saídas juntas ou somente um processo, como férias. Linhas de total do organograma, filial e empresa não são importadas.

## Estrutura

```text
projeto-hermes/
├── automations/
│   ├── __init__.py
│   ├── base.py                              # Contrato e constantes visuais
│   ├── ui.py                                # Tabela compartilhada e cores de status
│   ├── conciliacao_receita.py
│   ├── conciliacao_receita_diarias.py
│   ├── conciliacao_cupons_hospedes.py
│   ├── conferencia_rps_servicos_prestados.py
│   ├── lancamento_folha_pagamento.py
│   ├── relatorio_notas_debito.py
│   ├── notas_entrada_atrasadas.py
│   ├── conferencia_cupons.py
│   ├── conferencia_notas_servicos_tomados.py
│   ├── conferencia_contas_receber.py
│   ├── conferencia_contas_pagar.py
│   └── conferencia_custos_mercadoria.py
├── hermes_ui/
│   ├── app.py                               # Janela, navegação e componentes Flet
│   └── registry.py                          # Catálogo e adaptadores das automações
├── PROCESSOS AUTOMAÇÃO/                     # Planilhas de exemplo
├── main.py                                  # Ponto de entrada Flet
├── pyproject.toml
└── uv.lock
```

## Arquitetura

Cada automação mantém a leitura, as regras de conferência e os exportadores em seu próprio módulo. A camada visual:

1. É implementada em Flet dentro de `hermes_ui/app.py`.
2. Usa um único padrão de cards, filtros, tabela, gráfico, paginação e rodapé.
3. Localiza as automações pelo catálogo declarativo de `hermes_ui/registry.py`.
4. Executa leituras e exportações com `asyncio.to_thread()`, sem bloquear a interface.
5. Preserva as funções de análise e exportação existentes em `automations/`.

O `main.py` inicia o aplicativo Flet e o catálogo define a ordem do menu, os arquivos aceitos, as colunas e os formatos de saída.

## Padrão visual

Todas as telas devem seguir o mesmo padrão:

- Margem lateral de 30 px.
- Título de 26 px.
- Gráfico com 125 px de altura.
- Azul para ações principais.
- Cinza para ações secundárias.
- Verde `#21a67a` para resultados conciliados.
- Amarelo `#e0a83e` para informações ausentes.
- Vermelho `#dc5a5a` para divergências ou itens não escriturados.
- Tabelas Flet compartilhadas com rolagem horizontal e vertical.
- Cabeçalhos fixos, seleção de linha e rolagem horizontal e vertical.
- Paginação para resultados extensos.

O menu lateral inicia recolhido, mantendo o nome `HERMES` visível. Use o botão `☰` para abrir a lista de automações.

## Criar uma automação

Crie as funções de análise e exportação em um arquivo de `automations/`. Depois registre a tela em `hermes_ui/registry.py` com `AutomationSpec`:

```python
AutomationSpec(
    key="minha_automacao",
    name="Minha Automação",
    description="Explicação clara da conferência.",
    module="automations.minha_automacao",
    analyzer="analyze",
    rows_attribute="rows",
    columns=(
        Column("document", "Documento"),
        Column("status", "Resultado"),
    ),
)
```

O catálogo adiciona a automação ao menu e reutiliza toda a interface padrão.

## Processamento em segundo plano

As funções síncronas de processamento são executadas por `asyncio.to_thread()`. Exceções são tratadas pela camada visual e exibidas em um diálogo, sem encerrar o aplicativo.

## Exportações

As automações podem gerar:

- Excel com resumo, filtros, formatação monetária e detalhamento completo.
- PDF com resumo executivo da conferência.

Os arquivos de saída são escolhidos pelo usuário e não sobrescrevem automaticamente as planilhas originais.
