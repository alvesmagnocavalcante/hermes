# HERMES — Painel de Automação de Planilhas

Aplicação desktop modular para conferência, conciliação e geração de relatórios a partir de planilhas Excel e arquivos CSV.

O painel utiliza tema escuro, menu lateral recolhível, tabelas com filtros e paginação, gráficos de resumo, processamento em segundo plano e exportação para Excel ou PDF.

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
| Relatório de Notas de Débito | 1 ou mais | Consolida hotel, comprador, nota, emissão, item e valor. |
| Notas Fiscais de Entrada em Atraso | 2 | Compara Manifesto e Detalhe, aplicando prazo de 11 dias para o Ceará e 30 dias para outros estados. |
| Conferência dos Cupons | 3 | Compara Simphony, Fiscal e SEFAZ por chave e valor, distinguindo NFC-e e NF-e/Unknown. |
| Notas de Serviços Tomados | 5 | Compara Portal Nacional, prefeituras, CAP, BPM, hotel e ISS retido. |
| Contas a Receber | 6 | Confere clientes, notas a faturar e comissões contra o financeiro. |
| Contas a Pagar | 8 | Confere fornecedores, adiantamentos e impostos contra o financeiro. |
| Custos da Mercadoria Vendida | 4 | Compara entradas do CAP e saldos do inventário com a contabilidade. |

As planilhas usadas para validação ficam em `PROCESSOS AUTOMAÇÃO/`, organizadas por atividade.

### Atividade 2 — Folha de pagamento

Selecione simultaneamente os sete relatórios CSV: resumo mensal, relação de INSS, relação de FGTS, recibo de férias, líquido de férias, provisão de férias e provisão de 13º. Os nomes dos arquivos não são usados na identificação; o sistema reconhece cada relatório pelo conteúdo.

A planilha `Projeto 2 - DP CARMEL PADRÃO...xlsm` contém os de/para e as regras adotadas pelo departamento. Ela é carregada automaticamente da pasta de exemplo. Se uma versão atualizada da planilha for selecionada junto aos CSVs, essa versão passa a ser usada na análise.

O Excel exportado separa as saídas em `Folha_Mensal`, `Ferias`, `Provisao_Ferias`, `Provisao_13` e `Rateios_Mensais`, além das abas de resumo e detalhamento. Linhas de total do organograma, filial e empresa não são importadas.

## Estrutura

```text
projeto-hermes/
├── automations/
│   ├── __init__.py
│   ├── base.py                              # Contrato e constantes visuais
│   ├── ui.py                                # Tabela compartilhada e cores de status
│   ├── conciliacao_receita.py
│   ├── conciliacao_receita_diarias.py
│   ├── lancamento_folha_pagamento.py
│   ├── relatorio_notas_debito.py
│   ├── notas_entrada_atrasadas.py
│   ├── conferencia_cupons.py
│   ├── conferencia_notas_servicos_tomados.py
│   ├── conferencia_contas_receber.py
│   ├── conferencia_contas_pagar.py
│   └── conferencia_custos_mercadoria.py
├── PROCESSOS AUTOMAÇÃO/                     # Planilhas de exemplo
├── main.py                                  # Janela, navegação e threads
├── pyproject.toml
└── uv.lock
```

## Arquitetura

Cada automação:

1. Herda de `Automation`, definida em `automations/base.py`.
2. Implementa o método `render()`.
3. Exporta sua classe por meio da variável `AUTOMATION_CLASS`.
4. Mantém leitura, análise e exportação dentro do próprio módulo.
5. Executa operações demoradas com `self.app.run_background()` para não travar a interface.

O `main.py` descobre os módulos dinamicamente com `pkgutil` e adiciona as automações ao menu em ordem alfabética.

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
- Tabelas criadas com `create_result_table()` de `automations/ui.py`.
- Cabeçalhos fixos, seleção de linha e rolagem horizontal e vertical.
- Paginação para resultados extensos.

O menu lateral inicia recolhido, mantendo o nome `HERMES` visível. Use o botão `☰` para abrir a lista de automações.

## Criar uma automação

Crie um arquivo em `automations/`:

```python
import customtkinter as ctk

from automations.base import Automation
from automations.ui import TableColumn, create_result_table


class MinhaAutomation(Automation):
    name = "Minha Automação"

    def render(self) -> None:
        self.container.grid_columnconfigure(0, weight=1)
        self.container.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(
            self.container,
            text=self.name,
            font=ctk.CTkFont(size=26, weight="bold"),
        ).grid(row=0, column=0, padx=30, pady=(22, 10), sticky="w")

        self.table = create_result_table(
            self.container,
            (
                TableColumn("document", "Documento", 220),
                TableColumn("status", "Situação", 160),
            ),
            row=1,
        )


AUTOMATION_CLASS = MinhaAutomation
```

Na próxima inicialização, a nova automação aparecerá automaticamente no menu.

## Processamento em segundo plano

```python
self.app.run_background(
    tarefa,
    callback_sucesso,
    callback_erro,
)
```

Para atualizar o rodapé durante uma operação:

```python
self.app.report_progress("Processando planilhas...", 0.5)
```

Exceções lançadas pela tarefa são tratadas pelo painel e exibidas em uma mensagem visual, sem bloquear a janela principal.

## Exportações

As automações podem gerar:

- Excel com resumo, filtros, formatação monetária e detalhamento completo.
- PDF com resumo executivo da conferência.

Os arquivos de saída são escolhidos pelo usuário e não sobrescrevem automaticamente as planilhas originais.
