# HERMES — Manual do Usuário

Versão: 24 de agosto de 2026

## Sumário

1. [Finalidade](#1-finalidade)
2. [Acesso e preparação](#2-acesso-e-preparação)
3. [Procedimento padrão](#3-procedimento-padrão)
4. [Como interpretar a tela](#4-como-interpretar-a-tela)
5. [Automações](#5-automações)
6. [Exportações](#6-exportações)
7. [Mensagens de erro](#7-mensagens-de-erro)
8. [Segurança e armazenamento](#8-segurança-e-armazenamento)
9. [Checklist final](#9-checklist-final)

## 1. Finalidade

O HERMES automatiza conferências entre relatórios contábeis, financeiros, fiscais e operacionais. Ele identifica os arquivos selecionados, compara os dados e mostra quais registros estão conciliados e quais precisam de análise.

O sistema permite:

- comparar valores e documentos entre sistemas diferentes;
- identificar lançamentos ausentes;
- localizar diferenças por data, cliente, fornecedor, conta, nota, cupom ou RPS;
- consultar o resultado na tela;
- exportar relatório detalhado em Excel;
- exportar resumo em PDF;
- gerar o CSV da Folha de Pagamento para importação no CMFlex.

O HERMES não corrige nem altera os arquivos de origem. A decisão sobre ajustes contábeis, financeiros ou fiscais continua sendo do usuário responsável.

## 2. Acesso e preparação

### 2.1 Antes de abrir o HERMES

1. Gere todos os relatórios para o mesmo hotel, empresa e período.
2. Não apague cabeçalhos nem colunas.
3. Não reorganize, filtre ou salve os arquivos em outro formato antes da análise.
4. Confirme se os relatórios terminaram de ser gerados e não estão corrompidos.
5. Feche os arquivos no Excel quando for usar o aplicativo desktop.
6. Mantenha os arquivos originais disponíveis até concluir a exportação.

### 2.2 Acesso pelo navegador

Abra o endereço informado pela equipe responsável pelo HERMES. O processamento ocorre no computador ou servidor onde o Docker está executando, não dentro do navegador do usuário.

Se a página não abrir:

1. confirme se o computador servidor está ligado;
2. confirme se o HERMES está em execução;
3. verifique se você está conectado à mesma rede autorizada;
4. informe o endereço e a mensagem exibida ao suporte.

### 2.3 Hotéis disponíveis

As telas que possuem o campo **Hotel** permitem selecionar:

- Cumbuco;
- Magna;
- Taiba;
- Charme.

Use **Cumbuco** para relatórios anteriormente identificados como Wind. Wind e Cumbuco são tratados como o mesmo local no HERMES.

## 3. Procedimento padrão

1. Abra o HERMES.
2. Use o menu lateral para escolher a atividade.
3. Leia a orientação abaixo de **Arquivos necessários**.
4. Se houver o campo **Hotel**, selecione o hotel antes de carregar os arquivos.
5. Clique em **Arquivos**.
6. Selecione todos os relatórios solicitados na mesma operação.
7. Aguarde **Conferência concluída**.
8. Leia os indicadores e clique em **Resumo** quando precisar dos totais comparados.
9. Use **Exibir** para separar conciliados e pendências.
10. Use **Buscar nos resultados** para localizar um registro.
11. Analise as pendências e confirme os dados nas fontes de origem.
12. Escolha o formato no campo **Formato**.
13. Clique em **Exportar** e salve o resultado.
14. Clique em **Limpar** antes de iniciar outra conferência.

Os arquivos podem ser selecionados em qualquer ordem, salvo quando uma regra específica depender do nome do relatório. O HERMES normalmente identifica cada arquivo pelos cabeçalhos e pelo conteúdo.

## 4. Como interpretar a tela

### 4.1 Arquivos necessários

Mostra a quantidade e os tipos de relatórios exigidos. Depois do processamento, informa quantos arquivos foram reconhecidos. Posicione o mouse sobre essa informação para consultar os nomes selecionados.

### 4.2 Indicadores

Os cartões apresentam:

- quantidade total de registros;
- conciliados ou prontos;
- pendências;
- informativos;
- percentual de qualidade/conciliação.

Na Folha de Pagamento e nas Notas de Entrada, os rótulos são adaptados à atividade.

### 4.3 Resumo

O botão **Resumo** apresenta totais e comparações entre as fontes. Exemplo:

> Clientes: Balancete R$ 100.000,00 • Posição por cliente R$ 99.500,00 • Diferença R$ 500,00 • Divergente

O resumo ajuda a localizar qual bloco da conferência precisa de atenção. A análise definitiva deve considerar também as linhas detalhadas.

### 4.4 Cores

- **Verde:** conciliado, correto ou pronto para importação.
- **Amarelo:** informativo ou situação que precisa de contexto, como período incompleto, cancelamento ou ausência de movimento.
- **Vermelho:** divergência, ausência, atraso ou parametrização incompleta.

A cor não substitui a leitura do texto em **Resultado**, **Situação**, **Motivo** ou **Explicação**.

### 4.5 Filtros e pesquisa

- **Todos:** exibe todos os registros.
- **Conciliados:** exibe somente resultados considerados corretos.
- **Pendências:** exibe resultados que exigem análise.
- Na atividade de Notas de Entrada, o filtro permite `Em dia`, `Alerta`, `Em atraso` e `Todos`.
- A pesquisa verifica todas as colunas visíveis e não diferencia letras maiúsculas de minúsculas.
- Use as setas de paginação para navegar. A quantidade de linhas por página se adapta ao tamanho da janela.

### 4.6 Troca de hotel

Se o hotel for alterado depois de uma análise, o resultado anterior é limpo. Isso impede que um relatório seja exportado com o hotel errado.

### 4.7 Limpar

O botão **Limpar** remove o resultado da tela e libera os dados temporários da sessão. Ele não apaga:

- os arquivos originais;
- os relatórios que já foram baixados;
- arquivos existentes no computador do usuário.

## 5. Automações

## 5.1 Conciliação de Receita

### Objetivo

Comparar a receita diária registrada na Contabilidade com os lançamentos do Journal do Opera.

### Antes de começar

- escolha o hotel correto;
- gere os dois relatórios para o mesmo mês;
- confirme se o Journal possui todas as datas do período.

### Arquivos

1. Razão Analítico da Contabilidade.
2. Journal de Receita do Opera.

### Conferência executada

- Contabilidade: soma da coluna de movimento por data.
- Opera: soma de `CASHIER_DEBIT` por data de negócio.
- O resultado apresenta todos os dias do mês predominante, inclusive dias zerados.
- Diferença = Contabilidade menos Opera.
- Para divergências, o HERMES tenta apontar descrições do Journal que formem o valor.
- No hotel Magna, `Ajuste Taxa de Turismo` e `Taxa de Turismo` são desconsiderados. Essa exceção não é aplicada aos demais hotéis.

### Resultados

- **Conciliado:** diferença de até R$ 0,01.
- **Divergente:** diferença superior a R$ 0,01.
- **Diferença sem identificação automática:** o valor diverge, mas não foi encontrada combinação segura de descrições no Journal.

### Exportação

- Excel: `Resumo`, `Conciliação` e cópias das abas dos dois arquivos-base.
- PDF: resumo da conferência.
- Exemplo: `receita_charme_resultado.xlsx`.

## 5.2 Conciliação da Receita de Diárias

### Objetivo

Totalizar os códigos do Journal definidos como diária e diária média para o hotel selecionado.

### Arquivos

1. Planilha de Códigos de Transação.
2. Journal de Receita do Opera.

### Conferência executada

- lê a aba correspondente ao hotel;
- identifica os `TRX_CODE` marcados como diária e/ou diária média;
- conta os lançamentos e soma `CASHIER_DEBIT` por código;
- mostra também códigos configurados sem movimento;
- desconsidera o código `1011` somente no Charme.

### Resultados

- **Com movimento:** código encontrado no Journal.
- **Sem movimento:** código configurado sem lançamento no período; é informativo e não significa necessariamente erro.

### Exportação

- Excel: `Resumo` e `Detalhamento`.
- PDF: resumo por código.
- Exemplo: `diarias_cumbuco_resultado.xlsx`.

## 5.3 Lançamento da Folha de Pagamento

### Objetivo

Gerar lançamentos contábeis por centro de custo para folha mensal, férias, provisões, INSS e FGTS, removendo totalizadores e duplicidades.

### Cenário sem férias — 6 arquivos

1. Resumo mensal da folha.
2. Relação de INSS.
3. Relação de FGTS.
4. Relação mensal de IRRF.
5. Provisão de férias.
6. Provisão de 13º salário.

### Cenário com férias — 7 arquivos

1. Resumo mensal da folha.
2. Relação de INSS.
3. Relação de FGTS.
4. Recibo de férias.
5. Líquido de férias.
6. Provisão de férias.
7. Provisão de 13º salário.

Recibo e líquido de férias devem ser enviados juntos. No cenário com férias, a relação mensal de IRRF não substitui nenhum dos dois relatórios de férias.

### Parametrização

O modelo padrão é [modelo_folha.xlsm](../assets/folha/modelo_folha.xlsm). Se um modelo compatível for enviado junto com os relatórios, ele será usado somente naquela análise.

Eventos adicionais:

| Evento | Descrição | Débito | Crédito |
|---:|---|---|---|
| 271 | Ajuda de transporte estagiário | `3.02.02.01.01` | `2.01.01.01.01` |
| 311 | 13º Salário Adiantamento | `1.01.02.01.04` | `2.01.01.01.01` |
| 359 | Horas Férias Noturnas | `2.01.01.06.01` | `2.01.01.01.03` |

`REF PLANO ODONTOLÓGICO` e `REF PLANO DE SAÚDE` não geram lançamento automático. Esses valores dependem das notas fiscais e do preenchimento manual posterior.

### Resultados

- **Dados completos / Pronto para importar:** conta e centro de custo preenchidos.
- **De/para incompleto:** falta conta ou centro de custo; não importe antes de corrigir.
- **Totalizadores removidos:** linhas agregadas foram descartadas para evitar duplicidade.
- **Eventos desconsiderados:** eventos detalhados em fontes específicas foram retirados do resumo mensal.

### Exportação

- Excel: resumo, detalhamento e abas por processo.
- PDF: resumo executivo.
- CSV: arquivo de importação do CMFlex com 13 campos, sem cabeçalho e separado por `;`.

Não abra e salve novamente o CSV no Excel. Isso pode alterar separadores, codificação, casas decimais e zeros dos códigos.

### Validação obrigatória antes da importação

1. Empresa e competência.
2. Total de proventos e descontos.
3. Líquido a pagar.
4. Contas de débito e crédito.
5. Centros de custo.
6. Linhas com `De/para incompleto`.
7. Quantidade de funcionários e lançamentos de férias, quando houver.

## 5.4 Cupons Emitidos x Conta do Hóspede

### Objetivo

Confirmar se os cupons emitidos no BI/PDV foram lançados e efetivamente cobrados na conta do hóspede no Journal.

### Arquivos

1. Relatório BI/PDV.
2. Journal do Opera.
3. Planilha de correspondência/de-para dos `TRX_CODE`.

### Conferência executada

- identifica automaticamente hotel e de/para com maior correspondência;
- relaciona conta do PDV com `CHECK#` do Journal;
- compara data e valor líquido;
- procura cobrança em outra data quando há correspondência segura de valor.

### Resultados

- **Conciliado:** conta, data e valor localizados.
- **Conciliado - data diferente:** valor localizado em outra data.
- **Não cobrado:** o `CHECK#` existe, mas o valor líquido é zero.
- **Valor divergente:** cupom e Journal têm valores diferentes.
- **Lançado em outra data:** conta localizada, mas sem correspondência segura de valor.
- **Ausente na conta:** não existe `CHECK#` correspondente.
- **Journal não cobre a data:** o período selecionado não contém a emissão; gere outro Journal antes de concluir que existe erro.

### Exportação

- Excel: `Resumo` e `Conferencia`.
- PDF: resumo da conferência.

## 5.5 RPS de Serviços Prestados

### Objetivo

Confirmar se o RPS encerrado no Opera integrou no Fiscal do CMFlex e se a NFS-e foi emitida na Prefeitura com o mesmo valor.

### Arquivos

1. XML de encerramentos do Opera.
2. Relatório Fiscal do CMFlex.
3. Relatório da Prefeitura.

### Conferência executada

- compara RPS, data, tomador, situação e valor;
- identifica automaticamente o perfil de serviços mais compatível;
- reconhece os códigos configurados e as descrições adicionais do Opera;
- separa cancelamentos e documentos fora do período Fiscal.

### Resultados

- **Conciliado:** RPS e valor localizados nas três fontes.
- **Ausente no Fiscal:** encerrado no Opera sem integração no Fiscal.
- **Ausente na Prefeitura:** integrado no Fiscal sem NFS-e correspondente.
- **Ausente no Opera:** existe nas fontes fiscais, mas não no XML selecionado.
- **Valor divergente:** os valores das fontes não coincidem.
- **Fora do período do Fiscal:** a data do Opera não está coberta pelo relatório Fiscal; gere o Fiscal novamente com o período correto.
- **Cancelado:** cancelamento confirmado nas fontes correspondentes.
- **Inválido/irregular:** integrou, mas não possui situação fiscal válida.

### Serviços adicionais reconhecidos pelo nome

O reconhecimento não diferencia maiúsculas, acentos, espaços, barras e hífens.

- Diaria; Diaria Manual; Diaria No Show.
- Early Check In; Late Check Out.
- Cama Extra / Berco; Hospede Adicional.
- Ajuste Diaria Manual; Ajuste Early Check In; Ajuste Late Check Out.
- Ajuste Cama Extra / Berco; Ajuste Hospede Adicional.
- Serviço Abertura de Porta; Serviço Garçom Exclusivo.
- SPA-VINOPERFECT; SPA MASSAGEM; SPA-RESVERATROL LIFT; SPA-VINOPURE.
- SPA-TRATAMENTO CORPORAL RELAXANTE; SPA-CARMEL CORPORAL.
- SPA-CRUSCH CARBERNET CORPORAL; SPA MULTA CANCELAMENTO; SPA-VINOSOURCE.
- Ajuste Spa; Ajuste Shooting; Pulseira; Chave; Lanterna; Ajuste Lanterna.
- Rolha; Taxa Toalha; Ajuste Taxa Toalha; Shooting; Ajuste Rolha.
- Spa - Vinosource; Spa - Vinopure; Spa - Vinoperfct Facial.
- Spa - Vineactive Facial; Spa - Reveratrol Lift Facial.
- Spa - Relaxante Corporal; Spa - Fleur Vigne Corporal.
- Spa - Crushed Carbenet Corporal; Spa - Carmel Corporal.
- Ofuro; Multa Spa; Bolsa Personalizada Carmel.
- Ajustes individuais de todos os serviços de SPA, Ofuro, Multa Spa e Bolsa Personalizada Carmel.
- Diversos.

### Exportação

- Excel: `Resumo` e `Conferencia_RPS`.
- PDF: resumo do período e das situações.

## 5.6 Relatório de Notas de Débito

### Objetivo

Consolidar uma ou mais planilhas de notas de débito em um relatório único.

### Arquivos

- Uma ou mais planilhas de notas de débito em formato Excel.

### Resultado

Cada linha apresenta:

- hotel;
- comprador;
- número da nota;
- data de emissão;
- item;
- valor.

Use a pesquisa para localizar comprador, nota ou item. Confira se todas as planilhas selecionadas pertencem ao conjunto que deve ser consolidado.

### Exportação

- Excel detalhado.
- PDF de resumo.

## 5.7 Notas Fiscais de Entrada em Atraso

### Objetivo

Medir o prazo entre a emissão da nota de mercadoria e sua entrada no hotel, incluindo notas ainda não lançadas.

### Arquivos

1. Manifesto completo das notas.
2. Detalhe das notas recebidas/lançadas.

### Relacionamento

As notas são relacionadas por empresa e chave. Isso evita cruzar a mesma chave com empresa incorreta. Quando existem múltiplas entradas, é usada a primeira data válida.

### Prazos

| Local | Em dia | Alerta | Em atraso |
|---|---|---|---|
| Ceará | 0 a 5 dias | 6 a 10 dias | 11 dias ou mais |
| Outros estados | 0 a 19 dias | 20 a 30 dias | acima de 30 dias |

O intervalo considera data e horário e é arredondado para o dia inteiro mais próximo. Para nota sem entrada, o cálculo vai até a data atual.

### Colunas importantes

- **Lançamento:** `Lançada` ou `Não lançada`.
- **Dias:** tempo calculado.
- **Limite:** prazo usado para o estado.
- **Situação:** `Em dia`, `Alerta` ou `Em atraso`.

### Exportação

- Excel: `Resumo` e `Análise`.
- PDF: resumo por situação.

## 5.8 Conferência dos Cupons

### Objetivo

Comparar cupons NFC-e e notas NF-e entre Simphony, Fiscal do CMFlex e SEFAZ.

### Antes de começar

- escolha o hotel correto;
- use Cumbuco para relatórios anteriormente chamados Wind;
- gere as três fontes para o mesmo período.

### Arquivos

1. Relatório do Simphony.
2. Relatório Fiscal do CMFlex.
3. Consulta/relatório da SEFAZ.

### Resultados

- **Status Simphony — Aprovado:** documento válido no Simphony.
- **Status Simphony — Cancelado:** documento cancelado no Simphony.
- **Status Simphony — Ausente:** documento existe em outra fonte, mas não no Simphony.
- **Conciliado:** valores presentes e compatíveis.
- **Conciliado: cancelado:** cancelado no Simphony e Fiscal/SEFAZ vazios ou zerados.
- **Divergente: cancelamento:** cancelado no Simphony, mas Fiscal ou SEFAZ possui valor diferente de zero.
- **Ausente:** uma das integrações não foi localizada.
- **Divergente:** valores diferem mais de R$ 0,01.

Campos vazios no Fiscal e na SEFAZ são tratados como zero somente no cenário de cancelamento correspondente.

### Exportação

- Excel: `Resumo`, `Conciliação`, `Cupons_NFCe`, `Notas_NFe` e cópias das três planilhas-base.
- PDF: resumo com hotel.
- Exemplo: `cupons_cumbuco_resultado.xlsx`.

## 5.9 Notas de Serviços Tomados

### Objetivo

Conferir notas recebidas entre fontes externas, CAP, BPM, hotel e ISS retido.

### Arquivos mínimos

1. Um relatório CAP.
2. Um relatório Alterador/ISS retido.
3. Pelo menos uma fonte externa.

Na operação usual podem ser selecionados cinco arquivos: CAP, Portal Nacional, Prefeitura, ISS e BPM/relatório complementar. A quantidade de fontes externas varia conforme hotel e município.

### Fontes externas aceitas

- Portal Nacional;
- Prefeitura;
- relatórios municipais de NFS-e;
- CSV compatível;
- Excel ou relatório HTML fornecido com extensão `.xls`.

### Conferência executada

- consolida a mesma nota recebida de várias fontes;
- compara CNPJ, número e valor;
- remove o prefixo nacional do número quando necessário;
- quando CNPJ difere, só usa número e valor como contingência se existir correspondência única;
- verifica presença no CAP;
- verifica aprovação BPM;
- verifica hotel do CAP;
- compara valor bruto;
- verifica existência de fonte municipal;
- compara ISS somente quando retido/aplicável.

### Resultados

- **Conciliada:** todas as verificações aplicáveis estão corretas.
- **Ausente no CAP:** nota externa não encontrada no CAP.
- **Ausente nas fontes externas:** nota existe somente no CAP.
- **Ausente na Prefeitura:** existe no Portal, mas não existe em fonte municipal.
- **BPM pendente de aprovação:** nota no CAP ainda não aprovada.
- **Hotel divergente:** hotel do CAP não corresponde ao conjunto analisado.
- **Valor bruto divergente:** valor externo diferente do CAP.
- **Valor divergente entre fontes externas:** fontes externas discordam.
- **ISS retido ausente no CAP:** fonte externa informa retenção, mas CAP não possui o ISS.
- **ISS retido divergente:** valores de ISS diferem.
- **ISS retido ausente na prefeitura:** CAP possui ISS, mas a fonte municipal aplicável não informa retenção.

Regras importantes:

- nome do prestador com pequenas diferenças não deve, sozinho, gerar divergência;
- `ISS retido = Não` não gera cobrança de ISS no CAP;
- para relatórios de São Paulo, o ISS municipal não entra na comparação;
- uma nota somente no Portal continua pendente de Prefeitura.

### Exportação

- `Resumo`: indicadores e orientação.
- `Pendências`: motivo nas primeiras colunas.
- `Conciliadas`: registros sem divergência.
- `Base completa`: todas as colunas e fontes.
- PDF: resumo executivo.

## 5.10 Conferência do Contas a Receber

### Objetivo

Executar três conferências independentes entre Financeiro e Contabilidade.

### Arquivos — 6

1. Balancete por subcontas de clientes.
2. Posição por cliente.
3. Borderô de lançamento.
4. Razão Analítico de notas a faturar.
5. Agregados lançados.
6. Razão Analítico de comissões.

### Conferências

1. **Clientes:** Balancete x Posição por cliente.
2. **Notas a faturar:** valor absoluto do Borderô x débitos do Razão a faturar.
3. **Comissões:** valor absoluto dos Agregados x movimento absoluto do Razão.

Clientes relacionados são consolidados. Todas as variações identificadas como CVC formam um único total CVC antes da comparação. BRT e BWT permanecem separados.

### Resultados

- **Conciliado:** diferença de até R$ 0,01.
- **Divergente:** diferença superior a R$ 0,01.
- Cliente presente em apenas uma fonte aparece com zero na outra.

### Exportação

- Excel: `Resumo`, `Clientes` e cópias das seis planilhas-base.
- PDF: resumo das três conferências.
- Hotel no conteúdo e no nome, por exemplo `receber_cumbuco_resultado.xlsx`.

## 5.11 Conferência do Contas a Pagar

### Objetivo

Conferir fornecedores, adiantamentos e impostos entre Financeiro e Contabilidade.

### Arquivos — 8

1. Balancete de fornecedores.
2. Posição por fornecedor.
3. Balancete de adiantamentos.
4. Adiantamentos em aberto.
5. Agregado IRRF.
6. Agregado CSRF.
7. Agregado ISS.
8. Razão Analítico de impostos.

Os nomes dos três arquivos agregados precisam conter `IRRF`, `CSRF` ou `ISS`, pois os layouts são iguais e o nome identifica o imposto.

### Conferências

1. **Fornecedores:** Balancete x Posição por fornecedor.
2. **Adiantamentos:** Balancete x adiantamentos em aberto.
3. **IRRF:** agregado x movimentos credores da conta correspondente.
4. **CSRF:** agregado x PIS/COFINS/CSLL retidos.
5. **ISS:** agregado x ISS retido.

O HERMES consolida grupos comerciais conhecidos. Todas as variações CVC são somadas antes da comparação. BRT e BWT permanecem separados.

### Resultados

- **Conciliado:** diferença de até R$ 0,01.
- **Divergente:** diferença superior a R$ 0,01.
- Fornecedor/subconta presente somente em uma fonte aparece com zero na outra.

### Exportação

- Excel: `Resumo`, `Fornecedores e Adiantamentos` e cópias das oito planilhas-base.
- PDF: resumo das cinco conferências.
- Exemplo: `pagar_charme_resultado.xlsx`.

## 5.12 Custos da Mercadoria Vendida — CMV

### Objetivo

Executar separadamente a conferência de entradas de mercadoria e a conferência de saldo final dos estoques.

### Arquivos — 4

1. Documentos lançados por tipo de desembolso — CAP.
2. Razão Analítico das entradas de estoque.
3. Inventário físico e financeiro.
4. Razão Analítico dos estoques.

Os nomes precisam permitir a identificação dos quatro relatórios. Evite renomeá-los de forma que remova `Documentos Lançados`, `Razão Analítico Estoque AB`, `Inventário Físico` ou `Razão Analítico Estoques`.

### Conferência 1 — Entradas

- CAP x Contabilidade.
- Considera somente lançamentos cujo histórico contenha `Nota Fiscal`, `Mercadoria` e `Terceiros`.
- Não considera transferência, requisição, ajuste ou integração de custo.
- Grupos: Alimentos, Vinhos & Champanhe, Alcoólicos, Não Alcoólicos e Frigobar.

### Conferência 2 — Saldo final

- Inventário x saldo final da Contabilidade.
- As contas são relacionadas aos códigos de grupo do Inventário.
- Contas sem movimento continuam retornando o saldo final apresentado no Razão, mesmo quando o valor está em uma linha abaixo da descrição da conta.

As duas análises não devem ser somadas nem comparadas entre si.

### Resultados

- **Conciliado:** diferença de até R$ 0,01.
- **Divergente:** diferença superior a R$ 0,01.

### Exportação

- Excel: `Entradas`, `Saldo final` e cópias das quatro planilhas-base.
- PDF: seções separadas para as duas conferências.

## 6. Exportações

### 6.1 Excel

Use para auditoria detalhada, filtros, pesquisas e consulta das linhas.

Nas atividades abaixo, o Excel também inclui cópias consultáveis das planilhas analisadas:

- Conciliação de Receita;
- Conferência dos Cupons;
- Contas a Receber;
- Contas a Pagar;
- CMV.

As abas de análise ficam primeiro. As abas-base ficam depois e usam o nome do arquivo de origem. Elas preservam os valores para consulta, mas podem não reproduzir integralmente cores, objetos, macros e fórmulas do arquivo original.

### 6.2 PDF

Use para leitura rápida, aprovação ou compartilhamento do resumo. Quando for necessário investigar uma linha, use o Excel.

### 6.3 CSV

Disponível somente na Folha de Pagamento. É destinado à importação no CMFlex e deve ser validado antes do uso.

### 6.4 Nome do hotel

Quando a tela possui seletor de hotel, o hotel aparece no relatório e no nome do arquivo:

- `receita_<hotel>_resultado`;
- `diarias_<hotel>_resultado`;
- `cupons_<hotel>_resultado`;
- `receber_<hotel>_resultado`;
- `pagar_<hotel>_resultado`.

Se o nome estiver errado, refaça a análise escolhendo o hotel correto. Não renomeie apenas o arquivo, pois algumas regras também dependem da seleção.

## 7. Mensagens de erro

### Arquivo não reconhecido para esta conferência

Possíveis causas:

- atividade errada;
- cabeçalho alterado;
- arquivo incompleto;
- relatório salvo em formato diferente;
- foi selecionado um relatório que pertence a outra atividade.

Ação: gere novamente o relatório no sistema de origem e selecione todos os arquivos da atividade em uma única operação.

### Colunas obrigatórias não encontradas

O layout não contém os campos necessários. Não crie as colunas manualmente. Gere o relatório correto, sem remover títulos ou linhas de cabeçalho.

### Formato não reconhecido

Confirme a extensão e a origem. O HERMES aceita, conforme a atividade:

- `.xlsx`;
- `.xlsm`;
- `.xls`;
- `.xltx`;
- `.xltm`;
- `.csv`;
- `.xml`.

Um `.xls` que contém somente o índice de uma página web precisa ser exportado novamente como pasta de trabalho Excel completa.

### Faltam arquivos

Confira a quantidade indicada na tela. Não carregue os arquivos em etapas; faça uma única seleção com o conjunto completo.

### Dois arquivos foram identificados como o mesmo tipo

O conjunto contém relatório duplicado ou dois arquivos com o mesmo layout. Remova a duplicidade e confirme se não está faltando outro relatório.

### Período incompleto

Uma fonte não cobre todas as datas. Gere novamente o relatório com o mesmo intervalo das demais fontes.

### Nenhum registro utilizável

O arquivo pode estar vazio, filtrado, sem o período correto ou sem os tipos de lançamento reconhecidos.

### A tabela está vazia

1. escolha **Todos** em **Exibir**;
2. apague o conteúdo de **Buscar nos resultados**;
3. confirme os indicadores;
4. se o total for zero, gere novamente os arquivos.

### Erro ao exportar

- Feche o arquivo de destino se ele estiver aberto.
- Confirme permissão de gravação na pasta.
- No desktop, não mova nem exclua os arquivos-base antes de exportar.
- Tente salvar com outro nome.

### O conjunto excede o limite

O servidor limita quantidade e tamanho por operação. Não divida uma conferência que exige todos os relatórios. Solicite ao administrador a revisão do limite ou gere relatórios menores para um mesmo período válido.

## 8. Segurança e armazenamento

### 8.1 O que fica armazenado

- O HERMES não possui banco de dados nem histórico de conferências.
- Os arquivos enviados pelo navegador são temporários e removidos após a análise.
- Nas cinco atividades com abas-base, o último conjunto fica somente na memória da sessão até limpar, substituir, mudar de atividade ou encerrar/expirar a sessão.
- O resultado temporário da exportação é removido depois do download.
- O arquivo baixado permanece somente no local escolhido pelo usuário.

### 8.2 Boas práticas

- Use somente rede e computador autorizados.
- Não compartilhe relatórios por canais não aprovados.
- Salve o resultado em pasta com controle de acesso.
- Não deixe relatórios trabalhistas, fiscais ou de hóspedes em pastas públicas.
- Clique em **Limpar** ao concluir.
- Feche a sessão quando usar computador compartilhado.

## 9. Checklist final

Antes de considerar a atividade concluída:

1. O hotel está correto?
2. Todos os arquivos pertencem ao mesmo período?
3. A quantidade de arquivos reconhecidos está correta?
4. O resumo compara as fontes esperadas?
5. As linhas vermelhas foram verificadas?
6. As linhas amarelas foram interpretadas conforme a regra?
7. O nome do hotel no arquivo exportado está correto?
8. As abas-base necessárias estão presentes?
9. O arquivo foi salvo na pasta autorizada?
10. Na Folha, o CSV foi validado antes da importação?
11. O resultado foi limpo da tela após o uso?

## Suporte

Ao solicitar suporte, informe:

- nome da automação;
- hotel;
- período;
- nomes dos arquivos selecionados;
- mensagem completa exibida;
- número da nota, cupom, RPS, cliente, fornecedor ou conta usada como exemplo;
- relatório exportado pelo HERMES, quando permitido.

Não envie dados sensíveis fora dos canais autorizados.
