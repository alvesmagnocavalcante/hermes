# HERMES

## Manual do Usuário

**Painel de Automação de Planilhas**  
Versão: julho de 2026

---

## 1. Para que serve o HERMES

O HERMES compara planilhas de diferentes sistemas e mostra, de forma simples, o que está correto e o que precisa ser conferido.

Com ele, você pode:

- localizar valores ou documentos diferentes;
- identificar informações ausentes;
- consultar resultados na tela;
- gerar relatórios em Excel ou PDF;
- gerar o CSV da Folha de Pagamento para importação no CMFlex.

O HERMES não altera as planilhas selecionadas. Os arquivos originais continuam como estavam.

## 2. Como fazer uma conferência

1. Abra o HERMES pelo atalho.
2. Escolha a automação no menu esquerdo.
3. Confira o título da tela para ter certeza de que abriu a automação correta.
4. Escolha o **Hotel**, quando esse campo estiver disponível.
5. Clique em **Selecionar arquivos**.
6. Selecione todos os relatórios solicitados.
7. Aguarde a mensagem de conclusão.
8. Analise primeiro os resultados vermelhos e depois os amarelos.
9. Escolha **Excel**, **PDF** ou **CSV**, quando disponível.
10. Clique em **Exportar resultado** e escolha onde salvar.

Os arquivos podem ser selecionados em qualquer ordem. O HERMES identifica cada relatório pelo conteúdo, e não apenas pelo nome do arquivo.

## 3. Como entender a tela

### Arquivos reconhecidos

Após a seleção, o HERMES informa quantos arquivos conseguiu identificar. Se faltar algum, confirme se todos os relatórios da automação foram selecionados.

### Resumo

O resumo mostra os totais encontrados em cada sistema, a diferença e a situação da conferência.

Exemplo:

> Clientes: Balancete R$ 100.000,00 • Posição por cliente R$ 99.500,00 • Diferença R$ 500,00 • Divergente

### Cores

- **Verde:** informação correta, conciliada ou pronta.
- **Amarelo:** informação que merece atenção, mas pode não ser um erro.
- **Vermelho:** diferença, ausência ou dado incompleto que precisa ser conferido.

A legenda acima da tabela explica o significado das cores na automação selecionada.

### Tabela de resultados

A tabela mostra cada documento analisado e o motivo do resultado.

- Use **Exibir** para mostrar todos os registros, somente conciliados ou somente pendências.
- Use **Buscar nos resultados** para localizar nota, chave, fornecedor, cliente, hóspede, conta ou valor.
- Use as setas abaixo da tabela para mudar de página.

### Limpar

O botão **Limpar** apaga somente os dados exibidos na tela. Ele não exclui as planilhas originais nem os relatórios já exportados.

## 4. Automações disponíveis

### 4.1 Conciliação de Receita

**O que faz:** compara a receita da Contabilidade com a receita do Opera.

**Selecione:**

- Razão Analítico da Contabilidade;
- Journal de Receita do Opera.

**Como ler:** cada documento aparece como conciliado ou divergente. Uma divergência significa que o documento está ausente em uma fonte ou possui valores diferentes.

**Saídas:** Excel detalhado e PDF de resumo.

### 4.2 Conciliação da Receita de Diárias

**O que faz:** calcula a receita de diárias do Journal usando os códigos de transação definidos para cada hotel.

**Antes de começar:** escolha o hotel no campo exibido na tela.

**Selecione:**

- planilha de Códigos de Transação;
- Journal de Receita do Opera.

**Como ler:** a tabela mostra cada código, a quantidade de lançamentos e o valor encontrado. Um código sem movimento é uma informação, não necessariamente um erro.

**Saídas:** Excel detalhado e PDF de resumo.

### 4.3 Lançamento da Folha de Pagamento

**O que faz:** prepara os lançamentos contábeis da folha por centro de custo e remove linhas de total para evitar valores duplicados.

**Selecione:**

- resumo mensal da folha;
- relação de INSS;
- relação de FGTS;
- provisão de férias;
- provisão de 13º salário.

**Quando não houver férias:** selecione também a relação mensal de IRRF. Serão 6 arquivos no total.

**Quando houver férias:** selecione o recibo e o líquido de férias no lugar da relação mensal de IRRF. Serão 7 arquivos no total. O recibo e o líquido de férias devem ser selecionados juntos.

**Como ler:**

- **Pronto para importar:** o lançamento possui os dados necessários;
- **De/para incompleto:** alguma conta ou centro de custo precisa ser revisado;
- **Totalizadores removidos:** linhas de soma foram descartadas para não duplicar valores;
- **Líquido a pagar:** valor líquido identificado na folha.

**Saídas:** Excel para conferência, PDF de resumo e CSV para importação no CMFlex.

Os valores de plano odontológico e plano de saúde não são incluídos automaticamente. Eles devem ser preenchidos manualmente somente após o recebimento das notas fiscais.

Antes de importar o CSV, confira empresa, competência, contas, centros de custo e totais.

### 4.4 Cupons Emitidos x Conta do Hóspede

**O que faz:** verifica se os cupons emitidos foram lançados e cobrados na conta do hóspede.

**Selecione:**

- BI PDV;
- Journal do Opera;
- planilha de de/para.

**Como ler:** a tabela mostra cupom, data, hóspede, valor emitido, valor cobrado e o motivo de uma possível pendência.

Se aparecer **período incompleto**, o Journal não possui todas as datas dos cupons. Gere o Journal novamente com o período correto.

**Saídas:** Excel detalhado e PDF de resumo.

### 4.5 RPS de Serviços Prestados

**O que faz:** verifica se os RPS encerrados no Opera chegaram ao Fiscal do CMFlex e foram emitidos na Prefeitura.

**Selecione:**

- XML de encerramentos do Opera;
- relatório Fiscal do CMFlex;
- relatório da Prefeitura.

**Como ler:** a tabela compara o RPS e o valor nas três fontes. A coluna **Explicação** informa claramente o motivo de cada pendência.

Se aparecer **fora do período**, gere um relatório Fiscal que inclua a data indicada antes de considerar o registro como erro.

**Saídas:** Excel detalhado e PDF de resumo.

### 4.6 Relatório de Notas de Débito

**O que faz:** reúne as notas de débito em um relatório único e organizado.

**Selecione:** uma ou mais planilhas de notas de débito.

**Como ler:** confira hotel, comprador, número da nota, emissão, item e valor. Use a pesquisa para localizar uma nota específica.

**Saídas:** Excel e PDF.

### 4.7 Notas Fiscais de Entrada em Atraso

**O que faz:** calcula o tempo entre a emissão da nota de mercadoria e sua entrada no hotel.

**Selecione:**

- Manifesto de notas;
- Detalhe de notas recebidas.

**Prazos usados:**

- Ceará: atraso a partir de 11 dias;
- outros estados: atraso após 30 dias.

O cálculo considera os horários de emissão e entrada e arredonda o intervalo para o dia inteiro mais próximo.

**Como ler:** a tabela mostra emissão, entrada, quantidade de dias, prazo permitido e situação. Notas sem data ou não localizadas precisam ser verificadas nas planilhas originais.

**Saídas:** Excel detalhado e PDF de resumo.

### 4.8 Conferência dos Cupons

**O que faz:** verifica se cupons e notas do Simphony chegaram ao Fiscal do CMFlex e à SEFAZ.

**Antes de começar:** escolha o hotel no campo exibido na parte superior da tela. Use **Cumbuco** também para relatórios anteriormente identificados como Wind.

**Selecione:**

- relatório do Simphony;
- relatório Fiscal do CMFlex;
- relatório da SEFAZ.

**Como ler:** o HERMES compara chave, data e valor nas três fontes. A coluna **Status Simphony** informa se o documento está aprovado, cancelado ou ausente. A tabela diferencia **Cupom (NFC-e)** de **Nota (NF-e)** e informa onde o documento está ausente ou diferente. Um documento cancelado permanece visível e não entra no total do Simphony. Se os valores no Fiscal e na SEFAZ estiverem vazios ou zerados, o resultado será **Conciliado: cancelado**. Se alguma dessas fontes tiver valor diferente de zero, o resultado será **Divergente: cancelamento**.

**Saídas:** Excel detalhado e PDF de resumo com o hotel identificado. No Excel, cupons e notas também ficam em abas separadas. O nome do arquivo inclui o hotel, por exemplo: `cupons_cumbuco_resultado.xlsx`.

### 4.9 Notas de Serviços Tomados

**O que faz:** verifica se as notas recebidas estão no CAP, pertencem ao hotel correto, foram aprovadas no BPM e possuem o ISS correto.

**Selecione:**

- relatório do CAP;
- relatório do Portal Nacional;
- relatórios de Prefeitura ou ISS usados pelo hotel.

A quantidade de arquivos pode mudar conforme o hotel e o município.

**Como ler:**

- **Conciliada:** nota localizada com as informações corretas;
- **Ausente no CAP:** nota externa não encontrada no CAP;
- **BPM pendente:** nota existente no CAP, mas com aprovação do BPM ainda em andamento;
- **Hotel divergente:** hotel do CAP diferente do hotel analisado;
- **ISS divergente:** ISS da Prefeitura diferente do ISS do CAP.

O ISS só é considerado retido quando a fonte indicar retenção; se a Prefeitura informar **ISS retido = Não**, o valor do ISS não gera pendência no CAP. O prefixo nacional incluído em alguns números de NFS-e é removido automaticamente para localizar o número correspondente no CAP. Para notas de São Paulo, o ISS da planilha municipal não entra na comparação. Uma nota encontrada somente no Portal Nacional continua sinalizada como ausente na Prefeitura.

**Saídas:** PDF de resumo e Excel organizado em quatro abas:

- **Resumo:** totais e orientação de uso;
- **Pendências:** itens que precisam ser verificados, com o motivo nas primeiras colunas;
- **Conciliadas:** itens sem divergência;
- **Base completa:** todas as informações disponíveis.

### 4.10 Conferência do Contas a Receber

**O que faz:** confere clientes, notas a faturar e comissões entre Financeiro e Contabilidade.

**Antes de começar:** escolha o hotel no campo exibido na parte superior da tela. Use **Cumbuco** também para relatórios anteriormente identificados como Wind.

**Selecione:**

- Balancete por subcontas;
- Posição por cliente;
- Borderô de lançamento;
- Razão de notas a faturar;
- Agregados lançados;
- Razão de comissões.

**Como ler:** o resumo apresenta três verificações separadas:

- **Clientes:** Balancete x Posição por cliente;
- **Notas a faturar:** Borderô x Razão a faturar;
- **Comissões:** Agregados x Razão.

A tabela mostra quais clientes ou contas formam cada diferença.

**Para exportar:** escolha **Excel** e clique em **Exportar**. O hotel será mostrado no topo das abas **Resumo** e **Clientes**. O arquivo será salvo com um nome como `receber_cumbuco_resultado.xlsx`.

**Saídas:** Excel detalhado e PDF de resumo.

### 4.11 Conferência do Contas a Pagar

**O que faz:** confere fornecedores, adiantamentos e impostos entre Financeiro e Contabilidade.

**Antes de começar:** escolha o hotel no campo exibido na parte superior da tela. Use **Cumbuco** também para relatórios anteriormente identificados como Wind.

**Selecione:**

- Balancete de fornecedores;
- Posição por fornecedor;
- Adiantamentos em aberto;
- Balancete de adiantamentos;
- Agregados de IRRF, CSRF e ISS;
- Razão Analítico de impostos.

**Como ler:** o resumo mostra separadamente Fornecedores, Adiantamentos, IRRF, CSRF e ISS. Em cada grupo são exibidos os totais das duas fontes, a diferença e a situação.

**Para exportar:** escolha **Excel** e clique em **Exportar**. O hotel será mostrado no topo das abas **Resumo** e **Fornecedores e Adiantamentos**. O arquivo será salvo com um nome como `pagar_cumbuco_resultado.xlsx`.

**Saídas:** Excel detalhado e PDF de resumo.

### 4.12 Custos da Mercadoria Vendida

**O que faz:** confere entradas de mercadorias e saldos de estoque entre CAP, Inventário e Contabilidade.

**Selecione:**

- Documentos lançados por tipo de desembolso;
- Razão Analítico das entradas de estoque;
- Inventário físico e financeiro;
- Razão Analítico dos estoques.

**Como ler:**

- **Entradas:** compara CAP com Contabilidade;
- **Saldo final:** compara Inventário com Contabilidade.

A tabela mostra as contas ou grupos de estoque responsáveis pelas diferenças.

**Saídas:** Excel detalhado e PDF de resumo.

## 5. Qual arquivo exportar

### Excel

Use quando precisar analisar todos os registros, aplicar filtros ou guardar o resultado detalhado.

### PDF

Use quando precisar consultar ou compartilhar um resumo da conferência.

### CSV

Disponível na Folha de Pagamento. É o arquivo preparado para importação no CMFlex.

Não abra e salve novamente o CSV no Excel antes da importação. Isso pode alterar separadores e remover zeros dos códigos.

## 6. Se aparecer uma mensagem de erro

### Arquivo não reconhecido

Confirme se você abriu a automação correta e selecionou os relatórios pedidos naquela tela.

### Está faltando um arquivo

Clique novamente em **Selecionar arquivos** e escolha todos os relatórios necessários na mesma seleção.

### Coluna ou formato não reconhecido

Gere novamente o relatório no sistema de origem. Não apague cabeçalhos, não altere colunas e não converta o arquivo manualmente.

O HERMES aceita planilhas Excel nos formatos `.xlsx`, `.xlsm`, `.xls`, `.xltx` e `.xltm`, inclusive relatórios antigos ou relatórios HTML fornecidos pelo sistema com extensão `.xls`.

### Período incompleto

Uma das planilhas não cobre todo o período das demais. Gere novamente o relatório com as datas necessárias.

### A tabela está vazia

Selecione **Todos** no filtro **Exibir** e apague o texto do campo de busca.

### Não foi possível exportar

Feche o arquivo de destino caso ele esteja aberto no Excel ou em outro programa. Depois exporte novamente.

## 7. Cuidados antes de concluir

- Confirme o hotel, a empresa e o período dos relatórios.
- Use relatórios gerados para o mesmo período.
- Não altere as planilhas antes de carregá-las no HERMES.
- Verifique primeiro as linhas vermelhas.
- Leia a explicação apresentada na última coluna da tabela.
- Em Contas a Receber e Contas a Pagar, confirme o hotel no nome do arquivo exportado.
- Na Folha de Pagamento, confira os totais antes de importar o CSV.

## 8. Segurança das informações

O processamento acontece no computador onde o HERMES está aberto. Um novo arquivo só é criado quando você solicita uma exportação.

Não compartilhe planilhas financeiras, fiscais ou dados de hóspedes, fornecedores e colaboradores fora dos canais autorizados pela empresa.
