# HERMES

## Guia do Usuário

**Painel de Automação de Planilhas**  
Versão do guia: julho de 2026

---

## 1. Apresentação

O HERMES reúne conferências financeiras, fiscais, contábeis e operacionais em um único painel. O sistema compara relatórios de diferentes fontes, destaca valores coincidentes e aponta registros que precisam de análise.

O objetivo é reduzir o trabalho manual, facilitar a localização de diferenças e produzir arquivos organizados para conferência, registro ou importação.

O HERMES não altera os arquivos selecionados. As planilhas originais permanecem preservadas.

## 2. Como acessar e navegar

1. Abra o HERMES pelo atalho disponibilizado ou execute o aplicativo conforme a orientação da equipe responsável.
2. O menu de automações aparece na lateral esquerda.
3. Use o botão de menu abaixo do nome **HERMES** para abrir ou recolher a lista.
4. Quando o menu estiver recolhido, passe o mouse sobre um ícone para visualizar o nome da automação.
5. Selecione a automação desejada antes de carregar os relatórios.

O título exibido no alto da tela confirma qual conferência está ativa. Verifique esse título antes de selecionar os arquivos.

## 3. Fluxo padrão de utilização

O procedimento abaixo é utilizado na maioria das automações:

1. Gere os relatórios necessários nos sistemas de origem.
2. Evite alterar os nomes, colunas ou conteúdo dos arquivos antes da análise.
3. Abra a automação correspondente no menu lateral.
4. Leia a seção **Arquivos necessários**, exibida no alto da tela.
5. Clique em **Selecionar arquivos**.
6. Selecione todos os relatórios solicitados na mesma janela.
7. Aguarde a conclusão. O andamento é informado no rodapé.
8. Confira os indicadores, o resumo comparativo e a tabela.
9. Use o filtro e a pesquisa para localizar registros específicos.
10. Escolha o formato e clique em **Exportar resultado**.

Os arquivos podem ser selecionados em qualquer ordem. Quando um relatório não pertence à automação aberta, o HERMES apresenta uma orientação sobre os arquivos esperados.

## 4. Como interpretar a tela

### Arquivos reconhecidos

Após a leitura, o painel informa quantos arquivos foram reconhecidos. Se a quantidade for diferente da esperada, revise a seleção antes de continuar.

### Resumo comparativo

Nas conferências entre duas ou mais fontes, o resumo apresenta:

- o nome da verificação;
- o valor encontrado em cada fonte;
- a diferença entre os valores;
- o resultado da comparação.

Exemplo:

> Clientes: Balancete R$ 7.695.835,38 • Posição por cliente R$ 7.742.048,90 • Diferença -R$ 46.213,52 • Divergente

### Indicadores

Os cards mostram a quantidade total analisada e a distribuição dos resultados. Os nomes dos indicadores mudam conforme a finalidade da automação.

### Cores

Consulte sempre a legenda exibida acima da tabela. De forma geral:

- **Verde:** registro conciliado, correto ou pronto para utilização.
- **Amarelo:** informação complementar, cancelamento ou período incompleto.
- **Vermelho:** diferença, ausência ou dado incompleto que exige análise.

### Filtros, pesquisa e paginação

- Use **Exibir** para alternar entre todos os registros, conciliados e pendências.
- Use **Buscar nos resultados** para localizar documento, chave, fornecedor, cliente, hóspede, conta ou situação.
- Passe o mouse sobre um texto abreviado para visualizar o conteúdo completo.
- Use as setas abaixo da tabela para navegar entre as páginas.

### Limpar

O botão **Limpar** remove os resultados da tela e permite iniciar uma nova conferência. Os arquivos originais e os arquivos já exportados não são apagados.

## 5. Automações disponíveis

### 5.1 Conciliação de Receita

**Finalidade**  
Verificar se a receita registrada na Contabilidade coincide com os lançamentos de receita do Opera.

**Arquivos necessários**

- Razão Analítico da Contabilidade.
- Journal de Receita do Opera.

**Como interpretar**  
O resumo compara o total da Contabilidade com o total do Opera. Na tabela, cada documento é classificado como conciliado ou divergente. Uma diferença pode indicar ausência em uma das fontes ou valores distintos para o mesmo documento.

**Exportações**  
Excel detalhado e PDF de resumo.

### 5.2 Conciliação da Receita de Diárias

**Finalidade**  
Totalizar a receita de diárias do Journal utilizando somente os códigos de transação definidos para o hotel.

**Antes de selecionar os arquivos**  
Escolha o hotel no campo exibido ao lado do botão de seleção.

**Arquivos necessários**

- Planilha de Códigos de transação.
- Journal Opera - Receita.

O arquivo **BI PDV** não pertence a esta automação.

**Como interpretar**  
A tabela apresenta os códigos considerados como diária ou diária média, a quantidade de lançamentos e o valor encontrado no Journal. Códigos sem movimento aparecem como informação, não necessariamente como erro.

**Exportações**  
Excel detalhado e PDF de resumo.

### 5.3 Lançamento da Folha de Pagamento

**Finalidade**  
Transformar os relatórios do Departamento Pessoal em lançamentos contábeis por centro de custo, prontos para importação no CMFlex.

**Arquivos necessários**

- Resumo mensal da folha.
- Relação de INSS.
- Relação de FGTS.
- Recibo de férias.
- Líquido de férias.
- Provisão de férias.
- Provisão de 13º salário.

A planilha modelo do departamento pode ser selecionada junto com os relatórios quando for necessário atualizar o relacionamento de contas, eventos e centros de custo.

**Como interpretar**

- **Pronto para importar:** lançamento com os dados necessários preenchidos.
- **De/para incompleto:** lançamento que precisa de revisão antes da importação.
- **Totalizadores removidos:** linhas de total descartadas para evitar valores duplicados.
- **Líquido a pagar:** valor líquido identificado no resumo da folha.

**Exportações**

- **Excel:** arquivo de conferência com resumo e saídas separadas.
- **CSV:** arquivo no formato aceito pelo CMFlex.
- **PDF:** resumo da folha processada.

Antes de importar o CSV, confira a competência, as contas, os centros de custo e o total dos lançamentos.

### 5.4 Cupons Emitidos x Conta do Hóspede

**Finalidade**  
Verificar se os cupons emitidos pelo ponto de venda constam no Journal e foram cobrados na conta do hóspede.

**Arquivos necessários**

- BI PDV.
- Journal do Opera.
- Planilha de de/para.

**Como interpretar**  
O resumo compara o total do BI/PDV com o total localizado nas contas. A tabela informa cupom, data, quarto, hóspede, valor emitido, valor cobrado e motivo da pendência.

Um registro marcado como período incompleto indica que o Journal selecionado não cobre a data do cupom. Nesse caso, gere o Journal para o período correto antes de confirmar uma divergência.

**Exportações**  
Excel detalhado e PDF de resumo.

### 5.5 RPS de Serviços Prestados

**Finalidade**  
Verificar se os RPS encerrados no Opera integraram no Fiscal do CMFlex e foram emitidos na Prefeitura.

**Arquivos necessários**

- XML de encerramentos do Opera.
- Planilha Fiscal do CMFlex.
- Planilha da Prefeitura.

**Como interpretar**  
O painel compara Opera × Fiscal e Fiscal × Prefeitura. A tabela apresenta o RPS, a data, o tomador, os valores das três fontes, a NFS-e e a explicação do resultado.

Registros posteriores à última data do relatório Fiscal aparecem como **fora do período**. Eles devem ser reavaliados com um relatório Fiscal atualizado antes de serem tratados como erro de integração.

**Exportações**  
Excel detalhado e PDF com resumo e registros que exigem atenção.

### 5.6 Relatório de Notas de Débito

**Finalidade**  
Consolidar informações de notas de débito em um único relatório.

**Arquivos necessários**  
Uma ou mais planilhas de notas de débito.

**Como interpretar**  
O resultado reúne hotel, comprador, número da nota, emissão, item e valor. Utilize a pesquisa para localizar uma nota ou comprador específico.

**Exportações**  
Excel detalhado e PDF.

### 5.7 Notas Fiscais de Entrada em Atraso

**Finalidade**  
Conferir o tempo entre a emissão da nota fiscal de mercadoria e a entrada da nota no hotel.

**Arquivos necessários**

- Manifesto de notas.
- Detalhe de notas recebidas.

**Regras de prazo**

- Notas do Ceará: atraso a partir de 11 dias.
- Notas de outros estados: atraso a partir de 30 dias.

**Como interpretar**  
A tabela informa chave, empresa, fornecedor, estado, emissão, entrada, dias transcorridos, limite e situação. Notas não localizadas ou sem datas completas devem ser verificadas nas fontes originais.

**Exportações**  
Excel detalhado e PDF de resumo.

### 5.8 Conferência dos Cupons

**Finalidade**  
Verificar se os documentos do Simphony integraram no Fiscal do CMFlex e constam na SEFAZ.

**Arquivos necessários**

- Simphony.
- Fiscal do CMFlex.
- SEFAZ.

**Como interpretar**  
O painel compara Simphony × Fiscal e Fiscal × SEFAZ. A análise utiliza a chave fiscal, a data e o valor e distingue cupons NFC-e de notas NF-e.

O Excel exportado possui abas específicas:

- **Conciliação:** todos os documentos.
- **Cupons_NFCe:** somente cupons.
- **Notas_NFe:** somente notas.

**Exportações**  
Excel detalhado e PDF de resumo.

### 5.9 Notas de Serviços Tomados

**Finalidade**  
Conferir notas de serviços recebidas, sua escrituração no CAP, o hotel correspondente, a aprovação no BPM e o ISS retido.

**Arquivos necessários**

- Relatório do CAP.
- Portal Nacional.
- Relatórios de Prefeitura e ISS aplicáveis ao hotel analisado.

A quantidade de arquivos pode variar conforme o município e o processo realizado.

**Como interpretar**

- **Conciliada:** documento localizado com as informações esperadas.
- **Ausente no CAP:** nota encontrada em fonte externa, mas não localizada no CAP.
- **Não escriturada:** BPM ainda não aprovado.
- **Hotel divergente:** o hotel indicado no CAP não corresponde ao processo analisado.
- **ISS divergente:** valor da Prefeitura diferente do valor registrado no CAP.

O prestador pode aparecer com o CNPJ junto ao nome; o sistema trata essa informação durante a conferência.

**Exportações**  
Excel detalhado e PDF de resumo.

### 5.10 Conferência do Contas a Receber

**Finalidade**  
Verificar clientes a receber, notas a faturar e comissões comparando os relatórios operacionais, financeiros e contábeis.

**Arquivos necessários**

- Balancete por subcontas.
- Posição por cliente.
- Borderô de lançamento.
- Razão de notas a faturar.
- Agregados lançados.
- Razão de comissões.

**Como interpretar**  
O resumo apresenta três comparações:

- **Clientes:** Balancete × Posição por cliente.
- **Notas a faturar:** Borderô × Razão a faturar.
- **Comissões:** Agregados lançados × Razão.

A tabela permite verificar individualmente os clientes e as contas que formam a diferença.

**Exportações**  
Excel detalhado e PDF de resumo.

### 5.11 Conferência do Contas a Pagar

**Finalidade**  
Verificar se fornecedores, adiantamentos e impostos em aberto coincidem entre Financeiro e Contabilidade.

**Arquivos necessários**

- Balancete com subcontas de fornecedores.
- Posição por fornecedor.
- Adiantamentos em aberto.
- Balancete com subcontas de adiantamentos.
- Agregados lançados de IRRF.
- Agregados lançados de CSRF.
- Agregados lançados de ISS.
- Razão Analítico de impostos.

**Como interpretar**  
O resumo apresenta separadamente Fornecedores, Adiantamentos, IRRF, CSRF e ISS. Para cada grupo são exibidos o total do Financeiro, o total da Contabilidade, a diferença e o resultado.

A tabela detalha os fornecedores e subcontas que compõem cada conferência.

**Exportações**  
Excel detalhado e PDF de resumo.

### 5.12 Custos da Mercadoria Vendida

**Finalidade**  
Conferir as entradas de mercadorias e o saldo final dos estoques entre CAP, Inventário e Contabilidade.

**Arquivos necessários**

- Documentos lançados por tipo de desembolso.
- Razão Analítico das entradas de estoque.
- Inventário físico e financeiro.
- Razão Analítico dos estoques.

**Como interpretar**  
O resumo apresenta duas comparações:

- **Entradas:** CAP × Contabilidade.
- **Saldo final:** Inventário × Contabilidade.

A tabela detalha as contas ou grupos de estoque, os valores de cada fonte e a diferença encontrada.

**Exportações**  
Excel detalhado e PDF de resumo.

## 6. Formatos de exportação

### Excel

Indicado para conferência detalhada, aplicação de filtros, pesquisa, ordenação e arquivamento dos resultados.

### PDF

Indicado para consulta e compartilhamento do resumo da análise. Os PDFs apresentam somente títulos, informações essenciais, indicadores e tabelas.

### CSV

Disponível na Folha de Pagamento para importação no CMFlex. O arquivo deve ser utilizado sem alterações que modifiquem a ordem das colunas, os separadores ou os zeros dos códigos.

## 7. Mensagens e situações comuns

### Arquivo incompatível

O relatório selecionado não pertence à automação aberta. Confira o título da tela e a seção **Arquivos necessários**.

### Quantidade incorreta de arquivos

Selecione todos os relatórios solicitados ao mesmo tempo. Remova arquivos de outras atividades da seleção.

### Coluna não encontrada ou formato não reconhecido

Gere novamente o relatório no sistema de origem, sem excluir cabeçalhos, alterar colunas ou salvar em outro modelo.

### Período incompleto

Um dos relatórios não cobre todas as datas encontradas nas demais fontes. Gere novamente o arquivo com o período completo antes de confirmar a pendência.

### Nenhum registro exibido

Verifique o filtro **Exibir**, limpe o campo de pesquisa e confirme se a análise foi concluída.

### Não foi possível exportar

Confirme se o arquivo de destino não está aberto em outro programa e se a pasta permite gravação. Tente exportar novamente com outro nome.

## 8. Boas práticas

- Confirme a empresa, o hotel e o período antes da análise.
- Gere os relatórios na mesma data e para o mesmo intervalo.
- Mantenha os arquivos originais sem edições manuais.
- Analise primeiro as linhas vermelhas e depois as amarelas.
- Não considere uma ausência definitiva quando o período de uma fonte estiver incompleto.
- Antes de importar um CSV, confira os totais e os lançamentos indicados como prontos.
- Salve os resultados com nome que identifique empresa, competência e tipo de conferência.

## 9. Privacidade e segurança

O processamento é realizado localmente no computador em que o HERMES está sendo executado. O sistema lê os relatórios selecionados e cria um novo arquivo somente quando o usuário solicita uma exportação.

Não compartilhe relatórios financeiros, fiscais, dados de hóspedes, fornecedores ou colaboradores fora dos canais autorizados pela empresa.

