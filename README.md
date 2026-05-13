# SADPat — Sistema Auxiliar de Desfazimento Patrimonial - v15

Correção v15: manutenção da segregação das saídas por doação na aba **Registro SIAFI Web** e aprimoramento visual da interface inicial, com cabeçalho redistribuído, ampliação controlada da logo institucional e substituição das abas por botões de navegação com ícones. A estrutura, os cálculos e o layout da aba **Relatório Sintético** foram preservados.

Também foram mantidas as correções anteriores: parser de depreciação em modo estrito, desconsideração das saídas por transferência, manutenção dos resultados após processamento, geração do Relatório Sintético, Registro SIAFI Web, Excel e PDF.

# SADPat — Sistema Auxiliar de Desfazimento Patrimonial

Aplicativo Streamlit para cruzamento do **Relatório Sintético Patrimonial de Saídas** com o **Relatório de Depreciação Acumulada**, gerando relatório auxiliar para registro de baixas/desfazimento de bens móveis, com saídas em Excel e PDF.

## Principais melhorias desta versão

- Código reorganizado em camadas de leitura, tratamento, apuração, análise de consistência, exibição e exportação.
- Cache de leitura dos PDFs para reduzir reprocessamentos em reruns do Streamlit.
- Resultados preservados na tela após a geração dos arquivos Excel e PDF.
- Limpeza automática dos resultados somente quando os PDFs carregados são removidos ou substituídos.
- Layout institucional revisado, com cabeçalho mais equilibrado, cards de indicadores, abas e botões padronizados.
- Aba **Análise de Consistência** no aplicativo e no Excel, com verificações de duplicidade, mapeamento PCASP, descrições contaminadas por valores e coerência aritmética da apuração.
- Relatório Sintético padronizado com linha de total, valores no padrão brasileiro e descrição do grupo em letras maiúsculas.
- Exportação Excel com abas: **Relatório Sintético**, **Registro SIAFI Web**, **Memória de Cálculo**, **Análise de Consistência**, **Logs** e **Metadados**.
- Aba **Registro SIAFI Web** com segregação das saídas por doação na situação **IMB037** e dedução correspondente em **IMB025**.
- Exportação PDF com cabeçalho e rodapé institucionais, tabela compatível com o layout dos relatórios de entrada e rodapé da fonte SICPAT.

## Estrutura do projeto

```text
.
├── app.py
├── modelo_grupo_x_pcasp.xlsx
├── proplan_ufmg.jpg
├── requirements.txt
└── src
    ├── models.py
    ├── parsers.py
    ├── quality.py
    ├── reporting.py
    └── utils.py
```

## Como executar localmente

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Entradas esperadas

Enviar dois arquivos PDF da mesma competência:

1. Relatório Sintético Patrimonial de Saídas.
2. Relatório de Depreciação Acumulada.

A base `modelo_grupo_x_pcasp.xlsx` deve permanecer na raiz do projeto para permitir o relacionamento entre grupo patrimonial e conta contábil PCASP.

## Saídas geradas

- Arquivo Excel completo, com relatório final e abas auxiliares.
- PDF do Relatório Sintético.
- Visualização no app do Relatório Sintético, Registro SIAFI Web, Análise de Consistência, Memória de Cálculo, Logs e Metadados.

## Validação com PDFs reais - UG 153261 / abr-2026

Esta versão foi ajustada e validada com os arquivos reais:

- `153261_RMB_ABRIL_2026_UFMG.pdf`
- `153261_Depreciacao_acumulada_ABR_2026(1).pdf`

A validação confirmou a leitura da competência `abr/2026`, UG `153261 - HOSPITAL DAS CLÍNICAS`, e a compatibilização do total de saídas do Relatório Sintético Patrimonial com o total apurado no Relatório Sintético: R$ 1.468.278,10.

Ajuste técnico relevante: o parser do Relatório Sintético Patrimonial passou a priorizar o quadro sintético geral, ignorando os quadros auxiliares de entradas e saídas presentes no mesmo PDF. Esse ajuste evita dupla contagem e corrige a perda da parcela de transferências de saídas, especialmente observada no grupo 8.

## Regra de apuração v11 — saídas por transferência

A partir da v11, o aplicativo utiliza o quadro **Relatório Sintético Patrimonial de Saídas** para considerar apenas a coluna **SAÍDAS**, desconsiderando a coluna **TRANSFERE SAÍDAS**. Essa regra evita duplicidade informacional, pois o SICPAT possui relatório próprio de transferências com o valor dos bens transferidos e a depreciação acumulada associada.

No processamento, o aplicativo mantém a rastreabilidade do total original do SICPAT, registra em logs o montante de transferências desconsiderado e cruza a base das demais saídas com a linha **SAÍDAS (BAIXAS)** do Relatório de Depreciação Acumulada.


## Ajuste v11

- Corrigida a extração da linha `SAÍDAS (BAIXAS)` no relatório de depreciação acumulada.
- O parser agora prioriza exclusivamente a linha rotulada `SAÍDAS (BAIXAS)` e, no fallback, considera apenas linhas monetárias mensais completas.
- Cabeçalhos de grupo repetidos no final de página, sem dados da tabela, são ignorados para evitar duplicidade e captura indevida de números de página, datas ou exercício.
- Corrige o caso do grupo 52 em março/2026, em que o valor correto da depreciação associada à baixa é zero.

## Ajuste v13 - formatação numérica no Excel

A geração do arquivo Excel foi ajustada para aplicar formatação numérica brasileira às colunas monetárias e quantitativas dos relatórios e quadros de apuração. Os valores permanecem como números, preservando cálculos, filtros e totais, mas passam a ser exibidos com separador de milhar por ponto, separador decimal por vírgula e duas casas decimais, por exemplo: 1.234.567,80.

A correção abrange, entre outras, as colunas de Saídas/Grupo, Dep. Acumulada, Vlr. Liq. Contábil, Situação IMB010, Situação IMB025, saldos, totais e demais campos numéricos aplicáveis nas abas do Excel.


## Ajuste v14 — segregação de doações no Registro SIAFI Web

- O parser do Relatório Sintético Patrimonial de Saídas passou a armazenar a coluna **Doação** quando o quadro detalhado de saídas é identificado.
- A aba **Registro SIAFI Web** passa a exibir a coluna **Situação IMB037**, com o valor de doação por grupo/conta contábil.
- A coluna **Situação IMB025** passa a ser calculada como **Saídas/Grupo − Doação**, sem permitir valores negativos.
- Quando o valor de doação for superior ao total de saídas do grupo, o processamento continua, a IMB025 é limitada a zero e o fato é registrado nos logs.
- A aba **Relatório Sintético** permanece baseada na lógica aprovada de Saídas/Grupo, Dep. Acumulada e Vlr. Liq. Contábil, sem inclusão de novas colunas ou alteração visual.
- Foi removida uma chamada duplicada de competência no parser de depreciação e a chave de cache do parser foi atualizada para evitar reaproveitamento indevido de leituras antigas no Streamlit.

## Ajuste v15 — refinamento visual da interface

- Ampliação discreta da logo institucional com melhor enquadramento no cabeçalho.
- Redistribuição do cabeçalho para melhorar alinhamento, espaçamento e aproveitamento da largura da tela.
- Conversão das abas de resultados em botões de navegação horizontais com ícones ilustrativos.
- Inclusão de ícones discretos em pontos-chave da interface, como upload, orientações, geração e downloads.
- Preservação integral das regras de processamento, cálculos, validações e geração dos relatórios.

## Ajuste v16 — Relatório de Baixas em PDF multipágina

- O PDF exportado passa a se chamar **Relatório de Baixas**.
- O botão de download foi renomeado para **Exportar Relatório de Baixas em PDF**.
- O nome sugerido do arquivo exportado passa a iniciar por `relatorio_de_baixas`.
- O cabeçalho institucional do PDF foi padronizado com:
  - UNIVERSIDADE FEDERAL DE MINAS GERAIS;
  - SISTEMA AUXILIAR DE DESFAZIMENTO PATRIMONIAL - SADPat;
  - HOSPITAL DAS CLÍNICAS ou a unidade identificada no relatório;
  - RELATÓRIO AUXILIAR PARA REGISTRO CONTÁBIL DE BAIXAS DE BENS MÓVEIS.
- O PDF passa a consolidar os quadros **Relatório Sintético**, **Registro SIAFI Web** e **Logs e Metadados**, em páginas próprias.
- A regra de cálculo, a extração, as validações, o Excel e as estruturas das abas foram preservados.

## Ajuste v17 — Relatório Sintético e alinhamento de Logs e Metadados

- A nomenclatura do primeiro quadro foi padronizada como **Relatório Sintético** nos pontos de exibição ao usuário.
- O quadro 1 do **Relatório de Baixas em PDF** passou a ser identificado como **Relatório Sintético**.
- No quadro **Logs e Metadados** do PDF, os dados das colunas **Campo** e **Valor** passaram a ser alinhados à direita, preservando o cabeçalho centralizado.
- As regras de cálculo, extração, validação, Excel, Registro SIAFI Web e resultados gerados foram preservados.

## Ajuste v18 — revisão conservadora, nomenclatura e alinhamento

- Revisadas as mensagens internas da análise de consistência para substituir referências residuais à nomenclatura anterior por **Relatório Sintético**.
- Reforçado o alinhamento à direita do corpo das segunda e terceira colunas do quadro **Logs e Metadados** no PDF.
- Aplicado alinhamento equivalente na visualização do aplicativo e nas abas auxiliares **Logs** e **Metadados** do Excel, preservando o cabeçalho centralizado.
- Removidos artefatos temporários de cache/backup do pacote de distribuição para reduzir redundância e evitar conflito com versões anteriores.
- Mantidas as regras de extração, cálculo, validação, Registro SIAFI Web, Excel e PDF já aprovadas.
