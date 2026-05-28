# Arquitetura dos scripts

Este repositorio nao e um projeto de software generico.

O objetivo e bem especifico: pegar planilhas Excel existentes e gerar dashboards dentro dessas planilhas, usando scripts feitos para cada caso.

Cada dashboard pode ter regras, formulas, abas, graficos e tratamentos proprios. Por isso, a organizacao deve priorizar clareza e separacao por planilha, nao uma arquitetura grande demais.

## Objetivo

- Ler uma planilha Excel especifica.
- Criar ou recriar uma aba de dashboard dentro dela.
- Montar tabelas auxiliares, formulas, indicadores e graficos.
- Gerar uma nova versao da planilha em `outputs/`.
- Manter cada automacao facil de encontrar, rodar e alterar.

## Referencia visual

A primeira planilha trabalhada, `CQ Produto Acabado - Procytrat 2026`, deve ser usada como referencia de aparencia para os proximos dashboards.

Isso significa que os proximos dashboards devem seguir o mesmo padrao visual geral:

- layout organizado em paineis;
- cards de indicadores no topo;
- paleta de cores e estilo visual semelhantes;
- tabelas de apoio com visual consistente;
- graficos posicionados e formatados no mesmo estilo;
- dashboard como primeira aba ou em posicao de destaque, quando fizer sentido.

O que muda de uma planilha para outra deve ser principalmente a parte de dados: abas de origem, nomes dos campos, formulas, metricas, tabelas auxiliares, filtros e graficos especificos de cada arquivo.

Em outras palavras: manter o mesmo "padrãozinho" visual, adaptando a inteligencia do dashboard aos dados de cada planilha.

## Estrutura obrigatoria

Cada planilha/dashboard deve ficar em sua propria pasta dentro de `dashboards/`:

```text
dashboards/
  Nome_Da_Planilha/
    01_build_dashboard.py
    02_strip_charts.py
    03_recreate_native_charts.ps1
    04_cleanup_final.ps1
    config.json
    run.ps1
```

Use scripts numerados para manter clara a ordem de execucao.

## Papel de cada etapa

`01_build_dashboard.py`

Cria a aba do dashboard, tabelas, formulas, estilos e graficos iniciais.

`02_strip_charts.py`

Remove graficos criados pelo `openpyxl` quando for necessario recria-los como graficos nativos do Excel.

`03_recreate_native_charts.ps1`

Abre o arquivo pelo Excel/COM e recria graficos nativos, especialmente quando eles precisam ficar editaveis/moveis no Excel.

`04_cleanup_final.ps1`

Faz limpeza final, recalculo, validacoes e ajustes finais do arquivo.

`config.json`

Documenta configuracoes da planilha: nome, arquivo de entrada, aba do dashboard, ano padrao, saida esperada e parametros principais.

`run.ps1`

Executa a sequencia completa daquele dashboard.

## Regras para manter organizado

- Nao misture scripts de planilhas diferentes na raiz do repositorio.
- Nao crie uma abstracao compartilhada antes de existir repeticao real.
- Coloque em `comum/` apenas funcoes que forem reaproveitadas por mais de um dashboard.
- Mantenha regras especificas dentro da pasta da propria planilha.
- Gere arquivos sempre dentro de `outputs/Nome_Da_Planilha/`.
- Ao criar uma nova planilha, copie a estrutura de uma pasta existente e adapte os scripts.

## Como adicionar uma nova planilha

1. Crie uma pasta em `dashboards/` com o nome da planilha.
2. Adicione os scripts numerados seguindo o mesmo padrao.
3. Crie o `config.json`.
4. Ajuste o `run.ps1` da pasta.
5. Rode o dashboard pela raiz usando:

```powershell
.\gerar_dashboard.ps1 -Dashboard Nome_Da_Planilha
```

## O que evitar

- Transformar estes scripts em um framework generico.
- Criar muitas camadas de `src/`, `features/`, classes e interfaces sem necessidade.
- Colocar varios dashboards dentro de um unico script gigante.
- Salvar arquivos gerados ao lado dos scripts.

Este repositorio deve continuar simples: uma colecao organizada de automacoes especificas para gerar dashboards em planilhas especificas.
