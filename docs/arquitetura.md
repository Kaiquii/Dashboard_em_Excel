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

Em outras palavras: manter o mesmo "padraozinho" visual, adaptando a inteligencia do dashboard aos dados de cada planilha.

## Regras obrigatorias dos dashboards

Estas regras existem por causa das dificuldades que tivemos no dashboard `CQ Materia Prima - Fornecedores 2026`. Elas devem ser seguidas em todos os proximos dashboards:

1. Graficos nao podem ser fixos, imagens ou desenhos travados. Eles devem ser graficos padrao/nativos do Excel e devem poder ser movidos manualmente dentro da planilha.
2. O dashboard nao pode depender de dados fixos colados pelo Python. Cards, tabelas auxiliares, rankings, indicadores e graficos devem usar formulas do Excel ligadas as abas de origem, para atualizar quando os dados das planilhas forem alterados.
3. O campo de ano deve ser um filtro visivel no topo do dashboard, e todo o dashboard deve respeitar esse ano selecionado: indicadores, tabelas, rankings, vencimentos e graficos.

## Regra para graficos

Todo grafico gerado para uma planilha deve atender estes pontos:

- deve ser um grafico padrao/nativo do Excel, nao imagem ou desenho travado;
- deve ser movivel/editavel dentro do Excel;
- deve ter rotulos de dados visiveis;
- grafico de barras/colunas deve mostrar valores;
- grafico de rosca/pizza deve mostrar percentual e/ou categoria;
- se um grafico criado pelo `openpyxl` nao ficar movivel no Excel, ele deve ser recriado ou ajustado por uma etapa Excel/COM;
- antes de considerar um dashboard pronto, abrir no Excel e confirmar que os graficos podem ser movidos de lugar.

Essa regra vale para todos os dashboards novos.

## Regra para filtros e formulas

Todo dashboard deve continuar dinamico dentro do Excel.

- O ano base deve ser um filtro visivel no topo, seguindo o padrao `ANO` / `base` / `Selecionado`, com o valor selecionado na celula `W3` quando o layout usar esse padrao.
- Indicadores, tabelas auxiliares e graficos nao devem ser apenas valores fixos gerados pelo Python.
- O Python deve montar a estrutura, mas os numeros do dashboard devem vir de formulas do Excel ligadas as abas de origem.
- Quando os dados das abas da planilha forem alterados e o Excel recalcular, o dashboard tambem deve mudar.
- Graficos devem apontar para ranges alimentados por formulas, nao para valores colados manualmente.
- Tabelas auxiliares podem ficar ocultas, mas precisam existir para deixar claro de onde os indicadores e graficos saem.

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
