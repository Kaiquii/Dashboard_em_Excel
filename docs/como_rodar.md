# Como rodar

Para gerar uma planilha, coloque a planilha base/original em `inputs/`.

Exemplo, se for regenerar o dashboard de Procytrat:

```text
inputs\CQ_Produto_Acabado_Procytrat_2026\CQ Produto Acabado - Procytrat 2026.xlsm
```

Rode este comando na raiz do projeto:

```powershell
.\gerar_dashboard.ps1
```

O arquivo gerado vai sair aqui:

```text
outputs\CQ_Produto_Acabado_Procytrat_2026\
```

Essa pasta deve ficar com apenas 1 arquivo final.

Arquivo final:

```text
CQ Produto Acabado - Procytrat 2026 - Dashboard2-sem-simulados.xlsm
```

Resumo:

```text
inputs  = planilhas originais
outputs = planilhas geradas
```

A versao pronta/correta que usamos como referencia fica em `referencias/`.
Nao use a versao pronta como se fosse a base original, senao o gerador vai reconstruir um dashboard em cima de outro dashboard.

Se quiser rodar pelo Python manualmente nesta maquina, use o Python empacotado:

```powershell
& "$env:USERPROFILE\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" .\gerar_dashboard.py
```
