# Dashboard Excel

Scripts para gerar dashboards em planilhas especificas.

Este repositorio nao e um projeto generico. Ele e uma colecao organizada de scripts para pegar planilhas Excel especificas e gerar dashboards dentro delas.

Leia a arquitetura antes de adicionar novas planilhas:

[docs/arquitetura.md](docs/arquitetura.md)

Para gerar uma planilha, siga:

[docs/como_rodar.md](docs/como_rodar.md)

Para lembrar onde guardar cada tipo de planilha:

[docs/pastas.md](docs/pastas.md)

Comando rapido:

```powershell
.\gerar_dashboard.ps1
```

## Estrutura

- `dashboards/`: uma pasta por planilha/dashboard.
- `comum/`: espaco reservado para helpers compartilhados quando houver repeticao real.
- `outputs/`: arquivos gerados pelos scripts.

## Dashboard atual

O dashboard `CQ_Produto_Acabado_Procytrat_2026` fica em:

`dashboards/CQ_Produto_Acabado_Procytrat_2026/`

Ordem dos scripts:

1. `01_build_dashboard.py`
2. `02_strip_charts.py`
3. `03_recreate_native_charts.ps1`
4. `04_cleanup_final.ps1`

Para rodar tudo:

```powershell
.\dashboards\CQ_Produto_Acabado_Procytrat_2026\run.ps1
```
