# Atualização Frota Total — CarFast

## Objetivo
Atualizar a base `viaturas` com frota ativa e viaturas vendidas/históricas.

## Ficheiros necessários na pasta da app
- `stock.db`
- `frota 0805_v2.xlsx`
- `update_frota_lifecycle.py`
- `importar_frota_total.py`

## Passo 1 — Instalar dependência
```powershell
python -m pip install openpyxl
```

## Passo 2 — Atualizar estrutura da base
```powershell
python update_frota_lifecycle.py
```

## Passo 3 — Importar frota total
```powershell
python importar_frota_total.py
```

## O que o importador faz
- cria viaturas novas
- atualiza viaturas existentes
- marca viaturas vendidas como `Vendida`
- mantém vendidas no histórico
- não reativa viaturas vendidas
- guarda status Rentway
- guarda dados de venda quando disponíveis

## Estados usados
- Ativa
- Vendida

## Próximo passo na app
Adicionar filtros:
- Todas
- Ativas
- Vendidas
- Por marca
- Por modelo
- Por KM
- Por data compra
