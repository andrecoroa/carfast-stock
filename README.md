# CarFast Stock MVP — Fase 1

Aplicação simples de gestão de stocks para arrancar rapidamente.

## Funcionalidades
- Criar artigos de stock
- Registar entradas
- Registar saídas
- Consultar stock atual
- Definir stock mínimo
- Alerta visual quando o stock atual <= stock mínimo
- Associar saída a matrícula/viatura
- Histórico completo de movimentos
- Exportar artigos e movimentos para CSV

## Como executar

1. Instalar Python 3.10+
2. Abrir terminal dentro desta pasta
3. Executar:

```bash
pip install -r requirements.txt
python app.py
```

4. Abrir no browser:

```text
http://127.0.0.1:5000
```

## Estrutura
- `app.py` — aplicação Flask
- `stock.db` — base de dados SQLite, criada automaticamente
- `templates/` — páginas HTML
- `static/style.css` — estilos
- `exports/` — ficheiros exportados
