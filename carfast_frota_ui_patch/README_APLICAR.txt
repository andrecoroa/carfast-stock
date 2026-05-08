Atualização Frota — filtros + importação Excel Rentway

Inclui:
1) Página /viaturas com filtros e resumo da frota.
2) Página /viaturas/importar para carregar Excel Rentway na app.
3) Campos de lifecycle para viaturas vendidas/históricas.

Como aplicar:
1. Copiar templates/oficina/viaturas.html para a tua app.
2. Copiar templates/oficina/importar_frota.html para a tua app.
3. Abrir APP_SNIPPETS.py e copiar os blocos indicados para o teu app.py.
4. Garantir que requirements.txt tem openpyxl==3.1.5.
5. Fazer:
   git add .
   git commit -m "frota filtros importacao"
   git push
6. Render: Manual Deploy → Deploy latest commit.
