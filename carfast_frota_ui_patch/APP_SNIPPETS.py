# 1) Acrescentar no topo do app.py:
# from datetime import timedelta
# try:
#     from openpyxl import load_workbook
# except ImportError:
#     load_workbook = None

# 2) Dentro de init_db(), após criar tabela viaturas, acrescentar:
# lifecycle_columns = {
#   "estado_frota": "TEXT DEFAULT 'Ativa'", "data_venda": "TEXT", "valor_venda": "REAL",
#   "comprador": "TEXT", "motivo_venda": "TEXT", "data_baixa": "TEXT",
#   "ativa_operacional": "INTEGER DEFAULT 1", "current_status_rentway": "TEXT",
#   "status_rentway": "TEXT", "fleet": "TEXT", "rental_station": "TEXT", "grupo": "TEXT",
#   "categoria": "TEXT", "fornecedor": "TEXT", "valor_compra": "REAL",
#   "valor_compra_com_iva": "REAL", "valor_venda_com_iva": "REAL",
#   "documento_venda": "TEXT", "data_fatura_venda": "TEXT"
# }
# for col, typ in lifecycle_columns.items():
#     rows = c.execute("PRAGMA table_info(viaturas)").fetchall()
#     if not any(r[1] == col for r in rows):
#         c.execute(f"ALTER TABLE viaturas ADD COLUMN {col} {typ}")
# c.execute('''CREATE TABLE IF NOT EXISTS importacoes(id INTEGER PRIMARY KEY AUTOINCREMENT, tipo TEXT NOT NULL, filename TEXT, original_name TEXT, criadas INTEGER DEFAULT 0, atualizadas INTEGER DEFAULT 0, vendidas INTEGER DEFAULT 0, ativas INTEGER DEFAULT 0, ignoradas INTEGER DEFAULT 0, responsavel TEXT, observacoes TEXT, criado_em TEXT NOT NULL)''')

# 3) Substituir a rota /viaturas pela versão com filtros:
@app.route('/viaturas')
def viaturas():
    q=request.args.get('q','').strip(); estado=request.args.get('estado','Ativa'); marca=request.args.get('marca','').strip(); modelo=request.args.get('modelo','').strip(); km_min=request.args.get('km_min','').strip(); km_max=request.args.get('km_max','').strip()
    where=[]; params=[]
    if q: where.append('(matricula LIKE ? OR vin LIKE ?)'); params += [f'%{q}%',f'%{q}%']
    if estado and estado!='Todas': where.append("COALESCE(estado_frota,'Ativa')=?"); params.append(estado)
    if marca: where.append('marca=?'); params.append(marca)
    if modelo: where.append('modelo=?'); params.append(modelo)
    if km_min: where.append('km_atual>=?'); params.append(int(km_min))
    if km_max: where.append('km_atual<=?'); params.append(int(km_max))
    sql='WHERE '+' AND '.join(where) if where else ''
    with db() as c:
        rows=c.execute(f'SELECT * FROM viaturas {sql} ORDER BY matricula LIMIT 500',params).fetchall()
        marcas=c.execute("SELECT DISTINCT marca FROM viaturas WHERE marca IS NOT NULL AND marca<>'' ORDER BY marca").fetchall()
        modelos=c.execute("SELECT DISTINCT modelo FROM viaturas WHERE modelo IS NOT NULL AND modelo<>'' ORDER BY modelo").fetchall()
        resumo=c.execute("SELECT COUNT(*) total, SUM(CASE WHEN COALESCE(estado_frota,'Ativa')='Ativa' THEN 1 ELSE 0 END) ativas, SUM(CASE WHEN COALESCE(estado_frota,'Ativa')='Vendida' THEN 1 ELSE 0 END) vendidas, SUM(CASE WHEN COALESCE(estado_frota,'Ativa')='Em venda' THEN 1 ELSE 0 END) em_venda FROM viaturas").fetchone()
    return render_template('oficina/viaturas.html',viaturas=rows,marcas=marcas,modelos=modelos,filtros={'q':q,'estado':estado,'marca':marca,'modelo':modelo,'km_min':km_min,'km_max':km_max},resumo=resumo)

# 4) Acrescentar esta rota para menu de importação. A função de importação pode continuar por script nesta versão; aqui fica o histórico preparado:
@app.route('/viaturas/importar', methods=['GET','POST'])
def importar_frota_view():
    if request.method=='POST':
        flash('Upload recebido. A ligação direta ao importador fica no próximo passo.', 'success')
        return redirect(url_for('viaturas'))
    with db() as c:
        importacoes=c.execute('SELECT * FROM importacoes ORDER BY id DESC LIMIT 20').fetchall()
    return render_template('oficina/importar_frota.html',importacoes=importacoes)
