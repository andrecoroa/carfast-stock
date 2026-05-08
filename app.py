from flask import Flask, render_template, request, redirect, url_for, flash, send_file
import sqlite3
from pathlib import Path
from datetime import datetime
from werkzeug.utils import secure_filename

APP_DIR=Path(__file__).parent
DB_PATH=APP_DIR/'stock.db'
UPLOAD_DIR=APP_DIR/'uploads'; UPLOAD_DIR.mkdir(exist_ok=True)
app=Flask(__name__); app.secret_key='carfast-v2'
ESTADOS=['Aberto','Receção / ServiceBox','Diagnóstico Mecânico','Pendente Validação','FO Criada','Em Execução','Diagnóstico Pós','Concluído pelo Mecânico','Validação Administrativa','Fechado']

def now(): return datetime.now().isoformat(timespec='seconds')
def db():
    c=sqlite3.connect(DB_PATH); c.row_factory=sqlite3.Row; return c

def init_db():
    with db() as c:
        c.execute('''CREATE TABLE IF NOT EXISTS viaturas(id INTEGER PRIMARY KEY AUTOINCREMENT, matricula TEXT UNIQUE NOT NULL, vin TEXT, marca TEXT, modelo TEXT, versao TEXT, motorizacao TEXT, combustivel TEXT, caixa TEXT, ano INTEGER, data_compra TEXT, km_atual INTEGER, proxima_revisao_km INTEGER, proxima_revisao_data TEXT, ultima_revisao TEXT, campanhas_pendentes INTEGER DEFAULT 0, risco_tecnico TEXT DEFAULT 'Normal', observacoes TEXT, ativo INTEGER DEFAULT 1, criado_em TEXT NOT NULL)''')
        c.execute('''CREATE TABLE IF NOT EXISTS processos(id INTEGER PRIMARY KEY AUTOINCREMENT, numero_impro TEXT UNIQUE NOT NULL, viatura_id INTEGER NOT NULL, matricula TEXT NOT NULL, km_abertura INTEGER, responsavel TEXT, descricao_inicial TEXT, estado TEXT NOT NULL DEFAULT 'Aberto', data_abertura TEXT NOT NULL, data_fecho TEXT, observacoes TEXT)''')
        c.execute('''CREATE TABLE IF NOT EXISTS rececoes(id INTEGER PRIMARY KEY AUTOINCREMENT, processo_id INTEGER UNIQUE NOT NULL, servicebox_verificado INTEGER DEFAULT 0, campanhas_verificadas INTEGER DEFAULT 0, campanhas_pendentes INTEGER DEFAULT 0, plano_manutencao_verificado TEXT DEFAULT 'Não', historico_oem_verificado INTEGER DEFAULT 0, observacoes TEXT, atualizado_em TEXT NOT NULL)''')
        c.execute('''CREATE TABLE IF NOT EXISTS diagnosticos(id INTEGER PRIMARY KEY AUTOINCREMENT, processo_id INTEGER NOT NULL, viatura_id INTEGER NOT NULL, tipo TEXT NOT NULL, data_diagnostico TEXT NOT NULL, km_diagnostico INTEGER, n_manutencoes_ecu TEXT, intervencoes_reais TEXT, telecarregamento TEXT, ultima_manutencao_registada TEXT, manutencao_nao_registada TEXT, km_manutencao_anterior INTEGER, km_ate_manutencao INTEGER, diluicao_oleo REAL, carbono_oleo REAL, problema_identificado TEXT, causa_provavel TEXT, intervencao_recomendada TEXT, oficina_recomendada TEXT, pode_circular TEXT, prioridade TEXT, inconsistencias TEXT, observacoes TEXT, criado_em TEXT NOT NULL)''')
        c.execute('''CREATE TABLE IF NOT EXISTS folhas_obra(id INTEGER PRIMARY KEY AUTOINCREMENT, processo_id INTEGER UNIQUE NOT NULL, numero_fo TEXT, estado TEXT NOT NULL DEFAULT 'Aberta', responsavel TEXT, observacoes TEXT, criada_em TEXT NOT NULL, fechada_em TEXT)''')
        c.execute('''CREATE TABLE IF NOT EXISTS tarefas_fo(id INTEGER PRIMARY KEY AUTOINCREMENT, fo_id INTEGER NOT NULL, descricao TEXT NOT NULL, tipo_servico TEXT, responsavel TEXT, autorizada INTEGER DEFAULT 1, estado TEXT DEFAULT 'Pendente', observacoes TEXT, criado_em TEXT NOT NULL)''')
        c.execute('''CREATE TABLE IF NOT EXISTS validacoes(id INTEGER PRIMARY KEY AUTOINCREMENT, processo_id INTEGER UNIQUE NOT NULL, fo_fechada INTEGER DEFAULT 0, diagnostico_pos_anexado INTEGER DEFAULT 0, ficha_viatura_atualizada INTEGER DEFAULT 0, parametros_atualizados INTEGER DEFAULT 0, documentacao_completa INTEGER DEFAULT 0, observacoes TEXT, validado_por TEXT, atualizado_em TEXT NOT NULL)''')
        c.execute('''CREATE TABLE IF NOT EXISTS anexos(id INTEGER PRIMARY KEY AUTOINCREMENT, processo_id INTEGER NOT NULL, area TEXT NOT NULL, filename TEXT NOT NULL, original_name TEXT NOT NULL, criado_em TEXT NOT NULL)''')

def ensure_viatura(matricula, km=None):
    matricula=matricula.strip().upper()
    with db() as c:
        v=c.execute('SELECT * FROM viaturas WHERE matricula=?',(matricula,)).fetchone()
        if v:
            if km: c.execute('UPDATE viaturas SET km_atual=? WHERE id=?',(km,v['id']))
            return v['id']
        cur=c.execute('INSERT INTO viaturas(matricula,km_atual,criado_em) VALUES(?,?,?)',(matricula,km,now()))
        return cur.lastrowid

def processo(processo_id):
    with db() as c:
        return c.execute('''SELECT p.*,v.vin,v.marca,v.modelo,v.motorizacao FROM processos p JOIN viaturas v ON v.id=p.viatura_id WHERE p.id=?''',(processo_id,)).fetchone()

def uploads(processo_id, area, files):
    with db() as c:
        for f in files:
            if not f or not f.filename: continue
            original=f.filename; safe=secure_filename(original)
            name=f'{processo_id}_{area}_{datetime.now().strftime("%Y%m%d%H%M%S%f")}_{safe}'
            f.save(UPLOAD_DIR/name)
            c.execute('INSERT INTO anexos(processo_id,area,filename,original_name,criado_em) VALUES(?,?,?,?,?)',(processo_id,area,name,original,now()))

@app.route('/')
def home(): return redirect(url_for('processos'))

@app.route('/processos')
def processos():
    with db() as c:
        rows=c.execute('''SELECT p.*,v.vin,v.marca,v.modelo FROM processos p JOIN viaturas v ON v.id=p.viatura_id ORDER BY p.id DESC''').fetchall()
    return render_template('oficina/processos.html', processos=rows)

@app.route('/processos/novo', methods=['GET','POST'])
def novo_processo():
    if request.method=='POST':
        impro=request.form['numero_impro'].strip(); mat=request.form['matricula'].strip().upper(); km=int(request.form.get('km_abertura') or 0) or None
        vid=ensure_viatura(mat,km)
        try:
            with db() as c:
                cur=c.execute('''INSERT INTO processos(numero_impro,viatura_id,matricula,km_abertura,responsavel,descricao_inicial,estado,data_abertura) VALUES(?,?,?,?,?,?,?,?)''',(impro,vid,mat,km,request.form.get('responsavel',''),request.form.get('descricao_inicial',''),'Receção / ServiceBox',now()))
            flash('Processo aberto. Avança para Receção / ServiceBox.','success')
            return redirect(url_for('processo_detail', processo_id=cur.lastrowid))
        except sqlite3.IntegrityError:
            flash('Já existe processo com esse Nº Impro.','error')
    return render_template('oficina/novo_processo.html')

@app.route('/processos/<int:processo_id>')
def processo_detail(processo_id):
    p=processo(processo_id)
    with db() as c:
        rec=c.execute('SELECT * FROM rececoes WHERE processo_id=?',(processo_id,)).fetchone()
        diags=c.execute('SELECT * FROM diagnosticos WHERE processo_id=? ORDER BY id DESC',(processo_id,)).fetchall()
        fo=c.execute('SELECT * FROM folhas_obra WHERE processo_id=?',(processo_id,)).fetchone()
        tarefas=c.execute('SELECT * FROM tarefas_fo WHERE fo_id=? ORDER BY id',(fo['id'],)).fetchall() if fo else []
        val=c.execute('SELECT * FROM validacoes WHERE processo_id=?',(processo_id,)).fetchone()
        anexos=c.execute('SELECT * FROM anexos WHERE processo_id=? ORDER BY id DESC',(processo_id,)).fetchall()
    return render_template('oficina/processo_detail.html',p=p,rececao=rec,diagnosticos=diags,fo=fo,tarefas=tarefas,validacao=val,anexos=anexos,estados=ESTADOS)

@app.route('/processos/<int:processo_id>/estado', methods=['POST'])
def atualizar_estado(processo_id):
    with db() as c: c.execute('UPDATE processos SET estado=? WHERE id=?',(request.form['estado'],processo_id))
    return redirect(url_for('processo_detail', processo_id=processo_id))

@app.route('/processos/<int:processo_id>/rececao', methods=['GET','POST'])
def rececao(processo_id):
    p=processo(processo_id)
    if request.method=='POST':
        with db() as c:
            c.execute('''INSERT INTO rececoes(processo_id,servicebox_verificado,campanhas_verificadas,campanhas_pendentes,plano_manutencao_verificado,historico_oem_verificado,observacoes,atualizado_em) VALUES(?,?,?,?,?,?,?,?) ON CONFLICT(processo_id) DO UPDATE SET servicebox_verificado=excluded.servicebox_verificado,campanhas_verificadas=excluded.campanhas_verificadas,campanhas_pendentes=excluded.campanhas_pendentes,plano_manutencao_verificado=excluded.plano_manutencao_verificado,historico_oem_verificado=excluded.historico_oem_verificado,observacoes=excluded.observacoes,atualizado_em=excluded.atualizado_em''',(processo_id,1 if request.form.get('servicebox_verificado') else 0,1 if request.form.get('campanhas_verificadas') else 0,1 if request.form.get('campanhas_pendentes') else 0,request.form.get('plano_manutencao_verificado'),1 if request.form.get('historico_oem_verificado') else 0,request.form.get('observacoes'),now()))
            c.execute('UPDATE processos SET estado=? WHERE id=?',('Diagnóstico Mecânico',processo_id))
        uploads(processo_id,'rececao',request.files.getlist('anexos'))
        return redirect(url_for('processo_detail', processo_id=processo_id))
    with db() as c: r=c.execute('SELECT * FROM rececoes WHERE processo_id=?',(processo_id,)).fetchone()
    return render_template('oficina/rececao.html',p=p,r=r)

@app.route('/processos/<int:processo_id>/diagnostico/<tipo>', methods=['GET','POST'])
def diagnostico(processo_id,tipo):
    tipo_real={'pre':'Pré','mecanico':'Mecânico','pos':'Pós'}.get(tipo)
    p=processo(processo_id)
    if request.method=='POST':
        with db() as c:
            c.execute('''INSERT INTO diagnosticos(processo_id,viatura_id,tipo,data_diagnostico,km_diagnostico,n_manutencoes_ecu,intervencoes_reais,telecarregamento,ultima_manutencao_registada,manutencao_nao_registada,km_manutencao_anterior,km_ate_manutencao,diluicao_oleo,carbono_oleo,problema_identificado,causa_provavel,intervencao_recomendada,oficina_recomendada,pode_circular,prioridade,inconsistencias,observacoes,criado_em) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',(processo_id,p['viatura_id'],tipo_real,request.form.get('data_diagnostico') or now(),request.form.get('km_diagnostico') or None,request.form.get('n_manutencoes_ecu'),request.form.get('intervencoes_reais'),request.form.get('telecarregamento'),request.form.get('ultima_manutencao_registada'),request.form.get('manutencao_nao_registada'),request.form.get('km_manutencao_anterior') or None,request.form.get('km_ate_manutencao') or None,request.form.get('diluicao_oleo') or None,request.form.get('carbono_oleo') or None,request.form.get('problema_identificado'),request.form.get('causa_provavel'),request.form.get('intervencao_recomendada'),request.form.get('oficina_recomendada'),request.form.get('pode_circular'),request.form.get('prioridade'),request.form.get('inconsistencias'),request.form.get('observacoes'),now()))
            if tipo_real=='Mecânico': c.execute('UPDATE processos SET estado=? WHERE id=?',('Pendente Validação',processo_id))
            if tipo_real=='Pós': c.execute('UPDATE processos SET estado=? WHERE id=?',('Concluído pelo Mecânico',processo_id))
        uploads(processo_id,'diagnostico_'+tipo_real.lower(),request.files.getlist('anexos'))
        return redirect(url_for('processo_detail', processo_id=processo_id))
    return render_template('oficina/diagnostico.html',p=p,tipo=tipo_real)

@app.route('/processos/<int:processo_id>/fo/criar', methods=['POST'])
def criar_fo(processo_id):
    with db() as c:
        c.execute('INSERT OR IGNORE INTO folhas_obra(processo_id,numero_fo,estado,responsavel,observacoes,criada_em) VALUES(?,?,?,?,?,?)',(processo_id,request.form.get('numero_fo') or f'FO-{processo_id:04d}','Aberta',request.form.get('responsavel'),request.form.get('observacoes'),now()))
        c.execute('UPDATE processos SET estado=? WHERE id=?',('FO Criada',processo_id))
    return redirect(url_for('processo_detail', processo_id=processo_id))

@app.route('/fo/<int:fo_id>/tarefas/nova', methods=['POST'])
def nova_tarefa_fo(fo_id):
    with db() as c:
        fo=c.execute('SELECT * FROM folhas_obra WHERE id=?',(fo_id,)).fetchone()
        c.execute('INSERT INTO tarefas_fo(fo_id,descricao,tipo_servico,responsavel,autorizada,estado,observacoes,criado_em) VALUES(?,?,?,?,?,?,?,?)',(fo_id,request.form.get('descricao'),request.form.get('tipo_servico'),request.form.get('responsavel'),1 if request.form.get('autorizada') else 0,request.form.get('estado') or 'Pendente',request.form.get('observacoes'),now()))
    return redirect(url_for('processo_detail', processo_id=fo['processo_id']))

@app.route('/tarefas/<int:tarefa_id>/estado', methods=['POST'])
def atualizar_tarefa(tarefa_id):
    with db() as c:
        t=c.execute('SELECT t.*,f.processo_id FROM tarefas_fo t JOIN folhas_obra f ON f.id=t.fo_id WHERE t.id=?',(tarefa_id,)).fetchone()
        c.execute('UPDATE tarefas_fo SET estado=? WHERE id=?',(request.form.get('estado'),tarefa_id))
    return redirect(url_for('processo_detail', processo_id=t['processo_id']))

@app.route('/processos/<int:processo_id>/iniciar_execucao', methods=['POST'])
def iniciar_execucao(processo_id):
    with db() as c:
        fo=c.execute('SELECT * FROM folhas_obra WHERE processo_id=?',(processo_id,)).fetchone()
        if not fo:
            flash('Não é possível iniciar execução: não existe FO criada.','error'); return redirect(url_for('processo_detail', processo_id=processo_id))
        n=c.execute('SELECT COUNT(*) c FROM tarefas_fo WHERE fo_id=? AND autorizada=1',(fo['id'],)).fetchone()['c']
        if n==0:
            flash('Não é possível iniciar execução: a FO não tem tarefas autorizadas.','error'); return redirect(url_for('processo_detail', processo_id=processo_id))
        c.execute('UPDATE processos SET estado=? WHERE id=?',('Em Execução',processo_id))
    return redirect(url_for('processo_detail', processo_id=processo_id))

@app.route('/processos/<int:processo_id>/validacao', methods=['GET','POST'])
def validacao(processo_id):
    p=processo(processo_id)
    if request.method=='POST':
        vals=[1 if request.form.get(k) else 0 for k in ['fo_fechada','diagnostico_pos_anexado','ficha_viatura_atualizada','parametros_atualizados','documentacao_completa']]
        with db() as c:
            c.execute('''INSERT INTO validacoes(processo_id,fo_fechada,diagnostico_pos_anexado,ficha_viatura_atualizada,parametros_atualizados,documentacao_completa,observacoes,validado_por,atualizado_em) VALUES(?,?,?,?,?,?,?,?,?) ON CONFLICT(processo_id) DO UPDATE SET fo_fechada=excluded.fo_fechada,diagnostico_pos_anexado=excluded.diagnostico_pos_anexado,ficha_viatura_atualizada=excluded.ficha_viatura_atualizada,parametros_atualizados=excluded.parametros_atualizados,documentacao_completa=excluded.documentacao_completa,observacoes=excluded.observacoes,validado_por=excluded.validado_por,atualizado_em=excluded.atualizado_em''',(processo_id,*vals,request.form.get('observacoes'),request.form.get('validado_por'),now()))
            if all(vals): c.execute('UPDATE processos SET estado=?, data_fecho=? WHERE id=?',('Fechado',now(),processo_id))
            else: c.execute('UPDATE processos SET estado=? WHERE id=?',('Validação Administrativa',processo_id))
        return redirect(url_for('processo_detail', processo_id=processo_id))
    with db() as c: val=c.execute('SELECT * FROM validacoes WHERE processo_id=?',(processo_id,)).fetchone()
    return render_template('oficina/validacao.html',p=p,val=val)

@app.route('/processos/<int:processo_id>/comparacao')
def comparacao(processo_id):
    p=processo(processo_id)
    with db() as c: diags=c.execute('SELECT * FROM diagnosticos WHERE viatura_id=? ORDER BY data_diagnostico DESC,id DESC',(p['viatura_id'],)).fetchall()
    campos=[('km_diagnostico','Quilometragem'),('n_manutencoes_ecu','Nº manutenções ECU'),('intervencoes_reais','Intervenções reais'),('telecarregamento','Telecarregamento'),('ultima_manutencao_registada','Última manutenção registada'),('manutencao_nao_registada','Manutenção não registada'),('km_manutencao_anterior','Km manutenção anterior'),('km_ate_manutencao','Km até manutenção'),('diluicao_oleo','Diluição óleo'),('carbono_oleo','Carbono óleo')]
    return render_template('oficina/comparacao.html',p=p,diags=diags[:4],campos=campos)

@app.route('/viaturas')
def viaturas():
    with db() as c: rows=c.execute('SELECT * FROM viaturas ORDER BY matricula').fetchall()
    return render_template('oficina/viaturas.html',viaturas=rows)

@app.route('/viaturas/<int:viatura_id>', methods=['GET','POST'])
def viatura_detail(viatura_id):
    with db() as c:
        if request.method=='POST':
            c.execute('''UPDATE viaturas SET vin=?,marca=?,modelo=?,versao=?,motorizacao=?,combustivel=?,caixa=?,ano=?,data_compra=?,km_atual=?,proxima_revisao_km=?,proxima_revisao_data=?,ultima_revisao=?,campanhas_pendentes=?,risco_tecnico=?,observacoes=?,ativo=? WHERE id=?''',(request.form.get('vin'),request.form.get('marca'),request.form.get('modelo'),request.form.get('versao'),request.form.get('motorizacao'),request.form.get('combustivel'),request.form.get('caixa'),request.form.get('ano') or None,request.form.get('data_compra'),request.form.get('km_atual') or None,request.form.get('proxima_revisao_km') or None,request.form.get('proxima_revisao_data'),request.form.get('ultima_revisao'),1 if request.form.get('campanhas_pendentes') else 0,request.form.get('risco_tecnico'),request.form.get('observacoes'),1 if request.form.get('ativo') else 0,viatura_id))
            return redirect(url_for('viatura_detail',viatura_id=viatura_id))
        v=c.execute('SELECT * FROM viaturas WHERE id=?',(viatura_id,)).fetchone()
        processos_rows=c.execute('SELECT * FROM processos WHERE viatura_id=? ORDER BY id DESC',(viatura_id,)).fetchall()
        diags=c.execute('SELECT d.*,p.numero_impro FROM diagnosticos d JOIN processos p ON p.id=d.processo_id WHERE d.viatura_id=? ORDER BY d.data_diagnostico DESC,d.id DESC',(viatura_id,)).fetchall()
    return render_template('oficina/viatura_detail.html',v=v,processos=processos_rows,diagnosticos=diags)

@app.route('/uploads/<filename>')
def uploaded_file(filename): return send_file(UPLOAD_DIR/filename, as_attachment=True)

# redirects for old stock links kept simple
@app.route('/stock')
def stock_placeholder(): return '<h2>Stock mantido na versão anterior.</h2><p>Esta versão foca Oficina. O módulo stock volta a ser integrado na próxima consolidação.</p><p><a href="/processos">Voltar à Oficina</a></p>'

if __name__=='__main__':
    init_db(); app.run(debug=True,host='0.0.0.0',port=5000)
