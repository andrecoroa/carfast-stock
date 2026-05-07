from flask import Flask, render_template, request, redirect, url_for, send_file, flash
import sqlite3
from pathlib import Path
import csv
from datetime import datetime

APP_DIR = Path(__file__).parent
DB_PATH = APP_DIR / "stock.db"
EXPORT_DIR = APP_DIR / "exports"
EXPORT_DIR.mkdir(exist_ok=True)

app = Flask(__name__)
app.secret_key = "carfast-stock-mvp"


def db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    with db() as conn:
        conn.execute("""
        CREATE TABLE IF NOT EXISTS artigos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            codigo TEXT UNIQUE NOT NULL,
            nome TEXT NOT NULL,
            categoria TEXT NOT NULL,
            unidade TEXT NOT NULL DEFAULT 'unidade',
            localizacao TEXT NOT NULL DEFAULT 'Armazém OPO',
            fornecedor TEXT,
            preco_unitario REAL DEFAULT 0,
            stock_minimo INTEGER NOT NULL DEFAULT 0,
            ativo INTEGER NOT NULL DEFAULT 1,
            criado_em TEXT NOT NULL
        )
        """)
        conn.execute("""
        CREATE TABLE IF NOT EXISTS movimentos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            artigo_id INTEGER NOT NULL,
            tipo TEXT NOT NULL CHECK(tipo IN ('Entrada','Saída','Ajuste')),
            quantidade INTEGER NOT NULL,
            motivo TEXT,
            matricula TEXT,
            responsavel TEXT,
            observacoes TEXT,
            criado_em TEXT NOT NULL,
            FOREIGN KEY(artigo_id) REFERENCES artigos(id)
        )
        """)


def stock_atual(artigo_id):
    with db() as conn:
        rows = conn.execute(
            "SELECT tipo, quantidade FROM movimentos WHERE artigo_id = ? ORDER BY id",
            (artigo_id,)
        ).fetchall()
    total = 0
    for r in rows:
        if r["tipo"] == "Entrada":
            total += r["quantidade"]
        elif r["tipo"] == "Saída":
            total -= r["quantidade"]
        elif r["tipo"] == "Ajuste":
            total = r["quantidade"]
    return total


@app.route("/")
def index():
    with db() as conn:
        artigos = conn.execute("SELECT * FROM artigos ORDER BY nome").fetchall()
    enriched = []
    for a in artigos:
        atual = stock_atual(a["id"])
        enriched.append({**dict(a), "stock_atual": atual, "alerta": atual <= a["stock_minimo"]})
    return render_template("index.html", artigos=enriched)


@app.route("/artigos/novo", methods=["GET", "POST"])
def novo_artigo():
    if request.method == "POST":
        try:
            with db() as conn:
                conn.execute("""
                    INSERT INTO artigos
                    (codigo, nome, categoria, unidade, localizacao, fornecedor, preco_unitario, stock_minimo, criado_em)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    request.form["codigo"].strip(),
                    request.form["nome"].strip(),
                    request.form["categoria"],
                    request.form.get("unidade", "unidade").strip(),
                    request.form.get("localizacao", "Armazém OPO").strip(),
                    request.form.get("fornecedor", "").strip(),
                    float(request.form.get("preco_unitario") or 0),
                    int(request.form.get("stock_minimo") or 0),
                    datetime.now().isoformat(timespec="seconds")
                ))
            flash("Artigo criado com sucesso.", "success")
            return redirect(url_for("index"))
        except sqlite3.IntegrityError:
            flash("Já existe um artigo com esse código.", "error")
    return render_template("novo_artigo.html")


@app.route("/movimentos/novo", methods=["GET", "POST"])
def novo_movimento():
    with db() as conn:
        artigos = conn.execute("SELECT id, codigo, nome FROM artigos WHERE ativo = 1 ORDER BY nome").fetchall()

    if request.method == "POST":
        artigo_id = int(request.form["artigo_id"])
        tipo = request.form["tipo"]
        quantidade = int(request.form["quantidade"])

        if quantidade <= 0:
            flash("A quantidade tem de ser superior a zero.", "error")
            return render_template("novo_movimento.html", artigos=artigos)

        if tipo == "Saída" and stock_atual(artigo_id) < quantidade:
            flash("Stock insuficiente para esta saída.", "error")
            return render_template("novo_movimento.html", artigos=artigos)

        with db() as conn:
            conn.execute("""
                INSERT INTO movimentos
                (artigo_id, tipo, quantidade, motivo, matricula, responsavel, observacoes, criado_em)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                artigo_id,
                tipo,
                quantidade,
                request.form.get("motivo", "").strip(),
                request.form.get("matricula", "").strip().upper(),
                request.form.get("responsavel", "").strip(),
                request.form.get("observacoes", "").strip(),
                datetime.now().isoformat(timespec="seconds")
            ))
        flash("Movimento registado com sucesso.", "success")
        return redirect(url_for("index"))

    return render_template("novo_movimento.html", artigos=artigos)


@app.route("/movimentos")
def movimentos():
    with db() as conn:
        rows = conn.execute("""
            SELECT m.*, a.codigo, a.nome
            FROM movimentos m
            JOIN artigos a ON a.id = m.artigo_id
            ORDER BY m.id DESC
        """).fetchall()
    return render_template("movimentos.html", movimentos=rows)


@app.route("/export/artigos")
def export_artigos():
    path = EXPORT_DIR / "artigos.csv"
    with db() as conn, open(path, "w", newline="", encoding="utf-8-sig") as f:
        rows = conn.execute("SELECT * FROM artigos ORDER BY nome").fetchall()
        writer = csv.writer(f, delimiter=";")
        writer.writerow(["codigo", "nome", "categoria", "unidade", "localizacao", "fornecedor", "preco_unitario", "stock_minimo", "stock_atual"])
        for r in rows:
            writer.writerow([r["codigo"], r["nome"], r["categoria"], r["unidade"], r["localizacao"], r["fornecedor"], r["preco_unitario"], r["stock_minimo"], stock_atual(r["id"])])
    return send_file(path, as_attachment=True)


@app.route("/export/movimentos")
def export_movimentos():
    path = EXPORT_DIR / "movimentos.csv"
    with db() as conn, open(path, "w", newline="", encoding="utf-8-sig") as f:
        rows = conn.execute("""
            SELECT m.*, a.codigo, a.nome
            FROM movimentos m
            JOIN artigos a ON a.id = m.artigo_id
            ORDER BY m.id DESC
        """).fetchall()
        writer = csv.writer(f, delimiter=";")
        writer.writerow(["data", "codigo", "artigo", "tipo", "quantidade", "motivo", "matricula", "responsavel", "observacoes"])
        for r in rows:
            writer.writerow([r["criado_em"], r["codigo"], r["nome"], r["tipo"], r["quantidade"], r["motivo"], r["matricula"], r["responsavel"], r["observacoes"]])
    return send_file(path, as_attachment=True)


if __name__ == "__main__":
    init_db()
    app.run(debug=True)
