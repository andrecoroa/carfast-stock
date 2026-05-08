import sqlite3
from pathlib import Path

BASE_DIR = Path(__file__).parent
DB_PATH = BASE_DIR / "stock.db"

COLUMNS = {
    "estado_frota": "TEXT DEFAULT 'Ativa'",
    "data_venda": "TEXT",
    "valor_venda": "REAL",
    "comprador": "TEXT",
    "motivo_venda": "TEXT",
    "data_baixa": "TEXT",
    "ativa_operacional": "INTEGER DEFAULT 1",
    "current_status_rentway": "TEXT",
    "status_rentway": "TEXT",
    "fleet": "TEXT",
    "rental_station": "TEXT",
    "grupo": "TEXT",
    "categoria": "TEXT",
    "fornecedor": "TEXT",
    "valor_compra": "REAL",
    "valor_compra_com_iva": "REAL",
    "valor_venda_com_iva": "REAL",
    "documento_venda": "TEXT",
    "data_fatura_venda": "TEXT",
}

def column_exists(conn, table, column):
    rows = conn.execute(f"PRAGMA table_info({table})").fetchall()
    return any(r[1] == column for r in rows)

def main():
    if not DB_PATH.exists():
        raise SystemExit(f"Não encontrei a base de dados: {DB_PATH}")

    conn = sqlite3.connect(DB_PATH)

    for col, sql_type in COLUMNS.items():
        if not column_exists(conn, "viaturas", col):
            conn.execute(f"ALTER TABLE viaturas ADD COLUMN {col} {sql_type}")
            print(f"Coluna criada: {col}")
        else:
            print(f"Coluna já existe: {col}")

    conn.commit()
    conn.close()
    print("Atualização da estrutura da frota concluída.")

if __name__ == "__main__":
    main()
