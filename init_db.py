import csv
import sqlite3
import os
from pathlib import Path

def init_db():
    db_path = Path(os.getenv("DATABASE_PATH", Path(__file__).parent / "petlove.db"))
    csv_path = Path(os.getenv("BASE_DADOS_DIR", Path(__file__).parent / "Base de dados")) / "Fluxo de aprovação" / "Gestão de Perdas -Pet Love.csv"
    
    print(f"Lendo dados de {csv_path}...")
    if not csv_path.exists():
        print("CSV não encontrado.")
        return

    # Conectar ao banco de dados SQLite (cria se não existir)
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    # Identificar o delimitador do CSV
    delimiter = ','
    with open(csv_path, newline='', encoding='utf-8-sig', errors='ignore') as f:
        teste_f = f.read(1024)
        if ';' in teste_f and ',' not in teste_f:
            delimiter = ';'

    # Ler o CSV e obter os nomes das colunas originais
    with open(csv_path, newline='', encoding='utf-8-sig', errors='ignore') as f:
        reader = csv.DictReader(f, delimiter=delimiter)
        colunas_csv = list(reader.fieldnames)
        
        # Limpar nomes de colunas para serem válidos no SQLite
        colunas_db = []
        for col in colunas_csv:
            if not col:
                col = "coluna_vazia"
            # Remover caracteres especiais que atrapalham SQL
            col_db = col.strip().replace(" ", "_").replace(".", "").replace("-", "_").replace("(", "").replace(")", "").replace("/", "_")
            colunas_db.append(col_db)

        # Criar a tabela 'chamados'
        # Usamos TEXT para tudo pois o CSV é flexível
        colunas_sql = ", ".join([f'"{col}" TEXT' for col in colunas_db])
        colunas_sql += ', "Lista_Entrega_Cruzada" TEXT, "Status_da_Tratativa" TEXT, "Nome_do_cliente" TEXT, "Endereco_do_cliente" TEXT'
        
        cursor.execute(f'DROP TABLE IF EXISTS chamados')
        cursor.execute(f'CREATE TABLE chamados ({colunas_sql})')
        
        print(f"Tabela 'chamados' criada com colunas: {', '.join(colunas_db)}")
        
        # Inserir os dados
        placeholders = ", ".join(["?"] * len(colunas_db))
        sql_insert = f'INSERT INTO chamados VALUES ({placeholders})'
        
        linhas = []
        for row in reader:
            valores = [row.get(col, "") for col in colunas_csv]
            linhas.append(valores)
            
        cursor.executemany(sql_insert, linhas)
        conn.commit()
        print(f"{len(linhas)} registros inseridos no banco de dados com sucesso.")
        
    conn.close()
    print("Migração concluída.")

if __name__ == "__main__":
    init_db()
