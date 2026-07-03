import csv
import sqlite3
import os
from pathlib import Path

def sync_db():
    db_path = Path(os.getenv("DATABASE_PATH", Path(__file__).parent / "petlove.db"))
    csv_path = Path(os.getenv("BASE_DADOS_DIR", Path(__file__).parent / "Base de dados")) / "Fluxo de aprovação" / "Gestão de Perdas -Pet Love.csv"
    
    print(f"Sincronizando dados a partir de: {csv_path}")
    if not csv_path.exists():
        print("Arquivo CSV não encontrado.")
        return

    # Conectar ao banco de dados SQLite
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    # Identificar o delimitador do CSV analisando a primeira linha (cabeçalho)
    delimiter = ','
    with open(csv_path, newline='', encoding='utf-8-sig', errors='ignore') as f:
        first_line = f.readline()
        if first_line.count(';') > first_line.count(','):
            delimiter = ';'

    # Obter estrutura de colunas existentes ou criar
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='chamados'")
    tabela_existe = cursor.fetchone()

    with open(csv_path, newline='', encoding='utf-8-sig', errors='ignore') as f:
        reader = csv.DictReader(f, delimiter=delimiter)
        colunas_csv = list(reader.fieldnames)
        
        # Mapeamento de colunas do CSV para o DB
        colunas_db = []
        for col in colunas_csv:
            if not col:
                col = "coluna_vazia"
            col_db = col.strip().replace(" ", "_").replace(".", "").replace("-", "_").replace("(", "").replace(")", "").replace("/", "_")
            colunas_db.append(col_db)

        if not tabela_existe:
            colunas_sql = ", ".join([f'"{col}" TEXT' for col in colunas_db])
            cursor.execute(f'CREATE TABLE chamados ({colunas_sql})')
            print("Tabela 'chamados' criada pela primeira vez.")

        # Identificar as colunas corretas de ID, Responsável, Procedência e Tratativa no SQLite
        col_id_db = next((c for c in colunas_db if 'ID_do_Pedido' in c or 'ID_Pedido' in c), 'ID_do_Pedido')
        col_resp_db = next((c for c in colunas_db if 'Responsavel' in c), 'Responsavel')
        col_proc_db = next((c for c in colunas_db if 'Proced' in c), 'Procedência')
        col_trat_db = next((c for c in colunas_db if 'Tratativa' in c), 'Tratativa')

        # Buscar IDs existentes na tabela do SQLite e se têm responsável, procedência ou tratativa
        cursor.execute(f'SELECT "{col_id_db}", "{col_resp_db}", "{col_proc_db}", "{col_trat_db}" FROM chamados')
        rows_existentes = cursor.fetchall()
        
        # Criar dicionário {id: (responsavel, procedencia, tratativa)} para busca rápida
        dict_existentes = {}
        for r in rows_existentes:
            ped_id = str(r[0]).strip()
            resp = r[1]
            proc = r[2]
            trat = r[3]
            dict_existentes[ped_id] = (resp, proc, trat)

        novos_inseridos = 0
        atualizados = 0
        pula_tratados = 0

        # Preparar inserção
        placeholders = ", ".join(["?"] * len(colunas_db))
        sql_insert = f'INSERT INTO chamados VALUES ({placeholders})'

        # Preparar atualização
        update_sets = ", ".join([f'"{col}" = ?' for col in colunas_db if col != col_id_db])
        sql_update = f'UPDATE chamados SET {update_sets} WHERE "{col_id_db}" = ?'

        for row in reader:
            valores_csv = [row.get(col, "") for col in colunas_csv]
            
            # Achar o ID do pedido dessa linha do CSV
            idx_id = colunas_csv.index(next(c for c in colunas_csv if 'ID_do_Pedido' in c or 'ID_Pedido' in c))
            pedido_id = str(valores_csv[idx_id]).strip()

            if pedido_id in dict_existentes:
                # O chamado já existe no banco de dados.
                resp_existente, proc_existente, trat_existente = dict_existentes[pedido_id]
                
                # Se responsavel, tratativa e procedência forem nulos ou vazios (ou 'em analise'), podemos atualizar os dados (ainda pendente)
                is_resp_empty = not resp_existente or str(resp_existente).strip() == ''
                is_trat_empty = not trat_existente or str(trat_existente).strip() == ''
                is_proc_empty = not proc_existente or str(proc_existente).strip() == '' or str(proc_existente).strip().lower() == 'em analise'
                
                if is_resp_empty and is_trat_empty and is_proc_empty:
                    # Retirar o ID da lista de valores e colocar no final para a cláusula WHERE
                    valores_sem_id = []
                    for i, col in enumerate(colunas_csv):
                        if colunas_db[i] != col_id_db:
                            valores_sem_id.append(valores_csv[i])
                    valores_sem_id.append(pedido_id)
                    
                    cursor.execute(sql_update, valores_sem_id)
                    atualizados += 1
                else:
                    # Tem responsável, tratativa ou procedência, então o usuário já tratou/ajustou. PULAR!
                    pula_tratados += 1
            else:
                # É um chamado novo, inserir
                cursor.execute(sql_insert, valores_csv)
                novos_inseridos += 1

        conn.commit()
        conn.close()
        print(f"Sincronização concluída com sucesso:")
        print(f" - Novos chamados inseridos: {novos_inseridos}")
        print(f" - Chamados pendentes atualizados: {atualizados}")
        print(f" - Chamados tratados preservados (não alterados): {pula_tratados}")

if __name__ == "__main__":
    sync_db()
