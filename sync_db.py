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

    # Criar a tabela de historico se nao existir
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS historico_chamados (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            pedido_id TEXT,
            data_hora TEXT,
            usuario TEXT,
            campo TEXT,
            valor_antigo TEXT,
            valor_novo TEXT
        )
    ''')
    
    # Adicionar coluna Lista_Entrega_Cruzada se nao existir
    try:
        cursor.execute("ALTER TABLE chamados ADD COLUMN Lista_Entrega_Cruzada TEXT")
        conn.commit()
    except sqlite3.OperationalError:
        pass

    # Adicionar coluna Status_da_Tratativa se nao existir
    try:
        cursor.execute("ALTER TABLE chamados ADD COLUMN Status_da_Tratativa TEXT")
        conn.commit()
    except sqlite3.OperationalError:
        pass

    # Adicionar coluna Nome_do_cliente se nao existir
    try:
        cursor.execute("ALTER TABLE chamados ADD COLUMN Nome_do_cliente TEXT")
        conn.commit()
    except sqlite3.OperationalError:
        pass

    # Adicionar coluna Endereco_do_cliente se nao existir
    try:
        cursor.execute("ALTER TABLE chamados ADD COLUMN Endereco_do_cliente TEXT")
        conn.commit()
    except sqlite3.OperationalError:
        pass

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

        # Obter colunas reais da tabela (que incluem Lista_Entrega_Cruzada, Status_da_Tratativa, etc.)
        cursor.execute("PRAGMA table_info(chamados)")
        colunas_reais = [r[1] for r in cursor.fetchall()]

        # Obter todos os pedidos que já possuem histórico para preservá-los
        cursor.execute("SELECT DISTINCT pedido_id FROM historico_chamados")
        pedidos_com_historico = {str(r[0]).strip() for r in cursor.fetchall() if r[0]}

        # 1. DE-DUPLICAR O BANCO DE DADOS ATUAL (usando tabela temporária para segurança)
        cursor.execute("SELECT * FROM chamados")
        db_rows = cursor.fetchall()
        
        db_dict = {}
        for row in db_rows:
            row_dict = dict(zip(colunas_reais, row))
            ped_id = str(row_dict.get(col_id_db, "")).strip()
            if not ped_id:
                continue
                
            col_resp_name = next((c for c in colunas_reais if 'Responsavel' in c), 'Responsavel')
            col_proc_name = next((c for c in colunas_reais if 'Proced' in c), 'Procedência')
            col_trat_name = next((c for c in colunas_reais if 'Tratativa' in c), 'Tratativa')
            col_status_name = next((c for c in colunas_reais if 'Status' in c), 'Status_da_Tratativa')
            
            resp = str(row_dict.get(col_resp_name, "")).strip()
            proc = str(row_dict.get(col_proc_name, "")).strip()
            trat = str(row_dict.get(col_trat_name, "")).strip()
            status = str(row_dict.get(col_status_name, "")).strip()
            
            # Normalizar procedência
            proc_lower = proc.lower()
            if proc_lower in ["procedente", "procédente", "procedente"]:
                proc = "Procedente"
            elif proc_lower in ["não procedente", "nao procedente", "improcedente", "improcedente"]:
                proc = "Não Procedente"
            elif proc_lower in ["em analise", "em análise"]:
                proc = "Em Análise"
            
            is_treated = bool(resp or trat or (proc and proc.lower() not in ['em analise', 'em análise']) or (status and status.lower() == 'finalizado') or ped_id in pedidos_com_historico)
            
            if ped_id not in db_dict:
                db_dict[ped_id] = row
            else:
                existing_row = db_dict[ped_id]
                existing_dict = dict(zip(colunas_reais, existing_row))
                ext_resp = str(existing_dict.get(col_resp_name, "")).strip()
                ext_proc = str(existing_dict.get(col_proc_name, "")).strip()
                ext_trat = str(existing_dict.get(col_trat_name, "")).strip()
                ext_status = str(existing_dict.get(col_status_name, "")).strip()
                
                # Normalizar procedência existente
                ext_proc_lower = ext_proc.lower()
                if ext_proc_lower in ["procedente", "procédente", "procedente"]:
                    ext_proc = "Procedente"
                elif ext_proc_lower in ["não procedente", "nao procedente", "improcedente", "improcedente"]:
                    ext_proc = "Não Procedente"
                elif ext_proc_lower in ["em analise", "em análise"]:
                    ext_proc = "Em Análise"
                
                existing_treated = bool(ext_resp or ext_trat or (ext_proc and ext_proc.lower() not in ['em analise', 'em análise']) or (ext_status and ext_status.lower() == 'finalizado') or ped_id in pedidos_com_historico)
                
                if is_treated and not existing_treated:
                    db_dict[ped_id] = row

        # SEGURO: Usar tabela temporária em vez de DELETE direto
        colunas_reais_str = ", ".join([f'"{c}"' for c in colunas_reais])
        placeholders_reais = ", ".join(["?"] * len(colunas_reais))
        
        cursor.execute(f"DROP TABLE IF EXISTS chamados_staging")
        colunas_create = ", ".join([f'"{c}" TEXT' for c in colunas_reais])
        cursor.execute(f"CREATE TABLE chamados_staging ({colunas_create})")
        
        sql_insert_staging = f'INSERT INTO chamados_staging ({colunas_reais_str}) VALUES ({placeholders_reais})'
        for row in db_dict.values():
            cursor.execute(sql_insert_staging, row)
        
        # Verificar se a staging tem os mesmos registros antes de substituir
        cursor.execute("SELECT COUNT(*) FROM chamados_staging")
        staging_count = cursor.fetchone()[0]
        cursor.execute("SELECT COUNT(DISTINCT \"{}\") FROM chamados".format(col_id_db))
        original_unique_count = cursor.fetchone()[0]
        
        if staging_count >= original_unique_count * 0.9:  # Proteção: não perder mais de 10%
            cursor.execute("DROP TABLE chamados")
            cursor.execute("ALTER TABLE chamados_staging RENAME TO chamados")
            print(f"Limpeza concluída: Banco de dados de-duplicado para {len(db_dict)} registros únicos.")
        else:
            cursor.execute("DROP TABLE chamados_staging")
            print(f"ALERTA: De-duplicação abortada por segurança. Staging tinha {staging_count} vs {original_unique_count} registros únicos.")
            conn.close()
            return
        
        # Definir query para novos registros do CSV (14 colunas)
        placeholders = ", ".join(["?"] * len(colunas_db))
        colunas_csv_str = ", ".join([f'"{c}"' for c in colunas_db])
        sql_insert = f'INSERT INTO chamados ({colunas_csv_str}) VALUES ({placeholders})'

        # 2. DE-DUPLICAR O CSV QUE VEM DO ARQUIVO
        f.seek(0)
        # Pular cabeçalho
        next(f)
        csv_reader = csv.reader(f, delimiter=delimiter)
        
        csv_dedup = {}
        for row in csv_reader:
            if len(row) < len(colunas_csv):
                continue
            valores_csv = [val.strip() for val in row]
            
            # Identificar o ID do Pedido
            idx_id = colunas_csv.index(next(c for c in colunas_csv if 'ID_do_Pedido' in c or 'ID_Pedido' in c))
            pedido_id = valores_csv[idx_id]
            if not pedido_id:
                continue
                
            idx_resp = colunas_csv.index(next(c for c in colunas_csv if 'Responsavel' in c))
            idx_proc = colunas_csv.index(next(c for c in colunas_csv if 'Proced' in c))
            idx_trat = colunas_csv.index(next(c for c in colunas_csv if 'Tratativa' in c))
            
            resp = valores_csv[idx_resp]
            proc = valores_csv[idx_proc]
            trat = valores_csv[idx_trat]
            
            # Normalizar procedência do CSV
            proc_lower = proc.lower()
            if proc_lower in ["procedente", "procédente", "procedente"]:
                valores_csv[idx_proc] = "Procedente"
                proc = "Procedente"
            elif proc_lower in ["não procedente", "nao procedente", "improcedente", "improcedente"]:
                valores_csv[idx_proc] = "Não Procedente"
                proc = "Não Procedente"
            elif proc_lower in ["em analise", "em análise"]:
                valores_csv[idx_proc] = "Em Análise"
                proc = "Em Análise"
            
            is_treated = bool(resp or trat or (proc and proc.lower() not in ['em analise', 'em análise']))
            
            if pedido_id not in csv_dedup:
                csv_dedup[pedido_id] = valores_csv
            else:
                existing_val = csv_dedup[pedido_id]
                ext_resp = existing_val[idx_resp]
                ext_proc = existing_val[idx_proc]
                ext_trat = existing_val[idx_trat]
                
                # Normalizar procedência existente do CSV
                ext_proc_lower = ext_proc.lower()
                if ext_proc_lower in ["procedente", "procédente", "procedente"]:
                    ext_proc = "Procedente"
                elif ext_proc_lower in ["não procedente", "nao procedente", "improcedente", "improcedente"]:
                    ext_proc = "Não Procedente"
                elif ext_proc_lower in ["em analise", "em análise"]:
                    ext_proc = "Em Análise"
                
                existing_treated = bool(ext_resp or ext_trat or (ext_proc and ext_proc.lower() not in ['em analise', 'em análise']))
                
                if is_treated and not existing_treated:
                    csv_dedup[pedido_id] = valores_csv

        # 3. MERGE DOS DADOS DE-DUPLICADOS
        # Recarregar registros existentes para verificação
        col_status_db = "Status_da_Tratativa"
        cursor.execute(f'SELECT "{col_id_db}", "{col_resp_db}", "{col_proc_db}", "{col_trat_db}", "{col_status_db}" FROM chamados')
        rows_existentes = cursor.fetchall()
        dict_existentes = {}
        for r in rows_existentes:
            ped_id = str(r[0]).strip()
            dict_existentes[ped_id] = (r[1], r[2], r[3], r[4])

        novos_inseridos = 0
        atualizados = 0
        pula_tratados = 0

        # Preparar atualização
        update_sets = ", ".join([f'"{col}" = ?' for col in colunas_db if col != col_id_db])
        sql_update = f'UPDATE chamados SET {update_sets} WHERE "{col_id_db}" = ?'

        for pedido_id, valores_csv in csv_dedup.items():
            if pedido_id in dict_existentes:
                resp_existente, proc_existente, trat_existente, status_existente = dict_existentes[pedido_id]
                
                is_resp_empty = not resp_existente or str(resp_existente).strip() == ''
                is_trat_empty = not trat_existente or str(trat_existente).strip() == ''
                
                # Normalizar procedência existente para a verificação
                proc_existente_str = str(proc_existente).strip() if proc_existente else ''
                proc_existente_lower = proc_existente_str.lower()
                if proc_existente_lower in ["procedente", "procédente", "procedente"]:
                    proc_existente_str = "Procedente"
                elif proc_existente_lower in ["não procedente", "nao procedente", "improcedente", "improcedente"]:
                    proc_existente_str = "Não Procedente"
                elif proc_existente_lower in ["em analise", "em análise"]:
                    proc_existente_str = "Em Análise"
                
                is_proc_empty = not proc_existente_str or proc_existente_str == '' or proc_existente_str.lower() in ['em analise', 'em análise']
                is_status_finalizado = status_existente and str(status_existente).strip().lower() == 'finalizado'
                
                has_history = pedido_id in pedidos_com_historico
                
                # O registro já foi alterado pelo usuário se possuir histórico, responsável preenchido,
                # tratativa preenchida, procedência preenchida, ou status Finalizado
                if has_history or not is_resp_empty or not is_trat_empty or not is_proc_empty or is_status_finalizado:
                    pula_tratados += 1
                else:
                    # Retirar o ID da lista de valores e colocar no final para a cláusula WHERE
                    valores_sem_id = []
                    for i, col in enumerate(colunas_csv):
                        if colunas_db[i] != col_id_db:
                            valores_sem_id.append(valores_csv[i])
                    valores_sem_id.append(pedido_id)
                    
                    cursor.execute(sql_update, valores_sem_id)
                    atualizados += 1
            else:
                cursor.execute(sql_insert, valores_csv)
                novos_inseridos += 1

        conn.commit()
        conn.close()
        print(f"Sincronização concluída com de-duplicação:")
        print(f" - Novos chamados inseridos: {novos_inseridos}")
        print(f" - Chamados pendentes atualizados: {atualizados}")
        print(f" - Chamados tratados preservados (não alterados): {pula_tratados}")

if __name__ == "__main__":
    sync_db()
