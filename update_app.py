import re
from pathlib import Path

app_path = Path(__file__).parent / "app.py"

with open(app_path, "r", encoding="utf-8") as f:
    content = f.read()

# Novo código para substituir as funções de API
new_api_code = """
import sqlite3

def get_db_connection():
    db_path = Path(__file__).parent / "petlove.db"
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    return conn

def dict_from_row(row):
    d = dict(row)
    # Renomear as chaves de volta para o formato esperado pelo frontend
    mapping = {
        "Descri\\xe7\\xe3o_da_Reclama\\xe7\\xe3o": "Descri\\xe7\\xe3o da Reclama\\xe7\\xe3o",
        "Data_do_Carregamento": "Data do Carregamento",
        "ID_do_Pedido": "ID_do_Pedido",
        "Motorista": "Motorista",
        "Regional_2": "Regional_2",
        "Rota": "Rota",
        "Placa_do_ve\\xedculo": "Placa do ve\\xedculo",
        "Valor": "Valor",
        "Justificativa": "Justificativa",
        "Descri\\xe7\\xe3o_da_Diverg\\xeancia": "Descri\\xe7\\xe3o da Diverg\\xeancia",
        "Proced\\xeancia": "Proced\\xeancia",
        "Responsavel": "Responsavel",
        "Tratativa": "Tratativa",
        "Criado_por": "Criado por"
    }
    # Decodificar chaves corrompidas se necessário
    new_d = {}
    for k, v in d.items():
        # Vamos tentar associar pelo prefixo ou simplesmente substituir os underscores
        new_key = k.replace("_", " ")
        if "ID do Pedido" in new_key: new_key = "ID_do_Pedido"
        if "Regional 2" in new_key: new_key = "Regional_2"
        if "Criado por" in new_key: new_key = "Criado por"
        new_d[new_key] = v
    return new_d

@app.get("/api/fluxo/buscar/<pedido>")
@login_required
def fluxo_buscar(pedido):
    try:
        conn = get_db_connection()
        row = conn.execute("SELECT * FROM chamados WHERE ID_do_Pedido = ?", (str(pedido).strip(),)).fetchone()
        conn.close()
        
        if row:
            linha = dict_from_row(row)
            try:
                from exportar_petlove import buscar_lista_por_pedido
                linha['Lista_Entrega_Cruzada'] = buscar_lista_por_pedido(pedido) or "Não encontrada"
            except:
                linha['Lista_Entrega_Cruzada'] = "Erro"
            return jsonify(linha)
            
        return jsonify(erro="Pedido não encontrado no fluxo."), 404
    except Exception as e:
        return jsonify(erro=str(e)), 500

@app.post("/api/fluxo/atualizar")
@login_required
def fluxo_atualizar():
    dados = request.get_json(silent=True) or {}
    pedido = str(dados.get("pedido", "")).strip()
    
    if not pedido:
        return jsonify(erro="Pedido não informado."), 400
        
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        
        cursor.execute("SELECT * FROM chamados LIMIT 1")
        col_names = [description[0] for description in cursor.description]
        
        col_justificativa = next((c for c in col_names if 'Justificativa' in c), 'Justificativa')
        col_procedencia = next((c for c in col_names if 'Proced' in c), 'Procedência')
        col_tratativa = next((c for c in col_names if 'Tratativa' in c), 'Tratativa')
        col_responsavel = next((c for c in col_names if 'Responsavel' in c), 'Responsavel')
        col_divergencia = next((c for c in col_names if 'Diverg' in c), 'Descrição_da_Divergência')
        col_valor = next((c for c in col_names if 'Valor' in c), 'Valor')
        
        sql = f\"\"\"
            UPDATE chamados 
            SET "{col_justificativa}" = ?,
                "{col_procedencia}" = ?,
                "{col_tratativa}" = ?,
                "{col_responsavel}" = ?,
                "{col_divergencia}" = ?,
                "{col_valor}" = ?
            WHERE ID_do_Pedido = ?
        \"\"\"
        
        cursor.execute(sql, (
            dados.get("justificativa", ""),
            dados.get("procedencia", ""),
            dados.get("tratativa", ""),
            dados.get("responsavel", ""),
            dados.get("divergencia", ""),
            dados.get("valor", ""),
            pedido
        ))
        
        if cursor.rowcount == 0:
            conn.close()
            return jsonify(erro="Pedido não encontrado no fluxo."), 404
            
        conn.commit()
        conn.close()
        return jsonify(sucesso=True)
    except Exception as e:
        return jsonify(erro=str(e)), 500

@app.post("/api/fluxo/novo")
@login_required
def fluxo_novo():
    dados = request.get_json(silent=True) or {}
    pedido = dados.get("ID_do_Pedido")
    if not pedido:
        return jsonify(erro="ID do Pedido é obrigatório."), 400
        
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        
        cursor.execute("SELECT * FROM chamados LIMIT 1")
        col_names = [description[0] for description in cursor.description]
        
        placeholders = ", ".join(["?"] * len(col_names))
        sql = f"INSERT INTO chamados VALUES ({placeholders})"
        
        valores = []
        for col in col_names:
            key_space = col.replace("_", " ")
            if "ID do Pedido" in key_space: key_space = "ID_do_Pedido"
            if "Regional 2" in key_space: key_space = "Regional_2"
            
            valor = dados.get(key_space, dados.get(col, ""))
            valores.append(str(valor).strip())
            
        cursor.execute(sql, valores)
        conn.commit()
        conn.close()
        return jsonify(sucesso=True)
    except Exception as e:
        return jsonify(erro=str(e)), 500

@app.get("/api/fluxo/todos")
@login_required
def fluxo_todos():
    try:
        conn = get_db_connection()
        rows = conn.execute("SELECT * FROM chamados").fetchall()
        conn.close()
        
        linhas = [dict_from_row(row) for row in rows]
        
        perfil = session.get('perfil', 'Usuario')
        nome_usuario = session.get('nome', '')
        
        if perfil != 'Admin':
            linhas = [linha for linha in linhas if str(linha.get('Criado por', '')).strip().lower() == str(nome_usuario).strip().lower()]
            
        return jsonify(linhas)
    except Exception as e:
        return jsonify(erro=str(e)), 500

if __name__ == "__main__":
"""

# Replace
pattern = re.compile(r'@app\.get\("/api/fluxo/buscar/<pedido>"\).*?if __name__ == "__main__":', re.DOTALL)
match = pattern.search(content)

if match:
    new_content = content.replace(match.group(0), new_api_code.strip() + "\\n\\nif __name__ == \\\"__main__\\\":")
    with open(app_path, "w", encoding="utf-8") as f:
        f.write(new_content)
    print("app.py atualizado com sucesso!")
else:
    print("Match not found.")
