import os
from pathlib import Path
from uuid import uuid4

from flask import Flask, jsonify, render_template, request, send_from_directory, session, redirect, url_for
from functools import wraps
import csv
from exportar_petlove import consultar_lista, salvar_excel

BASE_DADOS_DIR = Path(os.getenv("BASE_DADOS_DIR", Path(__file__).parent / "Base de dados"))
DB_PATH = Path(os.getenv("DATABASE_PATH", Path(__file__).parent / "petlove.db"))

app = Flask(__name__)
app.secret_key = "petlove_jm_secret_key_2026"

def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'usuario' not in session:
            return redirect(url_for('login', proxima=request.path))
        return f(*args, **kwargs)
    return decorated_function

def buscar_lista_por_pedido(pedido):
    base_dir = BASE_DADOS_DIR
    if not base_dir.exists():
        return None
    
    # Busca em arquivos CSV
    for arquivo in base_dir.glob("*.csv"):
        with open(arquivo, newline='', encoding='utf-8-sig', errors='ignore') as f:
            leitor = csv.DictReader(f, delimiter=';')
            # Se não encontrar a coluna, tenta vírgula
            if not leitor.fieldnames or 'Pedido' not in leitor.fieldnames:
                f.seek(0)
                leitor = csv.DictReader(f, delimiter=',')
            if leitor.fieldnames and 'Pedido' in leitor.fieldnames:
                for linha in leitor:
                    if str(linha.get("Pedido", "")).strip() == pedido:
                        return str(linha.get("Lista_Entrega", "")).strip()
    
    # Busca em arquivos XLSX
    try:
        from openpyxl import load_workbook
        for arquivo in base_dir.glob("*.xlsx"):
            wb = load_workbook(arquivo, data_only=True, read_only=True)
            for sheet in wb.sheetnames:
                ws = wb[sheet]
                for idx, row in enumerate(ws.iter_rows(values_only=True)):
                    if idx == 0:
                        headers = [str(cell).strip() if cell else "" for cell in row]
                        continue
                    if 'Pedido' in headers and 'Lista_Entrega' in headers:
                        idx_pedido = headers.index('Pedido')
                        idx_lista = headers.index('Lista_Entrega')
                        if str(row[idx_pedido]).strip() == pedido:
                            return str(row[idx_lista]).strip()
    except Exception:
        pass

    return None

EXPORT_DIR = Path(__file__).parent / "exports"
EXPORT_DIR.mkdir(exist_ok=True)


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        email = request.form.get("email", "").strip()
        senha = request.form.get("senha", "").strip()
        
        csv_path = BASE_DADOS_DIR / "usuarios.csv"
        if csv_path.exists():
            with open(csv_path, newline='', encoding='utf-8', errors='ignore') as f:
                leitor = csv.DictReader(f)
                for user in leitor:
                    if user.get("Email") == email and user.get("Senha") == senha:
                        session['usuario'] = email
                        session['nome'] = user.get("Nome")
                        session['perfil'] = user.get("Perfil")
                        return redirect(request.args.get("proxima") or url_for('index'))
                        
        return render_template("login.html", erro="Credenciais inválidas")
        
    return render_template("login.html")

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for('login'))

@app.post("/api/solicitar-acesso")
def solicitar_acesso():
    import random
    import string
    dados = request.get_json(silent=True) or {}
    email = dados.get("email", "").strip()
    
    if not email or "@" not in email:
        return jsonify(erro="E-mail inválido."), 400
        
    csv_path = BASE_DADOS_DIR / "usuarios.csv"
    
    # Verifica se já existe
    if csv_path.exists():
        with open(csv_path, newline='', encoding='utf-8', errors='ignore') as f:
            leitor = csv.DictReader(f)
            for row in leitor:
                if row.get("Email") == email:
                    return jsonify(erro="Este e-mail já possui cadastro."), 400
                    
    # Gera senha de 6 caracteres aleatórios
    senha_gerada = ''.join(random.choices(string.ascii_letters + string.digits, k=6))
    
    # Extrai o nome a partir do email (antes do @)
    nome = email.split("@")[0].replace(".", " ").title()
    
    # Adiciona ao CSV
    modo = 'a' if csv_path.exists() else 'w'
    with open(csv_path, modo, newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        if modo == 'w':
            writer.writerow(["Email", "Senha", "Perfil", "Nome"])
        writer.writerow([email, senha_gerada, "Usuario", nome])
        
    return jsonify(sucesso=True, senha=senha_gerada)


@app.get("/")
@login_required
def index():
    return render_template("index.html", login_padrao=os.getenv("TMSLOG_LOGIN", ""), usuario=session.get('nome'), perfil=session.get('perfil'))


@app.post("/consultar")
def consultar():
    dados = request.get_json(silent=True) or {}
    numero_pedido = str(dados.get("pedido", "")).strip()
    numero_lista = str(dados.get("numero", "")).strip()
    login = str(dados.get("login", "")).strip()
    senha = str(dados.get("senha", "")).strip()
    if not login or not senha:
        return jsonify(erro="Preencha o login e senha corretamente."), 400
    if not numero_pedido and not numero_lista:
        return jsonify(erro="Preencha o Número do Pedido OU o Número da Lista."), 400
    try:
        if numero_pedido:
            lista_entrega = buscar_lista_por_pedido(numero_pedido)
            if not lista_entrega:
                return jsonify(erro=f"Pedido {numero_pedido} não encontrado na Base de dados. Verifique se o arquivo CSV/Excel está na pasta."), 404
        else:
            lista_entrega = numero_lista

        tabelas = consultar_lista(lista_entrega, login, senha)
        nome = f"lista_{lista_entrega}_{uuid4().hex[:8]}.xlsx"
        nomes = ["Resumo da carga", "Relação da carga"]
        salvar_excel(tabelas, EXPORT_DIR / nome, nomes)
        return jsonify(
            numero=numero_pedido or numero_lista,
            lista=lista_entrega,
            tabelas=tabelas,
            nomes_tabelas=nomes,
            download=f"/baixar/{nome}",
        )
    except Exception as exc:
        mensagem = str(exc)
        if "usuário ou senha inválido" in mensagem.lower():
            mensagem += " Verifique maiúsculas e minúsculas; a senha fornecida inicialmente possui 6 caracteres."
        return jsonify(erro=mensagem), 502


@app.get("/baixar/<path:nome>")
def baixar(nome):
    return send_from_directory(EXPORT_DIR, nome, as_attachment=True)


@app.get("/fluxo")
@login_required
def fluxo():
    return render_template("fluxo.html", login_padrao=os.getenv("TMSLOG_LOGIN", ""), usuario=session.get('nome'), perfil=session.get('perfil'))

import sqlite3

def get_db_connection():
    db_path = DB_PATH
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    return conn

def dict_from_row(row):
    d = dict(row)
    # Renomear as chaves de volta para o formato esperado pelo frontend
    mapping = {
        "Descri\xe7\xe3o_da_Reclama\xe7\xe3o": "Descri\xe7\xe3o da Reclama\xe7\xe3o",
        "Data_do_Carregamento": "Data do Carregamento",
        "ID_do_Pedido": "ID_do_Pedido",
        "Motorista": "Motorista",
        "Regional_2": "Regional_2",
        "Rota": "Rota",
        "Placa_do_ve\xedculo": "Placa do ve\xedculo",
        "Valor": "Valor",
        "Justificativa": "Justificativa",
        "Descri\xe7\xe3o_da_Diverg\xeancia": "Descri\xe7\xe3o da Diverg\xeancia",
        "Proced\xeancia": "Proced\xeancia",
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
        
        sql = f"""
            UPDATE chamados 
            SET "{col_justificativa}" = ?,
                "{col_procedencia}" = ?,
                "{col_tratativa}" = ?,
                "{col_responsavel}" = ?,
                "{col_divergencia}" = ?,
                "{col_valor}" = ?
            WHERE ID_do_Pedido = ?
        """
        
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

@app.get("/nova-acareacao")
@login_required
def nova_acareacao():
    return render_template("nova.html", usuario=session.get('nome'), perfil=session.get('perfil'))

@app.get("/gestao")
@login_required
def gestao():
    return render_template("gestao.html", usuario=session.get('nome'), perfil=session.get('perfil'))

@app.get("/dashboard")
@login_required
def dashboard():
    return render_template("dashboard.html", usuario=session.get('nome'), perfil=session.get('perfil'))

if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=False)
