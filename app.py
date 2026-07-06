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
    conn = sqlite3.connect(db_path, timeout=30.0)
    conn.row_factory = sqlite3.Row
    return conn

# Criar tabela de recuperação de senha se não existir
try:
    conn = get_db_connection()
    conn.execute('''
        CREATE TABLE IF NOT EXISTS recuperacao_senha (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            email TEXT,
            token TEXT,
            senha_provisoria TEXT,
            data_criacao TEXT
        )
    ''')
    conn.commit()
    conn.close()
except Exception as e:
    print(f"Erro ao inicializar tabela recuperacao_senha: {e}")

def atualizar_senha_csv(email, nova_senha):
    csv_path = BASE_DADOS_DIR / "usuarios.csv"
    if not csv_path.exists():
        return False
    
    linhas = []
    atualizado = False
    campos = ["Email", "Senha", "Perfil", "Nome"] # colunas padrão
    
    with open(csv_path, mode='r', newline='', encoding='utf-8', errors='ignore') as f:
        leitor = csv.DictReader(f)
        if leitor.fieldnames:
            campos = leitor.fieldnames
        for row in leitor:
            if row.get("Email") == email:
                row["Senha"] = nova_senha
                atualizado = True
            linhas.append(row)
            
    if atualizado:
        with open(csv_path, mode='w', newline='', encoding='utf-8') as f:
            escritor = csv.DictWriter(f, fieldnames=campos)
            escritor.writeheader()
            for row in linhas:
                escritor.writerow(row)
        return True
    return False

import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

def enviar_email_recuperacao(email, senha_provisoria, link_confirmacao):
    smtp_server = os.getenv("SMTP_SERVER", "smtp.gmail.com")
    try:
        smtp_port = int(os.getenv("SMTP_PORT", "587"))
    except:
        smtp_port = 587
    smtp_user = os.getenv("SMTP_USER", "")
    smtp_password = os.getenv("SMTP_PASSWORD", "")
    smtp_from = os.getenv("SMTP_FROM", smtp_user or "noreply@jmdistribuicao.com.br")

    assunto = "Recuperacao de Senha - Portal Petlove JM"
    corpo = f"""Olá,

Foi solicitada a recuperação de senha para sua conta no Portal Petlove JM.

Sua senha temporária gerada é: {senha_provisoria}

Para confirmar a recuperação de senha e redefinir para sua senha desejada, clique no link abaixo:
{link_confirmacao}

Atenção: Este link é obrigatório para ativar seu acesso e definir sua senha definitiva.

Se você não solicitou esta recuperação, por favor desconsidere este e-mail.

Atenciosamente,
Equipe JM Distribuição
"""

    if not smtp_user or not smtp_password:
        print("\n=== [SIMULACAO DE ENVIO DE E-MAIL] ===")
        print(f"Para: {email}")
        print(f"De: {smtp_from}")
        print(f"Assunto: {assunto}")
        print(f"Conteúdo:\n{corpo}")
        print("======================================\n")
        return True, "Simulado no console (sem credenciais SMTP configuradas)."

    try:
        msg = MIMEMultipart()
        msg['From'] = smtp_from
        msg['To'] = email
        msg['Subject'] = assunto
        msg.attach(MIMEText(corpo, 'plain', 'utf-8'))

        server = smtplib.SMTP(smtp_server, smtp_port, timeout=10)
        server.starttls()
        server.login(smtp_user, smtp_password)
        server.send_message(msg)
        server.quit()
        return True, "E-mail enviado com sucesso."
    except Exception as e:
        print(f"Erro ao enviar e-mail real: {e}")
        print("\n=== [FALLBACK - SIMULACAO DE E-MAIL] ===")
        print(f"Para: {email}")
        print(f"Conteúdo:\n{corpo}")
        print("========================================\n")
        return False, str(e)

@app.before_request
def verificar_forcar_redefinicao():
    if session.get('forcar_redefinicao'):
        # Rotas permitidas quando forçando redefinição
        rotas_permitidas = ['definir_senha', 'logout', 'static']
        if request.endpoint and request.endpoint not in rotas_permitidas:
            return redirect(url_for('definir_senha'))

@app.post("/api/esqueci-senha")
def api_esqueci_senha():
    import random
    import string
    import datetime
    dados = request.get_json(silent=True) or {}
    email = dados.get("email", "").strip()
    
    if not email or "@" not in email:
        return jsonify(erro="E-mail inválido."), 400
        
    csv_path = BASE_DADOS_DIR / "usuarios.csv"
    email_cadastrado = False
    
    # Verifica se e-mail está cadastrado
    if csv_path.exists():
        with open(csv_path, newline='', encoding='utf-8', errors='ignore') as f:
            leitor = csv.DictReader(f)
            for row in leitor:
                if row.get("Email") == email:
                    email_cadastrado = True
                    break
                    
    if not email_cadastrado:
        return jsonify(erro="Este e-mail não está cadastrado no sistema."), 404
        
    # Gera senha de 6 caracteres aleatórios
    senha_provisoria = ''.join(random.choices(string.ascii_letters + string.digits, k=6))
    token = str(uuid4())
    data_criacao = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    # Salva no SQLite
    try:
        conn = get_db_connection()
        conn.execute(
            "INSERT INTO recuperacao_senha (email, token, senha_provisoria, data_criacao) VALUES (?, ?, ?, ?)",
            (email, token, senha_provisoria, data_criacao)
        )
        conn.commit()
        conn.close()
    except Exception as e:
        return jsonify(erro=f"Erro ao salvar solicitação: {str(e)}"), 500
        
    # Cria o link de confirmação
    url_base = request.host_url.rstrip('/')
    link_confirmacao = f"{url_base}/confirmar-recuperacao?email={email}&token={token}"
    
    sucesso, msg = enviar_email_recuperacao(email, senha_provisoria, link_confirmacao)
    
    return jsonify(sucesso=True, mensagem=msg)

@app.route("/confirmar-recuperacao")
def confirmar_recuperacao():
    email = request.args.get("email", "").strip()
    token = request.args.get("token", "").strip()
    
    if not email or not token:
        return render_template("login.html", erro="Link de confirmação inválido ou incompleto.")
        
    try:
        conn = get_db_connection()
        row = conn.execute("SELECT * FROM recuperacao_senha WHERE email = ? AND token = ?", (email, token)).fetchone()
        
        if not row:
            conn.close()
            return render_template("login.html", erro="Token de confirmação inválido ou já utilizado.")
            
        senha_provisoria = row["senha_provisoria"]
        
        # Atualiza a senha no CSV para a provisória
        atualizar_senha_csv(email, senha_provisoria)
        
        # Busca perfil e nome no CSV para montar a sessão
        nome = email.split("@")[0].replace(".", " ").title()
        perfil = "Usuario"
        csv_path = BASE_DADOS_DIR / "usuarios.csv"
        if csv_path.exists():
            with open(csv_path, newline='', encoding='utf-8', errors='ignore') as f:
                leitor = csv.DictReader(f)
                for user in leitor:
                    if user.get("Email") == email:
                        nome = user.get("Nome", nome)
                        perfil = user.get("Perfil", perfil)
                        break
                        
        # Deleta a entrada de recuperação
        conn.execute("DELETE FROM recuperacao_senha WHERE email = ? AND token = ?", (email, token))
        conn.commit()
        conn.close()
        
        # Loga o usuário temporariamente e obriga a redefinir a senha
        session['usuario'] = email
        session['nome'] = nome
        session['perfil'] = perfil
        session['forcar_redefinicao'] = True
        
        return redirect(url_for('definir_senha'))
        
    except Exception as e:
        return render_template("login.html", erro=f"Erro interno ao confirmar recuperação: {str(e)}")

@app.route("/definir-senha", methods=["GET", "POST"])
def definir_senha():
    if not session.get('forcar_redefinicao'):
        return redirect(url_for('login'))
        
    email = session.get('usuario')
    
    if request.method == "POST":
        nova_senha = request.form.get("nova_senha", "").strip()
        confirmacao = request.form.get("confirmacao", "").strip()
        
        if not nova_senha:
            return render_template("definir_senha.html", erro="A senha não pode ser vazia.")
            
        if nova_senha != confirmacao:
            return render_template("definir_senha.html", erro="As senhas não coincidem.")
            
        # Atualiza a senha no CSV para a definitiva
        if atualizar_senha_csv(email, nova_senha):
            # Limpa flag de redefinição
            session.pop('forcar_redefinicao', None)
            return redirect(url_for('index'))
        else:
            return render_template("definir_senha.html", erro="Erro ao atualizar a senha no banco de dados.")
            
    return render_template("definir_senha.html")

@app.route("/usuarios")
@login_required
def usuarios_dashboard():
    if session.get("usuario") != "admin@jm.com":
        return render_template("login.html", erro="Acesso não autorizado. Apenas o administrador principal (admin@jm.com) pode gerenciar usuários."), 403
    return render_template("usuarios.html", usuario=session.get('nome'), perfil=session.get('perfil'))

@app.route("/api/usuarios")
@login_required
def api_listar_usuarios():
    if session.get("usuario") != "admin@jm.com":
        return jsonify(erro="Acesso negado. Apenas o administrador principal pode realizar esta ação."), 403
        
    csv_path = BASE_DADOS_DIR / "usuarios.csv"
    if not csv_path.exists():
        return jsonify([])
        
    usuarios = []
    with open(csv_path, newline='', encoding='utf-8', errors='ignore') as f:
        leitor = csv.DictReader(f)
        for row in leitor:
            usuarios.append({
                "Email": row.get("Email"),
                "Nome": row.get("Nome"),
                "Perfil": row.get("Perfil")
            })
    return jsonify(usuarios)

@app.route("/api/usuarios/salvar", methods=["POST"])
@login_required
def api_salvar_usuario():
    if session.get("usuario") != "admin@jm.com":
        return jsonify(erro="Acesso negado. Apenas o administrador principal pode realizar esta ação."), 403
        
    dados = request.get_json(silent=True) or {}
    email = dados.get("email", "").strip()
    nome = dados.get("nome", "").strip()
    perfil = dados.get("perfil", "").strip()
    senha = dados.get("senha", "").strip()
    
    if not email or not nome or not perfil:
        return jsonify(erro="Preencha os campos obrigatórios (E-mail, Nome e Perfil)."), 400
        
    if perfil not in ["Admin", "Usuario"]:
        return jsonify(erro="Perfil inválido."), 400
        
    csv_path = BASE_DADOS_DIR / "usuarios.csv"
    
    linhas = []
    editado = False
    campos = ["Email", "Senha", "Perfil", "Nome"]
    
    if csv_path.exists():
        with open(csv_path, newline='', encoding='utf-8', errors='ignore') as f:
            leitor = csv.DictReader(f)
            if leitor.fieldnames:
                campos = leitor.fieldnames
            for row in leitor:
                if row.get("Email") == email:
                    row["Nome"] = nome
                    row["Perfil"] = perfil
                    if senha:
                        row["Senha"] = senha
                    editado = True
                linhas.append(row)
                
    if not editado:
        # Novo usuário. Senha é obrigatória para novos usuários
        if not senha:
            return jsonify(erro="A senha é obrigatória para novos usuários."), 400
        linhas.append({
            "Email": email,
            "Senha": senha,
            "Perfil": perfil,
            "Nome": nome
        })
        
    # Grava de volta no CSV
    try:
        with open(csv_path, mode='w', newline='', encoding='utf-8') as f:
            escritor = csv.DictWriter(f, fieldnames=campos)
            escritor.writeheader()
            for row in linhas:
                escritor.writerow(row)
        return jsonify(sucesso=True)
    except Exception as e:
        return jsonify(erro=f"Erro ao salvar usuário: {str(e)}"), 500

@app.route("/api/usuarios/deletar", methods=["POST"])
@login_required
def api_deletar_usuario():
    if session.get("usuario") != "admin@jm.com":
        return jsonify(erro="Acesso negado. Apenas o administrador principal pode realizar esta ação."), 403
        
    dados = request.get_json(silent=True) or {}
    email = dados.get("email", "").strip()
    
    if not email:
        return jsonify(erro="E-mail do usuário não informado."), 400
        
    # Impedir que o administrador delete a si mesmo
    if email == session.get("usuario"):
        return jsonify(erro="Você não pode excluir sua própria conta."), 400
        
    csv_path = BASE_DADOS_DIR / "usuarios.csv"
    if not csv_path.exists():
        return jsonify(erro="Arquivo de usuários não encontrado."), 404
        
    linhas = []
    deletado = False
    campos = ["Email", "Senha", "Perfil", "Nome"]
    
    with open(csv_path, newline='', encoding='utf-8', errors='ignore') as f:
        leitor = csv.DictReader(f)
        if leitor.fieldnames:
            campos = leitor.fieldnames
        for row in leitor:
            if row.get("Email") == email:
                deletado = True
                continue
            linhas.append(row)
            
    if not deletado:
        return jsonify(erro="Usuário não encontrado."), 404
        
    try:
        with open(csv_path, mode='w', newline='', encoding='utf-8') as f:
            escritor = csv.DictWriter(f, fieldnames=campos)
            escritor.writeheader()
            for row in linhas:
                escritor.writerow(row)
        return jsonify(sucesso=True)
    except Exception as e:
        return jsonify(erro=f"Erro ao excluir usuário: {str(e)}"), 500

def sincronizar_pasta_tms():
    db_path = DB_PATH
    tms_dir = BASE_DADOS_DIR / "Relatorio TMS"
    
    if not tms_dir.exists():
        return 0, 0, "Pasta Relatorio TMS não encontrada."
        
    csv_files = list(tms_dir.glob("*.csv"))
    if not csv_files:
        return 0, 0, "Nenhum arquivo CSV encontrado na pasta Relatorio TMS."
        
    conn = sqlite3.connect(db_path, timeout=30.0)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    
    # Obter colunas reais da tabela chamados
    cursor.execute("SELECT * FROM chamados LIMIT 1")
    colunas_db = [description[0] for description in cursor.description]
    
    # Obter todos os chamados existentes na base
    cursor.execute("SELECT * FROM chamados")
    rows = cursor.fetchall()
    
    col_id_db = next((c for c in colunas_db if 'ID_do_Pedido' in c or 'ID_Pedido' in c), 'ID_do_Pedido')
    
    chamados_existentes = {}
    for r in rows:
        d = dict(r)
        key_map = {}
        for k in d.keys():
            k_clean = k.replace("_", " ").lower()
            if "id do pedido" in k_clean or "id_pedido" in k_clean:
                key_map["id"] = k
            elif "lista entrega cruzada" in k_clean or "lista_entrega_cruzada" in k_clean:
                key_map["lista"] = k
            elif "motorista" in k_clean:
                key_map["motorista"] = k
            elif "regional" in k_clean:
                key_map["regional"] = k
            elif "data do carregamento" in k_clean or "data_do_carregamento" in k_clean:
                key_map["data_carr"] = k
            elif "rota" in k_clean:
                key_map["rota"] = k
            elif "descricao da divergencia" in k_clean or "descrição da divergência" in k_clean or "descrição_da_divergência" in k_clean:
                key_map["divergencia"] = k
                
        chamados_existentes[str(d.get(key_map.get("id", col_id_db), "")).strip()] = (d, key_map)
        
    total_atualizados = 0
    total_lidos = 0
    
    for csv_file in csv_files:
        delimiter = ','
        with open(csv_file, newline='', encoding='utf-8-sig', errors='ignore') as f:
            first_line = f.readline()
            if first_line.count(';') > first_line.count(','):
                delimiter = ';'
            f.seek(0)
            
            leitor = csv.DictReader(f, delimiter=delimiter)
            if not leitor.fieldnames:
                continue
                
            for row in leitor:
                total_lidos += 1
                col_pedido = next((c for c in leitor.fieldnames if 'Pedido' in c), 'Pedido')
                pedido_id = str(row.get(col_pedido, "")).strip()
                
                if not pedido_id or pedido_id not in chamados_existentes:
                    continue
                    
                # O pedido existe na base! Vamos atualizar campos que estiverem em branco.
                chamado_dict, key_map = chamados_existentes[pedido_id]
                
                # Preparar valores do TMS
                data_carr_raw = row.get("Data_do_Carregamento_Lista", "")
                data_carr = data_carr_raw.split(" ")[0].strip() if " " in data_carr_raw else data_carr_raw.strip()
                
                motorista_raw = row.get("Motorista_Lista", "")
                motorista = motorista_raw.split(" - ", 1)[1].strip() if " - " in motorista_raw else motorista_raw.strip()
                
                filial_raw = row.get("Filial_Entrega", "")
                regional = filial_raw.strip()
                if "São Paulo" in filial_raw: regional = "JM SP"
                elif "Barueri" in filial_raw: regional = "JM BAR"
                elif "Santos" in filial_raw: regional = "JM SSZ"
                
                lista_entrega = row.get("Lista_Entrega", "").strip()
                ocorrencia = row.get("Ultima_Ocorrencia", "").strip()
                
                updates = {}
                
                col_lista = key_map.get("lista")
                if col_lista and not str(chamado_dict.get(col_lista, "")).strip():
                    updates[col_lista] = lista_entrega
                    
                col_mot = key_map.get("motorista")
                if col_mot and not str(chamado_dict.get(col_mot, "")).strip():
                    updates[col_mot] = motorista
                    
                col_reg = key_map.get("regional")
                if col_reg and not str(chamado_dict.get(col_reg, "")).strip():
                    updates[col_reg] = regional
                    
                col_data = key_map.get("data_carr")
                if col_data and not str(chamado_dict.get(col_data, "")).strip():
                    updates[col_data] = data_carr
                    
                col_rota = key_map.get("rota")
                if col_rota and not str(chamado_dict.get(col_rota, "")).strip():
                    updates[col_rota] = row.get("Rota_Entrega", "").strip()
                    
                col_diverg = key_map.get("divergencia")
                if col_diverg and not str(chamado_dict.get(col_diverg, "")).strip():
                    updates[col_diverg] = ocorrencia
                    
                if updates:
                    set_clause = ", ".join([f'"{k}" = ?' for k in updates.keys()])
                    sql = f'UPDATE chamados SET {set_clause} WHERE "{col_id_db}" = ?'
                    params = list(updates.values()) + [pedido_id]
                    cursor.execute(sql, params)
                    
                    # Registrar histórico
                    import datetime
                    data_hora = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                    for k, v in updates.items():
                        cursor.execute("""
                            INSERT INTO historico_chamados (pedido_id, data_hora, usuario, campo, valor_antigo, valor_novo)
                            VALUES (?, ?, ?, ?, ?, ?)
                        """, (pedido_id, data_hora, "Carga TMS", f"Atualizado: {k}", "", v))
                        
                    total_atualizados += 1
                    for k, v in updates.items():
                        chamado_dict[k] = v
                        
    conn.commit()
    conn.close()
    return total_atualizados, total_lidos, None

@app.route("/api/tms/sincronizar", methods=["POST"])
@login_required
def api_sincronizar_tms():
    if session.get("usuario") != "admin@jm.com":
        return jsonify(erro="Acesso negado. Apenas o administrador principal pode realizar esta ação."), 403
        
    try:
        total_novos, total_lidos, erro = sincronizar_pasta_tms()
        if erro:
            return jsonify(erro=erro), 400
        return jsonify(sucesso=True, total_novos=total_novos, total_lidos=total_lidos)
    except Exception as e:
        return jsonify(erro=str(e)), 500

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
                # Se a coluna existir e tiver valor no banco, usa ele. Caso contrário busca no CSV.
                colunas = row.keys() if hasattr(row, 'keys') else []
                if 'Lista_Entrega_Cruzada' in colunas and row['Lista_Entrega_Cruzada'] and str(row['Lista_Entrega_Cruzada']).strip():
                    linha['Lista_Entrega_Cruzada'] = str(row['Lista_Entrega_Cruzada']).strip()
                else:
                    linha['Lista_Entrega_Cruzada'] = buscar_lista_por_pedido(pedido) or "Não encontrada"
            except Exception as e:
                # Tenta chamar a função local buscando no CSV
                try:
                    linha['Lista_Entrega_Cruzada'] = buscar_lista_por_pedido(pedido) or "Não encontrada"
                except:
                    linha['Lista_Entrega_Cruzada'] = "Erro"
            
            # Adicionar histórico de modificações do pedido
            try:
                conn_hist = get_db_connection()
                hist_rows = conn_hist.execute("SELECT data_hora, usuario, campo, valor_antigo, valor_novo FROM historico_chamados WHERE pedido_id = ? ORDER BY id DESC", (str(pedido).strip(),)).fetchall()
                conn_hist.close()
                linha['Historico'] = [dict(r) for r in hist_rows]
            except Exception as e:
                linha['Historico'] = []
                
            return jsonify(linha)
            
        return jsonify(erro="Pedido não encontrado no fluxo."), 404
    except Exception as e:
        return jsonify(erro=str(e)), 500

@app.post("/api/fluxo/atualizar")
@login_required
def fluxo_atualizar():
    dados = request.get_json(silent=True) or {}
    pedido = str(dados.get("pedido", "")).strip()
    usuario = session.get('nome', 'Usuário')
    
    if not pedido:
        return jsonify(erro="Pedido não informado."), 400
        
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        
        # 1. Buscar valores antigos para o histórico
        cursor.execute("SELECT * FROM chamados WHERE ID_do_Pedido = ?", (pedido,))
        row_antiga = cursor.fetchone()
        if not row_antiga:
            conn.close()
            return jsonify(erro="Pedido não encontrado no fluxo."), 404
            
        col_names = [description[0] for description in cursor.description]
        col_justificativa = next((c for c in col_names if 'Justificativa' in c), 'Justificativa')
        col_procedencia = next((c for c in col_names if 'Proced' in c), 'Procedência')
        col_tratativa = next((c for c in col_names if 'Tratativa' in c), 'Tratativa')
        col_responsavel = next((c for c in col_names if 'Responsavel' in c), 'Responsavel')
        col_divergencia = next((c for c in col_names if 'Diverg' in c), 'Descrição_da_Divergência')
        col_valor = next((c for c in col_names if 'Valor' in c), 'Valor')
        col_status = next((c for c in col_names if 'Status' in c), 'Status_da_Tratativa')
        
        valores_antigos = dict(row_antiga)
        
        nova_justif = dados.get("justificativa", "")
        nova_proc = dados.get("procedencia", "")
        nova_trat = dados.get("tratativa", "")
        novo_resp = dados.get("responsavel", "")
        nova_diverg = dados.get("divergencia", "")
        novo_valor = dados.get("valor", "")
        
        proc_clean = nova_proc.strip().lower() if nova_proc else ""
        if proc_clean and proc_clean != "em analise" and nova_trat.strip():
            novo_status = "Finalizado"
        else:
            novo_status = "Em Andamento"
            
        # 2. Executar o UPDATE
        sql = f"""
            UPDATE chamados 
            SET "{col_justificativa}" = ?,
                "{col_procedencia}" = ?,
                "{col_tratativa}" = ?,
                "{col_responsavel}" = ?,
                "{col_divergencia}" = ?,
                "{col_valor}" = ?,
                "{col_status}" = ?
            WHERE ID_do_Pedido = ?
        """
        
        cursor.execute(sql, (
            nova_justif,
            nova_proc,
            nova_trat,
            novo_resp,
            nova_diverg,
            novo_valor,
            novo_status,
            pedido
        ))
        
        # 3. Registrar modificações no histórico
        import datetime
        data_hora = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        
        def log_mudanca(campo, valor_ant, valor_nov):
            val_ant_str = str(valor_ant or "").strip()
            val_nov_str = str(valor_nov or "").strip()
            if val_ant_str != val_nov_str:
                cursor.execute("""
                    INSERT INTO historico_chamados (pedido_id, data_hora, usuario, campo, valor_antigo, valor_novo)
                    VALUES (?, ?, ?, ?, ?, ?)
                """, (pedido, data_hora, usuario, campo, val_ant_str if val_ant_str else "Vazio", val_nov_str if val_nov_str else "Vazio"))

        log_mudanca("Justificativa", valores_antigos.get(col_justificativa), nova_justif)
        log_mudanca("Procedência", valores_antigos.get(col_procedencia), nova_proc)
        log_mudanca("Tratativa", valores_antigos.get(col_tratativa), nova_trat)
        log_mudanca("Responsável", valores_antigos.get(col_responsavel), novo_resp)
        log_mudanca("Divergência", valores_antigos.get(col_divergencia), nova_diverg)
        log_mudanca("Valor", valores_antigos.get(col_valor), novo_valor)
        log_mudanca("Status da Tratativa", valores_antigos.get(col_status), novo_status)
        
        conn.commit()
        conn.close()
        return jsonify(sucesso=True)
    except Exception as e:
        return jsonify(erro=str(e)), 500

@app.post("/api/fluxo/atualizar_romaneio")
@login_required
def fluxo_atualizar_romaneio():
    dados = request.get_json(silent=True) or {}
    pedido = str(dados.get("pedido", "")).strip()
    novo_romaneio = str(dados.get("romaneio", "")).strip()
    usuario = session.get('nome', 'Usuário')
    
    if not pedido:
        return jsonify(erro="Pedido não informado."), 400
        
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        
        cursor.execute("SELECT * FROM chamados WHERE ID_do_Pedido = ?", (pedido,))
        row = cursor.fetchone()
        if not row:
            conn.close()
            return jsonify(erro="Pedido não encontrado."), 404
            
        col_names = [description[0] for description in cursor.description]
        col_romaneio = next((c for c in col_names if 'Lista_Entrega_Cruzada' in c), 'Lista_Entrega_Cruzada')
        
        valores_antigos = dict(row)
        valor_antigo = valores_antigos.get(col_romaneio) or ""
        
        if str(valor_antigo).strip() == novo_romaneio:
            conn.close()
            return jsonify(sucesso=True)
            
        # Executa atualização
        cursor.execute(f'UPDATE chamados SET "{col_romaneio}" = ? WHERE ID_do_Pedido = ?', (novo_romaneio, pedido))
        
        # Registrar no histórico
        import datetime
        data_hora = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        cursor.execute("""
            INSERT INTO historico_chamados (pedido_id, data_hora, usuario, campo, valor_antigo, valor_novo)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (pedido, data_hora, usuario, "Nº Romaneio", valor_antigo if valor_antigo else "Vazio/Erro", novo_romaneio))
        
        conn.commit()
        conn.close()
        return jsonify(sucesso=True)
    except Exception as e:
        return jsonify(erro=str(e)), 500

def formatar_data_para_iso(dt_str):
    dt_str = str(dt_str).strip()
    if "/" in dt_str:
        partes = dt_str.split("/")
        if len(partes) == 3:
            d, m, y = partes
            if len(d) <= 2 and len(m) <= 2 and len(y) == 4:
                return f"{y}-{m.zfill(2)}-{d.zfill(2)}"
    return dt_str

@app.get("/api/fluxo/autocompletar/<pedido>")
@login_required
def fluxo_autocompletar(pedido):
    pedido = str(pedido).strip()
    resultado = {
        "Motorista": "",
        "Placa_do_veiculo": "",
        "Rota": "",
        "Regional_2": "",
        "Valor": "",
        "Data_do_Carregamento": ""
    }
    
    # 1. Buscar no banco SQLite
    try:
        conn = get_db_connection()
        row = conn.execute("SELECT * FROM chamados WHERE ID_do_Pedido = ?", (pedido,)).fetchone()
        conn.close()
        if row:
            colunas = row.keys() if hasattr(row, 'keys') else []
            col_mot = next((c for c in colunas if 'Motorista' in c), 'Motorista')
            col_placa = next((c for c in colunas if 'Placa' in c), 'Placa_do_veículo')
            col_rota = next((c for c in colunas if 'Rota' in c), 'Rota')
            col_reg = next((c for c in colunas if 'Regional' in c), 'Regional_2')
            col_val = next((c for c in colunas if 'Valor' in c), 'Valor')
            col_data = next((c for c in colunas if 'Data' in c), 'Data_do_Carregamento')
            
            if row[col_mot]: resultado["Motorista"] = str(row[col_mot]).strip()
            if row[col_placa]: resultado["Placa_do_veiculo"] = str(row[col_placa]).strip()
            if row[col_rota]: resultado["Rota"] = str(row[col_rota]).strip()
            if row[col_reg]: resultado["Regional_2"] = str(row[col_reg]).strip()
            if row[col_val]: resultado["Valor"] = str(row[col_val]).strip()
            if row[col_data]: resultado["Data_do_Carregamento"] = formatar_data_para_iso(row[col_data])
            
            if resultado["Motorista"] and resultado["Placa_do_veiculo"]:
                return jsonify(resultado)
    except:
        pass

    # 2. Buscar no Gestão de Perdas -Pet Love.csv
    try:
        csv_perdas = BASE_DADOS_DIR / "Fluxo de aprovação" / "Gestão de Perdas -Pet Love.csv"
        if csv_perdas.exists():
            delimiter = ','
            with open(csv_perdas, newline='', encoding='utf-8-sig', errors='ignore') as f:
                first_line = f.readline()
                if first_line.count(';') > first_line.count(','):
                    delimiter = ';'
                f.seek(0)
                reader = csv.DictReader(f, delimiter=delimiter)
                for row in reader:
                    col_id = next((c for c in reader.fieldnames if 'ID_do_Pedido' in c or 'ID_Pedido' in c), 'ID_do_Pedido')
                    if str(row.get(col_id, "")).strip() == pedido:
                        col_mot = next((c for c in reader.fieldnames if 'Motorista' in c), 'Motorista')
                        col_placa = next((c for c in reader.fieldnames if 'Placa' in c), 'Placa do veículo')
                        col_rota = next((c for c in reader.fieldnames if 'Rota' in c), 'Rota')
                        col_reg = next((c for c in reader.fieldnames if 'Regional' in c), 'Regional_2')
                        col_val = next((c for c in reader.fieldnames if 'Valor' in c), 'Valor')
                        col_data = next((c for c in reader.fieldnames if 'Carregamento' in c or 'Data' in c), 'Data do Carregamento')
                        
                        if row.get(col_mot) and not resultado["Motorista"]: resultado["Motorista"] = str(row[col_mot]).strip()
                        if row.get(col_placa) and not resultado["Placa_do_veiculo"]: resultado["Placa_do_veiculo"] = str(row[col_placa]).strip()
                        if row.get(col_rota) and not resultado["Rota"]: resultado["Rota"] = str(row[col_rota]).strip()
                        if row.get(col_reg) and not resultado["Regional_2"]: resultado["Regional_2"] = str(row[col_reg]).strip()
                        if row.get(col_val) and not resultado["Valor"]: resultado["Valor"] = str(row[col_val]).strip()
                        if row.get(col_data) and not resultado["Data_do_Carregamento"]: resultado["Data_do_Carregamento"] = formatar_data_para_iso(row[col_data])
                        break
    except:
        pass

    # 3. Buscar no Relatorio TMS.csv
    try:
        csv_tms = BASE_DADOS_DIR / "Relatorio TMS.csv"
        if csv_tms.exists():
            delimiter = ','
            with open(csv_tms, newline='', encoding='utf-8-sig', errors='ignore') as f:
                first_line = f.readline()
                if first_line.count(';') > first_line.count(','):
                    delimiter = ';'
                f.seek(0)
                reader = csv.DictReader(f, delimiter=delimiter)
                for row in reader:
                    col_id = next((c for c in reader.fieldnames if 'Pedido' in c), 'Pedido')
                    if str(row.get(col_id, "")).strip() == pedido:
                        col_mot = next((c for c in reader.fieldnames if 'Motorista' in c), 'Motorista_Lista')
                        col_rota = next((c for c in reader.fieldnames if 'Rota' in c), 'Rota_Entrega')
                        col_reg = next((c for c in reader.fieldnames if 'Filial' in c), 'Filial_Entrega')
                        col_data = next((c for c in reader.fieldnames if 'Carregamento' in c), 'Data_do_Carregamento_Lista')
                        
                        if row.get(col_mot) and not resultado["Motorista"]: resultado["Motorista"] = str(row[col_mot]).strip()
                        if row.get(col_rota) and not resultado["Rota"]: resultado["Rota"] = str(row[col_rota]).strip()
                        
                        if row.get(col_reg) and not resultado["Regional_2"]: 
                            reg_raw = str(row[col_reg]).strip()
                            if "São Paulo" in reg_raw: resultado["Regional_2"] = "JM SP"
                            elif "Barueri" in reg_raw: resultado["Regional_2"] = "JM BAR"
                            elif "Santos" in reg_raw: resultado["Regional_2"] = "JM SSZ"
                            else: resultado["Regional_2"] = reg_raw
                            
                        if row.get(col_data) and not resultado["Data_do_Carregamento"]:
                            dt = str(row[col_data]).strip()
                            resultado["Data_do_Carregamento"] = formatar_data_para_iso(dt.split(" ")[0])
                        break
    except:
        pass

    return jsonify(resultado)

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
        
        # Verificar duplicidade
        pedido_clean = str(pedido).strip()
        cursor.execute("SELECT * FROM chamados WHERE ID_do_Pedido = ?", (pedido_clean,))
        row_existente = cursor.fetchone()
        if row_existente:
            linha = dict_from_row(row_existente)
            criador = linha.get("Criado por") or "não informado"
            status = linha.get("Status da Tratativa") or "Em Andamento"
            proc = linha.get("Procedência") or "Em Análise"
            conn.close()
            return jsonify(erro=f"O pedido {pedido_clean} já possui uma solicitação ativa. Cadastrada por: {criador} | Status: {status} | Procedência: {proc}."), 409
            
        cursor.execute("SELECT * FROM chamados LIMIT 1")
        col_names = [description[0] for description in cursor.description]
        
        placeholders = ", ".join(["?"] * len(col_names))
        sql = f"INSERT INTO chamados VALUES ({placeholders})"
        
        valores = []
        for col in col_names:
            if col == "Status_da_Tratativa":
                valores.append("Em Andamento")
                continue
                
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
