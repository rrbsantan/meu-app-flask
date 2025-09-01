import json
from flask import Flask, render_template, request, redirect, url_for, session, send_from_directory
import os
import calendar
import csv

# --- CONFIGURAÇÃO DO APP ---
app = Flask(__name__)
app.secret_key = 'sua_chave_secreta_aqui' # Troque por uma chave mais segura

# --- CAMINHOS ---
BASE_DIR = os.path.abspath(os.path.dirname(__file__))
HISTORICO_ROTAS = os.path.join(BASE_DIR, '..', 'historico_rotas')
ARQUIVO_USUARIOS = os.path.join(BASE_DIR, 'usuarios.csv')

# --- FUNÇÕES AUXILIARES ---

def ler_usuarios():
    """Lê os usuários do arquivo CSV."""
    if not os.path.exists(ARQUIVO_USUARIOS):
        return []
    with open(ARQUIVO_USUARIOS, mode='r', newline='', encoding='utf-8') as csvfile:
        reader = csv.DictReader(csvfile)
        return list(reader)

def escrever_usuarios(usuarios):
    """Escreve a lista de usuários de volta no arquivo CSV."""
    with open(ARQUIVO_USUARIOS, mode='w', newline='', encoding='utf-8') as csvfile:
        fieldnames = ['email', 'senha', 'tipo', 'pasta', 'data_nascimento', 'nome_mae']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(usuarios)

# --- ROTAS DE AUTENTICAÇÃO E SESSÃO ---

@app.route('/', methods=['GET', 'POST'])
def login():
    if 'usuario' in session:
        return redirect(url_for('calendario_visual'))

    if request.method == 'POST':
        email = request.form['usuario'].strip().lower()
        senha = request.form['senha'].strip()
        
        usuarios = ler_usuarios()
        user = next((u for u in usuarios if u['email'].strip().lower() == email and u['senha'].strip() == senha), None)
        
        if user:
            session['usuario'] = user['email']
            session['tipo'] = user['tipo']
            session['pasta'] = user['pasta']
            return redirect(url_for('calendario_visual'))
        else:
            return render_template('login.html', erro='Usuário ou senha inválidos')
            
    return render_template('login.html')

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

# --- ROTAS DE RESET DE SENHA ---

@app.route('/reset-senha', methods=['GET', 'POST'])
def reset_senha_email():
    if request.method == 'POST':
        email = request.form['email'].strip().lower()
        usuarios = ler_usuarios()
        user = next((u for u in usuarios if u['email'].strip().lower() == email), None)
        
        if user:
            session['reset_email'] = email
            return redirect(url_for('reset_senha_perguntas'))
        else:
            return render_template('reset_senha_email.html', erro='E-mail não encontrado.')
            
    return render_template('reset_senha_email.html')

@app.route('/reset-senha/perguntas', methods=['GET', 'POST'])
def reset_senha_perguntas():
    if 'reset_email' not in session:
        return redirect(url_for('login'))

    if request.method == 'POST':
        data_nascimento = request.form['data_nascimento'].strip()
        nome_mae = request.form['nome_mae'].strip().lower()
        
        email = session['reset_email']
        usuarios = ler_usuarios()
        user = next((u for u in usuarios if u['email'].strip().lower() == email), None)
        
        # Correção para lidar com campos vazios no CSV
        csv_data_nascimento = (user.get('data_nascimento') or '').strip()
        csv_nome_mae = (user.get('nome_mae') or '').strip().lower()

        if user and csv_data_nascimento == data_nascimento and csv_nome_mae == nome_mae:
            session['reset_validado'] = True
            return redirect(url_for('reset_senha_nova'))
        else:
            # --- CÓDIGO DE DEPURAÇÃO ---
            if user:
                print("\n--- DEBUG: COMPARAÇÃO DE DADOS ---")
                print(f"Formulário (Nascimento): '{data_nascimento}'")
                print(f"CSV (Nascimento):        '{csv_data_nascimento}'")
                print(f"Formulário (Mãe):        '{nome_mae}'")
                print(f"CSV (Mãe):               '{csv_nome_mae}'")
                print("--- FIM DO DEBUG ---\n")
            # --- FIM DO CÓDIGO DE DEPURAÇÃO ---
            return render_template('reset_senha_perguntas.html', erro='Dados incorretos. Tente novamente.')

    return render_template('reset_senha_perguntas.html')

@app.route('/reset-senha/nova', methods=['GET', 'POST'])
def reset_senha_nova():
    if not session.get('reset_validado'):
        return redirect(url_for('login'))

    if request.method == 'POST':
        nova_senha = request.form['nova_senha']
        email = session['reset_email']
        
        usuarios = ler_usuarios()
        for user in usuarios:
            if user['email'].strip().lower() == email:
                user['senha'] = nova_senha
                break
        
        escrever_usuarios(usuarios)
        session.clear()
        return redirect(url_for('login', mensagem='Senha alterada com sucesso!'))

    return render_template('reset_senha_nova.html')

# --- ROTAS PRINCIPAIS DA APLICAÇÃO ---

@app.route('/calendario-visual')
def calendario_visual():
    if 'usuario' not in session:
        return redirect(url_for('login'))
        
    usuario = session['usuario']
    pasta_usuario = session.get('pasta')
    nome = usuario.split('@')[0].split('.')[0].capitalize()
    
    dias_disponiveis = set()
    if os.path.exists(HISTORICO_ROTAS):
        for mes_folder in os.listdir(HISTORICO_ROTAS):
            pasta_mes = os.path.join(HISTORICO_ROTAS, mes_folder)
            if not os.path.isdir(pasta_mes): continue
            
            for dia_folder in os.listdir(pasta_mes):
                pasta_dia = os.path.join(pasta_mes, dia_folder)
                if not os.path.isdir(pasta_dia): continue
                
                if session.get('tipo') == 'admin':
                    dias_disponiveis.add(dia_folder)
                elif pasta_usuario and os.path.exists(os.path.join(pasta_dia, pasta_usuario)):
                    dias_disponiveis.add(dia_folder)

    return render_template('calendario_visual.html', nome=nome, dias_json=json.dumps(sorted(list(dias_disponiveis))))

@app.route('/cards/<data>')
def cards(data):
    if 'usuario' not in session:
        return redirect(url_for('login'))
        
    usuario = session['usuario']
    tipo = session.get('tipo')
    pasta_usuario = session.get('pasta')
    
    try:
        ano, mes, dia = data.split('-')
        pasta_dia = os.path.join(HISTORICO_ROTAS, f"{ano}-{mes}", data)
        data_br = f"{dia}/{mes}/{ano}"
    except Exception:
        pasta_dia = ''
        data_br = data

    vendedores = []
    if os.path.exists(pasta_dia):
        pastas_coord = []
        if tipo == 'admin':
            pastas_coord = [d for d in os.listdir(pasta_dia) if os.path.isdir(os.path.join(pasta_dia, d))]
        elif pasta_usuario and os.path.exists(os.path.join(pasta_dia, pasta_usuario)):
            pastas_coord = [pasta_usuario]

        for coord in pastas_coord:
            pasta_coord_path = os.path.join(pasta_dia, coord)
            for f in os.listdir(pasta_coord_path):
                if f.endswith('.xlsx') and not f.endswith('_planejado.xlsx'):
                    nome_vend = f[:-5].replace('_', ' ').title()
                    
                    # Ler informações do relatório
                    info = {}
                    relatorio_path = os.path.join(pasta_coord_path, f.replace('.xlsx', '_relatorio.txt'))
                    if os.path.exists(relatorio_path):
                        with open(relatorio_path, encoding='utf-8') as rf:
                            for linha in rf:
                                if ':' in linha:
                                    chave, valor = linha.split(':', 1)
                                    info[chave.strip()] = valor.strip()

                    vendedores.append({
                        'nome': nome_vend,
                        'planilha_path': url_for('arquivo', filepath=f"{ano}-{mes}/{data}/{coord}/{f}"),
                        'mapa_path': url_for('arquivo', filepath=f"{ano}-{mes}/{data}/{coord}/{f.replace('.xlsx', '_mapa.html')}"),
                        'img_path': url_for('arquivo', filepath=f"{ano}-{mes}/{data}/{coord}/{f.replace('.xlsx', '_mapa.png')}"),
                        'info': info
                    })

    return render_template('cards.html', data_br=data_br, vendedores=vendedores)

@app.route('/arquivo/<path:filepath>')
def arquivo(filepath):
    """Serve arquivos estáticos das pastas de rotas."""
    if 'usuario' not in session:
        return "Acesso negado", 403
    return send_from_directory(HISTORICO_ROTAS, filepath)
