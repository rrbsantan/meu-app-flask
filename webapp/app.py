from flask import Flask, render_template, request, redirect, url_for, session
from functools import wraps
import os
import csv
import pandas as pd

app = Flask(__name__)
app.secret_key = 'sua_chave_secreta_aqui'

# Usuários de exemplo: nome, senha, tipo, pasta

def carregar_usuarios():
    usuarios = {}
    caminho_xlsx = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'logins.xlsx')
    df = pd.read_excel(caminho_xlsx)
    for _, row in df.iterrows():
        acesso = str(row['acesso'])
        if acesso == 'ADM':
            tipo = 'administrativo'
            pastas = None
        else:
            tipo = 'coordenador'
            pastas = [p.strip() for p in acesso.split(';')]
        usuarios[str(row['nome'])] = {
            'senha': str(row['senha']),
            'tipo': tipo,
            'pastas': pastas,
            'email': str(row['email']),
            'aniversario': str(row['aniversario'])
        }
    return usuarios

# Decorador para exigir login
def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'usuario' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated_function

@app.route('/', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        email = request.form.get('email')
        senha = request.form.get('senha')
        if not email or not senha:
            return render_template('login.html', erro='Preencha e-mail e senha.')
        usuarios = carregar_usuarios()
        user = None
        usuario_nome = None
        for nome, dados in usuarios.items():
            if dados['email'].lower() == email.lower() and dados['senha'] == senha:
                user = dados
                usuario_nome = nome
                break
        if user:
            session['usuario'] = usuario_nome
            session['tipo'] = user['tipo']
            session['pastas'] = user['pastas']
            session['email'] = user['email']
            session['aniversario'] = user['aniversario']
            return redirect(url_for('dashboard'))
        else:
            return render_template('login.html', erro='E-mail ou senha inválidos')
    return render_template('login.html')

@app.route('/dashboard', methods=['GET', 'POST'])
@login_required
def dashboard():
    tipo = session.get('tipo')
    pastas = session.get('pastas')
    if tipo == 'coordenador' and not pastas:
        pastas = []
    data = request.args.get('data')
    coord_selecionado = request.args.get('coordenador')
    arquivos = []
    coordenadores = []
    dias_com_arquivos = set()
    historico_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'historico_rotas')
    # Removida verificação e criação de pasta, restaurando funcionamento original
    if tipo == 'administrativo':
        meses = [m for m in os.listdir(historico_path) if os.path.isdir(os.path.join(historico_path, m))]
        for mes in meses:
            pasta_mes = os.path.join(historico_path, mes)
            for pasta_dia in os.listdir(pasta_mes):
                dia_path = os.path.join(pasta_mes, pasta_dia)
                if os.path.isdir(dia_path):
                    for coord in os.listdir(dia_path):
                        coord_path = os.path.join(dia_path, coord)
                        if os.path.isdir(coord_path):
                            if any(f.endswith('.xlsx') and not f.endswith('_planejado.xlsx') for f in os.listdir(coord_path)):
                                try:
                                    dias_com_arquivos.add(pasta_dia)
                                except Exception:
                                    pass
        # Se data selecionada, mostra coordenadores e arquivos do dia
        if data and '-' in data:
            ano, mes, dia = data.split('-')
            pasta_data = os.path.join(historico_path, f"{ano}-{mes}", f"{data}")
            if os.path.exists(pasta_data):
                coordenadores = [c for c in os.listdir(pasta_data) if os.path.isdir(os.path.join(pasta_data, c))]
                if coord_selecionado:
                    coord_path = os.path.join(pasta_data, coord_selecionado)
                    if os.path.exists(coord_path):
                        arquivos += _listar_arquivos(coord_path)
                else:
                    for coord in coordenadores:
                        coord_path = os.path.join(pasta_data, coord)
                        if os.path.exists(coord_path):
                            arquivos += _listar_arquivos(coord_path)
    elif tipo == 'coordenador' and pastas:
        meses = [m for m in os.listdir(historico_path) if os.path.isdir(os.path.join(historico_path, m))]
        for mes in meses:
            pasta_mes = os.path.join(historico_path, mes)
            for pasta_dia in os.listdir(pasta_mes):
                dia_path = os.path.join(pasta_mes, pasta_dia)
                if os.path.isdir(dia_path):
                    for pasta in pastas:
                        coord_path = os.path.join(dia_path, pasta)
                        if os.path.exists(coord_path):
                            if any(f.endswith('.xlsx') and not f.endswith('_planejado.xlsx') for f in os.listdir(coord_path)):
                                try:
                                    dias_com_arquivos.add(pasta_dia)
                                except Exception:
                                    pass
        # Se data selecionada, mostra só arquivos das suas pastas
        if data and '-' in data:
            ano, mes, dia = data.split('-')
            pasta_data = os.path.join(historico_path, f"{ano}-{mes}", f"{data}")
            if os.path.exists(pasta_data):
                for pasta in pastas:
                    coord_path = os.path.join(pasta_data, pasta)
                    if os.path.exists(coord_path):
                        arquivos += _listar_arquivos(coord_path)
    return render_template('dashboard.html', tipo=tipo, pastas=pastas, data=data if data else '', arquivos=arquivos, coordenadores=coordenadores, coord_selecionado=coord_selecionado, dias_com_arquivos=list(dias_com_arquivos))

# Função auxiliar para listar arquivos do dia
def _listar_arquivos(coord_path):
    # Monta cards completos por vendedor
    arquivos = []
    vendedores = set()
    for nome in os.listdir(coord_path):
        if nome.endswith('.xlsx') and not nome.endswith('_planejado.xlsx'):
            vendedores.add(nome[:-5])
    for v in vendedores:
        card = {}
        base_historico = os.path.relpath(coord_path, os.path.join(os.path.dirname(__file__), 'historico_rotas')).replace('\\', '/')
        card['planilha'] = f"{base_historico}/{v}.xlsx"
        card['mapa_html'] = f"{base_historico}/{v}_mapa.html"
        miniatura_path = os.path.join(coord_path, f"{v}_mapa.png")
        if os.path.exists(miniatura_path):
            card['miniatura'] = f"{base_historico}/{v}_mapa.png"
        else:
            card['miniatura'] = ""
        card['relatorio'] = os.path.join(coord_path, f"{v}_relatorio.txt")
        # Dados do relatório
        card_data = {
            "dist_real": "N/D", "dist_plan": "N/D",
            "custo_real": "N/D", "custo_plan": "N/D",
            "visitas_real": "N/D", "visitas_plan": "N/D",
            "dentro_real": "N/D", "fora_real": "N/D",
            "dentro_plan": "N/D", "fora_plan": "N/D",
            "periodo": "N/D"
        }
        if os.path.exists(card['relatorio']):
            with open(card['relatorio'], encoding='utf-8') as f:
                for linha in f:
                    if "Distancia total percorrida:" in linha and card_data["dist_real"] == "N/D":
                        card_data["dist_real"] = linha.split(":")[1].strip()
                    if "Distancia planejada:" in linha and card_data["dist_plan"] == "N/D":
                        card_data["dist_plan"] = linha.split(":")[1].strip()
                    if "Custo total estimado da gasolina:" in linha and card_data["custo_real"] == "N/D":
                        card_data["custo_real"] = linha.split(":")[1].strip()
                    if "Custo planejado:" in linha and card_data["custo_plan"] == "N/D":
                        card_data["custo_plan"] = linha.split(":")[1].strip()
                    if "Quantidade de visitas:" in linha and card_data["visitas_real"] == "N/D":
                        card_data["visitas_real"] = linha.split(":")[1].strip()
                    if "Visitas planejadas:" in linha and card_data["visitas_plan"] == "N/D":
                        card_data["visitas_plan"] = linha.split(":")[1].strip()
                    if "Clientes dentro do planejado:" in linha and card_data["dentro_real"] == "N/D":
                        card_data["dentro_real"] = linha.split(":")[1].strip()
                    if "Clientes fora do planejado:" in linha and card_data["fora_real"] == "N/D":
                        card_data["fora_real"] = linha.split(":")[1].strip()
                    if "Clientes dentro do planejado (planejado):" in linha and card_data["dentro_plan"] == "N/D":
                        card_data["dentro_plan"] = linha.split(":")[1].strip()
                    if "Clientes fora do planejado (planejado):" in linha and card_data["fora_plan"] == "N/D":
                        card_data["fora_plan"] = linha.split(":")[1].strip()
        # Período
        if os.path.exists(card['planilha']):
            try:
                import pandas as pd
                df_vend = pd.read_excel(card['planilha'])
                if not df_vend.empty and 'dEntrada' in df_vend and 'dSaida' in df_vend:
                    inicio = pd.to_datetime(df_vend['dEntrada'].min()).time()
                    fim = pd.to_datetime(df_vend['dSaida'].max()).time()
                    card_data["periodo"] = f"{inicio.strftime('%H:%M')} - {fim.strftime('%H:%M')}"
            except Exception:
                pass
        card['dados'] = card_data
        arquivos.append(card)
    return arquivos

# Rotas para baixar/ver arquivos
from flask import send_file, abort
import os

@app.route('/baixar/<path:caminho>')
@login_required
def baixar_arquivo(caminho):
    caminho_arquivo = os.path.join(os.path.dirname(__file__), 'historico_rotas', caminho)
    if not os.path.exists(caminho_arquivo):
        abort(404)
    return send_file(caminho_arquivo, as_attachment=True)

@app.route('/ver_mapa/<path:subpath>')
def ver_mapa(subpath):
    # Caminho relativo à pasta webapp
    caminho_arquivo = os.path.join(os.path.dirname(__file__), 'historico_rotas', subpath)
    if not os.path.exists(caminho_arquivo):
        abort(404)
    return send_file(caminho_arquivo)

@app.route('/ver_card/<path:caminho>')
@login_required
def ver_card(caminho):
    caminho_arquivo = os.path.join(os.path.dirname(__file__), 'historico_rotas', caminho)
    if not os.path.exists(caminho_arquivo):
        abort(404)
    return send_file(caminho_arquivo)

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

if __name__ == '__main__':
    app.run(debug=True)
