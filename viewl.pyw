import os
import pandas as pd
import folium
import openrouteservice
import webbrowser
from datetime import datetime, date
import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import Calendar
import subprocess
import sys
from threading import Thread
from PIL import Image, ImageTk
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
import unicodedata
import re
from itertools import permutations

# CONFIGURAÇÕES
CAMINHO_PLANILHA = r"C:\Users\ryannraphael-fhl\Desktop\System\Lat long\Vesrsão 1.5\visitas_mes.xlsx"
PASTA_SAIDA = r"C:\Users\ryannraphael-fhl\Desktop\System\Lat long\Vesrsão 1.5\historico_rotas"
# Nova configuração para rotas planejadas
CAMINHO_ROTAS_PLANEJADAS = r"C:\Users\ryannraphael-fhl\Desktop\System\Lat long\Vesrsão 1.5\Apoio\Rotas"
# Imagens
CAMINHO_APOIO = r"C:\Users\ryannraphael-fhl\Desktop\System\Lat long\Vesrsão 1.5\Apoio"
IMG_FUNDO = os.path.join(CAMINHO_APOIO, "fundo.png")
IMG_LOGO = os.path.join(CAMINHO_APOIO, "logo.png")
# Largura máxima do logo (px)
LOGO_MAX_WIDTH = 60
# Cor do título
TITULO_COLOR = "#0a2a66"

# Funcionários (planilha e aba)
# CAMINHO_FUNCIONARIOS = os.path.join(os.path.dirname(CAMINHO_PLANILHA), "Funcionarios.xlsx")
CAMINHO_FUNCIONARIOS = os.path.join(CAMINHO_APOIO, "Funcionarios.xlsx")
FUNCIONARIOS_SHEET = "Resultado"

# Ajustes visuais dos cards/lista (pode personalizar aqui)
CARDS_CANVAS_WIDTH = 720     # largura da área rolável dos cards (reduzido um pouco)
CARDS_CANVAS_HEIGHT = 480    # altura da área rolável dos cards
CARD_WIDTH = 220             # largura de cada card
CARD_HEIGHT = 235            # altura de cada card
THUMB_W, THUMB_H = 160, 100  # tamanho da miniatura do mapa no card
MAX_COLS = 3                 # quantidade de colunas de cards
CARD_PADX = 14               # espaçamento horizontal entre cards
CARD_PADY = 14               # espaçamento vertical entre cards
# Layout dos painéis (parametrizável)
PANEL_BG = "#ffffff"
PANEL_BORDER = "#d1d5db"
LEFT_PANEL_X = 10
LEFT_PANEL_Y = 80
LEFT_PANEL_W = 500
LEFT_PANEL_H = 580
GAP_PAINEL = 10
RIGHT_PANEL_W = 740
RIGHT_PANEL_H = 580

ORS_API_KEY = 'eyJvcmciOiI1YjNjZTM1OTc4NTExMTAwMDFjZjYyNDgiLCJpZCI6IjNhNTNmMGM1MDliYTRlNjZhZTlhMzQ2Mzg3N2IyMTQ1IiwiaCI6Im11cm11cjY0In0='
PRECO_MEDIO_GASOLINA = 5.50
CONSUMO_MEDIO_KM_L = 10.0

client = openrouteservice.Client(key=ORS_API_KEY)

def _norm_nome(nome: str) -> str:
    s = str(nome or "").strip()
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    return s.casefold()
    
def _clean_nome_vendedor(nome: str) -> str:
    # Remove conteúdo entre parênteses e espaços duplicados
    s = re.sub(r"\(.*?\)", "", str(nome or ""))
    s = " ".join(s.split())
    return s.replace("_", " ") # Garante que nomes com _ sejam tratados como espaço

def _to_float(val):
    if pd.isna(val):
        raise ValueError("valor NaN")
    s = str(val).strip().replace(",", ".")
    return float(s)

# --- NOVAS FUNÇÕES PARA ROTA PLANEJADA ---

def get_semana_mes(data):
    """Retorna 1 para semana 1/3/5 e 2 para semana 2/4."""
    semana_no_mes = (data.day - 1) // 7 + 1
    return 1 if semana_no_mes in [1, 3, 5] else 2

def carregar_rota_planejada(vendedor, data):
    """Carrega e filtra a rota planejada para um vendedor em uma data específica."""
    nome_vendedor_limpo = _clean_nome_vendedor(vendedor)
    nome_vendedor_norm = _norm_nome(nome_vendedor_limpo)
    caminho_arquivo = None
    print(f"--- Procurando rota planejada para: '{vendedor}' (Normalizado: '{nome_vendedor_norm}') em {data} ---")

    # Tenta encontrar o arquivo de forma mais robusta
    if os.path.exists(CAMINHO_ROTAS_PLANEJADAS):
        for f in os.listdir(CAMINHO_ROTAS_PLANEJADAS):
            if f.endswith('.xlsx'):
                nome_arquivo_base, _ = os.path.splitext(f)
                nome_arquivo_norm = _norm_nome(_clean_nome_vendedor(nome_arquivo_base))
                if nome_arquivo_norm == nome_vendedor_norm:
                    caminho_arquivo = os.path.join(CAMINHO_ROTAS_PLANEJADAS, f)
                    print(f"SUCESSO: Arquivo encontrado: '{f}'")
                    break

    if not caminho_arquivo:
        print(f"AVISO: Arquivo de rota planejada não encontrado para '{nome_vendedor_limpo}'")
        return pd.DataFrame()

    try:
        df = pd.read_excel(caminho_arquivo)
        df.columns = [c.strip() for c in df.columns] # Limpa espaços nos nomes das colunas

        semana_filtro = get_semana_mes(data)
        dia_semana_filtro = data.weekday() + 2  # 2=Segunda, ..., 6=Sexta

        df_filtrado = df[
            (df['ID_do_território_'].astype(str).str.strip().isin([str(semana_filtro)])) &
            (df['Dia_emitido_'] == dia_semana_filtro) &
            (~df['ID_do_território_'].astype(str).str.strip().isin(['erro', 'D']))
        ].copy()

        if df_filtrado.empty:
            print(f"INFO: Nenhum cliente planejado para '{nome_vendedor_limpo}' na data {data} (Semana tipo {semana_filtro}, Dia {dia_semana_filtro})")
            return pd.DataFrame()

        df_filtrado['Latitude_'] = pd.to_numeric(df_filtrado['Latitude_'], errors='coerce')
        df_filtrado['Longitude_'] = pd.to_numeric(df_filtrado['Longitude_'], errors='coerce')
        df_filtrado.dropna(subset=['Latitude_', 'Longitude_'], inplace=True)

        return df_filtrado
    except Exception as e:
        print(f"ERRO ao carregar rota planejada para '{nome_vendedor_limpo}': {e}")
        return pd.DataFrame()

def otimizar_rota(pontos, casa_coords=None):
    """Otimiza a ordem dos pontos usando o endpoint de otimização do ORS."""
    if not pontos:
        return [], 0

    # Jobs são os clientes a serem visitados
    jobs = [{'id': i, 'location': [p['lon'], p['lat']]} for i, p in enumerate(pontos)]
    
    # O veículo começa e termina na casa do vendedor
    if not casa_coords:
        # Se não tem casa, otimiza a partir do primeiro ponto da lista
        start_location = [pontos[0]['lon'], pontos[0]['lat']]
    else:
        start_location = [casa_coords[1], casa_coords[0]] # lon, lat

    vehicles = [{
        'id': 1,
        'profile': 'driving-car',
        'start': start_location,
        'end': start_location # Sempre retorna para o início
    }]

    try:
        # Faz a chamada única para otimização com timeout maior
        client_opt = openrouteservice.Client(key=ORS_API_KEY, timeout=30)
        result = client_opt.optimization(jobs=jobs, vehicles=vehicles)
        
        # Extrai a distância total da rota otimizada
        dist_planejada_m = result['summary']['distance']
        dist_planejada_km = dist_planejada_m / 1000

        # Extrai a ordem otimizada dos pontos
        optimized_steps = result['routes'][0]['steps']
        
        # Filtra apenas os 'jobs' (visitas) e mantém a ordem
        rota_otimizada = []
        for step in optimized_steps:
            if step['type'] == 'job':
                # O 'id' do job corresponde ao índice original na lista 'pontos'
                original_index = step['id']
                rota_otimizada.append(pontos[original_index])

        return rota_otimizada, dist_planejada_km

    except Exception as e:
        print(f"ERRO na otimização da rota planejada: {e}. Calculando rota ponto a ponto como fallback.")
        # Fallback: se a otimização falhar, calcula a rota ponto a ponto (não otimizada)
        dist_fallback_km = 0
        coords_fallback = [[p['lon'], p['lat']] for p in pontos]
        if casa_coords:
            coords_fallback.insert(0, [casa_coords[1], casa_coords[0]])
        
        if len(coords_fallback) > 1:
            try:
                route = client.directions(coords_fallback, profile='driving-car')
                dist_fallback_km = route['features'][0]['properties']['summary']['distance'] / 1000
            except Exception as e_dir:
                print(f"ERRO no fallback de direções: {e_dir}")
                return pontos, 0
        
        return pontos, dist_fallback_km

# -----------------------------------------

def processar_rota(df):
    distancia_total_km = 0
    pontos_rota = []
    linhas_rota = []
    distancias = [0]
    erros = []
    if df.empty:
        return 0, 0, 0, [], [], [], [], [], []

    # Valida e coleta todas as coordenadas primeiro
    coords_lon_lat = []
    for i in range(len(df)):
        try:
            lat, lon = _to_float(df.iloc[i]['nLatitude']), _to_float(df.iloc[i]['nLongitude'])
            if not (-90 <= lat <= 90 and -180 <= lon <= 180):
                erros.append(f"Coordenada inválida no índice {i}: {lat},{lon}")
                continue
            pontos_rota.append([lat, lon])
            coords_lon_lat.append((lon, lat))
        except (ValueError, TypeError):
            erros.append(f"Coordenada inválida ou ausente no índice {i}")
            continue
    
    # Se houver mais de um ponto válido, faz uma única chamada à API
    if len(coords_lon_lat) > 1:
        try:
            route = client.directions(coords_lon_lat, profile='driving-car', format='geojson')
            
            # Extrai a geometria completa da rota
            geometry = route['features'][0]['geometry']['coordinates']
            geometry_latlon = [[lat, lon] for lon, lat in geometry]
            linhas_rota.append(geometry_latlon)

            # Extrai a distância total e as distâncias de cada trecho
            distancia_total_km = route['features'][0]['properties']['summary']['distance'] / 1000
            
            # As distâncias dos trechos estão nos segmentos
            for segment in route['features'][0]['properties']['segments']:
                distancias.append(segment['distance'] / 1000)

        except Exception as e:
            erros.append(f"Erro na API de direções para a rota real: {e}")
            # Se a API falhar, a distância será 0 e os pontos ainda serão mostrados no mapa
            distancia_total_km = 0
            distancias = [0] * len(pontos_rota)
    else:
        distancias = [0] * len(pontos_rota)


    combustivel_gasto_litros = distancia_total_km / CONSUMO_MEDIO_KM_L
    custo_total_gasolina = combustivel_gasto_litros * PRECO_MEDIO_GASOLINA

    # Horários de chegada e saída
    hora_chegada = list(df['HoraEntrada'])
    hora_saida = list(df['dSaida']) if 'dSaida' in df.columns else [''] * len(df)
    if hora_saida and isinstance(hora_saida[0], pd.Timestamp):
        hora_saida = [h.time() for h in pd.to_datetime(hora_saida, errors='coerce')]
    return distancia_total_km, combustivel_gasto_litros, custo_total_gasolina, pontos_rota, linhas_rota, hora_chegada, hora_saida, distancias, erros


def listar_pastas(base):
    if not os.path.exists(base):
        return []
    return sorted([p for p in os.listdir(base) if os.path.isdir(os.path.join(base, p))])


def listar_arquivos(base, ext):
    if not os.path.exists(base):
        return []
    return sorted([f for f in os.listdir(base) if f.endswith(ext)])


def gerar_miniatura_mapa(html_path, img_path):
    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--window-size=600,400")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-software-rasterizer")
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-infobars")
    options.add_argument("--remote-debugging-port=0")
    driver = webdriver.Chrome(options=options)
    driver.get("file://" + os.path.abspath(html_path))
    time.sleep(2)
    driver.save_screenshot(img_path)
    driver.quit()


def primeiro_ultimo_nome(nome):
    partes = nome.strip().split()
    if len(partes) == 1:
        return partes[0]
    return f"{partes[0]} {partes[-1]}"


class RotasApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Gestão de Rotas - Visual")
        self.geometry("1280x700")
        self.resizable(False, False)

        # Cache de funcionários (nome normalizado -> (lat, lon))
        self.funcionarios_home = {}
        self._carregar_funcionarios()

        # Canvas de fundo
        self.canvas_bg = tk.Canvas(self, width=1280, height=700, highlightthickness=0, bd=0)
        self.canvas_bg.pack(fill="both", expand=True)

        # Fundo (imagem)
        self._carregar_background()

        # Logo (proporcional) no topo direito
        self._carregar_logo()

        # Título central
        self._desenhar_titulo()

        # Assinatura no rodapé (canto inferior esquerdo)
        self._desenhar_assinatura()

        # =============== Layout principal (sem um frame cobrindo a janela) ===============
        # Topo: botão e status (barra de progresso removida da tela inicial)
        self.frame_top = tk.Frame(self.canvas_bg)
        self.canvas_bg.create_window(10, 55, anchor="nw", window=self.frame_top)  # abaixo do título

        self.btn_extrair = ttk.Button(self.frame_top, text="Extrair Rotas da Planilha", command=self.extrair_rotas_thread)
        self.btn_extrair.grid(row=0, column=0, padx=(0, 10))

        # Mantém a barra de progresso somente para uso interno (oculta na UI inicial)
        self.progress = ttk.Progressbar(self.frame_top, orient="horizontal", length=350, mode="determinate")
        self.progress["value"] = 0
        # ocultar da tela inicial
        self.progress.grid(row=0, column=1, padx=(0, 10))
        self.progress.grid_remove()

        self.status_label = ttk.Label(self.frame_top, text="")
        self.status_label.grid(row=0, column=1, sticky="w")

        # Painel esquerdo (calendário + coordenadores) com fundo branco
        self.left_panel = tk.Frame(self.canvas_bg, bg=PANEL_BG, highlightbackground=PANEL_BORDER, highlightthickness=1, bd=0)
        self.left_panel.pack_propagate(False)
        self.left_panel.config(width=LEFT_PANEL_W, height=LEFT_PANEL_H)
        self.canvas_bg.create_window(LEFT_PANEL_X, LEFT_PANEL_Y, anchor="nw", window=self.left_panel)
        left_inner = tk.Frame(self.left_panel, bg=PANEL_BG)
        left_inner.pack(fill="both", expand=True, padx=10, pady=10)

        # Coluna 1: Calendário e coordenadores
        self.frame_col1 = tk.Frame(left_inner, bg=PANEL_BG)
        self.frame_col1.pack(fill="both", expand=True)

        tk.Label(self.frame_col1, text="Calendário", font=("Arial", 12, "bold"), bg=PANEL_BG).pack(anchor="w")
        self.cal = Calendar(self.frame_col1, selectmode='day', date_pattern='yyyy-mm-dd', font=("Arial", 12))
        self.cal.pack(pady=5)
        self.cal.bind("<<CalendarSelected>>", self.on_dia_select)
        tk.Label(self.frame_col1, text="Coordenadores", font=("Arial", 12, "bold"), bg=PANEL_BG).pack(pady=(10, 0), anchor="w")

        # Treeview para coordenadores
        coord_frame = tk.Frame(self.frame_col1, bg=PANEL_BG)
        coord_frame.pack()
        self.tree_coord = ttk.Treeview(coord_frame, columns=("coord",), show="headings", height=8)
        self.tree_coord.heading("coord", text="Nome do Coordenador")
        self.tree_coord.column("coord", width=260)
        self.tree_coord.pack(side=tk.LEFT, fill=tk.Y)
        scrollbar_coord = ttk.Scrollbar(coord_frame, orient="vertical", command=self.tree_coord.yview)
        scrollbar_coord.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree_coord.config(yscrollcommand=scrollbar_coord.set)
        self.tree_coord.bind("<<TreeviewSelect>>", self.on_coord_select)

        # Painel direito (lista de vendedores) com fundo branco
        right_x = LEFT_PANEL_X + LEFT_PANEL_W + GAP_PAINEL
        self.right_panel = tk.Frame(self.canvas_bg, bg=PANEL_BG, highlightbackground=PANEL_BORDER, highlightthickness=1, bd=0)
        self.right_panel.pack_propagate(False)
        self.right_panel.config(width=RIGHT_PANEL_W, height=RIGHT_PANEL_H)
        self.canvas_bg.create_window(right_x, LEFT_PANEL_Y, anchor="nw", window=self.right_panel)
        right_inner = tk.Frame(self.right_panel, bg=PANEL_BG)
        right_inner.pack(fill="both", expand=True, padx=10, pady=10)

        # Coluna 2: Vendedores com rolagem
        self.frame_vend_canvas = tk.Frame(right_inner, bg=PANEL_BG)
        self.frame_vend_canvas.pack(fill="both", expand=True)

        tk.Label(self.frame_vend_canvas, text="Vendedores", font=("Arial", 12, "bold"), bg=PANEL_BG).pack(anchor="w")
        self.canvas = tk.Canvas(self.frame_vend_canvas, width=CARDS_CANVAS_WIDTH, height=CARDS_CANVAS_HEIGHT, highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(self.frame_vend_canvas, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        self.frame_vendedores = tk.Frame(self.canvas)
        self.canvas.create_window((0, 0), window=self.frame_vendedores, anchor="nw")

        def _on_mousewheel(event):
            self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        self.canvas.bind_all("<MouseWheel>", _on_mousewheel)

        def _on_frame_configure(event):
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        self.frame_vendedores.bind("<Configure>", _on_frame_configure)

        # Estado
        self.dias_disponiveis = []
        self.dia_atual = None
        self.coord_atual = None
        self.vend_atual = None
        self.mes_atual = None

        self.atualizar_dias_disponiveis()
        self.after(500, self.marcar_dias_disponiveis)
        # Estado de extração
        self.extract_start_time = None
        self.extract_total = 0
        self.log_pb = None
        self.log_eta_label = None

    # --------------------------- Funcionários (casa) ---------------------------
    def _carregar_funcionarios(self):
        self.funcionarios_home = {}
        if not os.path.exists(CAMINHO_FUNCIONARIOS):
            return
        try:
            xl = pd.ExcelFile(CAMINHO_FUNCIONARIOS)
            # Procura a aba desejada de forma case-insensitive; se não achar, usa a primeira
            sheet = None
            for s in xl.sheet_names:
                if s.strip().lower() == str(FUNCIONARIOS_SHEET).strip().lower():
                    sheet = s
                    break
            if sheet is None:
                sheet = xl.sheet_names[0]
            df = xl.parse(sheet)
            # Normaliza nomes de colunas
            cols = {c.lower(): c for c in df.columns}
            nome_col = cols.get("nome") or "Nome"
            lat_col = cols.get("latitude") or "Latitude"
            lon_col = cols.get("longitude") or "Longitude"
            for _, row in df.iterrows():
                nome = _norm_nome(_clean_nome_vendedor(row.get(nome_col)))
                try:
                    lat = _to_float(row.get(lat_col))
                    lon = _to_float(row.get(lon_col))
                except Exception:
                    continue
                if -90 <= lat <= 90 and -180 <= lon <= 180:
                    self.funcionarios_home[nome] = (lat, lon)
        except Exception:
            pass

    # --------------------------- Fundo, logo e assinatura ---------------------------

    def _carregar_background(self):
        try:
            if os.path.exists(IMG_FUNDO):
                img = Image.open(IMG_FUNDO).resize((1280, 700))
                self.bg_photo = ImageTk.PhotoImage(img)
                self.canvas_bg.create_image(0, 0, image=self.bg_photo, anchor="nw")
            else:
                # fallback
                self.canvas_bg.configure(background="#e9eaee")
        except Exception:
            self.canvas_bg.configure(background="#e9eaee")

    def _carregar_logo(self):
        try:
            if os.path.exists(IMG_LOGO):
                img_raw = Image.open(IMG_LOGO)
                max_logo_width = LOGO_MAX_WIDTH
                w_percent = (max_logo_width / float(img_raw.size[0]))
                h_size = int((float(img_raw.size[1]) * float(w_percent)))
                img_resized = img_raw.resize((max_logo_width, h_size), Image.Resampling.LANCZOS)
                self.logo_photo = ImageTk.PhotoImage(img_resized)
                self.canvas_bg.create_image(1265, 18, image=self.logo_photo, anchor="ne")
        except Exception:
            pass

    def _desenhar_titulo(self):
        # Centralizado no topo
        self.canvas_bg.create_text(
            640, 18,
            anchor="n",
            text="Gerenciamento de rota",
            font=("Arial", 18, "bold"),
            fill=TITULO_COLOR
        )

    def _desenhar_assinatura(self):
        # Canto inferior esquerdo
        self.canvas_bg.create_text(
            10, 690,
            anchor="sw",
            text="© Develop br RRBS, Thinking by LDR & FHL",
            font=("Arial", 9, "italic"),
            fill="#1f2937"
        )

    # --------------------------- Fluxo de extração e exibição ---------------------------

    def extrair_rotas_thread(self):
        self.btn_extrair.config(state=tk.DISABLED)
        self.status_label.config(text="Processando rotas, aguarde...")
        self.progress["value"] = 0
        self.progress.update()
        self.log_janela = tk.Toplevel(self)
        self.log_janela.title("Progresso da Extração")
        # Barra e ETA
        top_frame = tk.Frame(self.log_janela)
        top_frame.pack(padx=10, pady=(10, 5), fill="x")
        self.log_pb = ttk.Progressbar(top_frame, orient="horizontal", length=360, mode="indeterminate")
        self.log_pb.pack(side="left", padx=(0, 10))
        self.log_eta_label = ttk.Label(top_frame, text="Tempo restante: —")
        self.log_eta_label.pack(side="left")
        # Log
        self.log_text = tk.Text(self.log_janela, width=80, height=18, font=("Consolas", 10))
        self.log_text.pack(padx=10, pady=(5, 10))
        # Inicia animação indeterminada até descobrir o total
        try:
            self.log_pb.start(10)
        except Exception:
            pass
        t = Thread(target=self.extrair_rotas)
        t.start()

    def extrair_rotas(self):
        try:
            df = pd.read_excel(CAMINHO_PLANILHA)
            df['nLatitude'] = pd.to_numeric(df['nLatitude'], errors='coerce')
            df['nLongitude'] = pd.to_numeric(df['nLongitude'], errors='coerce')
            df = df.dropna(subset=['nLatitude', 'nLongitude']).reset_index(drop=True)
            df['DataEntrada'] = pd.to_datetime(df['dEntrada']).dt.date
            df['HoraEntrada'] = pd.to_datetime(df['dEntrada']).dt.time

            total = df.groupby(['DataEntrada', 'Supervisor', 'Vendedor']).ngroups
            atual = 0
            # Configura barra determinada e cronômetro
            self.extract_total = total
            self.extract_start_time = time.time()
            self.after(0, self._init_progressbar_determinate)

            if total == 0:
                self.after(0, lambda: self.log_text.insert(tk.END, "Nenhum grupo de rota encontrado na planilha!\n"))
                self.after(0, lambda: self.status_label.config(text="Nada a extrair!"))
                self.after(0, lambda: self.btn_extrair.config(state=tk.NORMAL))
                return

            for data, df_data in df.groupby('DataEntrada'):
                mes = str(data)[:7]
                pasta_mes = os.path.join(PASTA_SAIDA, mes)
                pasta_data = os.path.join(pasta_mes, str(data))
                os.makedirs(pasta_data, exist_ok=True)
                for supervisor, df_sup in df_data.groupby('Supervisor'):
                    if pd.isna(supervisor):
                        atual += 1
                        self.after(0, self._update_extract_progress, atual, total)
                        continue
                    pasta_sup = os.path.join(pasta_data, str(supervisor))
                    os.makedirs(pasta_sup, exist_ok=True)
                    vendedores_list = list(df_sup.groupby('Vendedor'))
                    with ThreadPoolExecutor(max_workers=4) as executor:
                        futures = []
                        for vendedor, df_vend in vendedores_list:
                            futures.append(executor.submit(self.processar_vendedor, supervisor, vendedor, df_vend, pasta_sup, data, total, atual))
                        for future in as_completed(futures):
                            atual += 1
                            self.after(0, self._update_extract_progress, atual, total)
            self.after(0, lambda: self.progress.config(value=100))
            elapsed = time.time() - (self.extract_start_time or time.time())
            self.after(0, lambda: self.status_label.config(text=f"Extração concluída! ({self._format_duration(elapsed)})"))
            self.after(0, self._finish_extract_progress, elapsed)
            self.after(0, lambda: self.btn_extrair.config(state=tk.NORMAL))
            self.after(0, lambda: self.log_text.insert(tk.END, f"\nExtração concluída!\n"))
            self.after(0, lambda: self.atualizar_dias_disponiveis())
            self.after(0, lambda: self.marcar_dias_disponiveis())
        except Exception as e:
            self.after(0, lambda: self.status_label.config(text="Erro ao extrair rotas."))
            self.after(0, lambda: self.btn_extrair.config(state=tk.NORMAL))
            self.after(0, lambda: self.log_text.insert(tk.END, f"Erro: {e}\n"))

    def processar_vendedor(self, supervisor, vendedor, df_vend, pasta_sup, data, total, atual):
        if pd.isna(vendedor):
            return
        nome_base = f"{_clean_nome_vendedor(vendedor)}".replace(" ", "_")
        arquivo_excel = os.path.join(pasta_sup, f"{nome_base}.xlsx")
        arquivo_mapa = os.path.join(pasta_sup, f"{nome_base}_mapa.html")
        arquivo_txt = os.path.join(pasta_sup, f"{nome_base}_relatorio.txt")
        arquivo_excel_planejado = os.path.join(pasta_sup, f"{nome_base}_planejado.xlsx")

        if os.path.exists(arquivo_excel) and os.path.exists(arquivo_mapa) and os.path.exists(arquivo_txt):
            self.after(0, lambda: self.log_text.insert(tk.END, f"Já existe: {supervisor} - {vendedor} - {data}\n"))
            return
        df_vend = df_vend.sort_values('HoraEntrada').reset_index(drop=True)
        if len(df_vend) < 1:
            self.after(0, lambda: self.log_text.insert(tk.END, f"Poucos pontos: {supervisor} - {vendedor} - {data}\n"))
            return
        
        home_latlon = self.funcionarios_home.get(_norm_nome(_clean_nome_vendedor(vendedor)))

        # --- ROTA REAL (EXISTENTE) ---
        distancia_total_km, combustivel_gasto_litros, custo_total_gasolina, pontos_rota, linhas_rota, hora_chegada, hora_saida, distancias, erros = processar_rota(df_vend)
        
        coords_reais_com_casa = [(p[1], p[0]) for p in pontos_rota]
        if home_latlon and pontos_rota:
            coords_reais_com_casa.insert(0, (home_latlon[1], home_latlon[0]))
        
        route_real_full = None
        if len(coords_reais_com_casa) > 1:
            try:
                route_real_full = client.directions(coords_reais_com_casa, profile='driving-car', format='geojson')
                distancia_total_km = route_real_full['features'][0]['properties']['summary']['distance'] / 1000
                combustivel_gasto_litros = distancia_total_km / CONSUMO_MEDIO_KM_L
                custo_total_gasolina = combustivel_gasto_litros * PRECO_MEDIO_GASOLINA
            except Exception as e:
                erros.append(f"Erro ao recalcular rota real com casa: {e}")

        df_vend['HoraChegada'] = hora_chegada
        df_vend['HoraSaida'] = hora_saida
        df_vend['DistanciaKm'] = distancias
        df_vend.to_excel(arquivo_excel, index=False)

        # --- ROTA PLANEJADA (NOVO E SIMPLIFICADO) ---
        df_planejado = carregar_rota_planejada(vendedor, data)
        dist_planejada_km, comb_planejado_l, custo_planejado_g = 0, 0, 0
        pontos_planejados_otimizados = []
        rota_otimizada_dict = []
        rota_planejada_geo = None

        if not df_planejado.empty:
            pontos_para_otimizar = [{'lat': row['Latitude_'], 'lon': row['Longitude_'], 'nome': row['Descrição_']} for _, row in df_planejado.iterrows()]
            
            rota_otimizada_dict, _ = otimizar_rota(pontos_para_otimizar, home_latlon)
            pontos_planejados_otimizados = [[p['lat'], p['lon']] for p in rota_otimizada_dict]

            if pontos_planejados_otimizados:
                coords_planejados_com_casa = [(p['lon'], p['lat']) for p in rota_otimizada_dict]
                if home_latlon:
                    coords_planejados_com_casa.insert(0, (home_latlon[1], home_latlon[0]))
                
                if len(coords_planejados_com_casa) > 1:
                    try:
                        rota_planejada_geo = client.directions(coords_planejados_com_casa, profile='driving-car', format='geojson')
                        dist_planejada_km = rota_planejada_geo['features'][0]['properties']['summary']['distance'] / 1000
                        comb_planejado_l = dist_planejada_km / CONSUMO_MEDIO_KM_L
                        custo_planejado_g = comb_planejado_l * PRECO_MEDIO_GASOLINA
                    except Exception as e:
                        print(f"ERRO ao calcular rota planejada final para {vendedor}: {e}")

            if dist_planejada_km > 0:
                ordem_nomes = [p['nome'] for p in rota_otimizada_dict]
                df_planejado['ordem_visita'] = df_planejado['Descrição_'].apply(lambda x: ordem_nomes.index(x) + 1 if x in ordem_nomes else 999)
                df_planejado.sort_values('ordem_visita').to_excel(arquivo_excel_planejado, index=False)

        # --- COMPARAÇÃO E MAPEAMENTO DE CORES PARA CLIENTES EM COMUM ---
        nomes_realizados_limpos = set(df_vend['Terceiro'].apply(lambda x: _norm_nome(_clean_nome_vendedor(x))))
        nomes_planejados_limpos = set(df_planejado['Descrição_'].apply(lambda x: _norm_nome(_clean_nome_vendedor(x)))) if not df_planejado.empty else set()
        
        clientes_em_comum = nomes_realizados_limpos.intersection(nomes_planejados_limpos)
        dentro_planejado = len(clientes_em_comum)
        fora_planejado = len(nomes_realizados_limpos - nomes_planejados_limpos)

        # Paleta de cores para os clientes em comum
        cores_comuns = ['#FFA500', '#8A2BE2', '#00CED1', '#32CD32', '#FF1493', '#FFD700', '#ADFF2F', '#BA55D3', '#40E0D0']
        mapa_cores_comuns = {nome: cores_comuns[i % len(cores_comuns)] for i, nome in enumerate(clientes_em_comum)}

        # --- MAPA (ATUALIZADO) ---
        mapa = folium.Map(location=pontos_rota[0] if pontos_rota else ([0,0] if not home_latlon else home_latlon), zoom_start=12, tiles="CartoDB positron", control_scale=True)

        # ROTA REAL (AZUL)
        if route_real_full:
             geometry_real = route_real_full['features'][0]['geometry']['coordinates']
             geometry_real_latlon = [[lat, lon] for lon, lat in geometry_real]
             folium.PolyLine(geometry_real_latlon, color="#0984e3", weight=5, opacity=0.8, tooltip="Rota Realizada").add_to(mapa)
        else:
            for linha in linhas_rota:
                folium.PolyLine(linha, color="#0984e3", weight=5, opacity=0.8, tooltip="Rota Realizada").add_to(mapa)

        # ROTA PLANEJADA (VERMELHO)
        if rota_planejada_geo:
            geometry = rota_planejada_geo['features'][0]['geometry']['coordinates']
            geometry_latlon = [[lat, lon] for lon, lat in geometry]
            folium.PolyLine(geometry_latlon, color="#d63031", weight=4, opacity=0.7, dash_array="5,5", tooltip="Rota Planejada").add_to(mapa)

        # CASA DO VENDEDOR
        if home_latlon:
            folium.Marker(location=home_latlon, tooltip="Casa", icon=folium.Icon(color='darkpurple', icon='home', prefix='fa')).add_to(mapa)
            if pontos_rota:
                try:
                    route_casa_real = client.directions(((home_latlon[1], home_latlon[0]), (pontos_rota[0][1], pontos_rota[0][0])), profile='driving-car')
                    geometry = [[lat, lon] for lon, lat in route_casa_real['features'][0]['geometry']['coordinates']]
                    folium.PolyLine(geometry, color="#808080", weight=4, opacity=1, dash_array="5,5", tooltip="Casa -> 1ª Visita Real").add_to(mapa)
                except Exception: pass
            if pontos_planejados_otimizados:
                try:
                    route_casa_plan = client.directions(((home_latlon[1], home_latlon[0]), (pontos_planejados_otimizados[0][1], pontos_planejados_otimizados[0][0])), profile='driving-car')
                    geometry = [[lat, lon] for lon, lat in route_casa_plan['features'][0]['geometry']['coordinates']]
                    folium.PolyLine(geometry, color="#a0a0a0", weight=3, opacity=0.9, dash_array="5,5", tooltip="Casa -> 1ª Visita Planejada").add_to(mapa)
                except Exception: pass

        # PONTOS DE VISITA (REAL)
        for idx, ponto in enumerate(pontos_rota):
            cliente_info = df_vend.iloc[idx]
            nome_cliente = str(cliente_info['Terceiro'])
            nome_cliente_limpo = _norm_nome(_clean_nome_vendedor(nome_cliente))
            
            tempo_permanencia = "N/D"
            try:
                entrada = pd.to_datetime(cliente_info['HoraEntrada'], format='%H:%M:%S').time()
                saida = pd.to_datetime(cliente_info['dSaida']).time()
                delta = datetime.combine(date.min, saida) - datetime.combine(date.min, entrada)
                tempo_permanencia = str(delta)
            except Exception: pass
            popup_html = f"<b>Visita {idx+1} - {nome_cliente}</b><br>Permanência: {tempo_permanencia}"
            
            if nome_cliente_limpo in mapa_cores_comuns:
                cor_borda = cor_preench = mapa_cores_comuns[nome_cliente_limpo]
                tooltip_texto = f"Comum: {idx+1} - {nome_cliente}"
            else:
                cor_borda, cor_preench = ("#16a34a", "#86efac") if idx == 0 else ("#dc2626", "#fca5a5") if idx == len(pontos_rota) - 1 else ("#2563eb", "#93c5fd")
                tooltip_texto = f"Real: {idx+1} - {nome_cliente}"

            folium.CircleMarker(
                location=ponto, radius=7, color=cor_borda, weight=3, fill=True, fill_color=cor_preench, fill_opacity=0.9,
                tooltip=tooltip_texto, popup=folium.Popup(popup_html, max_width=250)
            ).add_to(mapa)

        # PONTOS DE VISITA (PLANEJADO)
        for idx, ponto in enumerate(pontos_planejados_otimizados):
            nome_cliente = rota_otimizada_dict[idx]['nome']
            nome_cliente_limpo = _norm_nome(_clean_nome_vendedor(nome_cliente))

            if nome_cliente_limpo in mapa_cores_comuns:
                cor_marcador = mapa_cores_comuns[nome_cliente_limpo]
                tooltip_texto = f"Comum: {idx+1} - {nome_cliente}"
                folium.RegularPolygonMarker(
                    location=ponto,
                    tooltip=tooltip_texto,
                    number_of_sides=4, # Quadrado
                    rotation=45,       # Para alinhar como um quadrado
                    radius=8,
                    color=cor_marcador,
                    fill_color=cor_marcador,
                    fill_opacity=0.8
                ).add_to(mapa)
            else:
                tooltip_texto = f"Planejado: {idx+1} - {nome_cliente}"
                folium.Marker(
                    location=ponto,
                    tooltip=tooltip_texto,
                    icon=folium.Icon(color='red', icon='flag', prefix='fa')
                ).add_to(mapa)

        # RESUMO NO MAPA
        resumo_html = f"""
        <b>Comparativo da Rota</b><br><b>Data:</b> {data}<br><hr>
        <b>Distância:</b> {distancia_total_km:.2f} km (Real) vs {dist_planejada_km:.2f} km (Planejado)<br>
        <b>Custo:</b> R$ {custo_total_gasolina:.2f} (Real) vs R$ {custo_planejado_g:.2f} (Planejado)<br>
        <b>Visitas:</b> {len(pontos_rota)} (Real) vs {len(pontos_planejados_otimizados)} (Planejado)<br>
        <b>Aderência:</b> {dentro_planejado} cliente(s) em comum.
        """
        if pontos_rota:
            folium.Marker(location=pontos_rota[0], popup=folium.Popup(resumo_html, max_width=320), icon=folium.Icon(color='blue', icon='info-sign')).add_to(mapa)

        bounds = [list(p) for p in pontos_rota] + [list(p) for p in pontos_planejados_otimizados]
        if home_latlon: bounds.append(list(home_latlon))
        if len(bounds) >= 1:
            try: mapa.fit_bounds(bounds, padding=(20, 20))
            except Exception: pass

        mapa.save(arquivo_mapa)
        
        # RELATÓRIO TXT (ATUALIZADO)
        with open(arquivo_txt, "w", encoding="utf-8") as f:
            f.write(f"Supervisor: {supervisor}\n")
            f.write(f"Vendedor: {vendedor}\n")
            f.write(f"Data: {data}\n")
            f.write("--- ROTA REALIZADA ---\n")
            f.write(f"Distancia total percorrida: {distancia_total_km:.2f} km\n")
            f.write(f"Combustivel total gasto: {combustivel_gasto_litros:.2f} litros\n")
            f.write(f"Custo total estimado da gasolina: R$ {custo_total_gasolina:.2f}\n")
            f.write(f"Quantidade de visitas: {len(pontos_rota)}\n")
            f.write("--- ROTA PLANEJADA ---\n")
            f.write(f"Distancia planejada: {dist_planejada_km:.2f} km\n")
            f.write(f"Combustivel planejado: {comb_planejado_l:.2f} litros\n")
            f.write(f"Custo planejado: R$ {custo_planejado_g:.2f}\n")
            f.write(f"Visitas planejadas: {len(pontos_planejados_otimizados)}\n")
            f.write("--- COMPARATIVO ---\n")
            f.write(f"Clientes dentro do planejado: {dentro_planejado}\n")
            f.write(f"Clientes fora do planejado: {fora_planejado}\n")

        self.after(0, lambda: self.log_text.insert(tk.END, f"OK: {supervisor} - {vendedor} - {data}\n"))

    def atualizar_dias_disponiveis(self):
        self.dias_disponiveis = []
        for mes in listar_pastas(PASTA_SAIDA):
            for dia in listar_pastas(os.path.join(PASTA_SAIDA, mes)):
                self.dias_disponiveis.append(dia)

    def marcar_dias_disponiveis(self):
        for dia in self.dias_disponiveis:
            try:
                self.cal.calevent_create(datetime.strptime(dia, "%Y-%m-%d"), 'Rota', 'rota')
            except Exception:
                pass

    def on_dia_select(self, event):
        dia = self.cal.get_date()
        for i in self.tree_coord.get_children():
            self.tree_coord.delete(i)
        self.mes_atual = dia[:7]
        self.dia_atual = dia
        self.coord_atual = None
        pasta_dia = os.path.join(PASTA_SAIDA, self.mes_atual, dia)
        coords = listar_pastas(pasta_dia)
        if not coords:
            self.tree_coord.insert("", "end", values=("Nenhum coordenador encontrado",))
            self.mostrar_vendedores([])
            return
        for c in coords:
            self.tree_coord.insert("", "end", values=(c,))
        self.mostrar_vendedores([])

    def on_coord_select(self, event):
        sel = self.tree_coord.selection()
        if not sel or not self.dia_atual:
            return
        coord = self.tree_coord.item(sel[0])['values'][0]
        if coord == "Nenhum coordenador encontrado":
            self.mostrar_vendedores([])
            return
        self.coord_atual = coord

        # Mostra barra de progresso/carregando
        loading = tk.Toplevel(self)
        loading.title("Carregando vendedores")
        loading.geometry("350x80")
        tk.Label(loading, text="Carregando informações dos vendedores...", font=("Arial", 12)).pack(pady=10)
        progress = ttk.Progressbar(loading, orient="horizontal", length=300, mode="determinate")
        progress.pack(pady=5)
        self.update()

        pasta_vend = os.path.join(PASTA_SAIDA, self.mes_atual, self.dia_atual, coord)
        vendedores = [f[:-5] for f in listar_arquivos(pasta_vend, '.xlsx') if not f.endswith('_planejado.xlsx')]
        vendedores_info = []
        total = len(vendedores)
        for idx, v in enumerate(vendedores):
            relatorio_path = os.path.join(pasta_vend, f"{v}_relatorio.txt")
            mapa_html = os.path.join(pasta_vend, f"{v}_mapa.html")
            
            # Dicionário para guardar os dados do card
            card_data = {
                "dist_real": "N/D", "dist_plan": "N/D",
                "custo_real": "N/D", "custo_plan": "N/D",
                "visitas_real": "N/D", "visitas_plan": "N/D",
                "dentro_plan": "N/D", "fora_plan": "N/D",
                "periodo": "N/D"
            }

            # Extrai dados do relatório
            if os.path.exists(relatorio_path):
                with open(relatorio_path, encoding='utf-8') as f:
                    for linha in f:
                        if "Distancia total percorrida:" in linha: card_data["dist_real"] = linha.split(":")[1].strip()
                        if "Distancia planejada:" in linha: card_data["dist_plan"] = linha.split(":")[1].strip()
                        if "Custo total estimado da gasolina:" in linha: card_data["custo_real"] = linha.split(":")[1].strip()
                        if "Custo planejado:" in linha: card_data["custo_plan"] = linha.split(":")[1].strip()
                        if "Quantidade de visitas:" in linha: card_data["visitas_real"] = linha.split(":")[1].strip()
                        if "Visitas planejadas:" in linha: card_data["visitas_plan"] = linha.split(":")[1].strip()
                        if "Clientes dentro do planejado:" in linha: card_data["dentro_plan"] = linha.split(":")[1].strip()
                        if "Clientes fora do planejado:" in linha: card_data["fora_plan"] = linha.split(":")[1].strip()

            # Calcula período
            excel_path = os.path.join(pasta_vend, f"{v}.xlsx")
            if os.path.exists(excel_path):
                try:
                    df_vend = pd.read_excel(excel_path)
                    if not df_vend.empty:
                        inicio = pd.to_datetime(df_vend['dEntrada'].min()).time()
                        fim = pd.to_datetime(df_vend['dSaida'].max()).time()
                        card_data["periodo"] = f"{inicio.strftime('%H:%M')} - {fim.strftime('%H:%M')}"

                except Exception: pass

            vendedores_info.append({
                "nome": v.replace("_", " "),
                "data": card_data,
                "mapa_html": mapa_html
            })
            progress["value"] = int((idx+1)*100/total) if total > 0 else 0
            loading.update()
        loading.destroy()
        self.mostrar_vendedores(vendedores_info)

    def mostrar_vendedores(self, vendedores_info):
        for widget in self.frame_vendedores.winfo_children():
            widget.destroy()
        if not vendedores_info:
            tk.Label(self.frame_vendedores, text="Nenhum vendedor encontrado", font=("Arial", 12, "italic"), fg="#888").pack(pady=20)
            return

        # Mostra carregando centralizado
        loading = tk.Toplevel(self)
        loading.title("Carregando mapas")
        loading.geometry("300x120")
        loading.transient(self)
        loading.grab_set()
        loading.resizable(False, False)
        tk.Label(loading, text="Gerando miniaturas dos mapas...", font=("Arial", 12)).pack(pady=(25,10))
        pb = ttk.Progressbar(loading, orient="horizontal", length=220, mode="determinate")
        pb.pack(pady=(0,10))
        pb["maximum"] = len(vendedores_info)
        self.update()

        # Gera todas as miniaturas antes de mostrar os cards
        def gerar_todas_miniaturas():
            for idx, info in enumerate(vendedores_info):
                mapa_html = info["mapa_html"]
                img_path = mapa_html.replace(".html", ".png")
                if not os.path.exists(img_path):
                    try:
                        gerar_miniatura_mapa(mapa_html, img_path)
                    except Exception:
                        pass
                pb["value"] = idx + 1
                loading.update()
            loading.destroy()

        Thread(target=gerar_todas_miniaturas).start()
        loading.wait_window()

        # Agora mostra todos os cards
        max_cols = MAX_COLS
        card_width = CARD_WIDTH
        card_height = CARD_HEIGHT
        for idx, info in enumerate(vendedores_info):
            nome, card_data, mapa_html = info["nome"], info["data"], info["mapa_html"]
            img_path = mapa_html.replace(".html", ".png")
            card = tk.Frame(self.frame_vendedores, bd=2, relief="ridge", bg="#ffffff",
                             highlightbackground="#0984e3", highlightcolor="#0984e3",
                             highlightthickness=2, width=card_width, height=card_height)
            card.grid(row=idx//max_cols, column=idx%max_cols, padx=CARD_PADX, pady=CARD_PADY)
            card.grid_propagate(False)
            tk.Label(card, text=primeiro_ultimo_nome(nome), font=("Arial", 11, "bold"),
                     bg="#ffffff", fg="#0984e3", wraplength=card_width-20).pack(pady=(4,2))
            
            # Miniatura do mapa
            if img_path and os.path.exists(img_path):
                img = Image.open(img_path).resize((THUMB_W, THUMB_H))
                photo = ImageTk.PhotoImage(img)
                lbl_img = tk.Label(card, image=photo, bg="#ffffff")
                lbl_img.image = photo
                lbl_img.pack(pady=2)
            else:
                tk.Label(card, text="(Sem miniatura)", font=("Arial", 10, "italic"),
                         bg="#ffffff", fg="#b2bec3", height=5).pack(pady=2)

            # Frame para os dados comparativos
            data_frame = tk.Frame(card, bg="#ffffff")
            data_frame.pack(pady=(2, 4), padx=5, fill="x")
            
            # Títulos
            tk.Label(data_frame, text="Real", font=("Arial", 8, "bold"), bg="#ffffff", fg="#0984e3").grid(row=0, column=1, sticky="ew")
            tk.Label(data_frame, text="Planejado", font=("Arial", 8, "bold"), bg="#ffffff", fg="#d63031").grid(row=0, column=2, sticky="ew")
            
            # Dados
            labels = ["Distância", "Custo", "Visitas", "Dentro", "Fora", "Período"]
            keys = ["dist_real", "dist_plan", "custo_real", "custo_plan", "visitas_real", "visitas_plan", "dentro_plan", "fora_plan"]
            
            tk.Label(data_frame, text="Distância:", font=("Arial", 8), bg="#ffffff", anchor="w").grid(row=1, column=0, sticky="w")
            tk.Label(data_frame, text=card_data.get("dist_real", "0"), font=("Arial", 8), bg="#ffffff").grid(row=1, column=1)
            tk.Label(data_frame, text=card_data.get("dist_plan", "0"), font=("Arial", 8), bg="#ffffff").grid(row=1, column=2)

            tk.Label(data_frame, text="Custo:", font=("Arial", 8), bg="#ffffff", anchor="w").grid(row=2, column=0, sticky="w")
            tk.Label(data_frame, text=card_data.get("custo_real", "0"), font=("Arial", 8), bg="#ffffff").grid(row=2, column=1)
            tk.Label(data_frame, text=card_data.get("custo_plan", "0"), font=("Arial", 8), bg="#ffffff").grid(row=2, column=2)

            tk.Label(data_frame, text="Visitas:", font=("Arial", 8), bg="#ffffff", anchor="w").grid(row=3, column=0, sticky="w")
            tk.Label(data_frame, text=card_data.get("visitas_real", "0"), font=("Arial", 8), bg="#ffffff").grid(row=3, column=1)
            tk.Label(data_frame, text=card_data.get("visitas_plan", "0"), font=("Arial", 8), bg="#ffffff").grid(row=3, column=2)

            tk.Label(data_frame, text="Dentro/Fora:", font=("Arial", 8), bg="#ffffff", anchor="w").grid(row=4, column=0, sticky="w")
            tk.Label(data_frame, text=f"{card_data.get('dentro_plan', '0')} / {card_data.get('fora_plan', '0')}", font=("Arial", 8), bg="#ffffff").grid(row=4, column=1, columnspan=2)

            tk.Label(data_frame, text="Período:", font=("Arial", 8), bg="#ffffff", anchor="w").grid(row=5, column=0, sticky="w")
            tk.Label(data_frame, text=card_data.get("periodo", "-"), font=("Arial", 8), bg="#ffffff").grid(row=5, column=1, columnspan=2)

            data_frame.grid_columnconfigure(1, weight=1)
            data_frame.grid_columnconfigure(2, weight=1)

            btn_frame = tk.Frame(card, bg="#ffffff")
            btn_frame.pack(pady=(0,4), fill="x", side="bottom")
            btn_mapa = ttk.Button(btn_frame, text="Mapa", width=7, command=lambda n=nome: self.abrir_mapa_vendedor(n))
            btn_mapa.pack(side=tk.LEFT, padx=2, expand=True)
            btn_plan = ttk.Button(btn_frame, text="Plan. Real", width=9, command=lambda n=nome: self.abrir_planilha_vendedor(n, 'real'))
            btn_plan.pack(side=tk.LEFT, padx=2, expand=True)
            btn_plan_plan = ttk.Button(btn_frame, text="Plan. Prev.", width=9, command=lambda n=nome: self.abrir_planilha_vendedor(n, 'planejado'))
            btn_plan_plan.pack(side=tk.LEFT, padx=2, expand=True)


    def abrir_mapa_vendedor(self, vend):
        vend_formatado = vend.replace(" ", "_")
        pasta_vend = os.path.join(PASTA_SAIDA, self.mes_atual, self.dia_atual, self.coord_atual)
        mapa_path = os.path.join(pasta_vend, f"{vend_formatado}_mapa.html")
        if os.path.exists(mapa_path):
            webbrowser.open('file://' + os.path.abspath(mapa_path))
        else:
            messagebox.showerror("Erro", "Mapa não encontrado!")

    def abrir_planilha_vendedor(self, vend, tipo):
        vend_formatado = vend.replace(" ", "_")
        pasta_vend = os.path.join(PASTA_SAIDA, self.mes_atual, self.dia_atual, self.coord_atual)
        
        if tipo == 'real':
            excel_path = os.path.join(pasta_vend, f"{vend_formatado}.xlsx")
            msg_erro = "Planilha da rota realizada não encontrada!"
        else: # planejado
            excel_path = os.path.join(pasta_vend, f"{vend_formatado}_planejado.xlsx")
            msg_erro = "Planilha da rota planejada não encontrada!"

        if os.path.exists(excel_path):
            if sys.platform.startswith('win'):
                os.startfile(excel_path)
            elif sys.platform.startswith('darwin'):
                subprocess.call(('open', excel_path))
            else:
                subprocess.call(('xdg-open', excel_path))
        else:
            messagebox.showerror("Erro", msg_erro)

    def _init_progressbar_determinate(self):
        # Troca para modo determinado quando o total é conhecido
        try:
            if self.log_pb:
                self.log_pb.stop()
                self.log_pb.config(mode="determinate", maximum=max(1, self.extract_total), value=0)
        except Exception:
            pass
        if self.log_eta_label:
            self.log_eta_label.config(text="Tempo restante: calculando...")

    def _update_extract_progress(self, atual, total):
        # Atualiza barra oculta principal (percentual)
        try:
            pct = int(atual * 100 / total) if total else 0
            self.progress.config(value=pct)
        except Exception:
            pass
        # Atualiza barra e ETA na janela de log
        try:
            if self.log_pb:
                self.log_pb["maximum"] = max(1, total)
                self.log_pb["value"] = min(atual, total)
            if self.log_eta_label and self.extract_start_time:
                elapsed = max(0.001, time.time() - self.extract_start_time)
                speed = atual / elapsed if elapsed > 0 else 0
                remaining = (total - atual) / speed if speed > 0 else 0
                self.log_eta_label.config(
                    text=f"Progresso: {int(atual*100/total) if total else 0}% ({atual}/{total}) | Tempo restante: {self._format_duration(remaining)}"
                )
        except Exception:
            pass

    def _finish_extract_progress(self, elapsed):
        try:
            if self.log_pb:
                self.log_pb["value"] = self.log_pb["maximum"]
            if self.log_eta_label:
                self.log_eta_label.config(text=f"Concluído em {self._format_duration(elapsed)}")
        except Exception:
            pass

    def _format_duration(self, seconds):
        try:
            s = int(seconds)
            h, r = divmod(s, 3600)
            m, s = divmod(r, 60)
            if h > 0:
                return f"{h:02d}:{m:02d}:{s:02d}"
            return f"{m:02d}:{s:02d}"
        except Exception:
            return "—"


if __name__ == "__main__":
    app = RotasApp()
    app.mainloop()