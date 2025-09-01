# WebApp Flora

Dashboard Flask para acompanhamento de rotas, KPIs e histórico de visitas.

## Estrutura
- `webapp/app.py`: Backend Flask
- `webapp/templates/`: Templates HTML
- `webapp/static/`: Imagens e arquivos estáticos
- `webapp/historico_rotas/`: Arquivos de histórico
- `logins.xlsx`: Planilha de usuários
- `requirements.txt`: Dependências do projeto

## Como rodar localmente
1. Crie o ambiente virtual:
   ```
   python -m venv .venv
   .venv\Scripts\activate
   ```
2. Instale as dependências:
   ```
   pip install -r requirements.txt
   ```
3. Execute o app:
   ```
   python webapp/app.py
   ```

## Deploy no Render
- Suba o projeto no GitHub (exceto `.venv`)
- Configure o comando de start: `gunicorn app:app` (dentro da pasta `webapp`)
- Render instala automaticamente as dependências do `requirements.txt`
