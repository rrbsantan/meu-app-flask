# WebApp Flora

Dashboard Flask para acompanhamento de rotas, KPIs e histórico de visitas.

## Estrutura
- `app.py`: Backend Flask
- `templates/`: Templates HTML
- `static/`: Imagens e arquivos estáticos
- `historico_rotas/`: Arquivos de histórico
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
   python app.py
   ```

## Deploy no Render
- Suba o projeto no GitHub (exceto `.venv`)
- Configure o comando de start: `gunicorn app:app`
- Render instala automaticamente as dependências do `requirements.txt`
