FROM python:3.11-slim

# Evita que o Python escreva arquivos .pyc
ENV PYTHONDONTWRITEBYTECODE=1
# Evita buffering de stdout/stderr
ENV PYTHONUNBUFFERED=1

# Configurações de caminhos de dados persistentes
ENV DATABASE_PATH=/app/data/petlove.db
ENV BASE_DADOS_DIR="/app/data/Base de dados"

WORKDIR /app

# Instala dependências
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copia o código da aplicação
COPY . .

# Prepara diretório de defaults para inicialização de volumes persistentes
RUN mkdir -p /app/defaults && \
    cp petlove.db /app/defaults/petlove.db && \
    cp -r "Base de dados" "/app/defaults/Base de dados"

# Torna o script de entrypoint executável
RUN chmod +x /app/entrypoint.sh

# Expõe a porta 5000
EXPOSE 5000

# Executa o entrypoint
ENTRYPOINT ["/bin/sh", "/app/entrypoint.sh"]
