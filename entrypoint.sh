#!/bin/sh

# Cria o diretório de dados se não existir
mkdir -p /app/data

# Se o banco de dados sqlite não existir no volume persistente, copia o padrão
if [ ! -f /app/data/petlove.db ]; then
    echo "Inicializando petlove.db no volume persistente..."
    cp /app/defaults/petlove.db /app/data/petlove.db
fi

# Sempre atualiza a Base de dados no volume persistente com os arquivos mais recentes do build
echo "Sincronizando arquivos da Base de dados..."
mkdir -p "/app/data/Base de dados"
cp -rf "/app/defaults/Base de dados/"* "/app/data/Base de dados/"

# Executa a sincronizacao segura do banco de dados SQLite
echo "Executando migracao/sincronizacao do banco de dados..."
python sync_db.py

# Executa o gunicorn
echo "Iniciando Gunicorn..."
exec gunicorn --bind 0.0.0.0:5000 --workers 4 --threads 2 app:app
