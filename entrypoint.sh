#!/bin/sh

# Cria o diretório de dados se não existir
mkdir -p /app/data

# Se o banco de dados sqlite não existir no volume persistente, copia o padrão
if [ ! -f /app/data/petlove.db ]; then
    echo "Inicializando petlove.db no volume persistente..."
    cp /app/defaults/petlove.db /app/data/petlove.db
fi

# Se a pasta de Base de dados não existir no volume persistente, copia a padrão
if [ ! -d "/app/data/Base de dados" ]; then
    echo "Inicializando Base de dados no volume persistente..."
    cp -r "/app/defaults/Base de dados" "/app/data/"
fi

# Executa o gunicorn
echo "Iniciando Gunicorn..."
exec gunicorn --bind 0.0.0.0:5000 app:app
