#!/bin/bash
# Setup do ambiente para o importador de preços

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

echo "============================================"
echo "Setup do Ambiente - Moto Faixxa"
echo "============================================"

# Criar venv se não existir
if [ ! -d "venv" ]; then
    echo ""
    echo "[1/3] Criando ambiente virtual..."
    python3 -m venv venv
else
    echo ""
    echo "[1/3] Ambiente virtual já existe"
fi

# Ativar venv e instalar dependências
echo ""
echo "[2/3] Instalando dependências..."
source venv/bin/activate
pip install -q -r requirements.txt
echo "Dependências instaladas: openpyxl, pymongo"

# Subir MongoDB
echo ""
echo "[3/3] Iniciando MongoDB via Docker..."
if command -v docker &> /dev/null; then
    docker compose up -d
    echo "MongoDB rodando em localhost:27017"
else
    echo "AVISO: Docker não encontrado. Instale o Docker e execute:"
    echo "  docker compose up -d"
fi

echo ""
echo "============================================"
echo "Setup concluído!"
echo ""
echo "Para executar a importação:"
echo "  ./run.sh"
echo "============================================"
