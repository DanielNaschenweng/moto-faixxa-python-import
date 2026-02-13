#!/bin/bash
# Executa o importador de preços
#
# Uso:
#   ./run.sh                    - Atualiza registros existentes (mantém IDs)
#   ./run.sh clean              - Limpa a base antes de inserir
#   ./run.sh debug=HAYABUSA     - Debug de um modelo específico (não salva no banco)
#   ./run.sh clean debug=MODELO - Combina opções

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

# Verificar se venv existe
if [ ! -d "venv" ]; then
    echo "Erro: Ambiente virtual não encontrado."
    echo "Execute primeiro: ./setup.sh"
    exit 1
fi

# Ativar venv
source venv/bin/activate

# Construir argumentos para Python
ARGS=""
for arg in "$@"; do
    case "$arg" in
        clean)
            ARGS="$ARGS --clean"
            ;;
        debug)
            ARGS="$ARGS --debug"
            ;;
        debug=*)
            ARGS="$ARGS --debug=${arg#debug=}"
            ;;
    esac
done

python import_precos.py $ARGS
