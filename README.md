# Importador de Preços - Moto Faixxa

Importa dados da planilha de preços de peças de moto para MongoDB.

## Requisitos

- Python 3.12+
- MongoDB rodando em Docker (localhost:27017)

## Instalação

```bash
# Criar ambiente virtual
python3 -m venv venv

# Ativar ambiente virtual
source venv/bin/activate

# Instalar dependências
pip install -r requirements.txt
```

## Configuração

Edite as variáveis no início do arquivo `import_precos.py` se necessário:

```python
MONGO_URI = "mongodb://root:password@localhost:27017"
DATABASE_NAME = "moto_faixxa"
COLLECTION_NAME = "precos"
EXCEL_PATH = "/home/daniel/projetos_sh/moto_faixxa/TABELA PREÇOS NOTE 09-2023(2).xlsx"
```

## Execução

```bash
# Ativar ambiente virtual (se não estiver ativo)
source venv/bin/activate

# Executar importação
python import_precos.py
```

## Estrutura da Planilha

A planilha deve ter as seguintes abas: SUZUKI, YAMAHA, HONDA, KAWASAKI, OUTRAS

Cada bloco de produto é separado por uma linha vazia (preço=0). Dentro do bloco:
- **Modelo**: primeira linha com modelo preenchido
- **Cor**: primeira linha com cor preenchida
- **Localização**: primeira linha com localização preenchida

## Estrutura do Documento MongoDB

```json
{
  "marca": "SUZUKI",
  "modelo": "650F 11",
  "cor": "PRETA",
  "peca": "FRONTAL",
  "elemento": null,
  "preco": 21.47,
  "localizacao": "64",
  "referencia": "78900000 5035",
  "data_importacao": "2026-01-15T..."
}
```

## Verificar Dados

```bash
# Acessar MongoDB
docker exec -it gsuite-user-store-mongodb-1 mongosh -u root -p password

# Comandos úteis
use moto_faixxa
db.precos.countDocuments()
db.precos.findOne()
db.precos.find({referencia: "78900000 5035"})
```
