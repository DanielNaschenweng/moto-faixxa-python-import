#!/usr/bin/env python3
"""
Script para importar dados da planilha de preços de peças de moto para MongoDB.
"""

from datetime import datetime
from openpyxl import load_workbook
from pymongo import MongoClient

# Configurações
MONGO_URI = "mongodb://root:password@localhost:27017"
DATABASE_NAME = "moto_faixxa"
COLLECTION_NAME = "precos"
EXCEL_PATH = "/home/daniel/projetos_sh/moto_faixxa/TABELA PREÇOS NOTE 09-2023(2).xlsx"


def conectar_mongodb():
    """Conecta ao MongoDB e retorna a coleção."""
    client = MongoClient(MONGO_URI)
    db = client[DATABASE_NAME]
    return db[COLLECTION_NAME], client


def eh_linha_separadora(row):
    """Verifica se é linha separadora (tudo vazio exceto preço=0)."""
    modelo = row[0]
    cor = row[1]
    kit = row[2]
    preco = row[4]
    loc = row[5]
    ref = row[6]

    # Tudo vazio exceto preço que pode ser 0
    sem_modelo = not modelo or not str(modelo).strip()
    sem_cor = not cor or not str(cor).strip()
    sem_kit = not kit or not str(kit).strip()
    sem_loc = not loc or not str(loc).strip()
    sem_ref = not ref or not str(ref).strip()
    preco_zero = preco == 0 or preco is None

    return sem_modelo and sem_cor and sem_kit and sem_loc and sem_ref and preco_zero


def processar_planilha(caminho_arquivo):
    """Lê a planilha e retorna lista de documentos para inserção."""
    wb = load_workbook(caminho_arquivo, data_only=True)
    documentos = []

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        marca = sheet_name.upper()

        print(f"Processando aba: {marca} ({ws.max_row} linhas)")

        # Primeiro passo: agrupar linhas por bloco (separados por linha vazia com preço=0)
        blocos = []
        bloco_atual = []

        for row in ws.iter_rows(min_row=2, values_only=True):
            if eh_linha_separadora(row):
                if bloco_atual:
                    blocos.append(bloco_atual)
                    bloco_atual = []
            else:
                bloco_atual.append(row)

        if bloco_atual:
            blocos.append(bloco_atual)

        # Segundo passo: processar cada bloco
        for bloco in blocos:
            # Encontrar primeiro modelo, primeira cor, primeira localização
            modelo = None
            cor = None
            loc = None

            for linha in bloco:
                if not modelo and linha[0] and str(linha[0]).strip():
                    modelo = str(linha[0]).strip()
                if not cor and linha[1] and str(linha[1]).strip():
                    cor = str(linha[1]).strip()
                if not loc and linha[5] and str(linha[5]).strip():
                    loc = str(linha[5]).strip()

            # Criar documentos para cada variante do bloco
            for linha in bloco:
                kit_conjunto = linha[2]
                elem = linha[3]
                preco = linha[4]
                referencia = linha[6]

                if kit_conjunto or (preco and preco != 0):
                    doc = {
                        "marca": marca,
                        "modelo": modelo,
                        "cor": cor,
                        "peca": str(kit_conjunto).strip() if kit_conjunto else None,
                        "elemento": str(elem).strip() if elem else None,
                        "preco": float(preco) if preco and isinstance(preco, (int, float)) else 0.0,
                        "localizacao": loc,
                        "referencia": str(referencia).strip() if referencia else None,
                        "data_importacao": datetime.now()
                    }
                    documentos.append(doc)

    wb.close()
    return documentos


def main():
    """Função principal."""
    print("=" * 60)
    print("Importador de Preços para MongoDB")
    print("=" * 60)

    # Conectar ao MongoDB
    print("\nConectando ao MongoDB...")
    try:
        colecao, client = conectar_mongodb()
        print(f"Conectado: {MONGO_URI}")
        print(f"Database: {DATABASE_NAME}")
        print(f"Collection: {COLLECTION_NAME}")
    except Exception as e:
        print(f"Erro ao conectar ao MongoDB: {e}")
        return

    # Processar planilha
    print(f"\nLendo planilha: {EXCEL_PATH}")
    documentos = processar_planilha(EXCEL_PATH)
    print(f"\nTotal de documentos para inserir: {len(documentos)}")

    if not documentos:
        print("Nenhum documento para inserir.")
        client.close()
        return

    # Limpar coleção existente (opcional)
    count_antes = colecao.count_documents({})
    if count_antes > 0:
        print(f"\nRemovendo {count_antes} documentos existentes...")
        colecao.delete_many({})

    # Inserir documentos em batch
    print("\nInserindo documentos...")
    try:
        resultado = colecao.insert_many(documentos)
        print(f"Documentos inseridos: {len(resultado.inserted_ids)}")
    except Exception as e:
        print(f"Erro ao inserir documentos: {e}")
        client.close()
        return

    # Verificar inserção
    count_depois = colecao.count_documents({})
    print(f"\nTotal de documentos na coleção: {count_depois}")

    # Mostrar exemplo
    print("\nExemplo de documento inserido:")
    exemplo = colecao.find_one()
    for chave, valor in exemplo.items():
        if chave != "_id":
            print(f"  {chave}: {valor}")

    # Estatísticas por marca
    print("\nDocumentos por marca:")
    pipeline = [{"$group": {"_id": "$marca", "total": {"$sum": 1}}}]
    for item in colecao.aggregate(pipeline):
        print(f"  {item['_id']}: {item['total']}")

    client.close()
    print("\n" + "=" * 60)
    print("Importação concluída com sucesso!")
    print("=" * 60)


if __name__ == "__main__":
    main()
