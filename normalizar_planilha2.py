
import csv
import pandas as pd
import sys
import os

def normalizar_planilha_csv(arquivo_origem, arquivo_saida=None):
    if arquivo_saida is None:
        base, _ = os.path.splitext(arquivo_origem)
        arquivo_saida = base + '_normalizada.csv'

    linhas_csv = []
    with open(arquivo_origem, encoding='cp1252', newline='') as f:
        leitor = csv.reader(f, delimiter=';')
        for row in leitor:
            linhas_csv.append(row)

    linhas = []
    dados_transacao = {}
    modo_ordem_servico = False
    cabecalho_itens = None
    for i, row in enumerate(linhas_csv):
        if not row or all(cell.strip() == "" for cell in row):
            continue  # ignora linhas em branco

        primeira_coluna = str(row[0]).strip() if len(row) > 0 else ""

        # --- NOVO: Detecta formato de ordem de serviço ---
        if (
            primeira_coluna.lower().startswith("número da ordem")
            or primeira_coluna.lower().startswith("nmero da ordem")
        ):
            modo_ordem_servico = True
            cabecalho_ordem = row
            continue

        if modo_ordem_servico:
            # Detecta linha principal de ordem de serviço
            if primeira_coluna != "" and not primeira_coluna.lower().startswith("número da ordem"):
                dados_transacao = {}
                for idx, col in enumerate(cabecalho_ordem):
                    if idx < len(row):
                        dados_transacao[col.strip()] = row[idx]
                continue
            # Detecta cabeçalho de itens
            if len(row) > 1 and row[1].strip().lower() in ["descrição do item", "descrio do item"]:
                cabecalho_itens = row
                continue
            # Detecta item
            if cabecalho_itens and primeira_coluna == "":
                item = {}
                for idx, col in enumerate(cabecalho_itens):
                    if idx < len(row):
                        item[col.strip()] = row[idx]
                nova_linha = dados_transacao.copy()
                nova_linha.update(item)
                linhas.append(nova_linha)
                continue
            # Detecta "Não possui itens" e ignora
            if cabecalho_itens and len(row) > 1 and "não possui itens" in row[1].lower():
                continue
            # Se linha vazia, reseta cabeçalho de itens
            if cabecalho_itens and all(cell.strip() == "" for cell in row):
                cabecalho_itens = None
                continue

        # --- Fluxo antigo (mantido para compatibilidade) ---
        # Detecta linha principal (descrição, veículo, fornecedor...)
        if primeira_coluna not in ("", "Item") and (len(row) > 8 and row[8] and row[8].strip().startswith("$")):
            while len(row) < 10:
                row.append(None)
            dados_transacao = {
                "Descrição": row[0],
                "Veículo": row[1],
                "Fornecedor": row[2],
                "Funcionário": row[3],
                "Viagem": row[4],
                "Data de cadastro": row[5],
                "Data de vencimento": row[6],
                "Data de pagamento": row[7],
                "Valor total (transação)": row[8],
                "Pago?": row[9],
            }
        # Detecta linha de item
        elif len(row) > 2 and row[2] != "Item" and dados_transacao:
            while len(row) < 6:
                row.append(None)
            nova_linha = dados_transacao.copy()
            nova_linha["ID"] = row[1]
            nova_linha["Item"] = row[2]
            nova_linha["Quantidade"] = row[3]
            nova_linha["Valor unitário"] = row[4]
            nova_linha["Valor total (item)"] = row[5]
            linhas.append(nova_linha)

    df_final = pd.DataFrame(linhas)
    df_final.to_csv(arquivo_saida, index=False, sep=';', encoding='cp1252')

    print(f"✅ Planilha normalizada criada: {arquivo_saida}")
    print(f"Total de linhas normalizadas: {len(df_final)}")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso: python normalizar_planilha2.py <arquivo_origem.csv> [arquivo_saida.csv]")
    else:
        normalizar_planilha_csv(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else None)
