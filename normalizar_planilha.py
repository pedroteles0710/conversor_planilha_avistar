import pandas as pd

# === CONFIGURAÇÃO ===
ARQUIVO_ORIGEM = "planilha_origem.xlsx"
ARQUIVO_SAIDA = "planilha_normalizada.xlsx"

# === LEITURA COMPLETA SEM CABEÇALHO ===
df_raw = pd.read_excel(ARQUIVO_ORIGEM, header=None)

# === LISTA QUE VAI ARMAZENAR AS LINHAS NORMALIZADAS ===
linhas = []

# Variável para guardar os dados da transação principal
dados_transacao = {}

for i, row in df_raw.iterrows():
    primeira_coluna = str(row[0]).strip() if pd.notna(row[0]) else ""

    # --- Detecta se é uma linha principal (Descrição + Veículo + Fornecedor...) ---
    if primeira_coluna not in ("nan", "", "Item"):  
        # Guardar os dados principais da transação
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

    # --- Detecta linhas de itens ---
    elif pd.notna(row[2]) and row[2] != "Item":
        # Linha de item, cria nova linha "normalizada"
        nova_linha = dados_transacao.copy()
        nova_linha["ID"] = row[1]
        nova_linha["Item"] = row[2]
        nova_linha["Quantidade"] = row[3]
        nova_linha["Valor unitário"] = row[4]
        nova_linha["Valor total (item)"] = row[5]
        linhas.append(nova_linha)

# === CRIAR DATAFRAME FINAL ===
df_final = pd.DataFrame(linhas)

# === SALVAR RESULTADO ===
df_final.to_excel(ARQUIVO_SAIDA, index=False)

print(f"✅ Planilha normalizada criada: {ARQUIVO_SAIDA}")
print(f"Total de linhas normalizadas: {len(df_final)}")
