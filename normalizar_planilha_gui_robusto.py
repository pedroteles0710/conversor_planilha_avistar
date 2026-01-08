import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os, csv, unicodedata

# --- utilitários ---
def detect_delimiter(path):
    try:
        with open(path, 'r', encoding='utf-8', errors='replace') as f:
            sample = ''.join([f.readline() for _ in range(10)])
        sniffer = csv.Sniffer()
        dialect = sniffer.sniff(sample, delimiters=";,\t|,\t|,\t|,\t|,\t|,\t|,\t|,\t|,\t|,\t|")
        return dialect.delimiter
    except Exception:
        return None

def normalize_text(s):
    s = "" if s is None else str(s)
    s = s.strip().replace("\n", " ").replace("\r", " ").replace("\xa0", " ")
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return s.lower()

def find_header_row_csv(path, sep):
    # lê primeiras linhas sem header e procura linha que pareça cabeçalho
    try:
        sample = pd.read_csv(path, sep=sep, header=None, nrows=20, encoding='utf-8-sig', engine='python')
    except Exception:
        sample = pd.read_csv(path, sep=sep, header=None, nrows=20, encoding='latin-1', engine='python')
    keywords = ["descr", "veic", "veiculo", "fornec", "funcion", "viagem", "cadast", "vencim", "pagam", "valor"]
    for i in range(min(len(sample), 15)):
        row = sample.iloc[i].astype(str).map(normalize_text).tolist()
        matches = sum(any(k in cell for k in keywords) for cell in row)
        if matches >= 2:
            return i
    return 0

def find_header_row_excel(path):
    try:
        sample = pd.read_excel(path, header=None, nrows=20)
    except Exception:
        return 0
    keywords = ["descr", "veic", "veiculo", "fornec", "funcion", "viagem", "cadast", "vencim", "pagam", "valor"]
    for i in range(min(len(sample), 15)):
        row = sample.iloc[i].astype(str).map(normalize_text).tolist()
        matches = sum(any(k in cell for k in keywords) for cell in row)
        if matches >= 2:
            return i
    return 0

def unique_column_names(cols):
    seen = {}
    out = []
    for c in cols:
        base = str(c)
        if base not in seen:
            seen[base] = 0
            out.append(base)
        else:
            seen[base] += 1
            out.append(f"{base}_{seen[base]}")
    return out

# --- função principal com interface ---
def normalizar_planilha():
    root = tk.Tk()
    root.withdraw()

    caminho = filedialog.askopenfilename(title="Selecione a planilha", filetypes=[("Planilhas", "*.xlsx *.xls *.csv")])
    if not caminho:
        messagebox.showinfo("Cancelado", "Nenhum arquivo selecionado.")
        return

    try:
        # lê arquivo tentando detectar cabeçalho corretamente
        if caminho.lower().endswith(".csv"):
            sep = detect_delimiter(caminho) or ";"
            header_row = find_header_row_csv(caminho, sep)
            try:
                df = pd.read_csv(caminho, sep=sep, header=header_row, encoding='utf-8-sig', engine='python')
            except Exception:
                df = pd.read_csv(caminho, sep=sep, header=header_row, encoding='latin-1', engine='python')
        else:
            header_row = find_header_row_excel(caminho)
            df = pd.read_excel(caminho, header=header_row)

        # limpa nomes de coluna
        original_cols = list(df.columns)
        cleaned_cols = [normalize_text(c).strip() for c in original_cols]
        # renomeia no dataframe para facilitar matching, mas guarda mapeamento para retornar nomes originais quando salvar
        mapping = {original_cols[i]: cleaned_cols[i] for i in range(len(original_cols))}
        df.columns = cleaned_cols

        # evita colunas duplicadas idênticas
        df.columns = unique_column_names(df.columns)

        # colunas alvo (nomes “esperados”)
        COLUNAS_PRINCIPAIS = [
            "Descrição", "Veículo", "Fornecedor", "Funcionário", "Viagem",
            "Data de cadastro", "Data de vencimento", "Data de pagamento",
            "Valor total", "Pago"
        ]
        COLUNAS_ITENS = ["ID", "Item", "Quantidade", "Valor unitário", "Valor total"]

        # função de match flexível (normalize as strings)
        def match_any(target, existing_cols):
            tn = normalize_text(target)
            for ec in existing_cols:
                if tn == ec or tn in ec or ec in tn:
                    return ec
            return None

        existing = list(df.columns)
        matched_principais = []
        for t in COLUNAS_PRINCIPAIS:
            m = match_any(t, existing)
            if m:
                matched_principais.append(m)
        matched_itens = []
        for t in COLUNAS_ITENS:
            m = match_any(t, existing)
            if m:
                matched_itens.append(m)

        # monta linhas finais
        linhas_final = []
        if matched_principais:
            grupos = df.groupby(matched_principais, dropna=False)
        else:
            # se não achou colunas principais, trata cada linha separadamente
            grupos = ((None, df.iloc[[i]]) for i in range(len(df)))

        for _, grupo in grupos:
            for _, linha_item in grupo.iterrows():
                nova = {}
                # adiciona colunas principais (se existirem)
                for col in matched_principais:
                    nova[col] = linha_item.get(col, None)
                # adiciona colunas de item (se existirem)
                for col in matched_itens:
                    nova[col] = linha_item.get(col, None)
                # adiciona também quaisquer outras colunas que existam e que possam ser úteis
                linhas_final.append(nova)

        if not linhas_final:
            messagebox.showwarning("Resultado vazio", "Nenhuma linha foi produzida — verifique o cabeçalho do arquivo.")
            return

        df_final = pd.DataFrame(linhas_final)

        # salva na mesma pasta com sufixo
        pasta = os.path.dirname(caminho)
        nome = os.path.splitext(os.path.basename(caminho))[0]
        saida = os.path.join(pasta, f"{nome}_normalizada.xlsx")
        df_final.to_excel(saida, index=False)
        messagebox.showinfo("Sucesso", f"Planilha normalizada salva em:\n{saida}\n\nColunas encontradas:\n{', '.join(original_cols)}")
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro:\n{e}")

if __name__ == "__main__":
    normalizar_planilha()
