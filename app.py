
from flask import Flask, send_file, jsonify, request, redirect, url_for, flash, render_template
import pandas as pd
import os
import tempfile

from werkzeug.utils import secure_filename
from datetime import datetime


from normalizar_planilha2 import normalizar_planilha_csv

app = Flask(__name__)
app.secret_key = 'supersecretkey'  # Necessário para flash messages


# Variáveis globais para armazenar o nome do último arquivo gerado
CSV_FILE = None
XLSX_FILE = None

# Serve o HTML do dashboard
@app.route('/')
def index():
    return render_template('index.html')

# Upload e processamento automático
@app.route('/upload', methods=['POST'])
def upload():
    global CSV_FILE, XLSX_FILE
    if 'file' not in request.files:
        flash('Nenhum arquivo enviado!')
        return redirect(url_for('index'))
    file = request.files['file']
    if file.filename == '':
        flash('Nenhum arquivo selecionado!')
        return redirect(url_for('index'))
    if file:
        filename = secure_filename(file.filename)
        temp_dir = tempfile.gettempdir()
        temp_path = os.path.join(temp_dir, filename)
        file.save(temp_path)
        # Gera nomes únicos com data e hora
        now = datetime.now().strftime('%Y%m%d_%H%M%S')
        base_name = os.path.splitext(filename)[0]
        csv_name = f"{base_name}_normalizada_{now}.csv"
        xlsx_name = f"{base_name}_normalizada_{now}.xlsx"
        CSV_FILE = csv_name
        XLSX_FILE = xlsx_name
        # Processa e gera CSV final na raiz do projeto
        normalizar_planilha_csv(temp_path, CSV_FILE)
        # Valida se o CSV gerado tem dados
        try:
            df = pd.read_csv(CSV_FILE, sep=';', encoding='cp1252')
            if df.empty or len(df.columns) == 0:
                os.remove(temp_path)
                os.remove(CSV_FILE)
                flash('Não foi possível normalizar a planilha: nenhum dado reconhecido. Verifique o formato do arquivo enviado.')
                return redirect(url_for('index'))
            df.to_excel(XLSX_FILE, index=False)
        except Exception as e:
            os.remove(temp_path)
            if os.path.exists(CSV_FILE):
                os.remove(CSV_FILE)
            flash(f'Erro ao processar a planilha: {str(e)}')
            return redirect(url_for('index'))
        os.remove(temp_path)
        flash('Arquivo processado com sucesso!')
        return redirect(url_for('index'))

# Serve CSV puro
@app.route('/csv')
def get_csv():
    global CSV_FILE
    if not CSV_FILE or not os.path.exists(CSV_FILE):
        flash('Nenhum arquivo CSV normalizado disponível!')
        return redirect(url_for('index'))
    return send_file(CSV_FILE, mimetype='text/csv', as_attachment=True, download_name=CSV_FILE)

# Serve Excel normalizado
@app.route('/xlsx')
def get_xlsx():
    global XLSX_FILE
    if not XLSX_FILE or not os.path.exists(XLSX_FILE):
        flash('Nenhum arquivo Excel normalizado disponível!')
        return redirect(url_for('index'))
    return send_file(XLSX_FILE, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name=XLSX_FILE)

# Serve CSV convertido em JSON
@app.route('/json')
def get_json():
    df = pd.read_csv(CSV_FILE, sep=';', encoding='cp1252')
    return jsonify(df.to_dict(orient='records'))

if __name__ == '__main__':
    app.run(debug=True)
