from datetime import datetime
from flask import Flask, request, send_file, abort
import openpyxl
import os

app = Flask(__name__)

current_dir = os.path.dirname(os.path.abspath(__file__))

def generate_neogen_order(data):
    # Lê o arquivo de template
    template_file = os.path.join(current_dir, 'assets/templates', 'neogen_order_template_v4.xlsx')
    wb = openpyxl.load_workbook(template_file)
    
    faturamentoSheet = wb['Faturamento']
    # Preencher as células de faturamento
    for item in data['data'][0]['faturamento']:
        cell = item['cell']
        value = item['value']
        faturamentoSheet[cell] = value
        
    amostrasSheet = wb['Amostras']
    # Preencher as células de amostras
    for item in data['data'][0]['amostra']:
        cell = item['cell']
        value = item['value']
        amostrasSheet[cell] = value
        
    # Gerar o nome do arquivo e salvar o Excel gerado
    filename = data['data'][0]['filename'] 
    file_path = os.path.join(current_dir, 'assets/orders' , filename)
    wb.save(file_path)
    
    return file_path

# Endpoint para gerar e devolver o arquivo
@app.route('/neogen_order', methods=['POST'])
def generate_neogen_order_endpoint():
    data = request.get_json()
    file_path = generate_neogen_order(data)
    
    try:
        # Enviar o arquivo como resposta
        return send_file(file_path, as_attachment=True)
    finally:
        # Deletar o arquivo após o envio
        if os.path.exists(file_path):
            os.remove(file_path)
            
@app.before_request
def check_request():
    token = request.headers.get('X-API-KEY')
    if request.remote_addr != '127.0.0.1' or token != os.environ.get('MS_API_KEY'): 
        return 'Acesso negado!!', 403

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
