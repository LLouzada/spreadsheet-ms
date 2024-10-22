import os
import sys
import openpyxl
from flask import Flask, request, jsonify

current_dir = os.path.dirname(os.path.abspath(__file__))

sys.path.insert(0, os.path.join(current_dir, 'lib', 'Flask-0.12'))
sys.path.insert(0, os.path.join(current_dir, 'lib', 'openpyxl-1.7.0'))

app = Flask(__name__)

def generate_neogen_order(data):
    # reads template file
    template_file = os.path.join(current_dir, 'assets/templates', 'neogen_order_template.xlsx')
    
    wb = openpyxl.load_workbook(template_file)
    amostrasSheet = wb['Amostras']
    faturamentoSheet = wb['Faturamento']

    for item in data['data'][0]['faturamento']:
        cell = item['cell']
        value = item['value']
        faturamentoSheet[cell] = value
        
    for item in data['data'][0]['amostra']:
        cell = item['cell']
        value = item['value']
        amostrasSheet[cell] = value
        
    #return the file
    file = os.path.join(current_dir, 'assets/files', 'neogen_order_nfa_' + data['filename'] + '_' + datetime.now().strftime("%Y%m%d%H%M%S") + '.xlsx')
    wb.save(file)
    return file

@app.route('/neogen_order', methods=['POST'])
def generate_neogen_order_endpoint():
    data = request.get_json()
    result = generate_neogen_order(data)
    file = generate_neogen_order(data)
    return jsonify(result)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)