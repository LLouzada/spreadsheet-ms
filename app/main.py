from datetime import datetime
from flask import Flask, request, send_file, abort
from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.cell.rich_text import TextBlock, CellRichText
from openpyxl.cell.text import InlineFont
import os

app = Flask(__name__)

current_dir = os.path.dirname(os.path.abspath(__file__))

def generate_neogen_order(data):
    template_file = os.path.join(current_dir, 'assets/templates', 'neogen_order_template_final.xlsx')
    wb = load_workbook(template_file)
    ws_faturamento = wb['Faturamento']

    # Aba Faturamento
    for item in data['data'][0]['faturamento']:
        cell = item['cell']
        value = item['value']
        ws_faturamento[cell].value = value

    # Aba Amostras
    ws_amostras = wb['Amostras']
    for item in data['data'][0]['amostra']:
        cell = item['cell']
        value = item['value']
        ws_amostras[cell].value = value
  
    # Formatação das orientações para o envio de amostras (openpyxl perde a formatação)
    normal_font = InlineFont(rFont='Arial', sz=12)
    bold_font = InlineFont(rFont='Arial', sz=12, b=True)
    red_font = InlineFont(rFont='Arial', sz=12, color='00FF0000', b=True)  
    
    first_line_text = "ORIENTAÇÕES PARA ENVIO DE AMOSTRAS\n\n"
    first_line_block = TextBlock(bold_font , first_line_text)
    
    second_line_prefix = "1º"
    second_line_pref_block = TextBlock(bold_font, second_line_prefix)
    second_line_text = " Preencher esta planilha em sua totalidade observando os campos solicitados nas abas \"Faturamento\" e \"Amostras\"\n"
    second_line_text_block = TextBlock(normal_font, second_line_text)
    
    third_line_prefix = "2º"
    third_line_pref_block = TextBlock(bold_font, third_line_prefix)
    third_line_text = " Enviar a planilha para os e-mails: "
    third_line_text_block = TextBlock(normal_font, third_line_text)
    third_line_suffix = " amostras@neogen.com e amartins@neogen.com\n"
    third_line_suffix_block = TextBlock(normal_font, third_line_suffix)
    
    fourth_line_prefix = "3º"
    fourth_line_pref_block = TextBlock(bold_font, fourth_line_prefix)
    fourth_line_text = " Postar as amostras juntamente com a planilha da aba \"Envio\" impressa para o endereço:\n\n"
    fourth_line_text_block = TextBlock(normal_font, fourth_line_text)
    
    fifth_line_text = "Av. Alexandrina das Chagas Moreira nº 964, Distrito Industrial, Pindamonhangaba -SP, CEP 12412-800, A/C Neogen Divisão Genômica"
    fifth_line_block = TextBlock(red_font, fifth_line_text)

    rich_text = CellRichText(first_line_block, second_line_pref_block, second_line_text_block, third_line_pref_block, third_line_text_block, third_line_suffix_block, fourth_line_pref_block, fourth_line_text_block, fifth_line_block)

    ws_amostras['B3'].value = rich_text
    
    # Formatação dos Textos das colunas da aba Amostras (openpyxl perde a formatação)
    normal_font = InlineFont(rFont='Arial', sz=11, color='00FFFFFF')
    bold_font = InlineFont(rFont='Arial', sz=11, b=True, color='00FFFFFF')
    
    first_col_name = "CÓDIGO DE BARRAS\n\n"
    first_col_name_block = TextBlock(bold_font, first_col_name)
    first_col_desc = "Presente no Cartão de Coleta de amostras\n(quando houver)"
    first_col_desc_block = TextBlock(normal_font, first_col_desc)
    
    rich_text = CellRichText(first_col_name_block, first_col_desc_block)
    ws_amostras['C6'].value = rich_text
    
    second_col_name = "IDENTIFICAÇÃO DO ANIMAL\n\n"
    second_col_name_block = TextBlock(bold_font, second_col_name)
    second_col_desc = "Registro da Associação / Ident. na Fazenda / Cód. do Animal"
    second_col_desc_block = TextBlock(normal_font, second_col_desc)
    
    rich_text = CellRichText(second_col_name_block, second_col_desc_block)
    ws_amostras['D6'].value = rich_text
    
    third_col_name = "ESPÉCIE e/ou RAÇA\n\n"
    third_col_name_block = TextBlock(bold_font, third_col_name)
    third_col_desc = "Preencha o campo abaixo indicando a espécie e/ou raça a que pertence a amostra"
    third_col_desc_block = TextBlock(normal_font, third_col_desc)
    
    rich_text = CellRichText(third_col_name_block, third_col_desc_block)
    ws_amostras['E6'].value = rich_text
    
    fourth_col_name = "SEXO\n\n"
    fourth_col_name_block = TextBlock(bold_font, fourth_col_name)
    fourth_col_desc = "Macho (M)\nFêmea (F)"
    fourth_col_desc_block = TextBlock(normal_font, fourth_col_desc)
    
    rich_text = CellRichText(fourth_col_name_block, fourth_col_desc_block)
    ws_amostras['F6'].value = rich_text
    
    fifth_col_name = "DATA DE NASCIMENTO\n\n"
    fifth_col_name_block = TextBlock(bold_font, fifth_col_name)
    fifth_col_desc = "Formato (DD/MM/AAAA)\nDia (DD) - 2 digitos\nMês (MM) - 2 digitos\nAno (AAAA) - 4 digitos"
    fifth_col_desc_block = TextBlock(normal_font, fifth_col_desc)
    
    rich_text = CellRichText(fifth_col_name_block, fifth_col_desc_block)
    ws_amostras['G6'].value = rich_text
    
    sixth_col_name = "TESTE\n\n"
    sixth_col_name_block = TextBlock(bold_font, sixth_col_name)
    sixth_col_desc = "Selecione o teste nos campos abaixo. Caso o teste solicitado não esteja presente na lista, preencha o campo com orientação do seu representante"
    sixth_col_desc_block = TextBlock(normal_font, sixth_col_desc)
    
    rich_text = CellRichText(sixth_col_name_block, sixth_col_desc_block)
    ws_amostras['H6'].value = rich_text
    
    seventh_col_name = "OPCIONAL\n\n"
    seventh_col_name_block = TextBlock(bold_font, seventh_col_name)
    seventh_col_desc = "Utilize este campo sob orientação do seu representante"
    seventh_col_desc_block = TextBlock(normal_font, seventh_col_desc)
    
    rich_text = CellRichText(seventh_col_name_block, seventh_col_desc_block)
    ws_amostras['N6'].value = rich_text
    ws_amostras['O6'].value = rich_text
    
    # Fixar a tab "Envio para impressão" como ativa
    ws_envio = wb['Envio para impressão']
    wb.active = ws_envio

    filename = data['data'][0]['filename']
    file_path = os.path.join(current_dir, 'assets/orders', filename)
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
        if os.path.exists(file_path):
            os.remove(file_path)
            
@app.before_request
def check_request():
    token = request.headers.get('X-API-KEY')
    if request.remote_addr != '127.0.0.1' or token != os.environ.get('MS_API_KEY'):
        return 'Acesso negado!!', 403

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
