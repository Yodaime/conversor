import datetime
from flask import Flask, request, send_file
import pandas as pd
import os
import re
import webbrowser
from threading import Timer

app = Flask(__name__)

# Função para abrir o navegador após o servidor começar
def open_browser():
    webbrowser.open_new("http://127.0.0.1:5000/")

def is_numeric(values):
    try:
        float(values.replace(',', '').replace('.', '').replace('-',''))  # Remove vírgulas/pontos para tentar converter
        return True
    except ValueError:
        return False

# Função para converter Excel para .txt
def excel_to_separated_excel(input_excel_file, output_excel_file, name, ender):
    # Ler o arquivo Excel
    df = pd.read_excel(input_excel_file, skiprows=1)
    colunas_excel = df.columns.tolist()



    # Cabeçalhos que serão inseridos nas primeiras duas linhas
    headers = [
       ["Nome", "Endereco", "Medidor", "Localizacao", "Tipo de Fluido", "Data", "Horario", 
               "Indice", "Indice antigo", "Consumo", "Unid._levant.", "Bateria", "Ha desmontagem/corte", 
               "Houve desmontagem/corte", "Ha vazamento", "Ha fraude magnetica", "Houve fraude magnetica", 
               "Ha medidor bloqueado", "Retorno de agua excedido"]
    ]

    # Lista para armazenar as linhas processadas
    processed_data = []



    # Adicionar as linhas de cabeçalho no início
    processed_data.extend(headers)

    # Função para mapear as palavras para as colunas adequadas
    def process_row(row_values):
        # Inicializar as colunas com valores vazios
        data_inicial = datetime.date.today().strftime("%d/%m/%Y")
        hora = datetime.datetime.now().strftime("%H:%M:%S")
        
        nome = endereco = medidor = localizacão = tipo_de_fluido = data = horario = indice = indice_antigo = consumo = unid_levant = bateria = ha_desmontagem = houve_desmontagem = ha_vazamento = ha_fraude_magnetica = houve_fraude_magnetica = ha_medidor_bloqueado = retorno_agua_excedido = ""
        
        col1 = str(row_values[0])  # Primeira coluna
        col2 = str(row_values[1])  # Segunda coluna
        col3 = str(row_values[2])
        col4 = str(row_values[3])
        col5 = str(row_values[4])   
        col6 = str(row_values[5])  
        col7 = str(row_values[6])
        col8 = str(row_values[7])
        col9 = str(row_values[8])
        col10 = str(row_values[9])
        col11 = str(row_values[10])

  
        # Iterar sobre os valores na linha e mapear para as colunas corretas
        for word in row_values:

          if colunas_excel[2] == "Nome.1" and colunas_excel[1] == "Tipo" :
              medidor = col8
              localizacão = col3
              tipo_de_fluido = col10 + " " + col11
            
            # elif word.isalpha():
            #     junto = word + row_values[1]
            #     match = re.match(r'([A-Za-z]+)(\d+)', junto)
            #     if match:
            #         unidade = match.group(1)  # Parte da palavra (ex: "SALA")
            #         ambiente = match.group(2)  # Parte do número (ex: "31")
            # elif is_numeric(word[-1]):
            #     matricula = word
            
            # if not is_numeric(col4) or col4.startswith("MED") or col4.startswith("BL") or col4.startswith("LOT") or col4.startswith("TOR") or col4.startswith("Bloco"):
            #     bloco = col4

            # else:
            #     medidor = word
            #     fracao_ideal = word

            # if is_numeric(col1):
            #     val = col1
            #     print(val)
            # if col1.isalpha():
            #     val2 = col1
            #     print(val2)

            # if not is_numeric(col1):
            #     val3 = col1
            #     print(val3)

          data = data_inicial
          horario = hora 
          nome = name
          endereco = ender
          indice = "0"
          indice_antigo = "0"
          consumo = "0"
          unid_levant = "l"
          bateria = "-"
          ha_desmontagem = "0"
          houve_desmontagem = "0"
          ha_vazamento = "0"
          ha_fraude_magnetica = "0"
          houve_fraude_magnetica = "0"
          ha_medidor_bloqueado = "0"
          retorno_agua_excedido = "0"

        return [ nome, endereco, medidor, localizacão, tipo_de_fluido, data, horario, indice, indice_antigo, consumo, unid_levant, 
                bateria, ha_desmontagem, houve_desmontagem, ha_vazamento, ha_fraude_magnetica, houve_fraude_magnetica, 
                ha_medidor_bloqueado, retorno_agua_excedido ]

    

    # Iterar sobre as linhas do DataFrame e processar cada linha
    for index, row in df.iterrows():
        line = ' '.join(map(str, row.values))
        words = line.split()
        processed_row = process_row(words)
        processed_data.append(processed_row)


    # Criar um novo DataFrame com as colunas reorganizadas
    processed_df = pd.DataFrame(processed_data)

    # Salvar o DataFrame em um novo arquivo Excel
    processed_df.to_excel(output_excel_file, index=False, header=False)
    print(f"Arquivo Excel {output_excel_file} gerado com sucesso.")
    
    return output_excel_file

def convert_excel_to_txt(output_excel_file, output_txt_file):
    # Ler o arquivo Excel
    df = pd.read_excel(output_excel_file)

    # Salvar o conteúdo do DataFrame em um arquivo .txt com tabulação como delimitador
    df.to_csv(output_txt_file, sep='\t', index=False)

    print(f"Arquivo {output_txt_file} gerado com sucesso.")

# Caminho do arquivo de entrada e de saída


# Executar a conversão
# Página inicial para upload do arquivo
@app.route('/')
def index():
    return '''
    <!doctype html>
<!doctype html>
<html lang="pt-br">
  <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">
    <title>Conversão de Arquivo Excel</title>
    <style>
      /* Estilo básico para centralizar o conteúdo */
      body {
        font-family: Arial, sans-serif;
        background-color: #f4f4f4;
        display: flex;
        justify-content: center;
        align-items: center;
        height: 100vh;
        margin: 0;
      }
      .container {
        background-color: #E0FFFF;
        border-radius: 10px;
        box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
        padding: 30px;
        max-width: 500px;
        width: 100%;
        text-align: center;
      }
      h1 {
        font-size: 24px;
        color: #333;
      }
      input[type="file"], input[type="text"] {
        padding: 10px;
        margin-bottom: 20px;
        border: 2px dashed #ccc;
        border-radius: 10px;
        cursor: pointer;
        width: 100%;
        color: #666;
      }
      input[type="submit"] {
        background-color: #007bff;
        color: white;
        border: none;
        padding: 15px 20px;
        border-radius: 5px;
        font-size: 16px;
        cursor: pointer;
        width: 100%;
      }
      input[type="submit"]:hover {
        background-color: #0056b3;
      }
      .note {
        font-size: 14px;
        color: #999;
      }
    </style>
  </head>
  <body>
    <div class="container">
      <h1>Converter Arquivo Excel</h1>
      <p>Envie seu arquivo Excel para reorganização e conversão no formato desejado.</p>
      <form action="/upload" method="post" enctype="multipart/form-data" id="uploadForm">
        <input type="file" name="file" id="fileInput" required><br>
        <label for="nome">Nome do Cliente:</label>
        <input type="text" name="nome" id="name" required><br>
        <label for="endereco">Endereço:</label>
    
        <input type="text" name="endereco" id="ender" required><br><br>
        <input type="submit" value="Converter e Baixar">
        <div class="note">Formato que será baixado: .txt</div>
      </form>
    </div>
  </body>
</html>


    '''

# Rota para upload e processamento do arquivo Excel
@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return "Nenhum arquivo foi enviado.", 400

    file = request.files['file']

    name = request.form.get('nome')
    ender = request.form.get('endereco')

    if file.filename == '':
        return "Nenhum arquivo foi selecionado.", 400

    # Salvar o arquivo Excel temporariamente
    input_excel_file = "temp_excel.xlsx"
    file.save(input_excel_file)

    # Nome dos arquivos de saída
    output_excel_file = "dados_convertidos.xlsx"
    
    output_txt_file = 'dados_convertidos.txt'  # Nome do arquivo de saída em .txt

    # Executar a conversão
    # Converter para Excel com colunas reorganizadas
    excel_to_separated_excel(input_excel_file, output_excel_file, name, ender)
    convert_excel_to_txt(output_excel_file, output_txt_file)


    # Fazer o download do arquivo Excel gerado
    return send_file(output_txt_file, as_attachment=True)

if __name__ == '__main__':
    # Abrir o navegador somente na primeira execução do servidor (não na reinicialização)
    if os.environ.get('WERKZEUG_RUN_MAIN') is None:
        Timer(1, open_browser).start()  # Abre o navegador após 1 segundo
    app.run(debug=True)