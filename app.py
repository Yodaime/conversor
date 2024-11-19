import datetime
from flask import Flask, redirect, request, send_file, url_for
import pandas as pd
import os
import re
import webbrowser
import math

from threading import Timer

app = Flask(__name__)

# Função para abrir o navegador após o servidor começar
def open_browser():
    webbrowser.open_new("http://127.0.0.1:5000/")

# Função para verificar se o valor é numérico
def is_numeric(values):
    try:
        float(values.replace(',', '').replace('.', '').replace('-',''))  # Remove vírgulas/pontos para tentar converter
        return True
    except ValueError:
        return False

# Função para processar planilha padrão (se a organização NÃO for "Prosper")
def excel_to_separated_excel(input_excel_file, output_excel_file):
    # Ler o arquivo Excel
    df = pd.read_excel(input_excel_file)


    # Cabeçalhos que serão inseridos nas primeiras duas linhas
    headers = [
        ["BLOCO", "", "UNIDADE", "", "", "AMBIENTE", "", "MEDIDOR"],
        ["Nome", "Tipo", "Nome", "Matrícula", "Fração Ideal", "Sigla", "Nome Completo", "Num Rádio", "Num. INMETRO", "Fluido", "Modelo"]
    ]

    processed_data = []

    processed_data.extend(headers)

    def process_row(row_values):
        # Inicializar as colunas com valores vazios
        bloco = unidade = ambiente = medidor = ""
        nome = tipo = Nome = matricula = fracao_ideal = sigla = nome_completo = num_radio = num_inmetro = fluido = modelo = ""
        
        col1 = str(row_values[0])  # Primeira coluna
        col2 = str(row_values[1])  # Segunda coluna
        col3 = str(row_values[2])
        col4 = str(row_values[3])
      
        # Iterar sobre os valores na linha e mapear para as colunas corretas
        for word in row_values:
            if word.startswith("MED"):
                bloco = word
            elif word.isalpha():
                junto = word + row_values[1]
                match = re.match(r'([A-Za-z]+)(\d+)', junto)
                if match:
                    unidade = match.group(1)  # Parte da palavra (ex: "SALA")
                    ambiente = match.group(2)  # Parte do número (ex: "31")
            elif is_numeric(word[-1]):
                matricula = word
            
            if not is_numeric(col4) or col4.startswith("MED") or col4.startswith("BL") or col4.startswith("LOT") or col4.startswith("TOR") or col4.startswith("Bloco"):
                bloco = col4

            else:
                medidor = word
                fracao_ideal = word

        Nome = "Único"
        sigla = "Água Fria"
        nome = "0,000"
        nome_completo = "1"
        
        return [bloco, unidade, ambiente, medidor, nome, tipo, Nome, matricula, 
                fracao_ideal, sigla, nome_completo, num_radio, num_inmetro, 
                fluido, modelo]

    for index, row in df.iterrows():
      line = ' '.join(map(str, row.values))
      words = line.split()
      processed_row = process_row(words)
      processed_data.append(processed_row)

    processed_df = pd.DataFrame(processed_data)
    processed_df.to_excel(output_excel_file, index=False, header=False)
    print(f"Arquivo Excel {output_excel_file} gerado com sucesso.")
    
    return output_excel_file

# Função para processar planilha da organização "Prosper"
def process_prosper_excel(input_excel_file, output_excel_file):
    # Ler o arquivo Excel
    # Cabeçalhos para o novo arquivo Excel


    with pd.ExcelFile(input_excel_file) as xls:
        # Encontrar a linha com o cabeçalho específico
        for sheet in xls.sheet_names:
            df_temp = pd.read_excel(xls, sheet)
            header_row = df_temp.apply(lambda row: row.astype(str).str.contains("Bloco").any(), axis=1).idxmax() or df_temp.apply(lambda row: row.astype(str).str.contains("Sala").any(), axis=1).idxmax()
            break  # Considerar apenas a primeira ocorrência do cabeçalho

    # Recarregar o DataFrame a partir da linha do cabeçalho encontrado
    df = pd.read_excel(input_excel_file, skiprows=header_row + 1)

    colunas_excel = df.columns.tolist()
    bl =""

    colum_values = pd.read_excel(input_excel_file)
    colum_values = colum_values.iterrows()
    # print(colunas_excel[3])

    start_row = None
    for i, row in df.iterrows():
        if "Bloco" in row.values or "Unidade" in row.values or "Leitura Ant." in row.values:
            start_row = i + 1  # Começa na próxima linha após os cabeçalhos
            break
        

    # Filtrar o DataFrame a partir da linha identificada
    df_filtered = df.iloc[start_row:].copy() if start_row is not None else df
    df_filtered = df_filtered[~df_filtered.apply(lambda row: row.astype(str).str.contains("TOTAL").any(), axis=1)]
  

    headers = [
        ["BLOCO", "", "UNIDADE", "", "", "AMBIENTE", "", "MEDIDOR"],
        ["Nome", "Tipo", "Nome", "Matrícula", "Fração Ideal", "Sigla", "Nome Completo", "Num Rádio", "Num. INMETRO", "Fluido", "Modelo"]
    ]

    processed_data = []

    processed_data.extend(headers)

    def process_row(row_values):
      bloco = unidade = ambiente = medidor = ""
      nome = tipo = Nome = matricula = fracao_ideal = sigla = nome_completo = num_radio = num_inmetro = fluido = modelo = ""
      col1 = str(row_values[0])
      col2 = str(row_values[1])
      col3 = str(row_values[2])
      col4 = str(row_values[3])
      col5 = str(row_values[4])
        
      for word in row_values:
        # print(word)
        if colunas_excel[0].startswith('Sala') and colunas_excel[1] == "Leitura Anterior" or colunas_excel[1] == "Leitura Ant." or colunas_excel[1] == "Leitura - Ant.":
          if word != '0' and word != '.':
            ambiente = col1.replace('.', '').replace('0', '')
            bloco = "1"
            unidade = "Unidade"
        if colunas_excel[0].startswith("Bloco") or colunas_excel[0].startswith('BL') or colunas_excel[0] == "BLOCO" or colunas_excel[0] == "BL":
          ambiente = col2.replace('.', '').replace('0', '')
          unidade = "Unidade"
          bloco = col1
        if colunas_excel[0].startswith('Unidade') or colunas_excel[0] == "UNIDADE" :
          if word != '0' and word != '.':
            ambiente = col1.replace('.', '').replace('0', '')
            unidade = "Unidade"
            bloco = "1"
            if any(word.startswith("TOTAL") for word in words):
              break
            
        if sigla == '':
          sigla = "Água Fria"



      tipo = "Único"
      Nome = "Único"
      nome = "0,000"
      nome_completo = "1"


      return [bloco, unidade, ambiente, medidor, nome, tipo, Nome, matricula, 
              fracao_ideal, sigla, nome_completo, num_radio, num_inmetro, 
              fluido, modelo]

    for index, row in df.iterrows():
        line = ' '.join(map(str, row.values))
        words = line.split()
      
        if any(word.startswith("TOTAL") for word in words):
          if processed_data:
            processed_data.pop()
            break

        processed_row = process_row(words)
        processed_data.append(processed_row)

    processed_df = pd.DataFrame(processed_data)
    processed_df.to_excel(output_excel_file, index=False, header=False)
    print(f"Arquivo Excel {output_excel_file} gerado com sucesso.")
    
    return output_excel_file

def process_olicon_excel(input_excel_file, output_excel_file):
    # Ler o arquivo Excel
    df = pd.read_excel(input_excel_file)

    # Cabeçalhos para o novo arquivo Excel
    headers = [
        ["BLOCO", "", "UNIDADE", "", "", "AMBIENTE", "", "MEDIDOR"],
        ["Nome", "Tipo", "Nome", "Matrícula", "Fração Ideal", "Sigla", "Nome Completo", "Num Rádio", "Num. INMETRO", "Fluido", "Modelo"]
    ]

    processed_data = []

    processed_data.extend(headers)

    def process_row(row_values):
        bloco = unidade = ambiente = medidor = ""
        nome = tipo = Nome = matricula = fracao_ideal = sigla = nome_completo = num_radio = num_inmetro = fluido = modelo = ""
        col1 = str(row_values[0])
        col2 = str(row_values[1])
        col3 = str(row_values[2])
        col4 = str(row_values[3])
        bloco = ""
        unidade = ""
        ambiente = ""
        medidor = ""
        Nome = "Único"
        sigla = "Água Fria"
        nome = "0,000"
        nome_completo = "1"
        tipo = ""
        fracao_ideal=""
        matricula=""
        num_radio=""
        num_inmetro=""
        fluido=""
        modelo=""

        # for word in col2:
        #     if word != '0' and word != '.':
        #       unidade = col2.replace('.', '').replace('0', '')

        
        
        return [bloco, unidade, ambiente, medidor, nome, tipo, Nome, matricula, 
                fracao_ideal, sigla, nome_completo, num_radio, num_inmetro, 
                fluido, modelo]

    for index, row in df.iterrows():
        line = ' '.join(map(str, row.values))
        words = line.split()
        processed_row = process_row(words)
        processed_data.append(processed_row)

    processed_df = pd.DataFrame(processed_data)
    processed_df.to_excel(output_excel_file, index=False, header=False)
    print(f"Arquivo Excel {output_excel_file} gerado com sucesso.")
    
    return output_excel_file

def excel_to_separated_excel1(input_excel_file, output_excel_file, name, ender):
    # Ler o arquivo Excel
    df = pd.read_excel(input_excel_file, skiprows=1)
    colunas_excel = df.columns.tolist()

    # Cabeçalhos que serão inseridos nas primeiras duas linhas
    headers = [
       ["Medidor", "Localizacao", "Tipo de Fluido", "Data", "Horario", 
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
              localizacão = col3.replace('.0', '')
              tipo_de_fluido = col10 + " " + col11
          
                # localizacão = localizacão.strip()
                # localizacão.replace('.', '')
                # print(localizacão)

        data = ender
        horario = hora 
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

        return [ medidor, localizacão, tipo_de_fluido, data, horario, indice, indice_antigo, consumo, unid_levant, 
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


@app.route('/')
def index():
    return '''
    <!doctype html>
    <html lang="pt-br">
      <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">
        <title>Escolha o Formato</title>
        <style>
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
          p {
            color: #666;
            margin-bottom: 20px;
          }
          input[type="file"] {
            padding: 10px;
            margin-bottom: 20px;
            border: 2px dashed #ccc;
            border-radius: 10px;
            cursor: pointer;
            width: 100%;
            color: #666;
          }
          input[type="file"]:hover {
            border-color: #007bff;
            background-color: #f8f8f8;
          }
          input[type="text"] {
            padding: 10px;
            margin-bottom: 20px;
            border: 1px solid #ccc;
            border-radius: 5px;
            width: 100%;
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
          <h1>Escolha o formato de conversão</h1>
          <form action="/select_format" method="post">
              <button name="format" value="excel" type="submit">Gerar arquivo Excel para medidores</button>
              <button name="format" value="txt" type="submit">Gerar arquivo TXT para subir leituras</button>
          </form>
        </div>
      </body>
    </html>
    '''

@app.route('/select_format', methods=['POST'])
def select_format():
    chosen_format = request.form['format']
    if chosen_format == "excel":
        return redirect(url_for('upload_excel'))
    elif chosen_format == "txt":
        return redirect(url_for('upload_txt'))
    

@app.route('/upload_excel')
def upload_excel():
    return '''
    <!doctype html>
    <html lang="pt-br">
      <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">
        <title>Conversão de Arquivo Excel</title>
        <style>
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
          p {
            color: #666;
            margin-bottom: 20px;
          }
          input[type="file"] {
            padding: 10px;
            margin-bottom: 20px;
            border: 2px dashed #ccc;
            border-radius: 10px;
            cursor: pointer;
            width: 100%;
            color: #666;
          }
          input[type="file"]:hover {
            border-color: #007bff;
            background-color: #f8f8f8;
          }
          input[type="text"] {
            padding: 10px;
            margin-bottom: 20px;
            border: 1px solid #ccc;
            border-radius: 5px;
            width: 100%;
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
          <p>Envie seu arquivo Excel para reorganização e conversão.</p>
          <form action="/process_excel" method="post" enctype="multipart/form-data" id="uploadForm">
            <input type="file" name="file" id="fileInput">
            <input type="text" name="organization" placeholder="Qual a organização?">
            <br/>
            <input type="submit" value="Converter e Baixar">
            <div class="note">Formato suportado: .xlsx</div>
            <div class="note">Observação!</div>
            <div class="note">Para a criação de novos medidores com este arquivo, abra-o e salve-o no formato ".xls"</div>
          </form>
        </div>
      </body>
    </html>
    '''

@app.route('/upload_txt')
def upload_txt():
    return '''
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
            <p>Envie seu arquivo Excel para reorganização e conversão.</p>
            <form action="/process_txt" method="post" enctype="multipart/form-data" id="uploadForm">
              <input type="file" name="file" id="fileInput" required><br>
              <label for="nome">Nome a ser salvo:</label>
              <input type="text" name="nome" id="name" required><br>
              <label for="endereco">Data da leitura:</label>
          
              <input type="text" name="endereco" id="ender" required><br><br>
              <input type="submit" value="Converter e Baixar">
              <div class="note">Formato que será baixado: .txt</div>
            </form>
          </div>
        </body>
      </html>
    '''

@app.route('/process_txt', methods=['POST'])
def process_txt():
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
    
    output_txt_file = f"{name}.txt"

    # Converter para Excel com colunas reorganizadas
    excel_to_separated_excel1(input_excel_file, output_excel_file, name, ender)
    convert_excel_to_txt(output_excel_file, output_txt_file)

    # Fazer o download do arquivo Excel gerado
    return send_file(output_txt_file, as_attachment=True)

# Rota para upload e processamento do arquivo Excel
@app.route('/process_excel', methods=['POST'])
def process_excel():
    if 'file' not in request.files or 'organization' not in request.form:
        return "Arquivo ou organização não fornecidos.", 400

    file = request.files['file']
    organization = request.form['organization']

    if file.filename == '':
        return "Nenhum arquivo foi selecionado.", 400

    # Salvar o arquivo Excel temporariamente
    input_excel_file = "temp_excel.xlsx"
    file.save(input_excel_file)

    # Nome do arquivo de saída
    output_excel_file = "dados_convertidos.xlsx"

    # Decidir qual função usar com base na organização
    if organization.lower() == "prosper":
        # Processar o arquivo para a organização "Prosper"
        output_excel_file = process_prosper_excel(input_excel_file, output_excel_file)
    elif organization.lower() == "olicon" or organization.lower() == "avulso" or organization.lower() == "":
        # Processar o arquivo para a organização "Olicon"
        output_excel_file = process_olicon_excel(input_excel_file, output_excel_file)
    elif organization.lower() == "tipo 1":
        # Processar o arquivo para a organização padrão
        output_excel_file = excel_to_separated_excel(input_excel_file, output_excel_file)

    # Fazer o download do arquivo Excel gerado
    return send_file(output_excel_file, as_attachment=True)

if __name__ == '__main__':
    if os.environ.get('WERKZEUG_RUN_MAIN') is None:
        Timer(1, open_browser).start()
    app.run(debug=True, host='0.0.0.0')
