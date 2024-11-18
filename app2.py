from flask import Flask, request, send_file, abort
import pandas as pd
import os, re
import webbrowser
from threading import Timer
from datetime import datetime


app = Flask(__name__)

# Função para abrir o navegador após o servidor começar
def open_browser():
    webbrowser.open_new("http://127.0.0.1:5000/")

# Função para gerar o arquivo .txt com as colunas solicitadas e preencher com valores padrão
def excel_to_custom_txt(input_excel_file, output_txt_file):
    # Ler o arquivo Excel
    df = pd.read_excel(input_excel_file)

    

    # Definir as colunas que serão usadas no arquivo .txt
    columns = ["Nome", "Endereco", "Medidor", "Localizacao", "Tipo de Fluido", "Data", "Horario", 
               "Indice", "Indice antigo", "Consumo", "Unid._levant.", "Bateria", "Ha desmontagem/corte", 
               "Houve desmontagem/corte", "Ha vazamento", "Ha fraude magnetica", "Houve fraude magnetica", 
               "Ha medidor bloqueado", "Retorno de agua excedido"]

    # Preencher colunas com valores padrão, se elas não existirem ou estiverem vazias
    # Preencher colunas com valores de data e hora atuais
    current_datetime = datetime.now()

    df["Nome"] = df.get("Nome", "Nome do Cliente")  # Valor padrão: "Água"
    df["Endereco"] = df.get("Endereco", "Rua Ipiranga, 195 - Vila Paiva")  # Valor padrão: "Água"
    df["Tipo de Fluido"] = df.get("Tipo de Fluido", "Água Fria")  # Valor padrão: "Água"
    df["Consumo"] = df.get("Consumo", "0")  # Valor padrão: "Água"
    df["Unid._levant."] = df.get("Unid._levant.", "l")  # Valor padrão: "Água"
    df["Bateria"] = df.get("Bateria", "-")  # Valor padrão: "Normal"
    df["Data"] = df.get("Data", current_datetime.strftime("%d/%m/%Y"))  # Exemplo de formato de data
    df["Horario"] = df.get("Horario", current_datetime.strftime("%H:%M:%S"))  # Exemplo de formato de horário
    df["Ha desmontagem/corte"] = df.get("Ha desmontagem/corte", "0")  # Valor padrão: "Não"
    df["Houve desmontagem/corte"] = df.get("Houve desmontagem/corte", "0")  # Valor padrão: "Não"
    df["Ha vazamento"] = df.get("Ha vazamento", "0")  # Valor padrão: "Não"
    df["Ha fraude magnetica"] = df.get("Ha fraude magnetica", "0")  # Valor padrão: "Não"
    df["Houve fraude magnetica"] = df.get("Houve fraude magnetica", "0")  # Valor padrão: "Não"
    df["Ha medidor bloqueado"] = df.get("Ha medidor bloqueado", "0")  # Valor padrão: "Não"
    df["Retorno de agua excedido"] = df.get("Retorno de agua excedido", "0")  # Valor padrão: "Não"

    df["Medidor"] = df.get("Nome")  # Capture os valores da coluna "Unidade"

  
    processed_data = []

    processed_data.extend(columns)
 
    def process_row(input_excel_file):
        # Inicializar as colunas com valores vazios
       
        Nome = Endereco = Medidor = Localizacao = Tipo_de_Fluido = sigla = Data = Horario = Indice = Indice_antigo = Consumo = ""
        
        col1 = str(input_excel_file[0])  # Primeira coluna
        col2 = str(input_excel_file[1])  # Segunda coluna
        col3 = str(input_excel_file[2])
        col4 = str(input_excel_file[3])
      
        print(input_excel_file)

        # Iterar sobre os valores na linha e mapear para as colunas corretas
        for word in input_excel_file:
            if word.startswith("MED"):
                bloco = word
            elif word.isalpha():
                junto = word + input_excel_file[1]
                match = re.match(r'([A-Za-z]+)(\d+)', junto)
                if match:
                    unidade = match.group(1)  # Parte da palavra (ex: "SALA")
                    ambiente = match.group(2)  # Parte do número (ex: "31")
         

            else:
                medidor = word
                fracao_ideal = word

            # if is_numeric(col1):
            #     val = col1
            #     print(val)
            # if col1.isalpha():
            #     val2 = col1
            #     print(val2)

            # if not is_numeric(col1):
            #     val3 = col1
            #     print(val3)

        Nome = "Único"
        sigla = "Água Fria"
   
      
     
        
        return [Nome, Endereco, Medidor, Localizacao, Tipo_de_Fluido, sigla, Data, Horario, Indice, Indice_antigo, Consumo]


    # Abrir o arquivo .txt para escrita
    with open(output_txt_file, 'w') as txt_file:
        # Escrever os cabeçalhos no arquivo .txt
        txt_file.write('\t'.join(columns) + '\n')  # Tabulação (\t) separa as colunas
        
        # Iterar sobre as linhas do DataFrame e escrever os dados no arquivo .txt
        for index, row in df.iterrows():
            # Criar uma linha com os valores separados por tabulação
            line = '\t'.join([str(row.get(col, "")) for col in columns]) + '\t' + str(row["Medidor"])

            txt_file.write(line + '\n')
            processed_row = process_row(columns)
            processed_data.append(processed_row)
    
    
    processed_df = pd.DataFrame(processed_data)
    print(f"Arquivo TXT {output_txt_file} gerado com sucesso.")
    return output_txt_file

# Página inicial para upload do arquivo e escolha de formato
@app.route('/')
def index():
    return '''
    <!doctype html>
    <html lang="pt-br">
      <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">
        <title>Conversão de Arquivo Excel para TXT</title>
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
          input[type="file"], input[type="radio"] {
            padding: 10px;
            margin-bottom: 20px;
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
          <h1>Converter Arquivo Excel para TXT</h1>
          <p>Envie seu arquivo Excel para gerar um arquivo .txt com colunas específicas.</p>
          <form action="/upload" method="post" enctype="multipart/form-data" id="uploadForm">
            <input type="file" name="file" id="fileInput" required><br>
            <input type="submit" value="Converter e Baixar">
            <div class="note">Formato suportado: .xlsx</div>
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

    if file.filename == '':
        return "Nenhum arquivo foi selecionado.", 400

    # Salvar o arquivo Excel temporariamente
    input_excel_file = "temp_excel.xlsx"
    file.save(input_excel_file)

    # Nome do arquivo de saída
    output_txt_file = "dados_convertidos.txt"

    # Converter o Excel para o arquivo .txt com as colunas solicitadas e valores padrão
    output_file = excel_to_custom_txt(input_excel_file, output_txt_file)

    # Fazer o download do arquivo gerado
    return send_file(output_file, as_attachment=True)

if __name__ == '__main__':
    if os.environ.get('WERKZEUG_RUN_MAIN') is None:
        Timer(1, open_browser).start()
    app.run(debug=True)
