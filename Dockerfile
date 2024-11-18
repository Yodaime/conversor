# Usa uma imagem base do Python
FROM python:3.9-slim

# Define o diretório de trabalho dentro do contêiner
WORKDIR /app

# Copia os arquivos do projeto para o contêiner
COPY . /app

# Instala as dependências do projeto
RUN pip install --no-cache-dir -r requirements.txt

# Exponha a porta que a aplicação usará
EXPOSE 5000

# Define o comando padrão para executar a aplicação
CMD ["python", "app.py"]

