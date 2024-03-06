# Use a imagem oficial do Python
FROM python:3.8

# Configurar o diretório de trabalho
WORKDIR /app

# Copiar os arquivos necessários para o diretório de trabalho
COPY requirements.txt .
COPY aplicacaoweb.py .
COPY index.html .

# Instalar as dependências
RUN pip install --upgrade pip
RUN pip install --no-cache-dir -r requirements.txt


# Expor a porta em que a aplicação Flask estará rodando
EXPOSE 5000

# Comando para executar a aplicação quando o contêiner for iniciado
CMD ["python", "aplicacaoweb.py"]