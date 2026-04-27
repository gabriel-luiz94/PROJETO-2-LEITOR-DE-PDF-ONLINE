# Dockerfile para o backend FastAPI (Leitor PDF Online)
FROM python:3.11-slim

# Evita prompts durante instalação de pacotes
ENV DEBIAN_FRONTEND=noninteractive
ENV PYTHONUNBUFFERED=1
ENV PORT=8080

WORKDIR /app

# Instala dependências do sistema necessárias para o PyMuPDF
RUN apt-get update && apt-get install -y --no-install-recommends \
    libmupdf-dev \
    mupdf \
    mupdf-tools \
    && rm -rf /var/lib/apt/lists/*

# Copia e instala dependências Python primeiro (cache de layer)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copia o código da aplicação
COPY app.py .
COPY static/ ./static/

# Expõe a porta interna
EXPOSE 8080

# Comando de start (usa variável PORT do ambiente)
CMD ["sh", "-c", "uvicorn app:app --host 0.0.0.0 --port ${PORT}"]
