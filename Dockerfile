# Usar imagen base con Python
FROM python:3.11-slim

# Instalar dependencias del sistema necesarias para Pillow y Node.js
RUN apt-get update && apt-get install -y \
    gcc \
    python3-dev \
    libjpeg-dev \
    zlib1g-dev \
    curl \
    && rm -rf /var/lib/apt/lists/*

# Instalar Node.js 20
RUN curl -fsSL https://deb.nodesource.com/setup_20.x | bash - && \
    apt-get install -y nodejs && \
    rm -rf /var/lib/apt/lists/*

# Establecer directorio de trabajo
WORKDIR /app

# Copiar archivos de dependencias
COPY requirements.txt package.json ./

# Instalar dependencias de Python
RUN pip install --no-cache-dir -r requirements.txt

# Instalar dependencias de Node.js
RUN npm install

# Copiar el resto de la aplicación
COPY . .

# Crear directorio para OTs generadas
RUN mkdir -p ordenes_generadas

# Exponer puerto
EXPOSE 10000

# Comando para iniciar la aplicación
CMD ["gunicorn", "--bind", "0.0.0.0:10000", "app:app"]
