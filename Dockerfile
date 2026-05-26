FROM python:3.11-slim

RUN apt-get update && apt-get install -y \
    gcc \
    python3-dev \
    libjpeg-dev \
    zlib1g-dev \
    curl \
    libreoffice \
    libreoffice-l10n-es \
    locales \
    && locale-gen es_PE.UTF-8 \
    && rm -rf /var/lib/apt/lists/*

RUN curl -fsSL https://deb.nodesource.com/setup_20.x | bash - && \
    apt-get install -y nodejs && \
    rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt package.json ./

RUN pip install --no-cache-dir -r requirements.txt

RUN npm install

COPY . .

RUN mkdir -p ordenes_generadas

EXPOSE 10000

CMD ["sh", "-c", "LANG=es_PE.UTF-8 LC_ALL=es_PE.UTF-8 gunicorn --bind 0.0.0.0:10000 app:app"]
