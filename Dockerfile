FROM python:3.11-slim

# Instala dependencias del sistema
RUN apt-get update && \
    apt-get install -y libreoffice fonts-liberation fontconfig texlive-xetex pandoc && \
    rm -rf /var/lib/apt/lists/*

# Copia los archivos del proyecto
WORKDIR /app
COPY . /app

# Instala dependencias de Python
RUN pip install --upgrade pip
RUN pip install -r requirements.txt

# Expone el puerto (ajusta si usas otro)
EXPOSE 5000

# Comando para iniciar la app
CMD ["python", "app.py"]