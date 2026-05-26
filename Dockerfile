# Dockerfile
FROM python:3.11-slim

WORKDIR /app

# Dépendances système
RUN apt-get update && apt-get install -y \
    poppler-utils \
    libgl1 \
    libglib2.0-0 \
    libsm6 \
    libxext6 \
    libxrender-dev \
    && rm -rf /var/lib/apt/lists/*

# Copier les requirements
COPY requirements.txt .

# Installer PyTorch CPU d'abord (plus léger, suffisant sur serveur Linux)
RUN pip install torch torchvision --index-url https://download.pytorch.org/whl/cpu

# Installer les autres dépendances
RUN pip install \
    easyocr==1.7.2 \
    PyMuPDF==1.27.2 \
    pdf2image==1.17.0 \
    ollama==0.6.1 \
    pillow==12.1.1 \
    opencv-python-headless==4.13.0.92 \
    fastapi \
    uvicorn \
    msal \
    requests \
    python-dotenv

# Copier le code
COPY *.py .

# Pré-télécharger les modèles EasyOCR au build
RUN python3 -c "import easyocr; easyocr.Reader(['fr', 'en', 'de'])"

CMD ["python3", "agent-compta.py"]