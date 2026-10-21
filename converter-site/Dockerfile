FROM python:3.11-slim

# LibreOffice is needed for PPT -> PDF
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-impress libreoffice-core fonts-liberation fonts-dejavu-core \
    && rm -rf /var/lib/apt/lists/*

# Hugging Face Spaces runs containers as user 1000
RUN useradd -m -u 1000 user
USER user
ENV PATH=/home/user/.local/bin:$PATH \
    HOME=/home/user \
    PORT=7860

WORKDIR /home/user/app

COPY --chown=user requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY --chown=user app.py .

EXPOSE 7860
CMD ["python", "app.py"]
