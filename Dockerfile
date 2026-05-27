FROM python:3.11-slim

# Install LibreOffice Writer (headless PDF conversion)
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-writer \
    libreoffice-java-common \
    default-jre-headless \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements_server.txt .
RUN pip install --no-cache-dir -r requirements_server.txt

# Copy source files and assets the server needs
COPY mcp_server.py .
COPY generate_receipts.py .
COPY "PayPal_Donation Receipt.docx" .
COPY zc_logo.png .

EXPOSE 8000

CMD ["python", "mcp_server.py"]
