FROM python:3.11-slim

# Install system dependencies if any are needed
RUN apt-get update && apt-get install -y --no-install-recommends \
    build-essential \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Copy requirements and install
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy all python files and config/templates
COPY accessibility.py .
COPY convert.py .
COPY main.py .
COPY pdfViewer.py .
COPY report.py .
COPY server.py .
COPY customers.json .

# Create uploads directory
RUN mkdir -p uploads

EXPOSE 8001

CMD ["uvicorn", "server:app", "--host", "0.0.0.0", "--port", "8001"]
