FROM python:3.11-slim

# Install system dependencies and LibreOffice (headless converter for Linux)
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Install Python dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy all backend source files
COPY . .

# Expose the Flask port
EXPOSE 5000

# Run the app via Gunicorn (production WSGI server)
CMD ["gunicorn", "--bind", "0.0.0.0:5000", "--timeout", "180", "app:app"]
