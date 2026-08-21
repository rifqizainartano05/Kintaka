FROM python:3.11-slim

WORKDIR /app

# Install system dependencies
RUN apt-get update && apt-get install -y \
    build-essential \
    libgl1-mesa-glx \
    libglib2.0-0 \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .

RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Expose port 8000
EXPOSE 8000

# Run Flask using Gunicorn (production server)
RUN pip install gunicorn
CMD ["gunicorn", "--bind", "0.0.0.0:8000", "--timeout", "120", "main:app"]
