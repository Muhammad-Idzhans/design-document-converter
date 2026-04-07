# Use a lightweight Python 3.11 Linux image
FROM python:3.11-slim

# Prevent Python from writing pyc files to disc and buffering stdout
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1

# CRITICAL for LibreOffice in Docker: Ensure it has a writable home directory
ENV HOME=/tmp

WORKDIR /app

# Install system dependencies required for document conversion
# - pandoc: required by pypandoc to convert markdown to docx
# - libreoffice: required to convert pptx/docx to pdf on Linux
# - fonts-liberation: ensures LibreOffice renders standard fonts correctly
RUN apt-get update && apt-get install -y \
    pandoc \
    libreoffice \
    fonts-liberation \
    && rm -rf /var/lib/apt/lists/*

# Copy and install Python dependencies
COPY requirements.txt /app/requirements.txt
RUN pip install --no-cache-dir -r requirements.txt

# Copy the rest of your application code (app.py, template, etc.)
COPY . /app

# Expose the port FastAPI will run on
EXPOSE 8000

# Start FastAPI using Uvicorn
# IMPORTANT: If your Python file is named main.py instead of app.py, change "app:app" to "main:app"
CMD ["uvicorn", "app:app", "--host", "0.0.0.0", "--port", "8000", "--workers", "4"]