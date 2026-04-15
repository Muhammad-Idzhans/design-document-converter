# Use a lightweight Python 3.11 Linux image
FROM python:3.11-slim

# Prevent Python from writing pyc files to disc and buffering stdout
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1

# Tell Debian not to prompt for user input during installation
ENV DEBIAN_FRONTEND=noninteractive

# CRITICAL for LibreOffice in Docker: Ensure it has a writable home directory
ENV HOME=/tmp

WORKDIR /app

# 1. Enable the "contrib" repository (where Microsoft fonts live in Debian)
# 2. Auto-accept the Microsoft End User License Agreement (EULA) so the build doesn't freeze
# 3. Install Pandoc, LibreOffice, and the fonts
RUN sed -i 's/Components: main/Components: main contrib/g' /etc/apt/sources.list.d/debian.sources 2>/dev/null || sed -i 's/main/main contrib/g' /etc/apt/sources.list \
    && apt-get update \
    && echo "ttf-mscorefonts-installer msttcorefonts/accepted-mscorefonts-eula select true" | debconf-set-selections \
    && apt-get install -y \
        pandoc \
        libreoffice \
        fonts-liberation \
        ttf-mscorefonts-installer \
        fontconfig \
    && fc-cache -f -v \
    && rm -rf /var/lib/apt/lists/*

# Copy custom Windows fonts (like Segoe UI) for high-fidelity LibreOffice rendering
COPY fonts/ /usr/share/fonts/truetype/custom/
RUN fc-cache -f -v


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