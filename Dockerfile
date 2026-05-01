FROM python:3.11-slim

# Install LibreOffice, fonts, and necessary libraries for document conversion
RUN apt-get update && apt-get install -y \
    libreoffice \
    fonts-liberation \
    libgdiplus \
    fontconfig \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Copy requirements and install dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy the rest of the application
COPY . .

# Create necessary directories
RUN mkdir -p static uploads

# Set environment variable for Render
ENV PORT=8000

# Start the application using a shell to handle the PORT variable
CMD ["sh", "-c", "uvicorn main:app --host 0.0.0.0 --port $PORT"]
