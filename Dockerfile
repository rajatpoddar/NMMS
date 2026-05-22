# ============================================
# NMMS Tracking Report - Production Server
# ============================================
FROM python:3.11-slim

LABEL maintainer="Nrega Bot Team"
LABEL description="NMMS Tracking Report - License Management Server"

# Install system dependencies for psycopg2
RUN apt-get update && apt-get install -y --no-install-recommends \
    libpq-dev \
    postgresql-client \
    && rm -rf /var/lib/apt/lists/*

# Set environment variables
ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

# Create non-root user for security
RUN addgroup --system --gid 1001 appuser && \
    adduser --system --uid 1001 --gid 1001 --no-create-home appuser

# Set working directory
WORKDIR /app

# Install Python dependencies
COPY requirements-server.txt requirements.txt
RUN pip install --no-cache-dir -r requirements.txt && \
    rm -rf /root/.cache/pip

# Copy application code
COPY server.py .
COPY scraper_worker.py .
COPY entrypoint.sh .

# Create output directories for extraction tasks
RUN mkdir -p /app/outputs /app/tasks && \
    chown appuser:appuser /app/outputs /app/tasks

# Make entrypoint executable
RUN chmod +x entrypoint.sh && chown appuser:appuser /app/*

# Switch to non-root user
USER appuser

# Expose the server port
EXPOSE 6667

# Health check
HEALTHCHECK --interval=30s --timeout=5s --start-period=15s --retries=3 \
    CMD python -c "import urllib.request; urllib.request.urlopen('http://localhost:6667/health')" || exit 1

# Run with gunicorn via entrypoint
ENTRYPOINT ["./entrypoint.sh"]
