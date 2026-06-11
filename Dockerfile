FROM python:3.12-slim

# Create a non-root user to run the application
RUN useradd --create-home --shell /bin/bash app

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY outlook_mcp/ ./outlook_mcp/
COPY server.py .

# Transfer ownership so non-root user can read the files
RUN chown -R app:app /app

USER app

ENV OUTLOOK_TOKEN=""
ENV OUTLOOK_PROXY=""
# SSL verification is enabled by default; set to "false" only in trusted dev environments
ENV OUTLOOK_VERIFY_SSL="true"

EXPOSE 8000

HEALTHCHECK --interval=30s --timeout=10s --start-period=5s --retries=3 \
    CMD python -c "import outlook_mcp.config; import sys; sys.exit(0 if outlook_mcp.config.OUTLOOK_TOKEN else 1)"

CMD ["python", "server.py"]
