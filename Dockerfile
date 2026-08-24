FROM python:3.14-slim

# Run as non-root
RUN useradd --create-home --uid 1000 appuser
WORKDIR /app

# LibreOffice renders slides to PDF for the visual-inspection tool.
# Common free fonts reduce font-substitution drift in renders.
RUN apt-get update && apt-get install -y --no-install-recommends \
        libreoffice-impress fonts-liberation fonts-dejavu \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .
USER appuser

# Remote MCP server defaults: streamable-http on 0.0.0.0:8000, one log line
# per event on stderr. Override via PPT_MCP_TRANSPORT / PPT_MCP_HOST /
# PPT_MCP_PORT / LOG_LEVEL.
ENV PPT_MCP_TRANSPORT=http \
    PPT_MCP_HOST=0.0.0.0 \
    PPT_MCP_PORT=8000 \
    LOG_LEVEL=INFO \
    PYTHONUNBUFFERED=1

EXPOSE 8000

ENTRYPOINT ["python", "ppt_mcp_server.py"]
