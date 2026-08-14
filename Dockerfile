FROM python:3.12-slim

# Run as non-root
RUN useradd --create-home --uid 1000 appuser
WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .
USER appuser

# Remote MCP server defaults: streamable-http on 0.0.0.0:8000.
# Override via PPT_MCP_TRANSPORT / PPT_MCP_HOST / PPT_MCP_PORT.
ENV PPT_MCP_TRANSPORT=http \
    PPT_MCP_HOST=0.0.0.0 \
    PPT_MCP_PORT=8000

EXPOSE 8000

ENTRYPOINT ["python", "ppt_mcp_server.py"]
