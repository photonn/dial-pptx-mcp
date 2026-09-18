FROM python:3.14-slim

# Run as non-root
RUN useradd --create-home --uid 1000 appuser
WORKDIR /app

# LibreOffice renders slides to PDF for the visual-inspection tool.
# Common free fonts reduce font-substitution drift in renders.
RUN apt-get update && apt-get install -y --no-install-recommends \
        libreoffice-impress fonts-liberation fonts-dejavu \
    && rm -rf /var/lib/apt/lists/*

# Every conversion runs in its own profile and its own $HOME (that isolation
# is what stops concurrent conversions fighting over the profile lock), so
# LibreOffice otherwise builds a user profile and a fontconfig cache from
# nothing every single time — on a one-slide deck that costs more than
# rendering the slide. Build both once here; each conversion starts from a
# private copy of them (visual_qa._seed_dir). soffice can exit non-zero on a
# trivial input while still having written the profile, hence the ';'.
RUN mkdir -p /opt/lo-seed/work /opt/lo-profile-template /opt/lo-cache-template \
    && printf '' > /opt/lo-seed/work/seed.txt \
    && HOME=/opt/lo-seed XDG_CACHE_HOME=/opt/lo-cache-template \
       soffice --headless --norestore \
               -env:UserInstallation=file:///opt/lo-profile-template \
               --convert-to pdf --outdir /opt/lo-seed/work \
               /opt/lo-seed/work/seed.txt \
       ; fc-cache -f \
    ; rm -rf /opt/lo-seed \
    && chown -R appuser:appuser /opt/lo-profile-template /opt/lo-cache-template

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .
USER appuser

# Remote MCP server defaults: streamable-http on 0.0.0.0:8000, one log line
# per event on stderr. Override via PPT_MCP_TRANSPORT / PPT_MCP_HOST /
# PPT_MCP_PORT / LOG_LEVEL.
# The two *_TEMPLATE paths are the profile and cache built above. Unset them
# (or point them at nothing) and conversions behave exactly as before, just
# slower; `convert_ok seeded_profile=` at LOG_LEVEL=DEBUG says which happened.
ENV PPT_MCP_TRANSPORT=http \
    PPT_MCP_HOST=0.0.0.0 \
    PPT_MCP_PORT=8000 \
    PPT_MCP_SOFFICE_PROFILE_TEMPLATE=/opt/lo-profile-template \
    PPT_MCP_SOFFICE_CACHE_TEMPLATE=/opt/lo-cache-template \
    LOG_LEVEL=INFO \
    PYTHONUNBUFFERED=1

EXPOSE 8000

ENTRYPOINT ["python", "ppt_mcp_server.py"]
