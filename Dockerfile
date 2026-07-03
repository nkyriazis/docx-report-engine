# ---------------------------------------------------------------
# Stage 1 — Node.js tools (mmdc + Puppeteer)
# ---------------------------------------------------------------
FROM node:20-slim AS node-tools

# Puppeteer system deps — Chromium won't run without these
RUN apt-get update && apt-get install -y --no-install-recommends \
    ca-certificates \
    fonts-liberation \
    fonts-noto-color-emoji \
    libasound2 \
    libatk-bridge2.0-0 \
    libatk1.0-0 \
    libcups2 \
    libdbus-1-3 \
    libdrm2 \
    libgbm1 \
    libgtk-3-0 \
    libnspr4 \
    libnss3 \
    libx11-xcb1 \
    libxcomposite1 \
    libxdamage1 \
    libxrandr2 \
    xdg-utils \
    && rm -rf /var/lib/apt/lists/*

# Install mermaid-cli globally — this also fetches Puppeteer/Chromium
RUN npm install -g @mermaid-js/mermaid-cli@latest \
    && npx puppeteer browsers install chrome

# ---------------------------------------------------------------
# Stage 2 — Final Python image
# ---------------------------------------------------------------
FROM python:3.11-slim AS builder

# Pandoc + build essentials (for pip wheels)
RUN apt-get update && apt-get install -y --no-install-recommends \
    pandoc \
    gcc \
    g++ \
    && rm -rf /var/lib/apt/lists/*

# Copy Python deps first for layer caching
COPY requirements.txt /tmp/requirements.txt
RUN pip install --no-cache-dir --prefix=/install -r /tmp/requirements.txt

# ---------------------------------------------------------------
# Stage 3 — Runtime (slim, no build tools)
# ---------------------------------------------------------------
FROM python:3.11-slim AS runtime

# Runtime system deps: pandoc + Chromium libs
RUN apt-get update && apt-get install -y --no-install-recommends \
    pandoc \
    ca-certificates \
    fonts-liberation \
    fonts-noto-color-emoji \
    libasound2 \
    libatk-bridge2.0-0 \
    libatk1.0-0 \
    libcups2 \
    libdbus-1-3 \
    libdrm2 \
    libgbm1 \
    libgtk-3-0 \
    libnspr4 \
    libnss3 \
    libx11-xcb1 \
    libxcomposite1 \
    libxdamage1 \
    libxrandr2 \
    xdg-utils \
    curl \
    && rm -rf /var/lib/apt/lists/*

# Copy Python packages from builder
COPY --from=builder /install /usr/local

# Copy Node.js + mmdc from node-tools (mmdc is an ES module, needs full node_modules)
COPY --from=node-tools /usr/local/lib/node_modules /usr/local/lib/node_modules
COPY --from=node-tools /usr/local/bin/node /usr/local/bin/node
RUN ln -s /usr/local/lib/node_modules/@mermaid-js/mermaid-cli/src/cli.js /usr/local/bin/mmdc

# Copy Chrome from node-tools (Puppeteer 2.x cache — need full dir for ICU data)
COPY --from=node-tools /root/.cache/puppeteer/chrome /opt/chromium/cache
RUN CHROME_DIR=$(find /opt/chromium/cache -name chrome-linux64 -type d | head -1) && \
    mkdir -p /opt/chromium/bin && \
    ln -sf "$CHROME_DIR/chrome" /opt/chromium/bin/chrome

# Set PATH so mmdc is found
ENV PATH="/usr/local/lib/node_modules/@mermaid-js/mermaid-cli/bin:/usr/local/bin:$PATH"
# Tell Puppeteer where Chromium lives
ENV PUPPETEER_EXECUTABLE_PATH="/opt/chromium/bin/chrome"

# Bake the engine into the image as a fallback; a project that mounts its
# repo at /app with an engine/ checkout (e.g. a git submodule) takes
# precedence via PYTHONPATH ordering.
COPY . /opt/engine
ENV PYTHONPATH="/app/engine:/opt/engine"

WORKDIR /app

# Default: run the build
ENTRYPOINT ["python3", "-m", "docx_builder"]
CMD []
