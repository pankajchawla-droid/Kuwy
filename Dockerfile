FROM node:20-slim

# System deps: GraphicsMagick + Ghostscript (used by pdf2pic to rasterize PDF pages)
# and Chromium + its runtime libs (used by Puppeteer to render HTML report pages)
RUN apt-get update && apt-get install -y \
    graphicsmagick \
    ghostscript \
    chromium \
    fonts-liberation \
    libnss3 libatk1.0-0 libatk-bridge2.0-0 libcups2 libdrm2 libxkbcommon0 \
    libxcomposite1 libxdamage1 libxfixes3 libxrandr2 libgbm1 libasound2 \
    --no-install-recommends && rm -rf /var/lib/apt/lists/*

ENV PUPPETEER_SKIP_CHROMIUM_DOWNLOAD=true \
    PUPPETEER_EXECUTABLE_PATH=/usr/bin/chromium

WORKDIR /app
COPY package.json ./
RUN npm install --omit=dev
COPY . .

ENV PORT=3000
EXPOSE 3000
CMD ["npm", "start"]
