FROM node:22-slim

WORKDIR /app

# better-sqlite3 native modülü için gerekli build araçları + Türkçe font (PDF)
RUN apt-get update && apt-get install -y --no-install-recommends \
    python3 make gcc g++ fonts-dejavu-core \
    && rm -rf /var/lib/apt/lists/*

COPY package*.json ./
RUN npm install

# Build araçlarını kaldır (imaj boyutunu küçültmek için)
RUN apt-get remove -y python3 make gcc g++ \
    && apt-get autoremove -y \
    && rm -rf /var/lib/apt/lists/*

COPY . .

RUN mkdir -p /app/data /app/public/uploads

EXPOSE 3000

CMD ["node", "server.js"]
