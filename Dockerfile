FROM node:20-alpine

WORKDIR /app

# Dependências primeiro (aproveita cache de layer)
COPY package*.json ./
RUN npm ci --omit=dev

# Código
COPY index.js ./

# Diretório de log
RUN mkdir -p /var/log/mcp

EXPOSE 8000

HEALTHCHECK --interval=30s --timeout=5s --retries=3 \
  CMD wget -qO- http://localhost:8000/health || exit 1

CMD ["node", "index.js"]
