FROM node:20-alpine

RUN apk add --no-cache wget

WORKDIR /app
COPY package*.json ./
RUN npm install --omit=dev
COPY index.js ./
RUN mkdir -p /var/log/mcp
EXPOSE 8000
HEALTHCHECK --interval=30s --timeout=5s --retries=3 \
  CMD wget -qO- http://localhost:8000/health || exit 1
CMD ["node", "index.js"]
