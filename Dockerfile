FROM node:22-alpine AS builder
WORKDIR /app

COPY package*.json ./
RUN npm ci

COPY . .
RUN npm run build

FROM node:22-alpine AS runtime
WORKDIR /app

ENV NODE_ENV=production
ENV SESSION_COOKIE_SECURE=false

COPY package*.json ./
RUN npm ci --omit=dev

COPY src ./src
COPY public ./public
COPY config ./config
COPY enablement-hub.json ./enablement-hub.json
COPY --from=builder /app/dist ./dist

EXPOSE 3000

CMD ["node", "--experimental-strip-types", "src/server.ts"]
