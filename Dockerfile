# syntax=docker/dockerfile:1

FROM node:20-alpine AS build

WORKDIR /app
COPY package.json package-lock.json ./
RUN npm ci

COPY tsconfig.json ./
COPY src ./src
RUN npm run build

FROM node:20-alpine

WORKDIR /app
COPY --from=build /app/package.json /app/package-lock.json ./
COPY --from=build /app/node_modules ./node_modules
COPY --from=build /app/dist ./dist

ENV TEAMS_MCP_TRANSPORT=http
ENV TEAMS_MCP_PORT=8888

EXPOSE 8888

CMD ["node", "dist/index.js"]
