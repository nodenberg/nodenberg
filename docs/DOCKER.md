# Docker Deployment Guide

This document explains how to deploy and operate the Nodenberg API Server with Docker.

## Prerequisites

- Docker 20.10 or later
- Docker Compose 2.0 or later

Verify your installation:

```bash
docker --version
docker compose version
```

## Quick Start

### 1. Configure environment variables

Create a `.env` file:

```bash
# Copy the example file
cp .env.example .env

# Generate and append a secure API key
echo "API_KEY=$(openssl rand -base64 32)" >> .env
```

### 2. Start with Docker Compose

```bash
# Build and start in detached mode
docker compose up -d

# Follow logs
docker compose logs -f
```

### 3. Verify the service

```bash
# Health check endpoint
curl http://localhost:3100/health

# Example response:
# {
#   "status": "ok",
#   "timestamp": "2025-12-14T...",
#   "service": "Nodenberg API Server",
#   "version": "0.0.1"
# }
```

## Docker Compose Commands

### Start and stop

```bash
# Start in detached mode
docker compose up -d

# Start in foreground
docker compose up

# Stop and remove containers
docker compose down

# Stop and remove containers and volumes
docker compose down -v
```

### Build

```bash
# Rebuild images
docker compose build

# Rebuild without cache
docker compose build --no-cache

# Build and start
docker compose up -d --build
```

### Logs

```bash
# Show all logs
docker compose logs

# Follow logs
docker compose logs -f

# Show last 100 lines
docker compose logs --tail=100
```

### Container management

```bash
# List running containers
docker compose ps

# Open a shell in the container
docker compose exec nodenberg-api bash

# List all containers (including stopped)
docker compose ps -a
```

## Manual Docker Build and Run

If you prefer not to use Docker Compose, run the container manually.

### Build image

```bash
docker build -t nodenberg-api:latest .
```

### Run container

```bash
docker run -d \
  --name nodenberg-api \
  -p 3100:3100 \
  -e API_KEY=your-secret-api-key-here \
  -e PORT=3100 \
  nodenberg-api:latest
```

### View logs

```bash
docker logs -f nodenberg-api
```

### Stop and remove

```bash
docker stop nodenberg-api
docker rm nodenberg-api
```

## Environment Variables

Variables used by Docker Compose:

| Variable | Default | Description |
| --- | --- | --- |
| `PORT` | `3100` | API server port |
| `API_KEY` | `default-secret-key-please-change-this` | API authentication key (required) |
| `NODE_ENV` | `production` | Runtime environment |

### Option 1: `.env` file

```bash
PORT=3100
API_KEY=your-secret-api-key-here
NODE_ENV=production
```

### Option 2: `docker-compose.yml`

```yaml
environment:
  - PORT=3100
  - API_KEY=your-secret-api-key-here
```

### Option 3: command line

```bash
API_KEY=your-key docker compose up -d
```

## Security

### Generate API keys

```bash
# Base64 (32 bytes)
openssl rand -base64 32

# Hex (32 bytes)
openssl rand -hex 32
```

### Protect `.env`

```bash
# Restrict file permissions
chmod 600 .env

# Confirm .env is ignored by Git
grep .env .gitignore
```

## Health Checks

Docker Compose includes a health check configuration:

```yaml
healthcheck:
  test: ["CMD", "curl", "-f", "http://localhost:3100/health"]
  interval: 30s
  timeout: 10s
  retries: 3
  start_period: 40s
```

Check health status:

```bash
# Quick status
docker compose ps

# Detailed health information
docker inspect nodenberg-api | jq '.[0].State.Health'
```

## Troubleshooting

### Container does not start

```bash
# Check logs
docker compose logs

# Inspect container details
docker inspect nodenberg-api
```

Common causes:

- Port `3100` is already in use.
- `API_KEY` is not set.
- Docker daemon is not running.

### Resolve port conflicts

```bash
# Find process using port 3100
lsof -i :3100
sudo netstat -tulpn | grep 3100

# Change host port in docker-compose.yml
ports:
  - "3200:3100"  # host:container
```

### LibreOffice issues

```bash
# Enter the container
docker compose exec nodenberg-api bash

# Verify LibreOffice
soffice --version

# Reinitialize LibreOffice
node scripts/init-libreoffice.js
```

### Memory issues

```bash
# Add memory limits in docker-compose.yml
services:
  nodenberg-api:
    deploy:
      resources:
        limits:
          memory: 2G
        reservations:
          memory: 1G
```

## Production Deployment

### Recommended Compose configuration

```yaml
# docker-compose.prod.yml
version: '3.8'

services:
  nodenberg-api:
    build: .
    container_name: nodenberg-api-prod
    ports:
      - "3100:3100"
    environment:
      - NODE_ENV=production
      - PORT=3100
      - API_KEY=${API_KEY}
    restart: always
    deploy:
      resources:
        limits:
          memory: 2G
          cpus: '1.0'
    healthcheck:
      test: ["CMD", "curl", "-f", "http://localhost:3100/health"]
      interval: 30s
      timeout: 10s
      retries: 3
      start_period: 40s
    logging:
      driver: "json-file"
      options:
        max-size: "10m"
        max-file: "3"
```

### Start in production mode

```bash
# Use production compose file
docker compose -f docker-compose.prod.yml up -d

# Or provide NODE_ENV at runtime
NODE_ENV=production docker compose up -d
```

## Updates and Maintenance

### Update application

```bash
# 1. Pull latest code
git pull

# 2. Rebuild image
docker compose build --no-cache

# 3. Restart containers
docker compose down
docker compose up -d

# 4. Confirm startup from logs
docker compose logs -f
docker compose logs -f nodenberg-api
```

### Backup

```bash
# Backup environment configuration
cp .env .env.backup

# Export image
docker save nodenberg-api:latest > nodenberg-api.tar

# Import image
docker load < nodenberg-api.tar
```

## External Access

### Reverse proxy (Nginx)

```nginx
# /etc/nginx/sites-available/nodenberg-api
server {
    listen 80;
    server_name api.yourdomain.com;

    location / {
        proxy_pass http://localhost:3100;
        proxy_http_version 1.1;
        proxy_set_header Upgrade $http_upgrade;
        proxy_set_header Connection 'upgrade';
        proxy_set_header Host $host;
        proxy_cache_bypass $http_upgrade;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
    }
}
```

### SSL/TLS (Let's Encrypt)

```bash
sudo certbot --nginx -d api.yourdomain.com
```

## Testing

### Run tests inside the container

```bash
# Enter container
docker compose exec nodenberg-api bash

# Run tests
npm test
```

### Test API from host

```bash
# Set API key
export API_KEY="your-secret-api-key"

# Health check
curl http://localhost:3100/health

# Get template info (requires API key)
curl -X POST \
  -H "Content-Type: application/json" \
  -H "X-API-Key: $API_KEY" \
  -d '{"templateBase64":"..."}' \
  http://localhost:3100/template/info
```

## Related Documents

- [API.md](API.md): API specification
- [README.md](README.md): Project overview
- [tests/README.md](tests/README.md): Test guide

## Best Practices

Do:

- Manage API keys through environment variables.
- Set `restart: always` in production.
- Configure log rotation.
- Keep health checks enabled.
- Set memory limits.

Do not:

- Hardcode API keys.
- Commit `.env` to Git.
- Run containers as root.
- Use the `latest` tag in production.
- Keep unlimited logs without rotation.

## Support

If you run into issues:

1. Check known issues at [GitHub Issues](https://github.com/nodenberg/nodenberg/issues).
2. Review logs with `docker compose logs`.
3. Check container status with `docker compose ps`.
4. Create a new issue if needed.

Last updated: 2025-12-14
