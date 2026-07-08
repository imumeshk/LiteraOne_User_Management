# Docker Build & Deployment Instructions

This guide provides detailed instructions on how to build, run, and troubleshoot the Litera One Manager Portal using Docker.

## Prerequisites

- [Docker](https://docs.docker.com/get-docker/) installed and running.
- [Docker Compose](https://docs.docker.com/compose/install/) (optional, but recommended for easier orchestration).

## 1. Environment Configuration

Before running the container, you need to configure your environment variables. 
Create a `.env` file in the root of your project based on the `.env.example` file.

```bash
cp .env.example .env
```

**Required Environment Variables for Docker:**
- `SESSION_SECRET`: Must be a strong, random string. Required for session encryption.
- `HTTPS_PROXY` / `HTTP_PROXY`: (Optional) Set these if your network requires a proxy to reach external endpoints (like Microsoft Graph).
- `NO_PROXY`: (Optional) Exclude localhost from proxy routing.

## 2. Using Docker Compose (Recommended)

The easiest way to build and run the application is using Docker Compose. The `docker-compose.yml` file is already configured.

### Build and Start

To build the image and start the container in detached mode (background):

```bash
docker compose up --build -d
```

### Stop the Container

To stop the running container:

```bash
docker compose down
```

### View Logs

To view the application logs:

```bash
docker compose logs -f
```

## 3. Manual Docker Build & Run (Without Compose)

If you prefer to use the standard Docker CLI, follow these steps.

### Build the Image

Build the Docker image and tag it as `litera-one-portal`:

```bash
docker build -t litera-one-portal .
```

### Run the Container

Run the image, passing the required environment variables and mapping port 3000:

```bash
docker run -d \
  --name litera-one-portal \
  -p 3000:3000 \
  -e SESSION_SECRET=your-very-strong-secret-key \
  litera-one-portal
```

*Note: For production, you may want to set `SESSION_COOKIE_SECURE=true` if running behind an HTTPS reverse proxy.*

## 4. Health Checks and Connectivity

The container exposes a health check endpoint at `/health`. Docker automatically monitors this.

To manually check if the container is healthy:
```bash
curl http://localhost:3000/health
```

### Troubleshooting Connectivity

If the portal shows `fetch failed` when attempting to sign in, the container likely cannot reach Microsoft login endpoints (often due to corporate firewalls/proxies).

You can run a standalone connectivity check using the provided utility in `docker-compose.yml`:

```bash
docker compose run --rm ms-connectivity-check
```

If it fails, ensure your `.env` file includes the correct `HTTPS_PROXY` settings and rebuild/restart the container.

## 5. Building for Production/Registry

To build and tag the image for a container registry (e.g., Azure Container Registry, Docker Hub):

```bash
# Build the image with your registry tag
docker build -t your-registry.azurecr.io/litera-one-portal:v1.0.0 .

# Log in to your registry
docker login your-registry.azurecr.io

# Push the image
docker push your-registry.azurecr.io/litera-one-portal:v1.0.0
```
