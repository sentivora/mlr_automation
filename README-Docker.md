# MLR Automation - Docker Setup for Mac

This guide will help you set up and run the MLR Automation application using Docker on a Mac machine.

## Prerequisites

1. **Docker Desktop for Mac**: Download and install from [https://www.docker.com/products/docker-desktop](https://www.docker.com/products/docker-desktop)
2. **Git**: To clone the repository (if needed)

## Quick Start

### Option 1: Using Docker Compose (Recommended)

1. **Clone or copy the project** to your Mac:
   ```bash
   # If cloning from git
   git clone <repository-url>
   cd "MLR Auto - Phase-01"
   
   # Or if you have the files, navigate to the project directory
   cd "/path/to/MLR Auto - Phase-01"
   ```

2. **Build and run the application**:
   ```bash
   docker-compose up --build
   ```

3. **Access the application**:
   Open your web browser and go to: `http://localhost:5000`

4. **Stop the application**:
   ```bash
   # Press Ctrl+C in the terminal, then run:
   docker-compose down
   ```

### Option 2: Using Docker Commands Directly

1. **Build the Docker image**:
   ```bash
   docker build -t mlr-automation .
   ```

2. **Run the container**:
   ```bash
   docker run -d \
     --name mlr-automation \
     -p 5000:5000 \
     -v "$(pwd)/uploads:/app/uploads" \
     -v "$(pwd)/outputs:/app/outputs" \
     mlr-automation
   ```

3. **Access the application**:
   Open your web browser and go to: `http://localhost:5000`

4. **Stop and remove the container**:
   ```bash
   docker stop mlr-automation
   docker rm mlr-automation
   ```

## File Persistence

The Docker setup includes volume mounts to ensure your files persist:
- **uploads/**: Your uploaded files
- **outputs/**: Generated PowerPoint presentations
- **static/**: Static assets (CSS, JS, images)
- **templates/**: HTML templates

## Troubleshooting

### Port Already in Use
If port 5000 is already in use, change it in the docker-compose.yml:
```yaml
ports:
  - "8080:5000"  # Use port 8080 instead
```

### Permission Issues on Mac
If you encounter permission issues with volumes:
```bash
# Make sure the directories exist and have proper permissions
mkdir -p uploads outputs
chmod 755 uploads outputs
```

### Docker Connection Issues on Windows
If you get "unable to get image" or "docker client must be run with elevated privileges" errors:

1. **Ensure Docker Desktop is running**:
   - Start Docker Desktop from the Start menu
   - Wait for it to fully initialize (whale icon in system tray should be stable)

2. **Run PowerShell as Administrator**:
   - Right-click on PowerShell and select "Run as administrator"
   - Navigate to your project directory
   - Try the docker commands again

3. **Check Docker service status**:
   ```powershell
   # Check if Docker is running
   docker version
   
   # If Docker Desktop isn't starting, restart it
   # Or restart the Docker service
   ```

4. **Alternative: Use WSL2**:
   - If you have WSL2 installed, you can run Docker commands from within WSL2
   - This often resolves permission issues on Windows

### View Container Logs
```bash
# Using docker-compose
docker-compose logs -f

# Using docker directly
docker logs -f mlr-automation
```

### Rebuild After Code Changes
```bash
# Stop the current container
docker-compose down

# Rebuild and start
docker-compose up --build
```

## Development Mode

For development with live code reloading:

1. **Modify docker-compose.yml** to mount the source code:
   ```yaml
   volumes:
     - .:/app
     - ./uploads:/app/uploads
     - ./outputs:/app/outputs
   environment:
     - FLASK_ENV=development
   ```

2. **Run with development settings**:
   ```bash
   docker-compose up --build
   ```

## System Requirements

- **RAM**: Minimum 2GB available for Docker
- **Storage**: At least 1GB free space
- **macOS**: 10.14 or later (for Docker Desktop)

## Security Notes

- Change the `SESSION_SECRET` in docker-compose.yml for production use
- The application runs on all interfaces (0.0.0.0) inside the container
- File uploads are limited to 500MB by default

## Support

If you encounter issues:
1. Check Docker Desktop is running
2. Verify you're in the correct directory
3. Check the container logs for error messages
4. Ensure no other services are using port 5000

## Useful Docker Commands

```bash
# View running containers
docker ps

# View all containers
docker ps -a

# Remove all stopped containers
docker container prune

# Remove unused images
docker image prune

# Enter the running container (for debugging)
docker exec -it mlr-automation /bin/bash
```