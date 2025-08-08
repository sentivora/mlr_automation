# VPS Deployment Guide for MLR Auto PowerPoint Generator

This guide will help you deploy the MLR Auto PowerPoint Generator application on a VPS using Docker.

## Prerequisites

- VPS with Ubuntu 20.04+ or similar Linux distribution
- Docker and Docker Compose installed
- At least 2GB RAM and 10GB storage
- Domain name (optional but recommended)
- SSL certificate (for HTTPS)

## Quick Start

### 1. Install Docker and Docker Compose

```bash
# Update system
sudo apt update && sudo apt upgrade -y

# Install Docker
curl -fsSL https://get.docker.com -o get-docker.sh
sudo sh get-docker.sh

# Add user to docker group
sudo usermod -aG docker $USER

# Install Docker Compose
sudo curl -L "https://github.com/docker/compose/releases/latest/download/docker-compose-$(uname -s)-$(uname -m)" -o /usr/local/bin/docker-compose
sudo chmod +x /usr/local/bin/docker-compose

# Logout and login again for group changes to take effect
```

### 2. Clone and Setup Application

```bash
# Clone your repository
git clone <your-repository-url>
cd MLR-Auto-Phase-01

# Copy environment file and configure
cp .env.example .env
nano .env  # Edit with your settings
```

### 3. Configure Environment Variables

Edit the `.env` file with your production settings:

```bash
# Generate a secure session secret
python3 -c "import secrets; print('SESSION_SECRET=' + secrets.token_hex(32))"

# Add the generated secret to your .env file
echo "SESSION_SECRET=your-generated-secret-here" >> .env
```

### 4. Build and Deploy

```bash
# Build and start the application
docker-compose up -d --build

# Check if containers are running
docker-compose ps

# View logs
docker-compose logs -f mlr-auto
```

### 5. Configure Reverse Proxy (Nginx)

Create an Nginx configuration for your domain:

```bash
sudo nano /etc/nginx/sites-available/mlr-auto
```

Add the following configuration:

```nginx
server {
    listen 80;
    server_name theexperimentai.org www.theexperimentai.org;
    
    # Redirect HTTP to HTTPS
    return 301 https://$server_name$request_uri;
}

server {
    listen 443 ssl http2;
    server_name theexperimentai.org www.theexperimentai.org;
    
    # SSL Configuration
    ssl_certificate /path/to/your/certificate.crt;
    ssl_certificate_key /path/to/your/private.key;
    ssl_protocols TLSv1.2 TLSv1.3;
    ssl_ciphers ECDHE-RSA-AES256-GCM-SHA512:DHE-RSA-AES256-GCM-SHA512:ECDHE-RSA-AES256-GCM-SHA384:DHE-RSA-AES256-GCM-SHA384;
    ssl_prefer_server_ciphers off;
    
    # Security headers
    add_header X-Frame-Options DENY;
    add_header X-Content-Type-Options nosniff;
    add_header X-XSS-Protection "1; mode=block";
    add_header Strict-Transport-Security "max-age=63072000; includeSubDomains; preload";
    
    # File upload size limit
    client_max_body_size 200M;
    
    location / {
        proxy_pass http://localhost:5000;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
        
        # Timeout settings for large file uploads
        proxy_connect_timeout 60s;
        proxy_send_timeout 60s;
        proxy_read_timeout 300s;
    }
}
```

Enable the site:

```bash
sudo ln -s /etc/nginx/sites-available/mlr-auto /etc/nginx/sites-enabled/
sudo nginx -t
sudo systemctl reload nginx
```

## Security Considerations

### 1. Firewall Configuration

```bash
# Install and configure UFW
sudo ufw enable
sudo ufw default deny incoming
sudo ufw default allow outgoing
sudo ufw allow ssh
sudo ufw allow 80/tcp
sudo ufw allow 443/tcp
```

### 2. SSL Certificate (Let's Encrypt)

```bash
# Install Certbot
sudo apt install certbot python3-certbot-nginx

# Get SSL certificate
sudo certbot --nginx -d theexperimentai.org -d www.theexperimentai.org

# Auto-renewal
sudo crontab -e
# Add: 0 12 * * * /usr/bin/certbot renew --quiet
```

### 3. Regular Updates

```bash
# Create update script
cat > update-app.sh << 'EOF'
#!/bin/bash
cd /path/to/your/app
git pull
docker-compose down
docker-compose up -d --build
docker system prune -f
EOF

chmod +x update-app.sh
```

## Monitoring and Maintenance

### 1. Health Checks

```bash
# Check application health
curl -f http://localhost:5000/ || echo "App is down!"

# Check container status
docker-compose ps

# View logs
docker-compose logs --tail=100 mlr-auto
```

### 2. Backup Strategy

```bash
# Backup uploads and outputs
tar -czf backup-$(date +%Y%m%d).tar.gz uploads/ outputs/

# Automated backup script
cat > backup.sh << 'EOF'
#!/bin/bash
BACKUP_DIR="/backup"
DATE=$(date +%Y%m%d_%H%M%S)
mkdir -p $BACKUP_DIR
tar -czf $BACKUP_DIR/mlr-auto-$DATE.tar.gz uploads/ outputs/
# Keep only last 7 days of backups
find $BACKUP_DIR -name "mlr-auto-*.tar.gz" -mtime +7 -delete
EOF

chmod +x backup.sh
# Add to crontab: 0 2 * * * /path/to/backup.sh
```

### 3. Log Rotation

Docker Compose is configured with log rotation, but you can also set up system-wide log rotation:

```bash
sudo nano /etc/logrotate.d/docker-containers
```

Add:

```
/var/lib/docker/containers/*/*.log {
    rotate 7
    daily
    compress
    size=1M
    missingok
    delaycompress
    copytruncate
}
```

## Troubleshooting

### Common Issues

1. **Container won't start**:
   ```bash
   docker-compose logs mlr-auto
   docker-compose down && docker-compose up -d
   ```

2. **Permission issues**:
   ```bash
   sudo chown -R 1000:1000 uploads/ outputs/
   ```

3. **Out of disk space**:
   ```bash
   docker system prune -a
   docker volume prune
   ```

4. **Memory issues**:
   ```bash
   # Check memory usage
   docker stats
   # Restart container
   docker-compose restart mlr-auto
   ```

### Performance Tuning

1. **Increase worker processes** (edit docker-compose.yml):
   ```yaml
   command: ["python", "-m", "gunicorn", "--bind", "0.0.0.0:5000", "--workers", "4", "--timeout", "120", "app:app"]
   ```

2. **Adjust memory limits** (edit docker-compose.yml):
   ```yaml
   deploy:
     resources:
       limits:
         memory: 2G
         cpus: '2.0'
   ```

## Support

For issues and support:
1. Check the application logs: `docker-compose logs mlr-auto`
2. Verify all environment variables are set correctly
3. Ensure sufficient disk space and memory
4. Check firewall and network connectivity

## Security Updates

Regularly update your system and Docker images:

```bash
# System updates
sudo apt update && sudo apt upgrade -y

# Rebuild Docker image with latest base image
docker-compose down
docker-compose build --no-cache
docker-compose up -d
```