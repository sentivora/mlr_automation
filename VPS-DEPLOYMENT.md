# VPS Deployment Guide for MLR Auto Phase-01

This guide addresses the common issue where VPS deployments return HTML error pages instead of JSON responses, causing `SyntaxError: Unexpected token '<'` in the frontend.

## Problem Description

The error occurs when:
- Frontend expects JSON responses from Flask API
- Web server (Nginx/Apache) intercepts requests and returns HTML error pages
- JavaScript tries to parse HTML as JSON, causing syntax errors

## Solution 1: Nginx Configuration (Recommended)

### Step 1: Apply Nginx Configuration

1. Copy the `nginx-api-config.conf` to your VPS:
```bash
scp nginx-api-config.conf user@your-vps:/etc/nginx/sites-available/mlr-auto
```

2. Create symbolic link to enable the site:
```bash
sudo ln -s /etc/nginx/sites-available/mlr-auto /etc/nginx/sites-enabled/
```

3. Remove default nginx configuration:
```bash
sudo rm /etc/nginx/sites-enabled/default
```

4. Test nginx configuration:
```bash
sudo nginx -t
```

5. Reload nginx:
```bash
sudo systemctl reload nginx
```

### Step 2: Verify Flask Application

Ensure your Flask app is running on the correct port:
```bash
# Check if Flask is running on port 5000
sudo netstat -tlnp | grep :5000

# If not running, start Flask app
cd /path/to/your/app
python app.py
```

## Solution 2: Apache Configuration

If using Apache instead of Nginx:

### Create Apache Virtual Host

Create `/etc/apache2/sites-available/mlr-auto.conf`:
```apache
<VirtualHost *:80>
    ServerName theexperimentai.org
    DocumentRoot /var/www/html/mlr-auto
    
    # Proxy Flask API requests
    ProxyPreserveHost On
    ProxyPass /upload http://127.0.0.1:5000/upload
    ProxyPassReverse /upload http://127.0.0.1:5000/upload
    ProxyPass /api/ http://127.0.0.1:5000/
    ProxyPassReverse /api/ http://127.0.0.1:5000/
    
    # Disable error documents for API routes
    <LocationMatch "^/(upload|api|convert-to-pdf|download)">
        ProxyErrorOverride Off
        ErrorDocument 400 "Bad Request"
        ErrorDocument 403 "Forbidden"
        ErrorDocument 404 "Not Found"
        ErrorDocument 500 "Internal Server Error"
    </LocationMatch>
    
    # Serve static files directly
    Alias /static /var/www/html/mlr-auto/static
    <Directory "/var/www/html/mlr-auto/static">
        Require all granted
    </Directory>
</VirtualHost>
```

Enable the site:
```bash
sudo a2ensite mlr-auto
sudo a2dissite 000-default
sudo systemctl reload apache2
```

## Solution 3: Direct Flask Deployment (No Web Server Proxy)

### Using Gunicorn

1. Install Gunicorn:
```bash
pip install gunicorn
```

2. Create Gunicorn configuration (`gunicorn.conf.py`):
```python
bind = "0.0.0.0:80"
workers = 4
worker_class = "sync"
worker_connections = 1000
timeout = 30
keepalive = 2
max_requests = 1000
max_requests_jitter = 100
preload_app = True
```

3. Run with Gunicorn:
```bash
sudo gunicorn -c gunicorn.conf.py app:app
```

### Using uWSGI

1. Install uWSGI:
```bash
pip install uwsgi
```

2. Create uWSGI configuration (`uwsgi.ini`):
```ini
[uwsgi]
module = app:app
master = true
processes = 4
socket = 0.0.0.0:80
protocol = http
die-on-term = true
vacuum = true
```

3. Run with uWSGI:
```bash
sudo uwsgi --ini uwsgi.ini
```

## Solution 4: Environment Variables for VPS

Create `.env` file on VPS:
```bash
# Flask Configuration
FLASK_ENV=production
FLASK_DEBUG=False
SECRET_KEY=your-production-secret-key

# File Upload Configuration
MAX_CONTENT_LENGTH=16777216
UPLOAD_FOLDER=/var/www/uploads
OUTPUT_FOLDER=/var/www/output

# Blob Storage (if using)
BLOB_READ_WRITE_TOKEN=your-blob-token

# CORS Configuration
CORS_ORIGINS=https://theexperimentai.org

# Logging
LOG_LEVEL=INFO
LOG_FILE=/var/log/mlr-auto/app.log
```

## Troubleshooting Steps

### Step 1: Identify Web Server

```bash
# Check what's running on port 80
sudo netstat -tlnp | grep :80

# Check nginx status
sudo systemctl status nginx

# Check apache status
sudo systemctl status apache2

# Check what web server is installed
which nginx
which apache2
```

### Step 2: Test Direct Flask Response

```bash
# Test Flask directly (bypass web server)
curl -X POST http://localhost:5000/upload \
  -H "Content-Type: multipart/form-data" \
  -H "Accept: application/json" \
  -F "file=@test.txt"
```

### Step 3: Test Through Web Server

```bash
# Test through web server
curl -X POST http://theexperimentai.org/upload \
  -H "Content-Type: multipart/form-data" \
  -H "Accept: application/json" \
  -F "file=@test.txt"
```

### Step 4: Check Error Logs

```bash
# Nginx error logs
sudo tail -f /var/log/nginx/error.log

# Apache error logs
sudo tail -f /var/log/apache2/error.log

# Flask application logs
tail -f /var/log/mlr-auto/app.log

# System logs
sudo journalctl -u nginx -f
sudo journalctl -u apache2 -f
```

### Step 5: Verify Content-Type Headers

```bash
# Check response headers
curl -I http://theexperimentai.org/upload

# Verbose curl to see full request/response
curl -v -X POST http://theexperimentai.org/upload \
  -H "Accept: application/json" \
  -F "file=@test.txt"
```

## Common Issues and Solutions

### Issue 1: 502 Bad Gateway
- **Cause**: Flask app not running or wrong port
- **Solution**: Start Flask app on correct port (5000)

### Issue 2: 403 Forbidden
- **Cause**: File permissions or SELinux
- **Solution**: 
  ```bash
  sudo chown -R www-data:www-data /var/www/html/mlr-auto
  sudo chmod -R 755 /var/www/html/mlr-auto
  sudo setsebool -P httpd_can_network_connect 1  # If SELinux enabled
  ```

### Issue 3: Still Getting HTML Responses
- **Cause**: Web server error pages not disabled
- **Solution**: Add `ProxyErrorOverride Off` (Apache) or `proxy_intercept_errors off` (Nginx)

### Issue 4: CORS Errors
- **Cause**: Missing CORS headers
- **Solution**: Ensure Flask-CORS is configured or add headers in web server config

## Testing Deployment

1. **Test File Upload**:
   ```bash
   curl -X POST http://theexperimentai.org/upload \
     -H "Accept: application/json" \
     -F "file=@sample.txt"
   ```

2. **Test Error Handling**:
   ```bash
   # Test with invalid file
   curl -X POST http://theexperimentai.org/upload \
     -H "Accept: application/json" \
     -F "file=@invalid.xyz"
   ```

3. **Test Frontend**:
   - Open browser developer tools
   - Go to Network tab
   - Upload a file through the web interface
   - Verify response Content-Type is `application/json`

## Security Considerations

1. **Firewall Configuration**:
   ```bash
   sudo ufw allow 80/tcp
   sudo ufw allow 443/tcp
   sudo ufw deny 5000/tcp  # Block direct Flask access
   ```

2. **SSL/HTTPS Setup**:
   ```bash
   sudo apt install certbot python3-certbot-nginx
   sudo certbot --nginx -d theexperimentai.org
   ```

3. **File Upload Security**:
   - Ensure upload directories are outside web root
   - Set proper file permissions
   - Validate file types and sizes

## Monitoring and Maintenance

1. **Log Rotation**:
   ```bash
   # Add to /etc/logrotate.d/mlr-auto
   /var/log/mlr-auto/*.log {
       daily
       missingok
       rotate 52
       compress
       delaycompress
       notifempty
       create 644 www-data www-data
   }
   ```

2. **Health Check Script**:
   ```bash
   #!/bin/bash
   # health-check.sh
   response=$(curl -s -o /dev/null -w "%{http_code}" http://localhost/upload)
   if [ $response -ne 405 ]; then  # 405 = Method Not Allowed (expected for GET)
       echo "Service is down"
       exit 1
   fi
   echo "Service is healthy"
   ```

This comprehensive guide should resolve the HTML response issues and provide a robust deployment strategy for your VPS.