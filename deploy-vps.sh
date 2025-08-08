#!/bin/bash

# MLR Auto Phase-01 VPS Deployment Script
# This script automates the deployment process on Ubuntu/Debian VPS

set -e  # Exit on any error

# Colors for output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

# Configuration variables
APP_NAME="mlr-auto"
APP_DIR="/var/www/html/$APP_NAME"
DOMAIN="your-domain.com"  # Change this to your actual domain
FLASK_PORT="5000"
NGINX_AVAILABLE="/etc/nginx/sites-available"
NGINX_ENABLED="/etc/nginx/sites-enabled"

# Function to print colored output
print_status() {
    echo -e "${BLUE}[INFO]${NC} $1"
}

print_success() {
    echo -e "${GREEN}[SUCCESS]${NC} $1"
}

print_warning() {
    echo -e "${YELLOW}[WARNING]${NC} $1"
}

print_error() {
    echo -e "${RED}[ERROR]${NC} $1"
}

# Function to check if command exists
command_exists() {
    command -v "$1" >/dev/null 2>&1
}

# Function to check if service is running
service_running() {
    systemctl is-active --quiet "$1"
}

# Function to backup existing configuration
backup_config() {
    if [ -f "$NGINX_ENABLED/default" ]; then
        print_status "Backing up default nginx configuration..."
        sudo cp "$NGINX_ENABLED/default" "$NGINX_ENABLED/default.backup.$(date +%Y%m%d_%H%M%S)"
    fi
}

# Function to install dependencies
install_dependencies() {
    print_status "Updating package list..."
    sudo apt update
    
    print_status "Installing required packages..."
    sudo apt install -y nginx python3 python3-pip python3-venv curl wget unzip
    
    print_success "Dependencies installed successfully"
}

# Function to setup application directory
setup_app_directory() {
    print_status "Setting up application directory..."
    
    # Create application directory
    sudo mkdir -p "$APP_DIR"
    sudo mkdir -p "$APP_DIR/static"
    sudo mkdir -p "$APP_DIR/templates"
    sudo mkdir -p "/var/www/uploads"
    sudo mkdir -p "/var/www/output"
    sudo mkdir -p "/var/log/$APP_NAME"
    
    # Set proper permissions
    sudo chown -R www-data:www-data "$APP_DIR"
    sudo chown -R www-data:www-data "/var/www/uploads"
    sudo chown -R www-data:www-data "/var/www/output"
    sudo chown -R www-data:www-data "/var/log/$APP_NAME"
    
    sudo chmod -R 755 "$APP_DIR"
    sudo chmod -R 755 "/var/www/uploads"
    sudo chmod -R 755 "/var/www/output"
    
    print_success "Application directory setup completed"
}

# Function to setup Python environment
setup_python_env() {
    print_status "Setting up Python virtual environment..."
    
    cd "$APP_DIR"
    sudo python3 -m venv venv
    sudo chown -R www-data:www-data venv
    
    print_success "Python environment setup completed"
}

# Function to configure nginx
configure_nginx() {
    print_status "Configuring Nginx..."
    
    # Copy nginx configuration
    if [ -f "nginx-vps-site.conf" ]; then
        # Replace domain placeholder
        sed "s/your-domain.com/$DOMAIN/g" nginx-vps-site.conf > "/tmp/$APP_NAME.conf"
        sudo mv "/tmp/$APP_NAME.conf" "$NGINX_AVAILABLE/$APP_NAME"
    else
        print_error "nginx-vps-site.conf not found in current directory"
        exit 1
    fi
    
    # Backup and remove default site
    backup_config
    sudo rm -f "$NGINX_ENABLED/default"
    
    # Enable our site
    sudo ln -sf "$NGINX_AVAILABLE/$APP_NAME" "$NGINX_ENABLED/$APP_NAME"
    
    # Test nginx configuration
    if sudo nginx -t; then
        print_success "Nginx configuration is valid"
    else
        print_error "Nginx configuration test failed"
        exit 1
    fi
    
    # Reload nginx
    sudo systemctl reload nginx
    print_success "Nginx configured and reloaded"
}

# Function to setup systemd service for Flask app
setup_flask_service() {
    print_status "Setting up Flask systemd service..."
    
    cat << EOF | sudo tee /etc/systemd/system/$APP_NAME.service
[Unit]
Description=MLR Auto Phase-01 Flask Application
After=network.target

[Service]
Type=simple
User=www-data
Group=www-data
WorkingDirectory=$APP_DIR
Environment=PATH=$APP_DIR/venv/bin
Environment=FLASK_ENV=production
Environment=FLASK_DEBUG=False
ExecStart=$APP_DIR/venv/bin/python app.py
Restart=always
RestartSec=3

# Logging
StandardOutput=append:/var/log/$APP_NAME/app.log
StandardError=append:/var/log/$APP_NAME/error.log

[Install]
WantedBy=multi-user.target
EOF

    # Reload systemd and enable service
    sudo systemctl daemon-reload
    sudo systemctl enable $APP_NAME
    
    print_success "Flask systemd service created"
}

# Function to setup firewall
setup_firewall() {
    print_status "Configuring firewall..."
    
    if command_exists ufw; then
        sudo ufw allow 22/tcp   # SSH
        sudo ufw allow 80/tcp   # HTTP
        sudo ufw allow 443/tcp  # HTTPS
        sudo ufw --force enable
        print_success "Firewall configured"
    else
        print_warning "UFW not installed, skipping firewall configuration"
    fi
}

# Function to test deployment
test_deployment() {
    print_status "Testing deployment..."
    
    # Test if nginx is running
    if service_running nginx; then
        print_success "Nginx is running"
    else
        print_error "Nginx is not running"
        return 1
    fi
    
    # Test if Flask app is running
    if service_running $APP_NAME; then
        print_success "Flask application is running"
    else
        print_warning "Flask application is not running (this is expected if app files are not deployed yet)"
    fi
    
    # Test HTTP response
    if curl -s -o /dev/null -w "%{http_code}" http://localhost | grep -q "200\|404\|502"; then
        print_success "HTTP server is responding"
    else
        print_error "HTTP server is not responding"
        return 1
    fi
    
    print_success "Basic deployment test completed"
}

# Function to display next steps
show_next_steps() {
    echo
    print_success "VPS deployment setup completed!"
    echo
    echo -e "${YELLOW}Next steps:${NC}"
    echo "1. Upload your application files to: $APP_DIR"
    echo "2. Install Python dependencies:"
    echo "   cd $APP_DIR"
    echo "   sudo -u www-data venv/bin/pip install -r requirements.txt"
    echo "3. Create .env file with your configuration"
    echo "4. Start the Flask application:"
    echo "   sudo systemctl start $APP_NAME"
    echo "5. Check application status:"
    echo "   sudo systemctl status $APP_NAME"
    echo "6. View logs:"
    echo "   sudo journalctl -u $APP_NAME -f"
    echo "   tail -f /var/log/$APP_NAME/app.log"
    echo
    echo -e "${YELLOW}For SSL/HTTPS setup:${NC}"
    echo "1. Install certbot: sudo apt install certbot python3-certbot-nginx"
    echo "2. Get SSL certificate: sudo certbot --nginx -d $DOMAIN"
    echo
    echo -e "${YELLOW}Troubleshooting:${NC}"
    echo "- Check nginx logs: sudo tail -f /var/log/nginx/error.log"
    echo "- Check application logs: sudo tail -f /var/log/$APP_NAME/error.log"
    echo "- Test Flask directly: curl http://localhost:$FLASK_PORT/info"
    echo "- Test through nginx: curl http://localhost/info"
}

# Main deployment function
main() {
    echo -e "${GREEN}MLR Auto Phase-01 VPS Deployment Script${NC}"
    echo "=========================================="
    echo
    
    # Check if running as root or with sudo
    if [ "$EUID" -eq 0 ]; then
        print_error "Please run this script as a regular user with sudo privileges, not as root"
        exit 1
    fi
    
    # Check if sudo is available
    if ! command_exists sudo; then
        print_error "sudo is required but not installed"
        exit 1
    fi
    
    # Prompt for domain name
    read -p "Enter your domain name (or IP address): " user_domain
    if [ -n "$user_domain" ]; then
        DOMAIN="$user_domain"
    fi
    
    print_status "Starting deployment for domain: $DOMAIN"
    
    # Run deployment steps
    install_dependencies
    setup_app_directory
    setup_python_env
    configure_nginx
    setup_flask_service
    setup_firewall
    test_deployment
    
    show_next_steps
}

# Run main function
main "$@"