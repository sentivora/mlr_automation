#!/bin/bash

# MLR Auto Phase-01 VPS Deployment Script (Gunicorn + Flask + Nginx)
# Run as a non-root sudo user!

set -e  # Exit on any error

# Colors
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

# Config vars
APP_NAME="mlr-auto"
APP_DIR="/var/www/html/$APP_NAME"
DOMAIN="theexperimentai.org"
FLASK_PORT="5000"
NGINX_AVAILABLE="/etc/nginx/sites-available"
NGINX_ENABLED="/etc/nginx/sites-enabled"
PY_USER="$(whoami)"

print_status()   { echo -e "${BLUE}[INFO]${NC} $1"; }
print_success()  { echo -e "${GREEN}[SUCCESS]${NC} $1"; }
print_warning()  { echo -e "${YELLOW}[WARNING]${NC} $1"; }
print_error()    { echo -e "${RED}[ERROR]${NC} $1"; }

command_exists() { command -v "$1" >/dev/null 2>&1; }
service_running(){ systemctl is-active --quiet "$1"; }

backup_config() {
    if [ -f "$NGINX_ENABLED/default" ]; then
        print_status "Backing up default nginx configuration..."
        sudo cp "$NGINX_ENABLED/default" "$NGINX_ENABLED/default.backup.$(date +%Y%m%d_%H%M%S)"
    fi
}

install_dependencies() {
    print_status "Updating packages..."
    sudo apt update
    print_status "Installing required packages..."
    sudo apt install -y nginx python3 python3-pip python3-venv curl git unzip
    print_success "Dependencies installed."
}

setup_app_directory() {
    print_status "Setting up app directory..."
    sudo mkdir -p "$APP_DIR/static" "$APP_DIR/templates" "/var/www/uploads" "/var/www/output" "/var/log/$APP_NAME"
    sudo chown -R $PY_USER:www-data "$APP_DIR" "/var/www/uploads" "/var/www/output" "/var/log/$APP_NAME"
    sudo chmod -R 755 "$APP_DIR" "/var/www/uploads" "/var/www/output"
    print_success "App directory ready."
}

setup_python_env() {
    print_status "Setting up Python venv..."
    cd "$APP_DIR"
    python3 -m venv venv
    print_success "Python venv created."
}

configure_nginx() {
    print_status "Configuring Nginx..."
    # Generate Nginx config (overwrite any old one)
    sudo tee "$NGINX_AVAILABLE/$APP_NAME" > /dev/null <<EOF
server {
    listen 80;
    server_name $DOMAIN www.$DOMAIN;
    location / {
        proxy_pass http://127.0.0.1:$FLASK_PORT;
        proxy_set_header Host \$host;
        proxy_set_header X-Real-IP \$remote_addr;
        proxy_set_header X-Forwarded-For \$proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto \$scheme;
    }
    client_max_body_size 200M;
}
EOF
    backup_config
    sudo rm -f "$NGINX_ENABLED/default"
    sudo ln -sf "$NGINX_AVAILABLE/$APP_NAME" "$NGINX_ENABLED/$APP_NAME"
    sudo nginx -t && print_success "Nginx config valid." || { print_error "Nginx config invalid."; exit 1; }
    sudo systemctl reload nginx
    print_success "Nginx reloaded."
}

setup_gunicorn_service() {
    print_status "Setting up Gunicorn systemd service..."
    sudo tee /etc/systemd/system/$APP_NAME.service > /dev/null <<EOF
[Unit]
Description=MLR Auto Flask App (Gunicorn)
After=network.target

[Service]
User=$PY_USER
Group=www-data
WorkingDirectory=$APP_DIR
Environment="PATH=$APP_DIR/venv/bin"
ExecStart=$APP_DIR/venv/bin/gunicorn -w 4 -b 127.0.0.1:$FLASK_PORT app:app
Restart=always
RestartSec=3

[Install]
WantedBy=multi-user.target
EOF

    sudo systemctl daemon-reload
    sudo systemctl enable $APP_NAME
    print_success "Gunicorn service created."
}

setup_firewall() {
    print_status "Configuring firewall..."
    if command_exists ufw; then
        sudo ufw allow 22/tcp   # SSH
        sudo ufw allow 80/tcp   # HTTP
        sudo ufw allow 443/tcp  # HTTPS
        sudo ufw --force enable
        print_success "Firewall configured."
    else
        print_warning "UFW not installed, skipping firewall config."
    fi
}

test_deployment() {
    print_status "Testing deployment..."
    service_running nginx && print_success "Nginx running." || print_error "Nginx not running."
    service_running $APP_NAME && print_success "Gunicorn app running." || print_warning "Gunicorn app not running (expected if app not yet uploaded)."
    if curl -s -o /dev/null -w "%{http_code}" http://localhost | grep -q "200\|404\|502"; then
        print_success "HTTP server is responding."
    else
        print_error "HTTP server not responding."
    fi
    print_success "Basic deployment test done."
}

show_next_steps() {
    echo
    print_success "VPS deployment setup completed!"
    echo
    echo -e "${YELLOW}Next steps:${NC}"
    echo "1. Upload your Flask app files to: $APP_DIR"
    echo "2. Install Python dependencies:"
    echo "   cd $APP_DIR"
    echo "   source venv/bin/activate"
    echo "   pip install -r requirements.txt"
    echo "3. Create your .env file if needed"
    echo "4. Start the Flask application:"
    echo "   sudo systemctl start $APP_NAME"
    echo "5. Check app status:"
    echo "   sudo systemctl status $APP_NAME"
    echo "6. Logs (journalctl):"
    echo "   sudo journalctl -u $APP_NAME -f"
    echo
    echo -e "${YELLOW}SSL Setup:${NC}"
    echo "   sudo apt install certbot python3-certbot-nginx"
    echo "   sudo certbot --nginx -d $DOMAIN -d www.$DOMAIN"
    echo
    echo -e "${YELLOW}Troubleshooting:${NC}"
    echo "   sudo tail -f /var/log/nginx/error.log"
    echo "   sudo journalctl -u $APP_NAME -f"
}

main() {
    echo -e "${GREEN}MLR Auto Phase-01 VPS Deploy Script${NC}"
    echo "=========================================="
    echo

    if [ "$EUID" -eq 0 ]; then
        print_error "Please run as a sudo user, not root."
        exit 1
    fi

    if ! command_exists sudo; then
        print_error "sudo is required but not installed"
        exit 1
    fi

    read -p "Enter your domain name (default: $DOMAIN): " user_domain
    if [ -n "$user_domain" ]; then DOMAIN="$user_domain"; fi

    print_status "Starting deployment for: $DOMAIN"
    install_dependencies
    setup_app_directory
    setup_python_env
    configure_nginx
    setup_gunicorn_service
    setup_firewall
    test_deployment
    show_next_steps
}

main "$@"
