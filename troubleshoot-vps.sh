#!/bin/bash

# MLR Auto Phase-01 VPS Troubleshooting Script
# This script helps identify web server interference and diagnose HTML response issues

set -e

# Colors for output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
MAGENTA='\033[0;35m'
CYAN='\033[0;36m'
NC='\033[0m' # No Color

# Configuration
FLASK_PORT="5000"
APP_NAME="mlr-auto"
TEST_FILE="/tmp/test_upload.txt"

# Function to print colored output
print_header() {
    echo -e "\n${CYAN}=== $1 ===${NC}"
}

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

print_result() {
    echo -e "${MAGENTA}[RESULT]${NC} $1"
}

# Function to check if command exists
command_exists() {
    command -v "$1" >/dev/null 2>&1
}

# Function to create test file
create_test_file() {
    echo "This is a test file for MLR Auto upload testing." > "$TEST_FILE"
}

# Function to cleanup test file
cleanup_test_file() {
    rm -f "$TEST_FILE"
}

# Function to check system information
check_system_info() {
    print_header "SYSTEM INFORMATION"
    
    print_status "Operating System:"
    if [ -f /etc/os-release ]; then
        . /etc/os-release
        echo "  Name: $NAME"
        echo "  Version: $VERSION"
    else
        uname -a
    fi
    
    print_status "Available Memory:"
    free -h
    
    print_status "Disk Space:"
    df -h /
    
    print_status "Current User:"
    whoami
    
    print_status "Network Interfaces:"
    ip addr show | grep -E "inet |UP|DOWN" | head -10
}

# Function to identify web servers
identify_web_servers() {
    print_header "WEB SERVER IDENTIFICATION"
    
    # Check what's running on port 80
    print_status "Checking what's running on port 80:"
    if command_exists netstat; then
        netstat -tlnp | grep :80 || echo "  Nothing found on port 80"
    elif command_exists ss; then
        ss -tlnp | grep :80 || echo "  Nothing found on port 80"
    else
        print_warning "Neither netstat nor ss available"
    fi
    
    # Check what's running on port 443
    print_status "Checking what's running on port 443:"
    if command_exists netstat; then
        netstat -tlnp | grep :443 || echo "  Nothing found on port 443"
    elif command_exists ss; then
        ss -tlnp | grep :443 || echo "  Nothing found on port 443"
    else
        print_warning "Neither netstat nor ss available"
    fi
    
    # Check installed web servers
    print_status "Checking installed web servers:"
    
    if command_exists nginx; then
        print_result "Nginx is installed: $(nginx -v 2>&1)"
        print_status "Nginx status:"
        systemctl is-active nginx && echo "  Running" || echo "  Not running"
        systemctl is-enabled nginx && echo "  Enabled" || echo "  Disabled"
    else
        echo "  Nginx: Not installed"
    fi
    
    if command_exists apache2; then
        print_result "Apache2 is installed: $(apache2 -v 2>&1 | head -1)"
        print_status "Apache2 status:"
        systemctl is-active apache2 && echo "  Running" || echo "  Not running"
        systemctl is-enabled apache2 && echo "  Enabled" || echo "  Disabled"
    elif command_exists httpd; then
        print_result "Apache (httpd) is installed: $(httpd -v 2>&1 | head -1)"
        print_status "Apache status:"
        systemctl is-active httpd && echo "  Running" || echo "  Not running"
        systemctl is-enabled httpd && echo "  Enabled" || echo "  Disabled"
    else
        echo "  Apache: Not installed"
    fi
    
    if command_exists lighttpd; then
        print_result "Lighttpd is installed"
        systemctl is-active lighttpd && echo "  Running" || echo "  Not running"
    else
        echo "  Lighttpd: Not installed"
    fi
}

# Function to check Flask application
check_flask_app() {
    print_header "FLASK APPLICATION STATUS"
    
    # Check if Flask is running on expected port
    print_status "Checking Flask on port $FLASK_PORT:"
    if command_exists netstat; then
        netstat -tlnp | grep ":$FLASK_PORT" || echo "  Flask not running on port $FLASK_PORT"
    elif command_exists ss; then
        ss -tlnp | grep ":$FLASK_PORT" || echo "  Flask not running on port $FLASK_PORT"
    fi
    
    # Check systemd service
    if systemctl list-units --type=service | grep -q "$APP_NAME"; then
        print_status "MLR Auto service status:"
        systemctl status "$APP_NAME" --no-pager -l || true
    else
        print_warning "MLR Auto systemd service not found"
    fi
    
    # Check for Python processes
    print_status "Python processes running:"
    ps aux | grep python | grep -v grep || echo "  No Python processes found"
    
    # Test direct Flask connection
    print_status "Testing direct Flask connection:"
    if curl -s --connect-timeout 5 "http://localhost:$FLASK_PORT/info" >/dev/null 2>&1; then
        print_success "Flask is responding on port $FLASK_PORT"
        curl -s "http://localhost:$FLASK_PORT/info" | head -3
    else
        print_error "Flask is not responding on port $FLASK_PORT"
    fi
}

# Function to test HTTP responses
test_http_responses() {
    print_header "HTTP RESPONSE TESTING"
    
    create_test_file
    
    # Test direct Flask upload (if running)
    print_status "Testing direct Flask upload:"
    if curl -s --connect-timeout 5 "http://localhost:$FLASK_PORT/info" >/dev/null 2>&1; then
        echo "  Testing POST to Flask directly..."
        response=$(curl -s -w "\nHTTP_CODE:%{http_code}\nCONTENT_TYPE:%{content_type}" \
            -X POST \
            -H "Accept: application/json" \
            -F "file=@$TEST_FILE" \
            "http://localhost:$FLASK_PORT/upload" 2>/dev/null || echo "ERROR")
        
        if [[ "$response" == *"ERROR"* ]]; then
            print_error "Direct Flask test failed"
        else
            echo "$response" | head -5
            http_code=$(echo "$response" | grep "HTTP_CODE:" | cut -d: -f2)
            content_type=$(echo "$response" | grep "CONTENT_TYPE:" | cut -d: -f2)
            print_result "HTTP Code: $http_code, Content-Type: $content_type"
        fi
    else
        print_warning "Flask not accessible, skipping direct test"
    fi
    
    # Test through web server
    print_status "Testing through web server (port 80):"
    echo "  Testing POST to web server..."
    response=$(curl -s -w "\nHTTP_CODE:%{http_code}\nCONTENT_TYPE:%{content_type}" \
        -X POST \
        -H "Accept: application/json" \
        -F "file=@$TEST_FILE" \
        "http://localhost/upload" 2>/dev/null || echo "ERROR")
    
    if [[ "$response" == *"ERROR"* ]]; then
        print_error "Web server test failed"
    else
        echo "$response" | head -5
        http_code=$(echo "$response" | grep "HTTP_CODE:" | cut -d: -f2)
        content_type=$(echo "$response" | grep "CONTENT_TYPE:" | cut -d: -f2)
        print_result "HTTP Code: $http_code, Content-Type: $content_type"
        
        # Check if response contains HTML
        if echo "$response" | grep -q "<!DOCTYPE\|<html\|<HTML"; then
            print_error "Response contains HTML! This is the source of the SyntaxError."
            echo "  First few lines of HTML response:"
            echo "$response" | grep -E "<!DOCTYPE|<html|<title|<body" | head -3
        else
            print_success "Response does not contain HTML"
        fi
    fi
    
    cleanup_test_file
}

# Function to check nginx configuration
check_nginx_config() {
    print_header "NGINX CONFIGURATION ANALYSIS"
    
    if ! command_exists nginx; then
        print_warning "Nginx not installed, skipping configuration check"
        return
    fi
    
    print_status "Nginx configuration test:"
    if nginx -t 2>&1; then
        print_success "Nginx configuration is valid"
    else
        print_error "Nginx configuration has errors"
    fi
    
    print_status "Active nginx sites:"
    if [ -d "/etc/nginx/sites-enabled" ]; then
        ls -la /etc/nginx/sites-enabled/ || echo "  No sites enabled"
    else
        echo "  sites-enabled directory not found"
    fi
    
    print_status "Checking for MLR Auto nginx configuration:"
    if [ -f "/etc/nginx/sites-enabled/mlr-auto" ]; then
        print_success "MLR Auto nginx config found"
        echo "  Checking proxy_intercept_errors setting:"
        if grep -q "proxy_intercept_errors off" /etc/nginx/sites-enabled/mlr-auto; then
            print_success "proxy_intercept_errors is set to off (correct)"
        else
            print_error "proxy_intercept_errors not found or not set to off"
        fi
    else
        print_error "MLR Auto nginx config not found in sites-enabled"
    fi
    
    print_status "Nginx error logs (last 10 lines):"
    if [ -f "/var/log/nginx/error.log" ]; then
        tail -10 /var/log/nginx/error.log || echo "  Cannot read error log"
    else
        echo "  Error log not found"
    fi
}

# Function to check Apache configuration
check_apache_config() {
    print_header "APACHE CONFIGURATION ANALYSIS"
    
    if ! command_exists apache2 && ! command_exists httpd; then
        print_warning "Apache not installed, skipping configuration check"
        return
    fi
    
    local apache_cmd="apache2"
    if command_exists httpd; then
        apache_cmd="httpd"
    fi
    
    print_status "Apache configuration test:"
    if $apache_cmd -t 2>&1; then
        print_success "Apache configuration is valid"
    else
        print_error "Apache configuration has errors"
    fi
    
    print_status "Active Apache sites:"
    if [ -d "/etc/apache2/sites-enabled" ]; then
        ls -la /etc/apache2/sites-enabled/ || echo "  No sites enabled"
    elif [ -d "/etc/httpd/conf.d" ]; then
        ls -la /etc/httpd/conf.d/ || echo "  No configurations found"
    fi
    
    print_status "Apache error logs (last 10 lines):"
    if [ -f "/var/log/apache2/error.log" ]; then
        tail -10 /var/log/apache2/error.log || echo "  Cannot read error log"
    elif [ -f "/var/log/httpd/error_log" ]; then
        tail -10 /var/log/httpd/error_log || echo "  Cannot read error log"
    else
        echo "  Error log not found"
    fi
}

# Function to check firewall and SELinux
check_security() {
    print_header "SECURITY CONFIGURATION"
    
    # Check firewall
    print_status "Firewall status:"
    if command_exists ufw; then
        ufw status || echo "  UFW status unknown"
    elif command_exists firewall-cmd; then
        firewall-cmd --state || echo "  Firewalld not running"
        firewall-cmd --list-all 2>/dev/null || echo "  Cannot list firewall rules"
    else
        echo "  No known firewall found"
    fi
    
    # Check SELinux
    print_status "SELinux status:"
    if command_exists getenforce; then
        getenforce
        if [ "$(getenforce)" = "Enforcing" ]; then
            print_warning "SELinux is enforcing - this might block HTTP connections"
            echo "  To allow HTTP connections: sudo setsebool -P httpd_can_network_connect 1"
        fi
    else
        echo "  SELinux not found (likely not RHEL/CentOS)"
    fi
}

# Function to provide recommendations
provide_recommendations() {
    print_header "RECOMMENDATIONS"
    
    echo -e "${YELLOW}Based on the analysis above, here are the recommended actions:${NC}\n"
    
    echo "1. ${CYAN}If Flask is not running:${NC}"
    echo "   - Check application logs: sudo journalctl -u mlr-auto -f"
    echo "   - Start the service: sudo systemctl start mlr-auto"
    echo "   - Check for Python errors in the application"
    echo
    
    echo "2. ${CYAN}If you're getting HTML responses instead of JSON:${NC}"
    echo "   - Apply the nginx configuration: sudo cp nginx-vps-site.conf /etc/nginx/sites-available/mlr-auto"
    echo "   - Enable the site: sudo ln -sf /etc/nginx/sites-available/mlr-auto /etc/nginx/sites-enabled/"
    echo "   - Remove default site: sudo rm -f /etc/nginx/sites-enabled/default"
    echo "   - Test config: sudo nginx -t"
    echo "   - Reload nginx: sudo systemctl reload nginx"
    echo
    
    echo "3. ${CYAN}If using Apache instead of Nginx:${NC}"
    echo "   - Disable Apache error pages for API routes"
    echo "   - Set ProxyErrorOverride Off for API locations"
    echo "   - Ensure mod_proxy and mod_proxy_http are enabled"
    echo
    
    echo "4. ${CYAN}If the problem persists:${NC}"
    echo "   - Try direct Flask deployment without web server proxy"
    echo "   - Use Gunicorn: pip install gunicorn && gunicorn -b 0.0.0.0:80 app:app"
    echo "   - Check if there's a load balancer or CDN in front of your server"
    echo
    
    echo "5. ${CYAN}For debugging:${NC}"
    echo "   - Monitor logs in real-time: sudo tail -f /var/log/nginx/error.log"
    echo "   - Test with curl: curl -v -X POST -H 'Accept: application/json' -F 'file=@test.txt' http://your-domain/upload"
    echo "   - Check browser developer tools Network tab for actual response content"
}

# Main function
main() {
    echo -e "${GREEN}MLR Auto Phase-01 VPS Troubleshooting Script${NC}"
    echo "============================================="
    echo -e "${YELLOW}This script will analyze your VPS configuration to identify why HTML is returned instead of JSON.${NC}\n"
    
    # Check if running with appropriate permissions
    if [ "$EUID" -eq 0 ]; then
        print_warning "Running as root. Some checks might not work as expected."
    fi
    
    # Run all checks
    check_system_info
    identify_web_servers
    check_flask_app
    test_http_responses
    check_nginx_config
    check_apache_config
    check_security
    provide_recommendations
    
    echo
    print_success "Troubleshooting analysis completed!"
    echo -e "${YELLOW}Please review the output above and follow the recommendations.${NC}"
    echo -e "${YELLOW}If you need further assistance, please share this output with your system administrator.${NC}"
}

# Trap to cleanup on exit
trap cleanup_test_file EXIT

# Run main function
main "$@"