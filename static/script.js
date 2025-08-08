// Feather icons initialization
feather.replace();

// Show file size error with user-friendly message
function showFileSizeError(actualSize) {
    const errorMessage = `File size (${actualSize} MB) exceeds the 200 MB limit for VPS Storage. Please choose a smaller file or compress your archive.`;
    
    // Create or update error alert
    let existingAlert = document.querySelector('.file-size-error');
    if (existingAlert) {
        existingAlert.remove();
    }
    
    const alertDiv = document.createElement('div');
    alertDiv.className = 'alert alert-danger alert-dismissible fade show file-size-error';
    alertDiv.innerHTML = `
        <i data-feather="alert-circle" class="me-2"></i>
        ${errorMessage}
        <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
    `;
    
    // Insert after the form
    const form = document.getElementById('uploadForm');
    form.parentNode.insertBefore(alertDiv, form.nextSibling);
    
    // Refresh feather icons
    feather.replace();
}

// Show current file size info
function showFileSizeInfo(fileName, fileSize) {
    const maxSize = 200; // Updated to 200MB for VPS Storage
    const sizePercentage = (fileSize / maxSize) * 100;
    
    // Remove existing info
    let existingInfo = document.querySelector('.file-size-info');
    if (existingInfo) {
        existingInfo.remove();
    }
    
    const infoDiv = document.createElement('div');
    infoDiv.className = 'alert alert-info file-size-info mt-3';
    
    const progressColor = sizePercentage > 80 ? 'bg-warning' : 'bg-success';
    
    infoDiv.innerHTML = `
        <div class="d-flex justify-content-between align-items-center mb-2">
            <span><i data-feather="upload-cloud" class="me-2"></i><strong>${fileName}</strong></span>
            <span class="badge ${sizePercentage > 80 ? 'bg-warning' : 'bg-success'}">${fileSize} MB / 200 MB</span>
        </div>
        <div class="progress" style="height: 8px;">
            <div class="progress-bar ${progressColor}" role="progressbar" style="width: ${Math.min(sizePercentage, 100)}%"></div>
        </div>
        <small class="text-muted"><i data-feather="info" class="me-1"></i>Files are stored securely in VPS Storage</small>
    `;
    
    // Insert after the file input
    const fileInputContainer = document.querySelector('.mb-4');
    fileInputContainer.appendChild(infoDiv);
    
    // Refresh feather icons
    feather.replace();
}

// File input handler
function handleFileInput() {
    const fileInput = document.getElementById('file');
    const uploadBtn = document.getElementById('uploadBtn');
    const uploadBtnText = document.getElementById('uploadBtnText');
    
    if (fileInput && fileInput.files.length > 0) {
        const file = fileInput.files[0];
        const fileName = file.name;
        const fileSize = (file.size / (1024 * 1024)).toFixed(2);
        
        // Validate file size (200MB limit for VPS Storage)
        const maxSizeBytes = 200 * 1024 * 1024; // 200MB in bytes
        if (file.size > maxSizeBytes) {
            showFileSizeError(fileSize);
            fileInput.value = '';
            return;
        }
        
        // Show file size info
        showFileSizeInfo(fileName, fileSize);
        
        // Update button text to show selected file
        uploadBtnText.textContent = `Upload ${fileName} (${fileSize} MB)`;
        uploadBtn.classList.add('btn-success');
        uploadBtn.classList.remove('btn-primary');
    }
}

// Form submission handler with direct file upload
function handleFormSubmission() {
    const form = document.getElementById('uploadForm');
    const progressContainer = document.getElementById('progressContainer');
    const progressBar = document.getElementById('progressBar');
    const uploadBtn = document.getElementById('uploadBtn');
    const fileInput = document.getElementById('file');
    
    if (form) {
        form.addEventListener('submit', async function(e) {
            e.preventDefault(); // Prevent default form submission
            
            if (!fileInput.files.length) {
                showError('Please select a file to upload.');
                return;
            }
            
            const file = fileInput.files[0];
            
            try {
                // Show progress indicator
                progressContainer.style.display = 'block';
                uploadBtn.disabled = true;
                uploadBtn.innerHTML = '<i data-feather="upload-cloud" class="me-2"></i>Uploading...';
                feather.replace();
                
                // Create FormData for direct file upload
                const formData = new FormData();
                formData.append('file', file);
                
                // Get annotation option from form
                const annotationOption = document.querySelector('input[name="annotation_option"]:checked');
                if (annotationOption) {
                    formData.append('annotation_option', annotationOption.value);
                }
                
                progressBar.style.width = '20%';
                
                // Upload file directly to server with explicit JSON headers
                const uploadResponse = await fetch('/upload', {
                    method: 'POST',
                    headers: {
                        'Accept': 'application/json'
                        // Note: Don't set Content-Type for FormData, let browser set it with boundary
                    },
                    body: formData
                });
                
                progressBar.style.width = '60%';
                
                if (!uploadResponse.ok) {
                    let errorMessage = 'Upload failed';
                    try {
                        // Check if response is HTML before trying to parse as JSON
                        const contentType = uploadResponse.headers.get('content-type');
                        if (contentType && contentType.includes('text/html')) {
                            const htmlText = await uploadResponse.text();
                            // Extract meaningful error from HTML if possible
                            const titleMatch = htmlText.match(/<title>(.*?)<\/title>/i);
                            errorMessage = titleMatch ? titleMatch[1] : `Server returned HTML error (${uploadResponse.status})`;
                        } else {
                            const error = await uploadResponse.json();
                            errorMessage = error.error || error.message || 'Upload failed';
                        }
                    } catch (parseError) {
                        // If response parsing fails completely, use status text
                        errorMessage = uploadResponse.statusText || `HTTP ${uploadResponse.status} error`;
                    }
                    throw new Error(errorMessage);
                }

                // Validate response is JSON before parsing
                const contentType = uploadResponse.headers.get('content-type');
                if (!contentType || !contentType.includes('application/json')) {
                    const responseText = await uploadResponse.text();
                    if (responseText.trim().startsWith('<!DOCTYPE') || responseText.trim().startsWith('<html')) {
                        throw new Error('Server returned HTML instead of JSON. Please check server configuration.');
                    }
                    throw new Error('Server returned non-JSON response: ' + contentType);
                }

                const result = await uploadResponse.json();
                progressBar.style.width = '100%';
                
                // Check if processing was successful
                if (result.success) {
                    // Show success message
                    showSuccess(`File uploaded and processed successfully! ${result.message || ''}`);
                    
                    // If we have a result URL, redirect after a delay
                    if (result.result_url) {
                        setTimeout(() => {
                            window.location.href = result.result_url;
                        }, 1500);
                    } else {
                        // Reset form
                        form.reset();
                        handleFileInput(); // Update UI
                    }
                } else {
                    // Show error message
                    throw new Error(result.message || 'Processing failed');
                }
                
            } catch (error) {
                console.error('Upload error:', error);
                showError(`Upload failed: ${error.message}`);
            } finally {
                // Reset UI
                setTimeout(() => {
                    progressContainer.style.display = 'none';
                    progressBar.style.width = '0%';
                    uploadBtn.disabled = false;
                    uploadBtn.innerHTML = '<i data-feather="upload-cloud" class="me-2"></i><span id="uploadBtnText">Upload and Convert</span>';
                    feather.replace();
                }, 2000);
            }
        });
    }
}

// Helper functions for showing messages
function showError(message) {
    showAlert(message, 'danger', 'alert-circle');
}

function showSuccess(message) {
    showAlert(message, 'success', 'check-circle');
}

function showAlert(message, type, icon) {
    // Remove existing alerts
    const existingAlerts = document.querySelectorAll('.upload-alert');
    existingAlerts.forEach(alert => alert.remove());
    
    // Format multi-line messages for better display
    const formattedMessage = message.replace(/\n/g, '<br>');
    
    const alertDiv = document.createElement('div');
    alertDiv.className = `alert alert-${type} alert-dismissible fade show upload-alert`;
    alertDiv.innerHTML = `
        <i data-feather="${icon}" class="me-2"></i>
        <div style="white-space: pre-line;">${formattedMessage}</div>
        <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
    `;
    
    // Insert after the form
    const form = document.getElementById('uploadForm');
    form.parentNode.insertBefore(alertDiv, form.nextSibling);
    
    // Refresh feather icons
    feather.replace();
    
    // Auto-remove success messages after 5 seconds
    if (type === 'success') {
        setTimeout(() => {
            if (alertDiv.parentNode) {
                alertDiv.remove();
            }
        }, 5000);
    }
}

// Drag and drop functionality
function setupDragAndDrop() {
    const fileInput = document.getElementById('file');
    const formContainer = document.querySelector('.card-body');
    
    if (fileInput && formContainer) {
        ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
            formContainer.addEventListener(eventName, preventDefaults, false);
        });
        
        function preventDefaults(e) {
            e.preventDefault();
            e.stopPropagation();
        }
        
        ['dragenter', 'dragover'].forEach(eventName => {
            formContainer.addEventListener(eventName, highlight, false);
        });
        
        ['dragleave', 'drop'].forEach(eventName => {
            formContainer.addEventListener(eventName, unhighlight, false);
        });
        
        function highlight(e) {
            formContainer.classList.add('border-primary');
        }
        
        function unhighlight(e) {
            formContainer.classList.remove('border-primary');
        }
        
        formContainer.addEventListener('drop', handleDrop, false);
        
        function handleDrop(e) {
            const dt = e.dataTransfer;
            const files = dt.files;
            
            if (files.length > 0) {
                fileInput.files = files;
                handleFileInput();
            }
        }
    }
}

// Theme toggle functionality
function initializeThemeToggle() {
    const themeToggle = document.getElementById('themeToggle');
    const themeIcon = document.getElementById('themeIcon');
    const htmlElement = document.documentElement;
    
    // Get saved theme or default to light
    const savedTheme = localStorage.getItem('theme') || 'light';
    htmlElement.setAttribute('data-bs-theme', savedTheme);
    updateThemeIcon(savedTheme);
    
    if (themeToggle) {
        themeToggle.addEventListener('click', function() {
            const currentTheme = htmlElement.getAttribute('data-bs-theme');
            const newTheme = currentTheme === 'dark' ? 'light' : 'dark';
            
            htmlElement.setAttribute('data-bs-theme', newTheme);
            localStorage.setItem('theme', newTheme);
            updateThemeIcon(newTheme);
            
            // Refresh feather icons after theme change
            setTimeout(() => {
                feather.replace();
            }, 50);
        });
    }
    
    function updateThemeIcon(theme) {
        if (themeIcon) {
            // Remove existing icon attribute
            themeIcon.removeAttribute('data-feather');
            
            // Set new icon based on theme
            if (theme === 'dark') {
                themeIcon.setAttribute('data-feather', 'moon');
                themeToggle.setAttribute('title', 'Switch to light theme');
            } else {
                themeIcon.setAttribute('data-feather', 'sun');
                themeToggle.setAttribute('title', 'Switch to dark theme');
            }
            
            // Refresh the specific icon
            feather.replace();
        }
    }
}

// Initialize all functionality when DOM is loaded
document.addEventListener('DOMContentLoaded', function() {
    // File input change listener
    const fileInput = document.getElementById('file');
    if (fileInput) {
        fileInput.addEventListener('change', handleFileInput);
    }
    
    // Form submission handler
    handleFormSubmission();
    
    // Drag and drop setup
    setupDragAndDrop();
    
    // Initialize theme toggle
    initializeThemeToggle();
    
    // Refresh feather icons
    feather.replace();
});

// Refresh icons when needed
function refreshIcons() {
    setTimeout(() => {
        feather.replace();
    }, 100);
}