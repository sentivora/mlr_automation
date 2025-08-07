// Feather icons initialization
feather.replace();

// File input handler
function handleFileInput() {
    const fileInput = document.getElementById('file');
    const uploadBtn = document.getElementById('uploadBtn');
    const uploadBtnText = document.getElementById('uploadBtnText');
    
    if (fileInput && fileInput.files.length > 0) {
        const file = fileInput.files[0];
        const fileName = file.name;
        const fileSize = (file.size / (1024 * 1024)).toFixed(2);
        
        // Validate file size (500MB limit)
        if (file.size > 500 * 1024 * 1024) {
            alert('File size exceeds 500MB limit. Please choose a smaller file.');
            fileInput.value = '';
            return;
        }
        
        // Update button text to show selected file
        uploadBtnText.textContent = `Upload ${fileName} (${fileSize} MB)`;
        uploadBtn.classList.add('btn-success');
        uploadBtn.classList.remove('btn-primary');
    }
}

// Form submission handler
function handleFormSubmission() {
    const form = document.getElementById('uploadForm');
    const progressContainer = document.getElementById('progressContainer');
    const progressBar = document.getElementById('progressBar');
    const uploadBtn = document.getElementById('uploadBtn');
    
    if (form) {
        form.addEventListener('submit', function(e) {
            // Show progress indicator
            progressContainer.style.display = 'block';
            uploadBtn.disabled = true;
            
            // Animate progress bar
            let progress = 0;
            const interval = setInterval(() => {
                progress += Math.random() * 15;
                if (progress > 90) {
                    progress = 90;
                    clearInterval(interval);
                }
                progressBar.style.width = progress + '%';
            }, 500);
        });
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
    
    // Refresh feather icons
    feather.replace();
});

// Refresh icons when needed
function refreshIcons() {
    setTimeout(() => {
        feather.replace();
    }, 100);
}