import os
import logging
import tempfile
from typing import Optional, Dict, Any
from werkzeug.utils import secure_filename

logger = logging.getLogger(__name__)

class UnifiedStorage:
    """Simplified storage system that only uses local file storage for VPS deployment."""
    
    def __init__(self):
        self.local_upload_dir = os.path.join(os.getcwd(), 'uploads')
        self.local_output_dir = os.path.join(os.getcwd(), 'outputs')
        
        # Ensure local directories exist
        os.makedirs(self.local_upload_dir, exist_ok=True)
        os.makedirs(self.local_output_dir, exist_ok=True)
        
        logger.info("UnifiedStorage initialized - Using local file storage only")
        
    def get_storage_info(self) -> Dict[str, Any]:
        """Get information about the current storage configuration."""
        return {
            'environment': 'local',
            'local_upload_dir': self.local_upload_dir,
            'local_output_dir': self.local_output_dir
        }
    
    async def upload_file(self, file_data: bytes, filename: str, content_type: str = 'application/octet-stream') -> Optional[Dict[str, Any]]:
        """Upload file to local filesystem."""
        filename = secure_filename(filename)
        
        try:
            local_path = os.path.join(self.local_upload_dir, filename)
            with open(local_path, 'wb') as f:
                f.write(file_data)
            
            logger.info(f"File saved locally: {local_path}")
            return {
                'url': f'/local-file/{filename}',
                'local_path': local_path,
                'filename': filename,
                'size': len(file_data),
                'storage_type': 'local'
            }
        except Exception as e:
            logger.error(f"Error saving file locally: {str(e)}")
            return None
    
    async def download_file(self, file_identifier: str) -> Optional[bytes]:
        """Download file from local filesystem."""
        # Check if it's a local file path or filename
        if file_identifier.startswith('/local-file/'):
            filename = file_identifier.replace('/local-file/', '')
            local_path = os.path.join(self.local_upload_dir, filename)
        elif file_identifier.startswith('/'):
            # Absolute path
            local_path = file_identifier
        else:
            # Assume it's a filename in uploads directory
            local_path = os.path.join(self.local_upload_dir, file_identifier)
        
        try:
            if os.path.exists(local_path):
                with open(local_path, 'rb') as f:
                    data = f.read()
                logger.info(f"File downloaded locally: {local_path} ({len(data)} bytes)")
                return data
            else:
                logger.error(f"Local file not found: {local_path}")
                return None
        except Exception as e:
            logger.error(f"Error reading local file: {str(e)}")
            return None
    
    async def delete_file(self, file_identifier: str) -> bool:
        """Delete file from local filesystem."""
        if file_identifier.startswith('/local-file/'):
            filename = file_identifier.replace('/local-file/', '')
            local_path = os.path.join(self.local_upload_dir, filename)
        else:
            local_path = os.path.join(self.local_upload_dir, file_identifier)
        
        try:
            if os.path.exists(local_path):
                os.remove(local_path)
                logger.info(f"Local file deleted: {local_path}")
                return True
            else:
                logger.warning(f"Local file not found for deletion: {local_path}")
                return False
        except Exception as e:
            logger.error(f"Error deleting local file: {str(e)}")
            return False
    
    def save_output_file(self, file_data: bytes, filename: str) -> Optional[str]:
        """Save output file (like generated PPTX) to local outputs directory."""
        filename = secure_filename(filename)
        
        try:
            local_path = os.path.join(self.local_output_dir, filename)
            with open(local_path, 'wb') as f:
                f.write(file_data)
            
            logger.info(f"Output file saved locally: {local_path}")
            return local_path
        except Exception as e:
            logger.error(f"Error saving output file locally: {str(e)}")
            return None
    
    def get_output_file_path(self, filename: str) -> Optional[str]:
        """Get the path to an output file."""
        filename = secure_filename(filename)
        local_path = os.path.join(self.local_output_dir, filename)
        return local_path if os.path.exists(local_path) else None
    
    async def get_file_info(self, file_identifier: str) -> Optional[Dict[str, Any]]:
        """Get file information from local storage."""
        # Get local file info - check both upload and output directories
        if file_identifier.startswith('/local-file/'):
            filename = file_identifier.replace('/local-file/', '')
            upload_path = os.path.join(self.local_upload_dir, filename)
            output_path = os.path.join(self.local_output_dir, filename)
        else:
            upload_path = os.path.join(self.local_upload_dir, file_identifier)
            output_path = os.path.join(self.local_output_dir, file_identifier)
        
        try:
            # Check upload directory first
            if os.path.exists(upload_path):
                stat = os.stat(upload_path)
                return {
                    'url': f'/local-file/{os.path.basename(upload_path)}',
                    'local_path': upload_path,
                    'size': stat.st_size,
                    'modified_time': stat.st_mtime,
                    'storage_type': 'local'
                }
            # Check output directory
            elif os.path.exists(output_path):
                stat = os.stat(output_path)
                return {
                    'url': f'/local-file/{os.path.basename(output_path)}',
                    'local_path': output_path,
                    'size': stat.st_size,
                    'modified_time': stat.st_mtime,
                    'storage_type': 'local'
                }
            else:
                return None
        except Exception as e:
            logger.error(f"Error getting local file info: {str(e)}")
            return None

# Global instance
unified_storage = None

def initialize_unified_storage():
    """Initialize unified storage system."""
    global unified_storage
    unified_storage = UnifiedStorage()
    return unified_storage