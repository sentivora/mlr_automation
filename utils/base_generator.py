"""
Base configuration and utilities for PowerPoint presentation generation.
Contains folder mapping, formatting logic, and common utilities.
"""

import os
import logging
from PIL import Image


class BaseGenerator:
    def __init__(self):
        self.logger = logging.getLogger(__name__)
        self.folder_mapping = {
            'ott': 'OTT',
            'vdxdesktopexpandable': 'DESKTOP EXPANDABLE',
            'vdxdesktopinframe': 'DESKTOP IN-FRAME',
            'vdxinstream': 'VDX DESKTOP INSTREAM',
            'vdxdesktopinstream': 'DESKTOP INSTREAM',
            'vdxmobileexpandable': 'MOBILE EXPANDABLE',
            'vdxmobileinframe': 'MOBILE IN-FRAME',
            'vdxmobileinstream': 'MOBILE INSTREAM',
            'ctv': 'CTV'
        }
    
    def _format_folder_name(self, folder_name):
        """Format folder name according to the specified logic."""
        # Split by both forward and backward slashes
        parts = folder_name.replace('\\', '/').split('/')
        
        if len(parts) == 1:
            # Single folder name
            base_name = parts[0].lower()
            return self.folder_mapping.get(base_name, folder_name)
        elif len(parts) == 2:
            # Folder with subfolder (e.g., "vdxdesktopexpandable/300x250")
            base_name = parts[0].lower()
            sub_name = parts[1].lower()
            
            # Special cases where we don't want the subfolder name
            if base_name in ['ott', 'ctv', 'vdxinstream', 'vdxdesktopinstream', 'vdxmobileinstream'] and sub_name == '1x10':
                return self.folder_mapping.get(base_name, parts[0])
            
            formatted_base = self.folder_mapping.get(base_name, parts[0])
            return f"{formatted_base} - {sub_name}"
        else:
            # Handle nested folders by using the last two parts
            base_name = parts[-2].lower()
            sub_name = parts[-1].lower()
            
            # Special cases where we don't want the subfolder name
            if base_name in ['ott', 'ctv', 'vdxinstream', 'vdxdesktopinstream', 'vdxmobileinstream'] and sub_name == '1x10':
                return self.folder_mapping.get(base_name, parts[-2])
            
            formatted_base = self.folder_mapping.get(base_name, parts[-2])
            return f"{formatted_base} - {sub_name}"
    
    def _extract_size_from_path(self, image_path):
        """Extract size information from image path for positioning logic."""
        path_lower = image_path.lower()
        
        # Check for different size patterns in the path
        if '970x250' in path_lower:
            return '970x250'
        elif '728x90' in path_lower:
            return '728x90'
        elif '300x250' in path_lower:
            return '300x250'
        elif '300x600' in path_lower:
            return '300x600'
        elif '160x600' in path_lower:
            return '160x600'
        elif '320x50' in path_lower:
            return '320x50'
        
        return None
    
    def _validate_image_dimensions(self, image_path, min_width=1900, min_height=1092):
        """Validate if image meets minimum dimension requirements."""
        try:
            with Image.open(image_path) as img:
                width, height = img.size
                if width >= min_width and height >= min_height:
                    self.logger.info(f"Image {os.path.basename(image_path)} dimensions: {width}x{height} - INCLUDED")
                    return True
                else:
                    self.logger.info(f"Image {os.path.basename(image_path)} dimensions: {width}x{height} - SKIPPED (too small)")
                    return False
        except Exception as e:
            self.logger.error(f"Error reading image {image_path}: {str(e)}")
            return False
    
    def _remove_placeholders(self, slide):
        """Remove all placeholder shapes from a slide to prevent unwanted text."""
        for shape in slide.placeholders:
            try:
                elem = shape.element
                elem.getparent().remove(elem)
            except:
                pass  # Ignore errors during placeholder removal
    
    def sort_images(self, img_path):
        """Sort images by priority: teaser first, then mainunit, then others."""
        filename = os.path.basename(img_path).lower()
        if 'teaser' in filename:
            return 0  # Teaser images come first
        elif 'mainunit' in filename:
            return 1  # Mainunit images come second
        else:
            return 2  # Other images come last