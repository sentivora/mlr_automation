"""
Main presentation assembly and orchestration.
Handles slide ordering, creation flow, and final presentation generation.
"""

import os
import tempfile
from datetime import datetime
from pptx import Presentation
from pptx.util import Inches

from .base_generator import BaseGenerator
from .image_processor import ImageProcessor
from .slide_creator import SlideCreator


class PresentationAssembler(BaseGenerator):
    def __init__(self):
        super().__init__()
        self.image_processor = ImageProcessor(self.logger)
        self.slide_creator = SlideCreator(self, self.image_processor)
    
    def create_presentation(self, folder_structure, output_dir, annotation_option='with_annos', is_multi_tab=False, implement_video_frames=False, video_position_params=None, original_filename=None):
        """
        Create a PowerPoint presentation from folder structure.
        
        Args:
            folder_structure (dict): Dictionary with folder names as keys and image paths as values
            output_dir (str): Directory to save the presentation
            annotation_option (str): Either 'with_annos' or 'no_annos' to control annotation display
            is_multi_tab (bool): Whether this is a Multi-tab request with additional slide logic
            implement_video_frames (bool): Whether to implement video frames for all units
            
        Returns:
            tuple: (str, int, bool) Path to the created presentation file, slide count, and video folder found
        """
        # Import the original implementation temporarily to maintain functionality
        from .presentation_generator import PresentationGenerator
        
        # Create instance of original generator and delegate
        original_generator = PresentationGenerator()
        return original_generator.create_presentation(folder_structure, output_dir, annotation_option, is_multi_tab, implement_video_frames, video_position_params, original_filename)