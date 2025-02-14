import logging
import os
import re

def setup_logging():
    """Configure logging settings"""
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.StreamHandler()
        ]
    )

def validate_filename(filename):
    """
    Validate filename against SharePoint restrictions
    """
    if not filename:
        return False
    
    # Check length
    if len(filename) > 128:
        return False
    
    # Check invalid characters
    invalid_chars = r'[<>:"/\\|?*]'
    if re.search(invalid_chars, filename):
        return False
    
    # Check if filename starts or ends with space or period
    if filename.startswith((' ', '.')) or filename.endswith((' ', '.')):
        return False
    
    return True

def sanitize_filename(filename):
    """
    Sanitize filename to meet SharePoint requirements
    """
    # Get file extension
    name, ext = os.path.splitext(filename)
    
    # Remove invalid characters
    name = re.sub(r'[<>:"/\\|?*]', '', name)
    
    # Remove leading/trailing spaces and periods
    name = name.strip(' .')
    
    # Truncate name if too long (accounting for extension)
    max_length = 128 - len(ext)
    if len(name) > max_length:
        name = name[:max_length]
    
    # Combine name and extension
    return f"{name}{ext}"
