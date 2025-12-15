import os
from pathlib import Path

# Supported image extensions
SUPPORTED_IMAGE_EXTENSIONS = {'.png', '.jpg', '.jpeg', '.gif', '.webp', '.bmp', '.tiff', '.tif'}
SUPPORTED_PDF_EXTENSIONS = {'.pdf'}
SUPPORTED_EXTENSIONS = SUPPORTED_IMAGE_EXTENSIONS | SUPPORTED_PDF_EXTENSIONS

def validate_pdf(uploaded_file):
    """Validate if uploaded file is a valid PDF"""
    try:
        # Check file extension
        if not uploaded_file.name.lower().endswith('.pdf'):
            return False
        
        # Check file size (max 50MB)
        if uploaded_file.size > 50 * 1024 * 1024:
            return False
        
        # Check PDF magic bytes
        uploaded_file.seek(0)
        header = uploaded_file.read(4)
        uploaded_file.seek(0)  # Reset file pointer
        
        if header != b'%PDF':
            return False
        
        return True
        
    except Exception:
        return False

def validate_image(uploaded_file):
    """Validate if uploaded file is a valid image"""
    try:
        # Check file extension
        ext = Path(uploaded_file.name).suffix.lower()
        if ext not in SUPPORTED_IMAGE_EXTENSIONS:
            return False
        
        # Check file size (max 50MB)
        if uploaded_file.size > 50 * 1024 * 1024:
            return False
        
        # Check image magic bytes
        uploaded_file.seek(0)
        header = uploaded_file.read(16)
        uploaded_file.seek(0)  # Reset file pointer
        
        # PNG magic bytes
        if header[:8] == b'\x89PNG\r\n\x1a\n':
            return True
        
        # JPEG magic bytes
        if header[:2] == b'\xff\xd8':
            return True
        
        # GIF magic bytes
        if header[:6] in (b'GIF87a', b'GIF89a'):
            return True
        
        # WebP magic bytes
        if header[:4] == b'RIFF' and header[8:12] == b'WEBP':
            return True
        
        # BMP magic bytes
        if header[:2] == b'BM':
            return True
        
        # TIFF magic bytes (little-endian and big-endian)
        if header[:4] in (b'II*\x00', b'MM\x00*'):
            return True
        
        return False
        
    except Exception:
        return False

def validate_file(uploaded_file):
    """Validate if uploaded file is a valid PDF or image"""
    ext = Path(uploaded_file.name).suffix.lower()
    
    if ext in SUPPORTED_PDF_EXTENSIONS:
        return validate_pdf(uploaded_file)
    elif ext in SUPPORTED_IMAGE_EXTENSIONS:
        return validate_image(uploaded_file)
    
    return False

def get_file_type(filename):
    """Get file type based on extension"""
    ext = Path(filename).suffix.lower()
    
    if ext in SUPPORTED_PDF_EXTENSIONS:
        return 'pdf'
    elif ext in SUPPORTED_IMAGE_EXTENSIONS:
        return 'image'
    
    return None

def get_supported_extensions_list():
    """Get list of all supported extensions for file uploader"""
    return list(SUPPORTED_EXTENSIONS)

def extract_filename(full_filename):
    """Extract filename without extension"""
    try:
        return Path(full_filename).stem
    except Exception:
        return "transcription"

def format_file_size(size_bytes):
    """Format file size in human readable format"""
    if size_bytes == 0:
        return "0 B"
    
    size_names = ["B", "KB", "MB", "GB"]
    i = 0
    while size_bytes >= 1024 and i < len(size_names) - 1:
        size_bytes /= 1024.0
        i += 1
    
    return f"{size_bytes:.1f} {size_names[i]}"

def clean_text_for_filename(text):
    """Clean text to be safe for use in filenames"""
    # Remove or replace invalid filename characters
    invalid_chars = '<>:"/\\|?*'
    for char in invalid_chars:
        text = text.replace(char, '_')
    
    # Limit length
    if len(text) > 100:
        text = text[:100]
    
    return text.strip()
