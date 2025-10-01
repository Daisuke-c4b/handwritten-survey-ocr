import os
from pathlib import Path

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
