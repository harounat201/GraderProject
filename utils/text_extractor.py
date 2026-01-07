import os

def extract_text_from_file(file_path):
    """
    Extracts text content from .txt or .docx files.
    Returns: String content.
    """
    ext = os.path.splitext(file_path)[1].lower()
    
    if ext == '.txt':
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                return f.read()
        except UnicodeDecodeError:
            # Fallback for other encodings
            with open(file_path, 'r', encoding='latin-1') as f:
                return f.read()
                
    elif ext == '.docx':
        try:
            import docx
            doc = docx.Document(file_path)
            full_text = []
            for para in doc.paragraphs:
                full_text.append(para.text)
            return '\n'.join(full_text)
        except Exception as e:
            return f"Error reading .docx file: {str(e)}"
            
    return None
