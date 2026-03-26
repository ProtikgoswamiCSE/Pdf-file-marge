import zipfile
import xml.etree.ElementTree as ET
import io
from fpdf import FPDF

def docx_to_text(docx_path):
    """Extracts text from a .docx file."""
    try:
        with zipfile.ZipFile(docx_path) as z:
            xml_content = z.read('word/document.xml')
            tree = ET.fromstring(xml_content)
            
            # Namespaces
            ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
            
            paragraphs = []
            for p in tree.findall('.//w:p', ns):
                texts = [node.text for node in p.findall('.//w:t', ns) if node.text]
                if texts:
                    paragraphs.append("".join(texts))
            
            return "\n".join(paragraphs)
    except Exception as e:
        return f"Error extracting text from DOCX: {str(e)}"

def text_to_pdf_buffer(text):
    """Converts plain text to a PDF in a BytesIO buffer using fpdf2."""
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("helvetica", size=12)
    
    # Simple multi-cell for text wrapping
    # We use encode-decode to handle potential latin-1 issues in default fonts
    clean_text = text.encode('latin-1', 'replace').decode('latin-1')
    pdf.multi_cell(0, 10, clean_text)
    
    buf = io.BytesIO()
    pdf_bytes = pdf.output()
    if isinstance(pdf_bytes, bytearray) or isinstance(pdf_bytes, bytes):
        buf.write(pdf_bytes)
    else:
        # Some versions might return string or something else, but output() 
        # usually returns bytes in newer fpdf2 versions when no dest is provided
        buf.write(pdf_bytes)
    
    buf.seek(0)
    return buf

def image_to_pdf_buffer(image_path):
    """Converts an image to a PDF in a BytesIO buffer using fpdf2."""
    pdf = FPDF()
    pdf.add_page()
    # fpdf2.image() can take a file path
    # We'll try to use it directly. Note: without Pillow, 
    # fpdf2 might only support JPEG/PNG if it can parse them natively.
    pdf.image(image_path, x=10, y=10, w=190)
    
    buf = io.BytesIO()
    pdf_bytes = pdf.output()
    buf.write(pdf_bytes)
    buf.seek(0)
    return buf
