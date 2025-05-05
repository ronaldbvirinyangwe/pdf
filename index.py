import sys
import os
import time

# Print setup information
print("Setting up dependencies...")

# Try importing the required libraries with better error handling
try:
    import pytesseract
    print("✓ Pytesseract successfully imported")
except ImportError:
    print("✗ Error importing pytesseract")
    print("\nFIX: Try running: pip install pytesseract")
    print("Also make sure Tesseract OCR is installed on your system:")
    print("- Mac: brew install tesseract")
    print("- Windows: Download from https://github.com/UB-Mannheim/tesseract/wiki")
    print("- Linux: sudo apt-get install tesseract-ocr")
    sys.exit(1)

try:
    from pdf2image import convert_from_path
    print("✓ pdf2image successfully imported")
except ImportError:
    print("✗ Error importing pdf2image")
    print("\nFIX: Try running: pip install pdf2image")
    print("For Windows users, also install poppler and add it to PATH")
    sys.exit(1)

try:
    import fitz  # PyMuPDF
    print("✓ PyMuPDF (fitz) successfully imported")
except ImportError:
    print("✗ Error importing PyMuPDF")
    print("\nFIX: Try running: pip install PyMuPDF")
    sys.exit(1)

from PIL import Image, ImageEnhance

print("\nAll dependencies loaded successfully!")

def pdf_to_images(pdf_path, dpi=300):
    """
    Convert PDF pages to images with robust error handling
    """
    images = []
    print(f"Converting PDF to images at {dpi} DPI...")
    
    try:
        # Handle platform-specific poppler path
        if os.name == 'nt':  # Windows
            try:
                # Try with several common Poppler installation paths
                poppler_paths = [
                    r'C:\Program Files\poppler-23.11.0\Library\bin',
                    r'C:\Program Files\poppler-xx.xx.0\Library\bin',
                    r'C:\poppler\bin'
                ]
                
                for path in poppler_paths:
                    if os.path.exists(path):
                        print(f"Found Poppler at: {path}")
                        pages = convert_from_path(pdf_path, dpi=dpi, poppler_path=path)
                        break
                else:
                    print("Poppler path not found, trying without explicit path...")
                    pages = convert_from_path(pdf_path, dpi=dpi)
            except Exception as e:
                print(f"Warning: {e}")
                print("Trying without explicit poppler path...")
                pages = convert_from_path(pdf_path, dpi=dpi)
        else:  # Linux/Mac
            pages = convert_from_path(pdf_path, dpi=dpi)
        
        print(f"Successfully converted {len(pages)} pages")
        
        for page in pages:
            images.append(page)
            
    except Exception as e:
        print(f"Error converting PDF to images: {e}")
        print("Please ensure the PDF file exists and is not corrupted.")
        sys.exit(1)
        
    return images

def preprocess_image(img):
    """
    Simple image preprocessing for OCR optimization
    """
    print("  Preprocessing image...", end="", flush=True)
    
    # Convert to grayscale
    img = img.convert('L')
    print(".", end="", flush=True)
    
    # Enhance contrast
    enhancer = ImageEnhance.Contrast(img)
    img = enhancer.enhance(2.0)
    print(".", end="", flush=True)
    
    # Sharpen
    enhancer = ImageEnhance.Sharpness(img)
    img = enhancer.enhance(2.0)
    print(".", end="", flush=True)
    
    # Enhance brightness
    enhancer = ImageEnhance.Brightness(img)
    img = enhancer.enhance(1.1)
    print(" Done!")
    
    return img

def extract_pdf_text(pdf_path, lang='eng', use_ocr_always=False):
    """
    Extract text from both text-based and image-based PDFs
    
    Parameters:
        pdf_path (str): Path to the PDF file
        lang (str): Language for OCR (default: 'eng')
        use_ocr_always (bool): Force OCR even if text extraction succeeds
    
    Returns:
        str: Extracted text from the PDF
    """
    text = ""
    
    # Try text-based extraction first (unless OCR is forced)
    if not use_ocr_always:
        try:
            print("\nAttempting text-based extraction...")
            with fitz.open(pdf_path) as doc:
                total_pages = len(doc)
                print(f"PDF has {total_pages} pages")
                
                for page_num, page in enumerate(doc):
                    print(f"  Processing page {page_num + 1}/{total_pages}", end="", flush=True)
                    page_text = page.get_text()
                    text += f"--- Page {page_num + 1} ---\n{page_text}\n\n"
                    print(" ✓")
                    
            # If we got meaningful text, return it
            if len(text.strip()) > 50:  # More than just page numbers/headers
                print("\nText extraction successful!")
                return text
            
            print("\nText extraction yielded minimal results (likely a scanned PDF)")
            print("Falling back to OCR...")
        except Exception as e:
            print(f"\nText extraction failed: {e}")
            print("Falling back to OCR...")
    else:
        print("\nForced OCR mode enabled, skipping text extraction...")
    
    # If text extraction failed or was skipped, use OCR
    print(f"\nStarting OCR processing with language: {lang}")
    
    # Check if tesseract is properly installed and accessible
    try:
        pytesseract.get_tesseract_version()
        print(f"  Tesseract version: {pytesseract.get_tesseract_version()}")
    except pytesseract.TesseractNotFoundError:
        print("\nERROR: Tesseract is not installed or not in PATH")
        print("Please install Tesseract OCR:")
        print("- Mac: brew install tesseract")
        print("- Windows: Download from https://github.com/UB-Mannheim/tesseract/wiki")
        print("- Linux: sudo apt-get install tesseract-ocr")
        sys.exit(1)
    
    # Convert PDF to images
    images = pdf_to_images(pdf_path)
    
    # Process each page
    for i, img in enumerate(images):
        print(f"\nProcessing page {i+1}/{len(images)}")
        start_time = time.time()
        
        # Preprocess image to improve OCR quality
        processed = preprocess_image(img)
        
        # OCR configuration for better results
        print("  Performing OCR...", end="", flush=True)
        config = f'--oem 1 --psm 3 -l {lang}'
        try:
            page_text = pytesseract.image_to_string(processed, config=config)
            text += f"--- Page {i + 1} ---\n{page_text}\n\n"
            elapsed = time.time() - start_time
            print(f" Done! ({elapsed:.2f} seconds)")
        except Exception as e:
            print(f" Failed! Error: {e}")
    
    return text

def analyze_document(pdf_path, output_txt=None, lang='eng', use_ocr_always=False):
    """
    Full document analysis workflow
    
    Parameters:
        pdf_path (str): Path to the PDF file
        output_txt (str): Path to save the extracted text (optional)
        lang (str): Language for OCR
        use_ocr_always (bool): Force OCR even if text extraction succeeds
    
    Returns:
        str: Extracted text from the PDF
    """
    print(f"\n{'='*60}")
    print(f"PDF TEXT EXTRACTION: {os.path.basename(pdf_path)}")
    print(f"{'='*60}")
    
    # Check if file exists
    if not os.path.exists(pdf_path):
        print(f"ERROR: File not found: {pdf_path}")
        sys.exit(1)
    
    # Extract text from PDF
    start_time = time.time()
    text = extract_pdf_text(pdf_path, lang=lang, use_ocr_always=use_ocr_always)
    total_time = time.time() - start_time
    
    # Save text to file if requested
    if output_txt:
        try:
            with open(output_txt, 'w', encoding='utf-8') as f:
                f.write(text)
            print(f"\nSaved extracted text to: {output_txt}")
        except Exception as e:
            print(f"\nError saving text to file: {e}")
    
    # Print summary
    print(f"\n{'-'*60}")
    print(f"EXTRACTION COMPLETE! Total time: {total_time:.2f} seconds")
    print(f"Total extracted text length: {len(text)} characters")
    print(f"{'-'*60}")
    
    return text

# Example usage
if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("\nUsage: python simple_pdf_ocr.py ventures_English_Gr_1.pdf ventures_English_Gr_1.txt [language] [force_ocr]")
        print("\nExample:")
        print("python simple_pdf_ocr.py document.pdf output.txt eng+fra True")
        print("\nParameters:")
        print("- pdf_file: Path to your PDF file (required)")
        print("- output_file: Path to save the text (optional)")
        print("- language: Language code for OCR, default is 'eng' (optional)")
        print("- force_ocr: Set to 'True' to force OCR even for text PDFs (optional)")
        sys.exit(1)
    
    # Parse command line arguments
    pdf_path = sys.argv[1]
    output_txt = sys.argv[2] if len(sys.argv) > 2 and sys.argv[2] != "None" else None
    lang = sys.argv[3] if len(sys.argv) > 3 else 'eng'
    use_ocr_always = sys.argv[4].lower() == 'true' if len(sys.argv) > 4 else False
    
    # Process the document
    text = analyze_document(pdf_path, output_txt, lang, use_ocr_always)
    
    # Print text preview
    print("\nExtracted Text Preview:")
    print("=" * 60)
    preview_length = min(1000, len(text))
    print(text[:preview_length])
    if len(text) > preview_length:
        print("...")
    print("=" * 60)