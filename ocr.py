"""
OCR module for extracting text from PDF files using Azure Document Intelligence
"""
import os
from io import BytesIO
from azure.core.credentials import AzureKeyCredential
from azure.ai.formrecognizer import DocumentAnalysisClient

def extract_text_from_pdf(file_bytes: bytes) -> str:
    """
    Extract text from a PDF file using Azure Document Intelligence (Form Recognizer).
    
    Args:
        file_bytes: PDF file as bytes
        
    Returns:
        Extracted text as string
        
    Raises:
        ValueError: If Azure credentials are not configured
        Exception: If OCR processing fails
    """
    # Get Azure Document Intelligence credentials
    endpoint = os.getenv("AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT")
    api_key = os.getenv("AZURE_DOCUMENT_INTELLIGENCE_KEY")
    
    if not endpoint or not api_key:
        raise ValueError(
            "Missing Azure Document Intelligence credentials. "
            "Please set AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT and "
            "AZURE_DOCUMENT_INTELLIGENCE_KEY in your .env file"
        )
    
    try:
        # Initialize the Document Analysis client
        document_analysis_client = DocumentAnalysisClient(
            endpoint=endpoint,
            credential=AzureKeyCredential(api_key)
        )
        
        # Convert bytes to BytesIO for the API
        pdf_stream = BytesIO(file_bytes)
        
        # Analyze the document using the prebuilt-read model
        poller = document_analysis_client.begin_analyze_document(
            "prebuilt-read", 
            document=pdf_stream
        )
        
        result = poller.result()
        
        # Extract all text content
        extracted_text = ""
        for page in result.pages:
            for line in page.lines:
                extracted_text += line.content + "\n"
        
        return extracted_text.strip()
    
    except Exception as e:
        raise Exception(f"OCR extraction failed: {str(e)}")


def extract_text_from_pdf_fallback(file_bytes: bytes) -> str:
    """
    Fallback method using PyPDF2 if Azure Document Intelligence is not available.
    This is less accurate but doesn't require Azure credentials.
    
    Args:
        file_bytes: PDF file as bytes
        
    Returns:
        Extracted text as string
    """
    try:
        import PyPDF2
        from io import BytesIO
        
        pdf_reader = PyPDF2.PdfReader(BytesIO(file_bytes))
        text = ""
        
        for page in pdf_reader.pages:
            text += page.extract_text() + "\n"
        
        return text.strip()
    
    except ImportError:
        raise ImportError(
            "PyPDF2 is not installed. Install it with: pip install PyPDF2"
        )
    except Exception as e:
        raise Exception(f"Fallback PDF extraction failed: {str(e)}")