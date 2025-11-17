import os
from azure.ai.documentintelligence import DocumentIntelligenceClient
from azure.ai.documentintelligence.models import AnalyzeDocumentRequest
from azure.core.credentials import AzureKeyCredential
from dotenv import load_dotenv

load_dotenv()

# Create client once
document_client = DocumentIntelligenceClient(
    endpoint=os.getenv("AZURE_DI_ENDPOINT"),
    credential=AzureKeyCredential(os.getenv("AZURE_DI_KEY"))
)

def extract_text_from_pdf(file_bytes: bytes) -> str:
    """
    Sends a PDF file's bytes to Azure Document Intelligence
    and returns extracted text.
    """

    poller = document_client.begin_analyze_document(
        model_id="prebuilt-read",
        analyze_request=file_bytes,                # ✔️ correct for your version
        content_type="application/pdf"             # ✔️ must be specified
    )

    result = poller.result()

    all_text = []

    for page in result.pages:
        if page.lines:
            for line in page.lines:
                all_text.append(line.content)

    return "\n".join(all_text)
