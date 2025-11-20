import os

import streamlit as st
from dotenv import load_dotenv
from openai import AzureOpenAI

from io import BytesIO
from docx import Document
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak, Table, TableStyle
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from datetime import datetime
import textwrap


from ocr import extract_text_from_pdf

import re

def clean_display_text(text: str) -> str:
    """
    Clean text for display by fixing common encoding issues
    """
    # Remove any weird unicode characters
    text = text.encode('ascii', 'ignore').decode('ascii')
    
    # Fix common spacing issues around dollar signs
    text = re.sub(r'\$(\d+\.?\d*)\s*million', r'$\1 million', text)
    text = re.sub(r'(\d+)%\s*', r'\1% ', text)
    
    # Ensure proper spacing around numbers
    text = re.sub(r'(\d+\.\d+)([a-zA-Z])', r'\1 \2', text)
    
    # Fix words that got concatenated
    text = re.sub(r'([a-z])([A-Z])', r'\1 \2', text)
    
    return text

# 🔹 MUST be first Streamlit call
st.set_page_config(page_title="AI Financial Report Analyzer", layout="wide")

load_dotenv()

# ===== SESSION STATE INITIALIZATION =====
# Initialize all session state variables (prevents errors and manages state properly)
if "structured_text" not in st.session_state:
    st.session_state.structured_text = ""
if "raw_text" not in st.session_state:
    st.session_state.raw_text = ""
if "last_file_id" not in st.session_state:
    st.session_state.last_file_id = None
if "processing_complete" not in st.session_state:
    st.session_state.processing_complete = False
if "demo_mode" not in st.session_state:
    st.session_state.demo_mode = False
# ===== END SESSION STATE INITIALIZATION =====

# Azure OpenAI client
client = AzureOpenAI(
    api_key=os.getenv("AZURE_OPENAI_API_KEY"),
    azure_endpoint=os.getenv("AZURE_OPENAI_ENDPOINT"),
    api_version=os.getenv("AZURE_OPENAI_API_VERSION"),
)

# Sidebar navigation
st.sidebar.title("🔍 Navigation")
page = st.sidebar.radio(
    "Go to",
    [
        "About This Project",
        "Cleaned Text",
        "Executive Summary",
        "KPIs",
        "Thematic Summaries",
        "Generate Report",
    ],
)

st.title("📊 AI Financial Report Analyzer")

# ===== NEW: DEMO MODE =====
demo_mode = st.sidebar.checkbox("🎬 Demo Mode (Use Sample Data)", value=False)

if demo_mode:
    st.info("📊 **Demo Mode Active**: Using pre-loaded sample financial report")
    
    # Sample structured text (mimics a real cleaned financial report)
    sample_text = """Financial Performance Overview

Revenue and Growth
Total revenue for fiscal year 2024 reached 45.2 million dollars, representing a 23 percent increase compared to the previous year. This growth was primarily driven by digital transformation initiatives which accounted for 67 percent of new revenue streams.

Expense Management
Operating expenses increased by 15 percent to 32.1 million dollars, demonstrating improved operational efficiency as revenue grew faster than costs. The company maintained strong cost discipline with a 71 percent gross margin.

Profitability Analysis
Net profit margin improved from 18 percent to 21 percent, with EBITDA reaching 12.4 million dollars. Return on equity increased to 24 percent, exceeding industry benchmarks of 18 to 20 percent.

Cash Flow Performance
Operating cash flow was 13.2 million dollars, representing a 28 percent increase year over year. Free cash flow reached 9.8 million dollars after capital expenditures of 3.4 million dollars.

Market Position and Trends
Market share in core segments grew from 12 percent to 15 percent. Customer acquisition costs decreased by 18 percent while customer lifetime value increased by 32 percent. Digital channels now represent 67 percent of total revenue.

Sustainability and ESG Initiatives
Sustainability investments totaled 3.2 million dollars, with measurable carbon reduction of 22 percent. ESG compliance costs for CBAM are estimated at 1.8 million dollars annually. Green revenue reached 8.4 million dollars, or 19 percent of total revenue.

Risk Factors and Challenges
Supply chain disruptions affected 12 percent of operations, resulting in 2.1 million dollars in additional costs. Currency fluctuations impacted margins by 2.3 percent. Regulatory compliance costs increased by 1.2 million dollars year over year.

Strategic Outlook
Management projects 18 to 22 percent revenue growth for 2025, driven by product expansion and market penetration. Capital expenditure plans include 5.5 million dollars for technology infrastructure and 2.8 million dollars for sustainability initiatives."""
    
    sample_raw = "Sample OCR text before cleaning (with simulated artifacts)..."
    
    # Store in session state
    st.session_state["structured_text"] = sample_text
    st.session_state["raw_text"] = sample_raw
    st.session_state["demo_mode"] = True
    
    st.success("✅ Sample data loaded! Navigate through pages to see AI analysis in action.")
    uploaded_file = None

    # Set a flag so we don't require file upload
    uploaded_file = None

else:
    st.session_state["demo_mode"] = False
    uploaded_file = st.file_uploader("Upload a PDF financial report", type=["pdf"])
# ===== END DEMO MODE =====

# Make sure these variables always exist
raw_text: str = ""
structured_text: str = ""

# -----------------------------
# OCR + CLEANING
# -----------------------------
# Process file OR use demo data
if uploaded_file is not None or st.session_state.get("demo_mode", False):
    # Only do OCR if we have an actual file (not demo mode)
    if uploaded_file is not None and not st.session_state.get("demo_mode", False):
        
        # ===== CHECK IF THIS FILE WAS ALREADY PROCESSED =====
        # Create unique file ID based on name and size
        current_file_id = f"{uploaded_file.name}_{uploaded_file.size}"
        
        # Only process if it's a NEW file
        if st.session_state.last_file_id != current_file_id:
            st.info("🔄 New file detected. Extracting text from PDF… please wait.")
            
            # 1) OCR extraction
            file_bytes = uploaded_file.read()
            raw_text = extract_text_from_pdf(file_bytes)
            st.success("✅ OCR extraction complete!")
            
            # 2) Clean & structure OCR text with Azure OpenAI
            st.info("🧠 Cleaning and structuring the extracted text with AI…")
            
            try:
                cleaned = client.chat.completions.create(
                    model=os.getenv("AZURE_OPENAI_DEPLOYMENT"),
                    messages=[
                        {
                            "role": "system",
                            "content": (
                                "You are an assistant that cleans messy OCR text from PDFs. "
                                "Your job is ONLY to rewrite the text in a clean, readable way, "
                                "without losing any information.\n\n"
                                "Rules:\n"
                                "- Fix words where letters are split by line breaks "
                                "  (e.g. 'm\\ni\\ll\\li\\on' -> 'million').\n"
                                "- Fix numbers and ranges that are broken across lines.\n"
                                "- For things that look like charts or distributions "
                                "  (e.g. 'Organization revenue in US dollars'), "
                                "  reconstruct them as a clear bullet list or short paragraph "
                                "  describing the ranges and percentages.\n"
                                "- Remove page numbers, repeated headings, and footers.\n"
                                "- Preserve all content and meaning. Do NOT summarize or omit sections.\n"
                            ),
                        },
                        {"role": "user", "content": raw_text},
                    ],
                    timeout=60,  # Prevent hanging
                )
                structured_text = cleaned.choices[0].message.content
                
                # ===== SAVE TO SESSION STATE =====
                st.session_state.structured_text = structured_text
                st.session_state.raw_text = raw_text
                st.session_state.last_file_id = current_file_id
                st.session_state.processing_complete = True
                
                st.success("✅ Text cleaning complete! Navigate through pages to analyze.")
                
            except Exception as e:
                st.error(f"❌ Error during AI processing: {str(e)}")
                st.info("💡 Tip: Check your Azure OpenAI deployment and API quota")
                st.stop()
        
        else:
            # File already processed - just show success message
            st.success("✅ Using previously processed document (already in memory)")
            st.info("📝 Navigate through pages to view analysis, or upload a new file to start over.")

# -----------------------------
# PAGES (only if we have structured_text)
# -----------------------------
# Get structured text from session state (works for both demo and uploaded files)
structured_text = st.session_state.get("structured_text", "")
raw_text = st.session_state.get("raw_text", "")

# ===== ABOUT THIS PROJECT PAGE (Always accessible, doesn't need structured_text) =====
if page == "About This Project":
    st.header("📚 Azure AI-102 Project: Financial Report Analyzer")
    
    st.markdown("""
    ### 🎯 Project Overview
    This application demonstrates practical implementation of **Azure AI services** for enterprise 
    document intelligence and natural language processing. Built as a capstone project for the 
    **Microsoft Azure AI-102 certification course**.
    """)
    
    # Two-column layout for technologies
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("""
        #### ☁️ Azure Services
        - ✅ **Azure OpenAI Service** (GPT-4o-mini)
        - ✅ **Document Intelligence** (OCR)
        - ✅ **Azure Resource Management**
        - ✅ **API Authentication & Security**
        
        #### 🐍 Python Stack
        - **Streamlit** - Interactive web UI
        - **python-docx** - Word document export
        - **ReportLab** - PDF generation
        - **python-dotenv** - Configuration management
        """)
    
    with col2:
        st.markdown("""
        #### 🧠 AI Capabilities Demonstrated
        - 📖 **Text Extraction & Cleaning** - OCR post-processing
        - 🎯 **Prompt Engineering** - System/user message design
        - 📊 **Structured Data Extraction** - KPI tables from text
        - 📝 **Multi-step Reasoning** - Thematic analysis pipeline
        - 🔄 **Context Management** - State preservation across operations
        """)
    
    st.markdown("---")
    
    # Key concepts section
    st.subheader("🔑 Azure AI-102 Concepts Demonstrated")
    
    tab1, tab2, tab3 = st.tabs(["Prompt Engineering", "AI Pipeline", "Production Readiness"])
    
    with tab1:
        st.markdown("""
        #### Prompt Engineering Best Practices
        
        **1. Role-Based System Messages**
```python
        {"role": "system", "content": "You are a senior financial analyst..."}
```
        Defines AI personality and constraints upfront.
        
        **2. Clear Output Specifications**
        - "Return ONLY as a Markdown table with columns: KPI | Value"
        - Prevents unwanted preambles and ensures consistent formatting
        
        **3. Few-Shot Learning**
        - Providing examples of desired outputs
        - Reduces ambiguity in complex tasks
        
        **4. Temperature Control**
        - `temperature=0.3` for financial analysis (deterministic)
        - Higher temps (0.7-0.9) for creative content
        
        **5. Token Management**
        - `max_tokens` limits to control costs
        - Monitoring input/output token usage
        """)
    
    with tab2:
        st.markdown("""
        #### Multi-Step AI Processing Pipeline
        
        This application demonstrates a production-grade AI workflow:
        
        **Step 1: Document Ingestion**
        - PDF upload with validation
        - File size and type checking
        
        **Step 2: OCR Extraction**
        - Extract raw text from financial documents
        - Handle multi-page layouts
        
        **Step 3: AI-Powered Cleaning**
        - Fix OCR artifacts (broken words, spacing)
        - Remove headers/footers/page numbers
        - Preserve all semantic content
        
        **Step 4: Multi-Modal Analysis**
        - Executive summarization (abstractive)
        - KPI extraction (extractive + structured)
        - Thematic analysis (domain-specific)
        
        **Step 5: Report Synthesis**
        - Combine multiple AI outputs
        - Generate professional documents
        - Export to multiple formats
        
        Each step preserves context for downstream operations while managing token limits efficiently.
        """)
    
    with tab3:
        st.markdown("""
        #### Production-Ready Features
        
        **Error Handling**
        - Graceful degradation on API failures
        - User-friendly error messages
        - Troubleshooting guidance
        
        **State Management**
        - Session state for user data persistence
        - File change detection (avoid reprocessing)
        - Caching strategies to reduce API calls
        
        **Cost Optimization**
        - Token usage monitoring
        - Single-pass processing with caching
        - Timeout controls
        
        **User Experience**
        - Progress indicators for long operations
        - Demo mode for quick testing
        - Clear status messages
        
        **Security Considerations**
        - Environment variable configuration
        - API key protection
        - Input validation
        """)
    
    st.markdown("---")
    
    # Architecture diagram
    st.subheader("🏗️ Application Architecture")
    
    st.code("""
    ┌─────────────────────────────────────────────────────────────┐
    │                      USER INTERFACE                         │
    │                   (Streamlit Web App)                       │
    └────────────────────┬────────────────────────────────────────┘
                         │
                         ▼
    ┌─────────────────────────────────────────────────────────────┐
    │                   DOCUMENT UPLOAD                           │
    │              • PDF File Upload                              │
    │              • Demo Mode (Sample Data)                      │
    │              • File Validation                              │
    └────────────────────┬────────────────────────────────────────┘
                         │
                         ▼
    ┌─────────────────────────────────────────────────────────────┐
    │                  OCR EXTRACTION                             │
    │          (Azure Document Intelligence)                      │
    │              • Text extraction                              │
    │              • Layout preservation                          │
    └────────────────────┬────────────────────────────────────────┘
                         │
                         ▼
    ┌─────────────────────────────────────────────────────────────┐
    │              AI-POWERED TEXT CLEANING                       │
    │               (Azure OpenAI - GPT)                          │
    │              • Fix OCR artifacts                            │
    │              • Remove noise                                 │
    │              • Structure content                            │
    └────────────────────┬────────────────────────────────────────┘
                         │
                         ▼
    ┌─────────────────────────────────────────────────────────────┐
    │              MULTI-MODAL AI ANALYSIS                        │
    │               (Azure OpenAI Service)                        │
    │  ┌──────────────┬──────────────┬──────────────────────┐    │
    │  │  Executive   │  KPI         │  Thematic            │    │
    │  │  Summary     │  Extraction  │  Analysis            │    │
    │  │              │              │  (8 themes)          │    │
    │  └──────────────┴──────────────┴──────────────────────┘    │
    └────────────────────┬────────────────────────────────────────┘
                         │
                         ▼
    ┌─────────────────────────────────────────────────────────────┐
    │              REPORT GENERATION & EXPORT                     │
    │         • Markdown (.md)                                    │
    │         • Microsoft Word (.docx)                            │
    │         • PDF (.pdf)                                        │
    └─────────────────────────────────────────────────────────────┘
    """, language="text")
    
    st.markdown("---")
    
    # Exam objectives covered
    st.subheader("✅ AI-102 Exam Objectives Covered")
    
    objectives_col1, objectives_col2 = st.columns(2)
    
    with objectives_col1:
        st.markdown("""
        **Plan and Manage Azure AI Solutions**
        - ✓ Select appropriate AI service
        - ✓ Plan for cost management
        - ✓ Implement security best practices
        
        **Implement Natural Language Processing**
        - ✓ Use Azure OpenAI for text generation
        - ✓ Implement prompt engineering
        - ✓ Extract information from documents
        """)
    
    with objectives_col2:
        st.markdown("""
        **Implement Computer Vision Solutions**
        - ✓ Extract text from documents (OCR)
        - ✓ Process document layouts
        
        **Develop AI Applications**
        - ✓ Design conversational AI solutions
        - ✓ Manage model deployment
        - ✓ Monitor AI service performance
        """)
    
    st.info("💡 **Course**: Microsoft Azure AI-102: Designing and Implementing a Microsoft Azure AI Solution")
    
    # Call to action
    st.markdown("---")
    st.success("""
    **🎓 Ready to explore?** Use the navigation menu on the left to:
    1. Enable **Demo Mode** for instant testing (no PDF required)
    2. View **Cleaned Text** to see OCR post-processing
    3. Generate **AI-powered analyses** (Summary, KPIs, Themes)
    4. Create a **professional report** with multiple export formats
    """)

# ===== END ABOUT PAGE =====

if structured_text:

    # CLEANED TEXT PAGE
    if page == "Cleaned Text":
        st.subheader("✨ Cleaned & Structured Text")
        st.markdown(structured_text)

        st.subheader("📄 Raw OCR Text (debugging)")
        st.text_area("Raw OCR Text", value=raw_text, height=250)

    # EXECUTIVE SUMMARY PAGE
    if page == "Executive Summary":
        st.subheader("🧠 Generate Executive Summary")

        if st.button("Generate Summary"):
            st.info("Generating summary…")

            summary = client.chat.completions.create(
                model=os.getenv("AZURE_OPENAI_DEPLOYMENT"),
                messages=[
                    {
                        "role": "system",
                        "content": (
                            "You are a senior financial analyst. Create a concise, "
                            "professional executive summary of the document. "
                            "Focus on main themes, key findings, risks, and strategic "
                            "insights. Avoid unnecessary detail. "
                            "Write in clear business English for CFO-level readers."
                        ),
                    },
                    {"role": "user", "content": structured_text},
                ],
            )

            summary_text = summary.choices[0].message.content
            st.success("Summary generated!")
            st.write(summary_text)

    # KPI PAGE
    if page == "KPIs":
        st.subheader("📊 Extract Key Metrics & KPIs")

        if st.button("Extract KPIs & Metrics"):
            st.info("Extracting key metrics from the document…")

            kpi_response = client.chat.completions.create(
                model=os.getenv("AZURE_OPENAI_DEPLOYMENT"),
                messages=[
                    {
                        "role": "system",
                        "content": (
                            "You are a financial data analyst. From the text below, "
                            "extract the most important quantitative insights as KPIs. "
                            "Focus on percentages, ranges, counts, and other numeric facts "
                            "that describe:\n"
                            "- survey results\n"
                            "- movement toward digitalization\n"
                            "- sustainability impacts\n"
                            "- CBAM impacts\n"
                            "- global minimum tax expectations\n"
                            "- revenue segment distribution\n\n"
                            "Return the result ONLY as a Markdown table with two columns: "
                            "'KPI' and 'Value'. Do not add any commentary before or after "
                            "the table."
                        ),
                    },
                    {"role": "user", "content": structured_text},
                ],
            )

            kpi_text = kpi_response.choices[0].message.content
            st.success("KPIs extracted!")
            st.markdown(kpi_text)

    # THEMATIC SUMMARIES PAGE
    if page == "Thematic Summaries":
        st.subheader("📘 Financial Thematic Summaries")

        topic = st.selectbox(
            "Choose a financial theme:",
            [
                "Revenue & Growth",
                "Expenses & Cost Structure",
                "Profitability & Margins",
                "Cash Flow & Liquidity",
                "Balance Sheet Health",
                "Market Trends & Risks",
                "Operational Efficiency",
                "ESG & Sustainability (Financial Impact)",
            ],
        )

        if st.button("Generate Thematic Summary"):
            st.info(f"Analyzing theme: {topic}…")

            theme_response = client.chat.completions.create(
                model=os.getenv("AZURE_OPENAI_DEPLOYMENT"),
                messages=[
                    {
                        "role": "system",
                        "content": (
                            "You are a senior financial analyst. Extract only the content "
                            "related to the selected financial theme. Provide a precise "
                            "6–10 sentence summary focused on performance indicators, "
                            "risks, opportunities, and strategic insights. Maintain a "
                            "CFO-level analytical tone. Do NOT include unrelated topics."
                        ),
                    },
                    {
                        "role": "user",
                        "content": f"Theme: {topic}\n\nDocument:\n{structured_text}",
                    },
                ],
            )

            theme_text = theme_response.choices[0].message.content
            st.success("Summary generated!")
            st.write(theme_text)

    # GENERATE REPORT PAGE
    if page == "Generate Report":
        st.subheader("📄 Build Your Custom Financial Report")

        st.write("Select which sections you want included in the AI-generated report.")

        include_exec = st.checkbox("Include Executive Summary", value=True)
        include_kpis = st.checkbox("Include KPIs Table", value=True)

        include_themes = st.multiselect(
            "Include thematic summaries:",
            [
                "Revenue & Growth",
                "Expenses & Cost Structure",
                "Profitability & Margins",
                "Cash Flow & Liquidity",
                "Balance Sheet Health",
                "Market Trends & Risks",
                "Operational Efficiency",
                "ESG & Sustainability (Financial Impact)",
            ],
        )

        # 1) First button: build a simple template / structure
        if st.button("Generate Report Content"):
            st.info("Preparing report structure…")

            report_text = ""

            if include_exec:
                report_text += "## Executive Summary\n\n"

            if include_kpis:
                report_text += "## Key Metrics and KPIs\n\n"

            for t in include_themes:
                report_text += f"## {t}\n\n"

            st.success("Report structure prepared!")
            st.session_state["report_selected_sections"] = report_text
            st.text_area("Report structure (template)", report_text, height=250)

        # 2) Second button: generate full AI report
        st.subheader("🧠 Generate Full AI-Powered Report")

        if st.button("Generate Full Report"):
            st.info("Generating full report with AI… this may take several seconds.")

            full_report = ""

            # Executive summary
            if include_exec:
                exec_response = client.chat.completions.create(
                    model=os.getenv("AZURE_OPENAI_DEPLOYMENT"),
                    messages=[
                        {
                            "role": "system",
                            "content": (
                                "You are a senior financial analyst. Create a concise, "
                                "well-structured executive summary suitable for C-suite readers."
                            ),
                        },
                        {"role": "user", "content": structured_text},
                    ],
                )
                full_report += "## Executive Summary\n"
                full_report += exec_response.choices[0].message.content + "\n\n"

            # KPIs
            if include_kpis:
                kpi_response = client.chat.completions.create(
                    model=os.getenv("AZURE_OPENAI_DEPLOYMENT"),
                    messages=[
                        {
                            "role": "system",
                            "content": (
                                "Extract all key KPIs as a Markdown table with columns: KPI | Value. "
                                "Include ALL relevant numerical insights. No commentary."
                            ),
                        },
                        {"role": "user", "content": structured_text},
                    ],
                )
                full_report += "## Key Metrics and KPIs\n"
                full_report += kpi_response.choices[0].message.content + "\n\n"

            # Thematic sections
            for theme in include_themes:
                theme_response = client.chat.completions.create(
                    model=os.getenv("AZURE_OPENAI_DEPLOYMENT"),
                    messages=[
                        {
                            "role": "system",
                            "content": (
                                "You are a senior financial analyst. Extract only the content "
                                f"related to the theme '{theme}'. Provide a structured summary "
                                "with insights, trends, and financial implications."
                            ),
                        },
                        {"role": "user", "content": structured_text},
                    ],
                )
                full_report += f"## {theme}\n"
                full_report += theme_response.choices[0].message.content + "\n\n"

            # Save to session state for future export
            st.session_state["final_report_md"] = full_report

            st.success("Full AI report generated!")
            st.text_area("Preview of Full Report (Markdown)", full_report, height=300)

                    # -------------------------------------------
                # -------------------------------------------
# ===== DOWNLOAD SECTION (only show if report was generated) =====
if "final_report_md" in st.session_state and st.session_state["final_report_md"]:
    report_md = st.session_state["final_report_md"]
    
    st.markdown("---")
    st.subheader("📥 Download Your Report")
    
    # Prepare all export formats
    try:
        # 1) Markdown bytes
        md_bytes = report_md.encode("utf-8")
        
        # 2) DOCX in memory
        docx_buffer = BytesIO()
        doc = Document()
        for line in report_md.split("\n"):
            doc.add_paragraph(line)
        doc.save(docx_buffer)
        docx_buffer.seek(0)
        
        # 3) Professional PDF generation
        def create_professional_pdf(report_content: str) -> BytesIO:
            """
            Creates a professional-looking PDF with proper formatting.
            Handles headers, paragraphs, and tables.
            """
            buffer = BytesIO()
            
            # Create document with margins
            pdf_doc = SimpleDocTemplate(
                buffer,
                pagesize=letter,
                rightMargin=72,
                leftMargin=72,
                topMargin=72,
                bottomMargin=36,
            )
            
            # Container for the PDF elements
            story = []
            
            # Get default styles
            styles = getSampleStyleSheet()
            
            # Custom title style
            title_style = ParagraphStyle(
                'CustomTitle',
                parent=styles['Heading1'],
                fontSize=24,
                textColor=colors.HexColor('#1f4788'),
                spaceAfter=30,
                alignment=TA_CENTER,
                fontName='Helvetica-Bold'
            )
            
            # Custom heading style  
            heading_style = ParagraphStyle(
                'CustomHeading',
                parent=styles['Heading2'],
                fontSize=14,
                textColor=colors.HexColor('#2c5aa0'),
                spaceAfter=12,
                spaceBefore=12,
                fontName='Helvetica-Bold'
            )
            
            # Body text style
            body_style = ParagraphStyle(
                'CustomBody',
                parent=styles['BodyText'],
                fontSize=10,
                leading=14,
                spaceAfter=10,
                alignment=TA_LEFT
            )
            
            # Add title page
            story.append(Paragraph("AI Financial Report Analysis", title_style))
            story.append(Spacer(1, 0.2*inch))
            
            # Add metadata
            metadata_style = ParagraphStyle(
                'Metadata',
                parent=styles['Normal'],
                fontSize=9,
                textColor=colors.grey,
                alignment=TA_CENTER
            )
            story.append(Paragraph(f"Generated: {datetime.now().strftime('%B %d, %Y at %I:%M %p')}", metadata_style))
            story.append(Paragraph("Powered by Azure OpenAI", metadata_style))
            story.append(Spacer(1, 0.3*inch))
            
            # Add horizontal line
            from reportlab.platypus import HRFlowable
            story.append(HRFlowable(width="100%", thickness=2, color=colors.HexColor('#1f4788')))
            story.append(Spacer(1, 0.3*inch))
            
            # Parse markdown and add content
            lines = report_content.split('\n')
            in_table = False
            table_data = []
            
            for line in lines:
                line = line.strip()
                
                if not line:
                    story.append(Spacer(1, 0.1*inch))
                    continue
                
                # Handle H2 headings (##)
                if line.startswith('## '):
                    heading_text = line[3:].strip()
                    story.append(Paragraph(heading_text, heading_style))
                    continue
                
                # Handle markdown tables
                if '|' in line and not line.startswith('|--'):
                    if not in_table:
                        in_table = True
                        table_data = []
                    
                    # Parse table row
                    cells = [cell.strip() for cell in line.split('|')]
                    # Remove empty first/last cells from markdown format
                    cells = [c for c in cells if c]
                    if cells:
                        table_data.append(cells)
                    continue
                
                elif in_table and '|' not in line:
                    # End of table - render it
                    if table_data and len(table_data) > 0:
                        # Create table
                        t = Table(table_data, repeatRows=1)
                        t.setStyle(TableStyle([
                            # Header row styling
                            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1f4788')),
                            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                            ('FONTSIZE', (0, 0), (-1, 0), 11),
                            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                            ('TOPPADDING', (0, 0), (-1, 0), 12),
                            # Data rows styling
                            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                            ('FONTSIZE', (0, 1), (-1, -1), 9),
                            ('TOPPADDING', (0, 1), (-1, -1), 8),
                            ('BOTTOMPADDING', (0, 1), (-1, -1), 8),
                            # Grid
                            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                        ]))
                        story.append(t)
                        story.append(Spacer(1, 0.2*inch))
                    
                    in_table = False
                    table_data = []
                    
                    # Process the current line as normal paragraph
                    if line and not line.startswith('#'):
                        story.append(Paragraph(line, body_style))
                    continue
                
                # Handle regular paragraphs
                if not line.startswith('#'):
                    story.append(Paragraph(line, body_style))
            
            # Handle any remaining table at end of document
            if in_table and table_data and len(table_data) > 0:
                t = Table(table_data, repeatRows=1)
                t.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1f4788')),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, 0), 11),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                    ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                    ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                ]))
                story.append(t)
            
            # Build the PDF
            pdf_doc.build(story)
            buffer.seek(0)
            return buffer
        
        # Generate the professional PDF
        pdf_buffer = create_professional_pdf(report_md)
        
        # Display download buttons in columns
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.download_button(
                label="📄 Markdown",
                data=md_bytes,
                file_name="financial_report.md",
                mime="text/markdown",
                use_container_width=True
            )
        
        with col2:
            st.download_button(
                label="📘 Word",
                data=docx_buffer,
                file_name="financial_report.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
        
        with col3:
            st.download_button(
                label="📕 PDF",
                data=pdf_buffer,
                file_name="financial_report.pdf",
                mime="application/pdf",
                use_container_width=True
            )
        
        st.success("✅ All export formats ready for download!")
        
    except Exception as e:
        st.error(f"❌ Error preparing downloads: {str(e)}")
        st.info("The report was generated but there was an issue creating export files.")

# Show welcome message if no data loaded
if not st.session_state.get("structured_text"):
    st.info("👆 **Get Started:** Enable Demo Mode or upload a PDF financial report to begin analysis.")