import os
import streamlit as st
from dotenv import load_dotenv
from openai import AzureOpenAI
from io import BytesIO
from docx import Document
from reportlab.lib.pagesizes import letter, A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from datetime import datetime
import re
from ocr import extract_text_from_pdf

# 🔹 MUST be first Streamlit call
st.set_page_config(page_title="AI Financial Report Analyzer", layout="wide")
load_dotenv()

# ===== SESSION STATE INITIALIZATION =====
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
# ===== ERROR-SAFE AI CALL WRAPPER =====
def safe_ai_call(system_prompt: str, user_content: str, operation_name: str, max_tokens: int = 2000) -> str:
    """
    Wrapper for Azure OpenAI calls with comprehensive error handling.
    """
    try:
        response = client.chat.completions.create(
            model=os.getenv("AZURE_OPENAI_DEPLOYMENT"),
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_content}
            ],
            max_tokens=max_tokens,
            temperature=0.3,
            timeout=60,
        )
        return response.choices[0].message.content
    
    except Exception as e:
        error_type = type(e).__name__
        st.error(f"❌ **{operation_name} Failed**")
        
        if "rate_limit" in str(e).lower() or "429" in str(e):
            st.warning("**Rate Limit Reached** - Wait a few minutes and try again")
        elif "authentication" in str(e).lower() or "401" in str(e):
            st.warning("**Authentication Error** - Check your AZURE_OPENAI_API_KEY")
        elif "not found" in str(e).lower() or "404" in str(e):
            st.warning("**Deployment Not Found** - Check your AZURE_OPENAI_DEPLOYMENT name")
        elif "timeout" in str(e).lower():
            st.warning("**Request Timeout** - Try again with a shorter document")
        else:
            st.warning(f"**Unexpected Error: {error_type}** - {str(e)}")
        
        with st.expander("🔧 Troubleshooting Tips"):
            st.code(f"""
# Check your .env file contains:
AZURE_OPENAI_API_KEY=your-key-here
AZURE_OPENAI_ENDPOINT=https://your-resource.openai.azure.com/
AZURE_OPENAI_DEPLOYMENT=your-deployment-name
AZURE_OPENAI_API_VERSION=2024-02-15-preview

# Current values (masked):
AZURE_OPENAI_API_KEY={'*' * 20 if os.getenv('AZURE_OPENAI_API_KEY') else 'NOT SET'}
AZURE_OPENAI_ENDPOINT={os.getenv('AZURE_OPENAI_ENDPOINT', 'NOT SET')}
AZURE_OPENAI_DEPLOYMENT={os.getenv('AZURE_OPENAI_DEPLOYMENT', 'NOT SET')}
            """)
        
        return ""
# ===== END ERROR HANDLER =====

# ===== TOKEN ESTIMATION & COST TRACKING =====
def estimate_tokens(text: str) -> int:
    """Rough token estimation: 1 token ≈ 4 characters"""
    return len(text) // 4

def get_deployment_model() -> str:
    """
    Auto-detect which Azure OpenAI model is being used.
    Returns the deployment name from environment variables.
    """
    deployment = os.getenv("AZURE_OPENAI_DEPLOYMENT", "unknown")
    return deployment.lower()

def estimate_cost(tokens: int, model_type: str = None) -> float:
    """
    Estimate Azure OpenAI API cost based on the actual deployed model.
    Automatically detects model from AZURE_OPENAI_DEPLOYMENT if not specified.
    
    Pricing:
    - GPT-4o Mini: ~$0.000375 per 1K tokens (avg)
    - GPT-4o: ~$0.0075 per 1K tokens (avg)
    - GPT-4 Turbo: ~$0.045 per 1K tokens (avg)
    - GPT-3.5 Turbo: ~$0.00175 per 1K tokens (avg)
    """
    # Auto-detect model if not specified
    if model_type is None:
        model_type = get_deployment_model()
    else:
        model_type = model_type.lower()
    
    # Determine cost based on model
    if "4o-mini" in model_type or "gpt-4o-mini" in model_type:
        cost_per_1k = 0.000375
        model_name = "GPT-4o Mini"
    elif "4o" in model_type or "gpt-4o" in model_type:
        cost_per_1k = 0.0075
        model_name = "GPT-4o"
    elif "gpt-4-turbo" in model_type or "gpt4-turbo" in model_type:
        cost_per_1k = 0.045
        model_name = "GPT-4 Turbo"
    elif "gpt-4" in model_type or "gpt4" in model_type:
        cost_per_1k = 0.045
        model_name = "GPT-4"
    elif "gpt-35" in model_type or "gpt-3.5" in model_type:
        cost_per_1k = 0.00175
        model_name = "GPT-3.5 Turbo"
    else:
        cost_per_1k = 0.002
        model_name = "Unknown Model"
    
    # Store model name in session state for display
    st.session_state["detected_model"] = model_name
    st.session_state["cost_per_1k"] = cost_per_1k
    
    return (tokens / 1000) * cost_per_1k

def format_large_number(num: int) -> str:
    """Format numbers with commas"""
    return f"{num:,}"

def clean_ai_output(text: str) -> str:
    """
    Post-process AI output to fix common formatting issues.
    This is a safety net if the AI doesn't follow spacing rules.
    """
    # Fix: $3.2million -> $3.2 million
    text = re.sub(r'\$(\d+\.?\d*)(million|billion|thousand)', r'$\1 \2', text)
    
    # Fix: 22percent -> 22 percent
    text = re.sub(r'(\d+\.?\d*)percent', r'\1 percent', text)
    
    # Fix: number concatenated with word (generic)
    text = re.sub(r'(\d)([a-zA-Z])', r'\1 \2', text)
    
    return text
# ===== END TOKEN TRACKING =====

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

# ===== SIDEBAR METRICS =====
if st.session_state.get("structured_text"):
    st.sidebar.markdown("---")
    st.sidebar.subheader("📊 Document Stats")
    
    # Show detected model
    detected_model = st.session_state.get("detected_model", "Detecting...")
    cost_per_1k = st.session_state.get("cost_per_1k", 0)
    
    st.sidebar.info(f"""
    **🤖 Model:** {detected_model}  
    **💰 Rate:** ${cost_per_1k:.6f} per 1K tokens
    """)
    
    structured_text = st.session_state.get("structured_text", "")
    doc_tokens = estimate_tokens(structured_text)
    estimated_cost = estimate_cost(doc_tokens * 2)
    char_count = len(structured_text)
    word_count = len(structured_text.split())
    
    st.sidebar.metric("Document Tokens", format_large_number(doc_tokens))
    st.sidebar.metric("Est. Cost per Analysis", f"${estimated_cost:.4f}")
    st.sidebar.metric("Characters", format_large_number(char_count))
    st.sidebar.metric("Words (approx)", format_large_number(word_count))
    
    if doc_tokens > 8000:
        st.sidebar.warning("⚠️ Large document")
    elif doc_tokens > 4000:
        st.sidebar.info("💡 Medium-sized document")
    else:
        st.sidebar.success("✅ Optimal size")
# ===== END SIDEBAR METRICS =====

st.title("📊 AI Financial Report Analyzer")

# ===== DEMO MODE =====
demo_mode = st.sidebar.checkbox("🎬 Demo Mode (Use Sample Data)", value=False)

if demo_mode:
    st.info("📊 **Demo Mode Active**: Using pre-loaded sample financial report")
    
    sample_text = """Financial Performance Overview

Revenue and Growth
Total revenue for fiscal year 2024 reached $45.2 million, representing a 23 percent increase compared to the previous year. This growth was primarily driven by digital transformation initiatives which accounted for 67 percent of new revenue streams. Recurring revenue grew to 58 percent of total revenue, up from 43 percent in 2023.

Expense Management
Operating expenses increased by 15 percent to $32.1 million, demonstrating improved operational efficiency as revenue grew faster than costs. The company maintained strong cost discipline with a 71 percent gross margin, up from 68 percent in the prior year.

Profitability Analysis
Net profit margin improved from 18 percent to 21 percent, with EBITDA reaching $12.4 million. Return on equity increased to 24 percent, exceeding industry benchmarks of 18 to 20 percent.

Cash Flow Performance
Operating cash flow was $13.2 million, representing a 28 percent  increase  year  over  year. Free  cash  flow  reached  $9.8 million after capital expenditures of $3.4 million. The company ended the year with $18.5 million in cash and equivalents.

Market Position and Trends
Market share in core segments grew from 12 percent to 15 percent. Customer acquisition costs decreased by 18 percent while customer lifetime value increased by 32 percent. Digital channels now represent 67 percent of total revenue, up from 45 percent in 2023.

Sustainability and ESG Initiatives
Sustainability investments totaled $3.2 million, with measurable carbon reduction of 22 percent. ESG compliance costs for CBAM are estimated at $1.8 million annually. Green revenue reached $8.4 million, or 19 percent of total revenue.

Risk Factors and Challenges
Supply chain disruptions affected 12 percent of operations, resulting in $2.1 million in additional costs. Currency fluctuations impacted margins by 2.3 percent. Regulatory compliance costs increased by $1.2 million year over year.

Global Tax Considerations
The company is preparing for global minimum tax implementation, estimating $0.8 million to $1.2 million in additional tax liability. Effective tax rate was 24 percent, slightly above the industry average of 22 percent.

Strategic Outlook
Management projects 18 to 22 percent revenue growth for 2025, driven by product expansion and market penetration. Capital expenditure plans include $5.5 million for technology infrastructure and $2.8 million for sustainability initiatives. The company expects to maintain profit margins above 20 percent while investing in growth."""
    
    sample_raw = "Sample OCR text before cleaning (with simulated artifacts)..."
    
    st.session_state["structured_text"] = sample_text
    st.session_state["raw_text"] = sample_raw
    st.session_state["demo_mode"] = True
    
    st.success("✅ Sample data loaded! Navigate through pages to see AI analysis in action.")
    uploaded_file = None

else:
    st.session_state["demo_mode"] = False
    uploaded_file = st.file_uploader("Upload a PDF financial report", type=["pdf"])
# ===== END DEMO MODE =====

# ===== FILE PROCESSING =====
if uploaded_file is not None or st.session_state.get("demo_mode", False):
    if uploaded_file is not None and not st.session_state.get("demo_mode", False):
        current_file_id = f"{uploaded_file.name}_{uploaded_file.size}"
        
        if st.session_state.last_file_id != current_file_id:
            st.info("🔄 New file detected. Extracting text from PDF...")
            
            try:
                file_bytes = uploaded_file.read()
                raw_text = extract_text_from_pdf(file_bytes)
                st.success("✅ OCR extraction complete!")
                
                st.info("🧠 Cleaning and structuring the extracted text with AI...")
                
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
                                "- Fix words where letters are split by line breaks.\n"
                                "- Fix numbers and ranges that are broken across lines.\n"
                                "- Remove page numbers, repeated headings, and footers.\n"
                                "- Preserve all content and meaning. Do NOT summarize.\n"
                                "- Ensure proper spacing: '$3.2 million' NOT '$3.2million'\n"
                            ),
                        },
                        {"role": "user", "content": raw_text},
                    ],
                    timeout=60,
                )
                structured_text = cleaned.choices[0].message.content
                
                st.session_state.structured_text = structured_text
                st.session_state.raw_text = raw_text
                st.session_state.last_file_id = current_file_id
                st.session_state.processing_complete = True
                
                st.success("✅ Text cleaning complete!")
                
            except Exception as e:
                st.error(f"❌ Error during processing: {str(e)}")
                st.stop()
        else:
            st.success("✅ Using previously processed document")
# ===== END FILE PROCESSING =====

# Get text from session state
structured_text = st.session_state.get("structured_text", "")
raw_text = st.session_state.get("raw_text", "")

# ===== ABOUT PAGE =====
if page == "About This Project":
    st.header("📚 Azure AI-102 Project: Financial Report Analyzer")
    
    st.markdown("""
    ### 🎯 Project Overview
    This application demonstrates practical implementation of **Azure AI services** for enterprise 
    document intelligence and natural language processing. Built as a capstone project for the 
    **Microsoft Azure AI-102 certification course**.
    """)
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown(f"""
        #### ☁️ Azure Services
        - ✅ **Azure OpenAI Service**
        - 🤖 Model: `{get_deployment_model().upper()}`
        - ✅ Document Intelligence (OCR)
        - ✅ API Authentication
        
        #### 🐍 Python Stack
        - Streamlit - Web UI
        - python-docx - Word export
        - ReportLab - PDF generation
        """)
    
    with col2:
        st.markdown("""
        #### 🧠 AI Capabilities
        - 📖 Text Extraction & Cleaning
        - 🎯 Prompt Engineering
        - 📊 Structured Data Extraction
        - 📝 Multi-step Reasoning
        - 🔄 Context Management
        """)
    
    st.markdown("---")
    st.subheader("🔑 Azure AI-102 Concepts")
    
    tab1, tab2, tab3 = st.tabs(["Prompt Engineering", "AI Pipeline", "Production"])
    
    with tab1:
        st.markdown("""
        #### Prompt Engineering Best Practices
        
        **1. Role-Based System Messages**
        - Defines AI personality upfront
        
        **2. Clear Output Specifications**
        - "Return ONLY as a Markdown table"
        
        **3. Temperature Control**
        - 0.3 for financial analysis (deterministic)
        
        **4. Token Management**
        - Monitor usage to control costs
        """)
    
    with tab2:
        st.markdown("""
        #### Multi-Step AI Pipeline
        
        1. Document Ingestion (PDF upload)
        2. OCR Extraction (text from images)
        3. AI Cleaning (fix artifacts)
        4. Multi-Modal Analysis (summaries, KPIs)
        5. Report Synthesis (combine outputs)
        """)
    
    with tab3:
        st.markdown("""
        #### Production-Ready Features
        
        - Error handling with friendly messages
        - State management (avoid reprocessing)
        - Cost optimization (token monitoring)
        - Progress indicators
        - Demo mode for testing
        - Auto-model detection
        """)
    
    st.markdown("---")
    st.subheader("🏗️ Architecture")
    st.code("""
    USER → Upload PDF → OCR → AI Cleaning → Analysis → Export
    """, language="text")
    
    st.info("💡 **Course**: Microsoft Azure AI-102: Designing and Implementing a Microsoft Azure AI Solution")

# ===== OTHER PAGES =====
if structured_text:
    
    if page == "Cleaned Text":
        st.subheader("✨ Cleaned & Structured Text")
        st.markdown(structured_text)
        st.subheader("📄 Raw OCR Text")
        st.text_area("Raw", value=raw_text, height=250)
    
    if page == "Executive Summary":
        st.subheader("🧠 Generate Executive Summary")
        if st.button("Generate Summary"):
            with st.spinner("🧠 Generating..."):
                summary_text = safe_ai_call(
                    system_prompt=(
                        "You are a senior financial analyst. Create a concise executive summary. "
                        "FORMATTING RULES:\n"
                        "- Always add a space before 'million' (write '$3.2 million' NOT '$3.2million')\n"
                        "- Always add a space before 'percent' (write '22 percent' NOT '22percent')\n"
                        "- Never concatenate numbers with words\n"
                        "- Use proper spacing in all financial figures\n"
                        "Write in clear, professional business English for C-suite readers."
                    ),
                    user_content=structured_text,
                    operation_name="Executive Summary Generation"
                )
            if summary_text:
                summary_text = clean_ai_output(summary_text)
                st.success("✅ Summary generated!")
                st.write(summary_text)
    
    if page == "KPIs":
        st.subheader("📊 Extract Key Metrics")
        if st.button("Extract KPIs"):
            with st.spinner("📊 Extracting..."):
                kpi_text = safe_ai_call(
                    system_prompt=(
                        "Extract key financial KPIs as a Markdown table with two columns: KPI | Value\n\n"
                        "FORMATTING RULES:\n"
                        "- Use proper spacing: '$3.2 million' NOT '$3.2million'\n"
                        "- Use proper spacing: '22 percent' NOT '22percent'\n"
                        "- Keep numbers and units separate\n"
                        "Return ONLY the table, no commentary."
                    ),
                    user_content=structured_text,
                    operation_name="KPI Extraction"
                )
            if kpi_text:
                kpi_text = clean_ai_output(kpi_text)
                st.success("✅ KPIs extracted!")
                st.markdown(kpi_text)
    
    if page == "Thematic Summaries":
        st.subheader("📘 Thematic Analysis")
        topic = st.selectbox("Choose theme:", [
            "Revenue & Growth",
            "Expenses & Cost Structure",
            "Profitability & Margins",
            "Cash Flow & Liquidity"
        ])
        if st.button("Generate Thematic Summary"):
            with st.spinner(f"🔍 Analyzing {topic}..."):
                theme_text = safe_ai_call(
                    system_prompt=(
                        f"You are a senior financial analyst. Extract content related to '{topic}' "
                        "and provide focused financial insights.\n\n"
                        "FORMATTING RULES:\n"
                        "- Always use proper spacing in numbers: '$3.2 million' NOT '$3.2million'\n"
                        "- Write '22 percent' NOT '22percent'\n"
                        "- Keep all financial figures properly spaced"
                    ),
                    user_content=structured_text,
                    operation_name=f"Thematic Analysis: {topic}"
                )
            if theme_text:
                theme_text = clean_ai_output(theme_text)
                st.success("✅ Summary generated!")
                st.write(theme_text)
    
    if page == "Generate Report":
        st.subheader("📄 Build Custom Report")
        
        include_exec = st.checkbox("Executive Summary", value=True)
        include_kpis = st.checkbox("KPIs Table", value=True)
        include_themes = st.multiselect("Thematic summaries:", [
            "Revenue & Growth",
            "Profitability & Margins",
            "Cash Flow & Liquidity"
        ])
        
        if st.button("Generate Full Report"):
            num_ops = sum([include_exec, include_kpis, len(include_themes)])
            total_tokens = estimate_tokens(structured_text) * num_ops * 2
            total_cost = estimate_cost(total_tokens)
            
            st.info(f"🧠 Generating with {num_ops} AI operations...")
            st.caption(f"📊 ~{format_large_number(total_tokens)} tokens | ${total_cost:.4f}")
            
            full_report = ""
            
            if include_exec:
                with st.spinner("Executive summary..."):
                    exec_text = safe_ai_call(
                        system_prompt=(
                            "Create executive summary for C-suite readers.\n"
                            "CRITICAL: Use proper spacing in all numbers:\n"
                            "- '$3.2 million' NOT '$3.2million'\n"
                            "- '22 percent' NOT '22percent'"
                        ),
                        user_content=structured_text,
                        operation_name="Report: Executive Summary"
                    )
                if exec_text:
                    exec_text = clean_ai_output(exec_text)
                    full_report += "## Executive Summary\n" + exec_text + "\n\n"
            
            if include_kpis:
                with st.spinner("Extracting KPIs..."):
                    kpi_text = safe_ai_call(
                        system_prompt=(
                            "Extract KPIs as table: KPI | Value\n"
                            "Use proper spacing: '$3.2 million' not '$3.2million'"
                        ),
                        user_content=structured_text,
                        operation_name="Report: KPIs"
                    )
                if kpi_text:
                    kpi_text = clean_ai_output(kpi_text)
                    full_report += "## Key Metrics\n" + kpi_text + "\n\n"
            
            for theme in include_themes:
                with st.spinner(f"Analyzing {theme}..."):
                    theme_text = safe_ai_call(
                        system_prompt=(
                            f"Analyze '{theme}' with financial insights.\n"
                            "CRITICAL: Proper spacing in numbers: '$3.2 million' NOT '$3.2million'"
                        ),
                        user_content=structured_text,
                        operation_name=f"Report: {theme}"
                    )
                if theme_text:
                    theme_text = clean_ai_output(theme_text)
                    full_report += f"## {theme}\n" + theme_text + "\n\n"
            
            st.session_state["final_report_md"] = full_report
            st.success("✅ Report generated!")
            st.text_area("Preview", full_report, height=300)

# ===== DOWNLOAD SECTION =====
if "final_report_md" in st.session_state and st.session_state["final_report_md"]:
    report_md = st.session_state["final_report_md"]
    
    st.markdown("---")
    st.subheader("📥 Download Your Report")
    
    try:
        # Markdown
        md_bytes = report_md.encode("utf-8")
        
        # DOCX
        docx_buffer = BytesIO()
        doc = Document()
        for line in report_md.split("\n"):
            doc.add_paragraph(line)
        doc.save(docx_buffer)
        docx_buffer.seek(0)
        
        # PDF
        def create_professional_pdf(content: str) -> BytesIO:
            buffer = BytesIO()
            pdf_doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=72, leftMargin=72, topMargin=72, bottomMargin=36)
            story = []
            styles = getSampleStyleSheet()
            
            title_style = ParagraphStyle('Title', parent=styles['Heading1'], fontSize=24, 
                                        textColor=colors.HexColor('#1f4788'), alignment=TA_CENTER, fontName='Helvetica-Bold')
            heading_style = ParagraphStyle('Heading', parent=styles['Heading2'], fontSize=14,
                                          textColor=colors.HexColor('#2c5aa0'), fontName='Helvetica-Bold', spaceAfter=12)
            body_style = ParagraphStyle('Body', parent=styles['BodyText'], fontSize=10, leading=14, spaceAfter=10)
            metadata_style = ParagraphStyle('Metadata', parent=styles['Normal'], fontSize=9, 
                                           textColor=colors.grey, alignment=TA_CENTER)
            
            # Title page
            story.append(Paragraph("AI Financial Report Analysis", title_style))
            story.append(Spacer(1, 0.2*inch))
            story.append(Paragraph(f"Generated: {datetime.now().strftime('%B %d, %Y at %I:%M %p')}", metadata_style))
            story.append(Paragraph("Powered by Azure OpenAI", metadata_style))
            story.append(Spacer(1, 0.3*inch))
            story.append(HRFlowable(width="100%", thickness=2, color=colors.HexColor('#1f4788')))
            story.append(Spacer(1, 0.3*inch))
            
            # Content
            lines = content.split('\n')
            in_table = False
            table_data = []
            
            for line in lines:
                line = line.strip()
                
                if not line:
                    story.append(Spacer(1, 0.1*inch))
                    continue
                
                if line.startswith('## '):
                    story.append(Paragraph(line[3:], heading_style))
                    continue
                
                # Handle markdown tables
                if '|' in line and not line.startswith('|--'):
                    if not in_table:
                        in_table = True
                        table_data = []
                    cells = [cell.strip() for cell in line.split('|')]
                    cells = [c for c in cells if c]
                    if cells:
                        table_data.append(cells)
                    continue
                
                elif in_table and '|' not in line:
                    if table_data and len(table_data) > 0:
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
                        story.append(Spacer(1, 0.2*inch))
                    in_table = False
                    table_data = []
                    if line and not line.startswith('#'):
                        story.append(Paragraph(line, body_style))
                    continue
                
                if line and not line.startswith('#'):
                    story.append(Paragraph(line, body_style))
            
            # Handle remaining table
            if in_table and table_data and len(table_data) > 0:
                t = Table(table_data, repeatRows=1)
                t.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1f4788')),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                ]))
                story.append(t)
            
            pdf_doc.build(story)
            buffer.seek(0)
            return buffer
        
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