import os

import streamlit as st
from dotenv import load_dotenv
from openai import AzureOpenAI

from io import BytesIO
from docx import Document
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
import textwrap


from ocr import extract_text_from_pdf

# 🔹 MUST be first Streamlit call
st.set_page_config(page_title="AI Financial Report Analyzer", layout="wide")

load_dotenv()

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
        "Cleaned Text",
        "Executive Summary",
        "KPIs",
        "Thematic Summaries",
        "Generate Report",
    ],
)

st.title("📊 AI Financial Report Analyzer")

uploaded_file = st.file_uploader("Upload a PDF financial report", type=["pdf"])

# Make sure these variables always exist
raw_text: str = ""
structured_text: str = ""

# -----------------------------
# OCR + CLEANING
# -----------------------------
if uploaded_file is not None:
    st.info("Extracting text from PDF… please wait.")

    # 1) OCR extraction
    file_bytes = uploaded_file.read()
    raw_text = extract_text_from_pdf(file_bytes)

    st.success("OCR extraction complete!")

    # 2) Clean & structure OCR text with Azure OpenAI
    st.info("Cleaning and structuring the extracted text with AI…")

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
    )

    structured_text = cleaned.choices[0].message.content

# -----------------------------
# PAGES (only if we have structured_text)
# -----------------------------
if structured_text:

    # CLEANED TEXT PAGE
    if page == "Cleaned Text":
        st.subheader("✨ Cleaned & Structured Text")
        st.write(structured_text)

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
        # 📥 DOWNLOAD REPORT: MD / DOCX / PDF
        # -------------------------------------------
        if "final_report_md" in st.session_state:
            report_md = st.session_state["final_report_md"]

            # 1) Markdown bytes
            md_bytes = report_md.encode("utf-8")

            # 2) DOCX in memory
            docx_buffer = BytesIO()
            doc = Document()
            for line in report_md.split("\n"):
                doc.add_paragraph(line)
            doc.save(docx_buffer)
            docx_buffer.seek(0)

            # 3) PDF in memory – slightly nicer formatting
            pdf_buffer = BytesIO()
            c = canvas.Canvas(pdf_buffer, pagesize=A4)
            width, height = A4

        # margins
        left_margin = 50
        top_margin = height - 60
        bottom_margin = 60

        # start text object
        textobject = c.beginText(left_margin, top_margin)

        for line in report_md.split("\n"):
            # headings
            if line.startswith("## "):
                # draw previous text run
                c.setFont("Times-Bold", 14)
                heading = line[3:].strip()
                c.drawString(left_margin, textobject.getY(), heading)
                # move cursor down a bit
                textobject.setTextOrigin(left_margin, textobject.getY() - 24)
                textobject.setFont("Times-Roman", 11)
                continue

            # simple blank line
            if line.strip() == "":
                textobject.textLine("")
                continue

            # normal paragraph text, wrapped
            wrapped = textwrap.wrap(line, width=90)
            if not wrapped:
                textobject.textLine("")
            else:
                for subline in wrapped:
                    textobject.textLine(subline)

            # simple page break
            if textobject.getY() < bottom_margin:
                c.drawText(textobject)
                c.showPage()
                textobject = c.beginText(left_margin, top_margin)
                textobject.setFont("Times-Roman", 11)

        # finish last page
        c.drawText(textobject)
        c.showPage()
        c.save()
        pdf_buffer.seek(0)


        # Download buttons
        st.download_button(
                label="⬇️ Download Report as Markdown (.md)",
                data=md_bytes,
                file_name="financial_report.md",
                mime="text/markdown",
            )

        st.download_button(
                label="⬇️ Download Report as Word (.docx)",
                data=docx_buffer,
                file_name="financial_report.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )

        st.download_button(
                label="⬇️ Download Report as PDF (.pdf)",
                data=pdf_buffer,
                file_name="financial_report.pdf",
                mime="application/pdf",
            )


else:
    st.info("Please upload a PDF financial report to begin.")
