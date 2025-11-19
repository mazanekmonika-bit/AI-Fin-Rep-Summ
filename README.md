📊 AI Financial Report Analyzer

Azure AI-102 Certification Project
Intelligent document analysis using Azure OpenAI and Document Intelligence

🎯 Project Overview

This application demonstrates enterprise-grade implementation of Azure AI services for automated financial document analysis. Built as a capstone project for the Microsoft Azure AI-102 certification course, it showcases practical skills in AI orchestration, prompt engineering, and production-ready deployment.

Key Features

📄 PDF Text Extraction - OCR-powered document ingestion
🧠 AI-Powered Cleaning - Removes artifacts and restructures messy OCR output
📊 Financial Analysis - Executive summaries, KPI extraction, thematic insights
💰 Cost Monitoring - Real-time token usage and API cost tracking
📥 Multi-Format Export - Generate reports in Markdown, Word, and PDF
🎬 Demo Mode - Test with pre-loaded sample data

🏗️ Architecture

USER UPLOAD → PDF → OCR Extraction → AI Cleaning → Multi-Step Analysis → Export
                                            ↓
                            Azure OpenAI (GPT-4/GPT-4o/GPT-3.5)
☁️ Azure Services Used

Service	Purpose

Azure OpenAI	Text cleaning, summarization, analysis
Document Intelligence	OCR extraction from PDFs
API Management	Authentication and rate limiting

🚀 Getting Started

Prerequisites
Python 3.8+
Azure OpenAI service deployment
Valid API credentials
Installation
Clone the repository
bash
   git clone <your-repo-url>
   cd financial-report-analyzer
Install dependencies
bash
   pip install -r requirements.txt
Configure environment variables Create a .env file in the project root:
env
   AZURE_OPENAI_API_KEY=your-api-key-here
   AZURE_OPENAI_ENDPOINT=https://your-resource.openai.azure.com/
   AZURE_OPENAI_DEPLOYMENT=your-deployment-name
   AZURE_OPENAI_API_VERSION=2024-10-21
   
Run the application
bash
   streamlit run app.py
Access the app Open your browser to http://localhost:8501
📦 Dependencies
streamlit>=1.28.0
python-dotenv>=1.0.0
openai>=1.0.0
python-docx>=0.8.11
reportlab>=4.0.0
🎓 Azure AI-102 Concepts Demonstrated
1. Prompt Engineering
Role-based system messages for consistent AI behavior
Clear output specifications (e.g., "Return ONLY as Markdown table")
Temperature control (0.3 for deterministic financial analysis)
Context management for multi-step reasoning
2. Production Best Practices
✅ Comprehensive error handling with user-friendly messages
✅ State management to avoid redundant API calls
✅ Token monitoring and cost optimization
✅ Auto-detection of deployed models
✅ Progress indicators for long operations
✅ Timeout handling for large documents
3. Multi-Modal AI Pipeline
Document Ingestion - PDF upload via Streamlit
OCR Extraction - Text extraction from images/PDFs
AI Cleaning - Fix formatting artifacts and line breaks
Parallel Analysis - Generate summaries, KPIs, themes
Report Synthesis - Combine outputs into cohesive document
🛠️ Usage Guide
Demo Mode (Recommended for Testing)
Check "🎬 Demo Mode" in the sidebar
Explore pre-loaded sample financial data
Navigate through analysis pages
Generate and export reports
Production Mode
Upload a PDF financial report
Wait for OCR extraction and AI cleaning
Navigate through pages:
Cleaned Text - View processed document
Executive Summary - Generate C-suite overview
KPIs - Extract key metrics as table
Thematic Summaries - Analyze specific topics
Generate Report - Build custom multi-section report
Download in your preferred format (MD/DOCX/PDF)
💡 Features in Detail
Intelligent Text Cleaning
Fixes words split across line breaks
Preserves numerical ranges and financial figures
Removes page numbers and repeated headers
Ensures proper spacing (e.g., "$3.2 million" not "$3.2million")
Cost Optimization
Real-time token estimation
Automatic model detection (GPT-4o Mini, GPT-4o, GPT-4 Turbo, GPT-3.5)
Per-operation cost preview
Document size warnings for large files
Error Handling
Graceful degradation for API failures
Specific guidance for common errors:
Rate limiting (429)
Authentication issues (401)
Deployment not found (404)
Request timeouts
Troubleshooting tips with environment variable checks
📊 Supported Models
The app auto-detects your Azure OpenAI deployment:
Model	Cost per 1K Tokens	Best For
GPT-4o Mini	$0.000375	High-volume analysis
GPT-4o	$0.0075	Balanced performance
GPT-4 Turbo	$0.045	Complex reasoning
GPT-3.5 Turbo	$0.00175	Simple tasks
🔒 Security Considerations
API keys stored in .env (never commit to version control)
Masked credentials in troubleshooting displays
No data persistence (stateless session management)
Rate limiting protection
📝 License
This project was created for educational purposes as part of the Azure AI-102 certification course.
🙏 Acknowledgments
Microsoft Azure for AI services
Streamlit for rapid web app development
Azure AI-102 Course for certification preparation
📧 Contact
For questions about this project, please refer to the course materials or Azure documentation.
Built with ❤️ for 🐍
