# 📊 AI Financial Report Summarizer

Upload a financial report (PDF or image) and get:
- Executive summary  
- KPI table (Revenue, COGS, Margin, Opex, EBITDA, Cash)  
- Top 3 variances  
- Action items  

## Stack
- Azure Document Intelligence (OCR)
- Azure OpenAI (summarization)
- Streamlit (UI)

## Quick Start
```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
cp .env .env       # fill in your values
streamlit run app.py

