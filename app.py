import os
from dotenv import load_dotenv
import streamlit as st

load_dotenv()

st.set_page_config(page_title="AI Financial Report Summarizer", layout="centered")
st.title("📊 AI Financial Report Summarizer")

st.write("Repo setup is working. Next we’ll hook up Azure services.")
st.code(
    f"DI endpoint: {os.getenv('AZURE_DI_ENDPOINT') or 'not set'}\n"
    f"AOAI endpoint: {os.getenv('AZURE_OPENAI_ENDPOINT') or 'not set'}",
    language="bash",
)

st.success("If you can see this page, Streamlit is installed and running.")

