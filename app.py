import os
import streamlit as st
import pdfplumber
import requests
import email
import yaml
from email import policy
from email.parser import BytesParser
import io
import zipfile

CONFIG_PATH = r"E:\Hackathon\E-MAIL MANAGEMENT\conf\config.yml"

def load_config():
    with open(CONFIG_PATH, "r") as f:
        return yaml.safe_load(f)
    
config = load_config()
os.environ["GOOGLE_API_KEY"] = config['GOOGLE_API_KEY']

# ---------------------------
# Helper: Extract text from PDF
# ---------------------------
def extract_pdf_text(file):
    text = ""
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            text += page.extract_text() + "\n"
    return text.strip()

# ---------------------------
# Helper: Extract text from .eml
# ---------------------------
def extract_eml_text(file):
    msg = BytesParser(policy=policy.default).parse(file)

    text_parts = []

    # Walk through all parts of the email
    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            if content_type == "text/plain":
                text_parts.append(part.get_content())
    else:
        text_parts.append(msg.get_content())

    return "\n".join(text_parts).strip()

# ---------------------------
# Helper: Call Google AI API
# ---------------------------
def call_google_ai(api_key, model, prompt):
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={api_key}"
    headers = {"Content-Type": "application/json"}
    data = {"contents": [{"parts": [{"text": prompt}]}]}
    response = requests.post(url, headers=headers, json=data)
    if response.status_code == 200:
        return response.json()["candidates"][0]["content"]["parts"][0]["text"]
    else:
        return f"Error: {response.text}"

# ---------------------------
# Helper: Call Grok AI API (placeholder until real endpoint available)
# ---------------------------
def call_grok_ai(api_key, model, prompt):
    url = f"https://api.x.ai/{model}"
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
    data = {"messages": [{"role": "user", "content": prompt}]}
    response = requests.post(url, headers=headers, json=data)
    if response.status_code == 200:
        return response.json().get("choices", [{}])[0].get("message", {}).get("content", "")
    else:
        return f"Error: {response.text}"

# ---------------------------
# Streamlit UI
# ---------------------------
# st.set_page_config(page_title="Automatic Email Reader & Responder", layout="wide")
st.set_page_config(page_title="Automatic Email Reader & Responder")
st.title("📧 Automatic Email Reading and Response (EML + Prompt Chaining)")

# Upload EML file
uploaded_file = st.file_uploader("Upload Outlook Mail Chain (.eml)", type="eml")

# Upload PDF
# uploaded_file = st.file_uploader("Upload Outlook Mail Chain (PDF)", type="pdf")

# Select Provider
provider = st.selectbox("Choose AI Provider", ["Google AI", "Grok AI"])

# Model Selection
if provider == "Google AI":
    model = st.selectbox("Select Google AI Model", ["gemini-1.5-flash", "gemini-1.5-pro"])
else:
    model = st.selectbox("Select Grok AI Model", ["grok-beta", "grok-1"])

# API Key input
# api_key = st.text_input("Enter API Key", type="password")
api_key = os.environ["GOOGLE_API_KEY"]

# Process Button
if uploaded_file and api_key and st.button("Process Mail Chain"):
    with st.spinner("Extracting and analyzing mail chain..."):
        # Step 1: Extract PDF Text
        mail_text = extract_eml_text(uploaded_file)
        print("Mail text : \n" , mail_text)

        # Define prompts
        prompt_summary = f"""
        You are an assistant that analyzes emails.
        Read the following email chain and provide a **detailed summary** of all messages.
        
        Email Chain:
        {mail_text}
        """

        prompt_key_info = f"""
        You are an assistant that extracts structured information.
        From the following email chain, **list key information** including:
        - Decisions made
        - Dates and deadlines
        - People involved (names, roles if mentioned)
        - Tasks and responsibilities
        
        Email Chain:
        {mail_text}
        """

        prompt_response = f"""
        You are an assistant that drafts professional emails.
        Based on the following email chain, **generate a polite, professional response email** 
        that acknowledges key points, confirms understanding, and provides next steps.
        
        Email Chain:
        {mail_text}
        """

        # Call AI based on provider
        if provider == "Google AI":
            summary = call_google_ai(api_key, model, prompt_summary)
            key_info = call_google_ai(api_key, model, prompt_key_info)
            response_mail = call_google_ai(api_key, model, prompt_response)
        else:
            summary = call_grok_ai(api_key, model, prompt_summary)
            key_info = call_grok_ai(api_key, model, prompt_key_info)
            response_mail = call_grok_ai(api_key, model, prompt_response)

    # Display Results
    st.subheader("📜 Detailed Summary of Mail Chain")
    st.write(summary)

    st.subheader("🔑 Key Information")
    st.write(key_info)

    st.subheader("✉️ Draft Response Mail")
    st.write(response_mail)

    files = {
        "SUMMARY.txt": summary,
        "KEY_INFORMATION.txt": key_info,
        "RESPONSE_MAIL.txt": response_mail,
    }

    # Create in-memory ZIP
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zf:
        for filename, content in files.items():
            zf.writestr(filename, content)
    zip_buffer.seek(0)

    # Option to download all content
    st.download_button(
        label="Download All Files (ZIP)",
        data=zip_buffer,
        file_name="files.zip",
        mime="application/zip"
    )
