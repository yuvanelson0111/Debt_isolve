import streamlit as st
import pandas as pd
import re
import subprocess
import smtplib
from email.mime.text import MIMEText
from datetime import datetime
from thefuzz import process
from PyPDF2 import PdfReader
import threading
import speech_recognition as sr
import asyncio
import edge_tts
import os
import uuid
import pygame

st.set_page_config(page_title="💰 Voilà Debt Assistant", layout="wide")

def speak_text(text):
    def truncate_response(text):
        lines = text.splitlines()
        return '\n'.join(lines[:2])

    async def run_tts(text, audio_path):
        communicate = edge_tts.Communicate(text=truncate_response(text), voice="en-IN-NeerjaNeural")
        await communicate.save(audio_path)
        try:
            # Play audio using pygame
            pygame.mixer.init()
            pygame.mixer.music.load(audio_path)
            pygame.mixer.music.play()
            while pygame.mixer.music.get_busy():
                pygame.time.Clock().tick(10)
            pygame.mixer.quit()
        finally:
            try:
                os.remove(audio_path)
            except Exception as e:
                print(f"Failed to delete {audio_path}: {e}")

    audio_path = f"voice_{uuid.uuid4().hex}.mp3"
    threading.Thread(target=lambda: asyncio.run(run_tts(text, audio_path))).start()

def recognize_once():
    recognizer = sr.Recognizer()
    mic = sr.Microphone()
    with mic as source:
        recognizer.adjust_for_ambient_noise(source)
        st.info("🎙️ Speak now...")
        try:
            audio = recognizer.listen(source, timeout=6, phrase_time_limit=6)
            text = recognizer.recognize_google(audio)
            st.success(f"✅ You said: {text}")
            return text
        except sr.WaitTimeoutError:
            st.warning("⌛ Timeout. No speech detected.")
        except sr.UnknownValueError:
            st.warning("❌ Could not understand audio.")
        except sr.RequestError:
            st.error("❌ Google STT request failed.")
    return ""

# === CSS ===
st.markdown("""
<style>
.stChatMessage {
    background-color: #e6f0ff;
    border-radius: 10px;
    padding: 8px;
    margin-bottom: 10px;
    max-width: 80%;
    word-wrap: break-word;
}
.stChatMessage.user {
    background-color: #d1f7c4;
    align-self: flex-end;
    margin-left: auto;
}
.stChatMessage.bot {
    background-color: #f0f0f0;
    align-self: flex-start;
    margin-right: auto;
}
.chat-container {
    display: flex;
    flex-direction: column;
    gap: 10px;
}
.footer-tag {
    font-size: 11px;
    color: gray;
    margin-top: 20px;
    text-align: center;
}
section[data-testid="stChatInput"] {
    display: flex;
    justify-content: center;
}
section[data-testid="stChatInput"] > div {
    width: 60%;
}
</style>
""", unsafe_allow_html=True)

loan_and_debt_pdf = r"D:\yuva\working_chatbot\Encyclopedia of Loans and Debt Processes for Banking Chatbots.pdf"
USER_XLSX_PATH = r"D:\yuva\working_chatbot\new_customer.xlsx"
LOAN_DETAILS_XLSX_PATH = r"D:\yuva\working_chatbot\loan_accounts.xlsx"
OLLAMA_PATH = r"C:/Users/vmuser3189/AppData/Local/Programs/Ollama/ollama.exe"
OLLAMA_MODEL = "qwen2.5"
SMTP_EMAIL = "yuvanelson0@gmail.com"
SMTP_PASSWORD = "ojak nzfm qcnv zlnx"
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587

def load_pdf_chunks(pdf_path):
    reader = PdfReader(pdf_path)
    full_text = ""
    for page in reader.pages:
        text = page.extract_text()
        if text:
            full_text += text + "\n\n"
    chunks = [p.strip() for p in full_text.split('\n\n') if p.strip()]
    return chunks

pdf_chunks = load_pdf_chunks(loan_and_debt_pdf)

def call_ollama(prompt):
    try:
        result = subprocess.run(
            [OLLAMA_PATH, "run", OLLAMA_MODEL],
            input=prompt,
            text=True,
            capture_output=True,
            encoding="utf-8",
            errors="replace"
        )
        return result.stdout.strip()
    except Exception as e:
        return f"Error: {str(e)}"

def find_relevant_chunk(question, chunks):
    matches = process.extract(question, chunks, limit=5)
    best_match, score = matches[0]
    if score > 60:
        return best_match
    else:
        return "\n\n".join([m[0] for m in matches if m[1] > 50])

def answer_from_pdf(question):
    chunk = find_relevant_chunk(question, pdf_chunks)
    if not chunk:
        return "Sorry, I couldn't find relevant information in the document. Could you please rephrase your question?"
    prompt = f"""You are a helpful assistant answering questions about loans and debt using the following document excerpt:

\"\"\"{chunk}\"\"\"

Question: {question}

Answer in clear, concise language."""
    answer = call_ollama(prompt)
    if not answer:
        return "Sorry, I couldn't generate an answer at this time."
    return answer

def send_email(to_email, subject, body):
    try:
        msg = MIMEText(body)
        msg["Subject"] = subject
        msg["From"] = SMTP_EMAIL
        msg["To"] = to_email
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(SMTP_EMAIL, SMTP_PASSWORD)
        server.sendmail(SMTP_EMAIL, to_email, msg.as_string())
        server.quit()
        return True
    except Exception:
        return False

def load_user_data(xlsx_path):
    try:
        df = pd.read_excel(xlsx_path, dtype=str)
        df.columns = df.columns.str.strip()
        return df
    except Exception as e:
        st.error(f"Could not load user data: {e}")
        return pd.DataFrame()

user_df = load_user_data(USER_XLSX_PATH)

def is_valid_user_id(user_id):
    return not user_df[user_df['user_id'] == user_id].empty

def is_valid_account_number(user_id, account_number):
    return not user_df[(user_df['user_id'] == user_id) & (user_df['Account_no'] == account_number)].empty

def get_account_details(user_id, account_number):
    row = user_df[(user_df['user_id'] == user_id) & (user_df['Account_no'] == account_number)]
    if not row.empty:
        details = row.squeeze().to_dict()
        labels = {
            "user_id": "User Id",
            "Account_no": "Account No",
            "name": "Name",
            "email": "Email",
            "phone_number": "Phone Number",
            "dob": "Dob",
            "kyc_status": "Kyc Status",
            "address": "Address"
        }
        details_str = "\n".join([f"- **{labels[k]}:** {details.get(k, '')}" for k in labels])
        return f"\n\n{details_str}"
    else:
        return "Account details not found."

def format_loan_details(loans_df):
    display_cols = [
        "loan_id", "loan_type", "loan_amount", "disbursed_amount", "disbursement_date",
        "interest_rate", "tenure_months", "emi_amount", "emi_due_date", "emi_frequency",
        "outstanding_principal", "status", "collateral_type", "collateral_value",
        "prepayment_allowed", "prepayment_charges"
    ]
    col_names = [
        "Loan ID", "Type", "Amount", "Disbursed", "Disb. Date", "Interest",
        "Tenure (months)", "EMI", "EMI Due", "EMI Freq.", "Outstanding", "Status",
        "Collateral", "Collat. Value", "Prepay Allowed", "Prepay Charges"
    ]
    table = loans_df[display_cols]
    table.columns = col_names
    return "**Loan Details**\n\n" + table.to_markdown(index=False)

def load_loan_data(xlsx_path):
    try:
        df = pd.read_excel(xlsx_path, dtype=str)
        df.columns = df.columns.str.strip()
        return df
    except Exception as e:
        st.error(f"Could not load loan details: {e}")
        return pd.DataFrame()

loan_df = load_loan_data(LOAN_DETAILS_XLSX_PATH)

def get_loans_by_user(user_id):
    return loan_df[loan_df['user_id'] == user_id]

loan_field_map = {
    "loan amount": "loan_amount",
    "amount disbursed": "disbursed_amount",
    "disbursed amount": "disbursed_amount",
    "disbursement date": "disbursement_date",
    "emi months": "tenure_months",
    "tenure": "tenure_months",
    "tenure months": "tenure_months",
    "emi": "emi_amount",
    "emi amount": "emi_amount",
    "emi due date": "emi_due_date",
    "emi frequency": "emi_frequency",
    "outstanding": "outstanding_principal",
    "principal": "outstanding_principal",
    "interest": "interest_rate",
    "interest rate": "interest_rate",
    "collateral": "collateral_type",
    "collateral value": "collateral_value",
    "prepayment charges": "prepayment_charges",
    "prepayment allowed": "prepayment_allowed",
    "status": "status",
    "closure date": "closure_date",
}

def answer_loan_question(user_id, question):
    user_loans = loan_df[loan_df['user_id'] == user_id]
    if user_loans.empty:
        return "You have no active loans on record."

    lower_q = question.lower()
    choices = list(loan_field_map.keys())
    best_match, score = process.extractOne(lower_q, choices)
    found_field = loan_field_map[best_match] if score > 60 else None
    field_label = best_match.title() if found_field else None

    # Even if no specific field, send all context through LLM
    context = user_loans.to_string(index=False)
    prompt = (
        "You are a helpful banking assistant. "
        "Answer the user's question using only the following loan data:\n\n"
        f"{context}\n\n"
        f"Question: {question}\n\n"
        "Answer in a single, concise and complete sentence."
    )
    return call_ollama(prompt)

session_defaults = {
    "messages": [],
    "step": "awaiting_help_type", 
    "email_requested": False,
    "email": "",
    "user_id": "",
    "account_number": ""
}
for key, value in session_defaults.items():
    if key not in st.session_state:
        st.session_state[key] = value

left_col, right_col = st.columns([1,3])

with st.sidebar:
    st.markdown("## About Voilà Banking Assistant")
    st.markdown("""
Voilà Debt Assistant helps you understand and manage your debts and loans.

- Ask questions related to debt, loans, repayments, refinancing, and more.
- Or, retrieve your account details securely.
- Answers are provided based on a comprehensive debt encyclopedia document or your account records.
- Optionally get summaries emailed to you.

Powered by AI and your debt knowledge base.
""")
    st.markdown('<div class="footer-tag">Designed and powered by AiSolve</div>', unsafe_allow_html=True)

with right_col:
    st.title("💰 Voilà Banking Assistant")
    st.markdown('<div class="chat-container">', unsafe_allow_html=True)
    for msg in st.session_state.messages:
        css_class = "user" if msg["role"] == "user" else "bot"
        st.markdown(f'<div class="stChatMessage {css_class}">{msg["content"]}</div>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    input_col1, input_col2 = st.columns([3,1])
    user_input = None

    with input_col1:
        user_input = st.chat_input("Type your request...")

    with input_col2:
        if st.button("🎤 To Speak "):
            user_input = recognize_once()
            if user_input:
                st.info(f"You said: {user_input}")

    if user_input:
        st.session_state.messages.append({"role": "user", "content": user_input})
        step = st.session_state.step
        response = ""

        if step == "awaiting_help_type":
            response = (
                "Hello! I'm ISOLVE Ai, your Banking assistant. "
                "How can I help you  today?\n\n"
                "- **Account Details**\n- **Financial Question**\n\n"
                "Please type 'account details' or 'financial question'."
            )
            st.session_state.step = "awaiting_help_type_response"
            speak_text(response)

        elif step == "awaiting_help_type_response":
            if "account" in user_input.lower() or "account status" in user_input.lower() or "user" in user_input.lower() or "status" in user_input.lower() or "account details" in user_input.lower():
                response = "Sure! Please enter your User ID."
                speak_text(response)
                st.session_state.step = "awaiting_user_id"
            elif "financial" in user_input.lower() or "loan" in user_input.lower() or "debt" in user_input.lower():
                response = "Great! Please Speak out or type your financial or debt-related question."
                speak_text(response)
                st.session_state.step = "answering_financial"
            else:
                response = "Please specify if you need help with 'account details' or a 'financial question'."
                speak_text(response)

        elif step == "awaiting_user_id":
            user_id = user_input.strip().replace(" ","")
            if is_valid_user_id(user_id):
                st.session_state.user_id = user_id
                name_row = user_df.loc[user_df['user_id'] == user_id, 'name']
                if not name_row.empty:
                    name = name_row.iloc[0]
                else:
                    name = "User"
                response = f"{name}, Please enter your Account Number."
                speak_text(response)
                st.session_state.step = "awaiting_account_number"
            else:
                response = "User ID not found. Please re-enter a valid User ID."
                speak_text(response)

        elif step == "awaiting_account_number":
            account_number = user_input.strip().replace(" ","")
            user_id = st.session_state.user_id
            if is_valid_account_number(user_id, account_number):
                st.session_state.account_number = account_number
                details = get_account_details(user_id, account_number)
                response = f"Here are your account details:\n\n{details}\n\nIf you have any questions about your loans (like EMI, outstanding, status, interest rate, collateral, etc.), please ask now."
                speak_text(response)
                st.session_state.email_requested = True
                st.session_state.step = "awaiting_loan_question"
            else:
                response = "Account number not found for the provided User ID. Please re-enter a valid account number."
                speak_text(response)

        elif step == "awaiting_loan_question":
            user_id = st.session_state.user_id

            if "all detail" in user_input.lower() and "email" in user_input.lower():
                account_number = st.session_state.account_number
                acc_details = get_account_details(user_id, account_number)
                loan_details_df = get_loans_by_user(user_id)
                if not loan_details_df.empty:
                    loan_details = format_loan_details(loan_details_df)
                else:
                    loan_details = "No loan details found."

                email_body = f"**Account Details:**\n{acc_details}\n\n{loan_details}"
                response = "Please provide your email address to receive all your details."
                speak_text(response)
                st.session_state.step = "await_email_all"
                st.session_state.email_all_body = email_body

            else:
                response = answer_loan_question(user_id, user_input)
                response += "\n\nYou can ask More Loan Questions, or ask for All Details to be Emailed."
                speak_text(response)

        elif step == "await_email_all":
            if re.fullmatch(r"[^@\s]+@[^@\s]+\.[a-zA-Z]+", user_input):
                st.session_state.email = user_input
                sent = send_email(st.session_state.email, "Your Debt Assistant Full Details", st.session_state.email_all_body)
                if sent:
                    response = f"✅ All your details were emailed successfully to {st.session_state.email}!"
                    speak_text(response)
                else:
                    response = "Sorry, I couldn't send the email. Please try again later."
                    speak_text(response)
                st.session_state.step = "awaiting_loan_question"
            else:
                response = "That doesn't look like a valid email address. Please enter a valid email."
                speak_text(response)

        elif step == "account_details_done":
            response = "If you want all your details emailed to you, type 'I want all the details to email'."
            speak_text(response)

        elif step == "answering_financial":
            response = answer_from_pdf(user_input)
            if not st.session_state.email_requested:
                response += "\n\nIf you want, you can ask for all your details to be emailed."
                speak_text(response)
                st.session_state.email_requested = True

        if response:
            st.session_state.messages.append({"role": "bot", "content": response})
            st.rerun()