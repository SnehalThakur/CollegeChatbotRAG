from io import BytesIO

import streamlit as st
from PIL import Image
import os
import datetime
import pandas as pd
from dotenv import load_dotenv
from langchain.chains.question_answering import load_qa_chain
from langchain_community.vectorstores import FAISS
from langchain_core.prompts import PromptTemplate
from langchain_google_genai import ChatGoogleGenerativeAI, GoogleGenerativeAIEmbeddings
from langchain.text_splitter import RecursiveCharacterTextSplitter
import google.generativeai as genai

import ExcelDataParserJson
from JsonToTextFile import TimetableProcessor

load_dotenv()
os.getenv("GOOGLE_API_KEY")
genai.configure(api_key=os.getenv("GOOGLE_API_KEY"))
# Set page config to use wide layout
st.set_page_config(
    layout="wide",
    page_title="SAHAYAK - College Assistant",
    page_icon=":mortar_board:"
)


def get_text_chunks(text):
    text_splitter = RecursiveCharacterTextSplitter(chunk_size=10000, chunk_overlap=1000)
    chunks = text_splitter.split_text(text)
    return chunks


def get_vector_store(text_chunks, folder_name):
    embeddings = GoogleGenerativeAIEmbeddings(model="models/embedding-001")

    if not os.path.exists(folder_name):
        vector_store = FAISS.from_texts(text_chunks, embedding=embeddings)
        vector_store.save_local(folder_name)


def get_conversational_chain():
    prompt_template = """
    Answer the question as detailed as possible from the provided context, make sure to provide all the details, if the answer is not in
    provided context just say, "answer is not available in the context", don't provide the wrong answer\n\n
    Context:\n {context}?\n
    Question: \n{question}\n

    Answer:
    """
    model = ChatGoogleGenerativeAI(model="gemini-1.5-flash", temperature=0.3)
    prompt = PromptTemplate(template=prompt_template, input_variables=["context", "question"])
    chain = load_qa_chain(model, chain_type="stuff", prompt=prompt)
    return chain


def getModelResponse(user_question,
                     textFilePath=r"C:\Users\snehal\PycharmProjects\ChatbotRAG\timetable_structured.txt"):
    embeddings = GoogleGenerativeAIEmbeddings(model="models/embedding-001")
    vector_store_folder = f"faiss_index_timetable"

    # try:
    #     new_db = FAISS.load_local(vector_store_folder, embeddings, allow_dangerous_deserialization=True)
    # except (KeyError, FileNotFoundError) as e:
    # st.warning(f"Error loading vector store: {str(e)}. Attempting to recreate the index.")

    raw_text = open(textFilePath).read()
    if raw_text:
        text_chunks = get_text_chunks(raw_text)
        get_vector_store(text_chunks, vector_store_folder)
        new_db = FAISS.load_local(vector_store_folder, embeddings, allow_dangerous_deserialization=True)
    else:
        st.error("Unable to process the PDF. Please check the file path and try again.")
        return

    docs = new_db.similarity_search(user_question)
    chain = get_conversational_chain()
    response = chain({"input_documents": docs, "question": user_question}, return_only_outputs=True)
    print("response['output_text'] - ", response["output_text"])
    return response["output_text"]


# Custom College-Themed CSS
st.markdown("""
<style>
    /* Main Theme Colors */
    :root {
        --primary: #2C3E50;
        --secondary: #E74C3C;
        --accent: #3498DB;
        --light: #ECF0F1;
        --dark: #2C3E50;
        --success: #27AE60;
    }

    /* Main Container Styling */
    .stApp {
        background-color: #f5f7fa;
        background-image: linear-gradient(315deg, #f5f7fa 0%, #e4e8f0 74%);
    }

    /* Header Styling */
    .st-emotion-cache-10trblm {
        color: var(--primary);
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        border-bottom: 2px solid var(--accent);
        padding-bottom: 10px;
    }

    /* Container Borders */
    [data-testid="stVerticalBlock"] > [style*="flex-direction: column"] > [data-testid="stVerticalBlock"] {
        border-radius: 15px;
        box-shadow: 0 4px 8px rgba(0,0,0,0.1);
        background-color: white;
    }

    /* Button Styling */
    .stButton>button {
        border: none;
        border-radius: 8px;
        font-weight: 600;
        transition: all 0.3s ease;
    }

    .stButton>button:first-of-type {
        background-color: var(--accent);
        color: white;
    }

    .stButton>button:first-of-type:hover {
        background-color: #2980B9;
        transform: translateY(-2px);
        box-shadow: 0 4px 8px rgba(52, 152, 219, 0.3);
    }

    /* Chat Bubbles */
    .user-message {
        background-color: var(--accent);
        color: white;
        padding: 12px 16px;
        border-radius: 18px 18px 0 18px;
        margin: 8px 0;
        max-width: 80%;
        margin-left: auto;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }

    .bot-message {
        background-color: var(--light);
        color: var(--dark);
        padding: 12px 16px;
        border-radius: 18px 18px 18px 0;
        margin: 8px 0;
        max-width: 80%;
        margin-right: auto;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        border: 1px solid #ddd;
    }

    /* Notification Cards */
    .notification-card {
        background-color: white;
        border-left: 4px solid var(--secondary);
        padding: 12px;
        margin: 8px 0;
        border-radius: 0 8px 8px 0;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        transition: all 0.3s ease;
    }

    .notification-card:hover {
        transform: translateX(5px);
        box-shadow: 0 4px 8px rgba(0,0,0,0.1);
    }

    /* Input Fields */
    .stTextArea textarea, .stTextInput input {
        border-radius: 8px !important;
        border: 1px solid #ddd !important;
        padding: 10px !important;
    }

    /* Tab Styling */
    .st-bd {
        border-radius: 8px !important;
    }

    /* College Badge */
    .college-badge {
        background-color: var(--primary);
        color: white;
        padding: 8px 16px;
        border-radius: 20px;
        font-weight: bold;
        display: inline-block;
        margin-bottom: 15px;
        font-size: 0.8em;
    }
</style>
""", unsafe_allow_html=True)

# Initialize session state variables
if 'notifications' not in st.session_state:
    st.session_state.notifications = []
if 'chat_history' not in st.session_state:
    st.session_state.chat_history = []
if 'authenticated' not in st.session_state:
    st.session_state.authenticated = False
if 'show_login' not in st.session_state:
    st.session_state.show_login = False

# Load college logo
try:
    st.image(r"C:\Users\snehal\PycharmProjects\ChatbotRAG\clg.png")
except FileNotFoundError:
    st.warning("College logo not found")
    college_logo = None

# Create a three-column layout (20% | 55% | 25%)
col1, col2, col3 = st.columns([2, 5.5, 2.5])

# Left Column - User Profile Panel
with col1:
    with st.container(height=750, border=False):
        st.markdown('<div class="college-badge">YCCE College Portal</div>', unsafe_allow_html=True)

        # Authentication Section
        if not st.session_state.authenticated:
            if st.button("🔑 Sign In / Sign Up",
                         use_container_width=True,
                         type="primary",
                         help="Access all features by signing in"):
                st.session_state.show_login = not st.session_state.show_login
                st.rerun()

            if st.session_state.show_login:
                with st.expander("Account Access", expanded=True):
                    auth_tab1, auth_tab2 = st.tabs(["Sign In", "Sign Up"])

                    with auth_tab1:
                        with st.form("signin_form"):
                            email = st.text_input("Email", placeholder="faculty@ycce.edu")
                            password = st.text_input("Password", type="password")
                            if st.form_submit_button("Sign In", use_container_width=True):
                                if email and password:
                                    st.session_state.authenticated = True
                                    st.session_state.show_login = False
                                    st.rerun()

                    with auth_tab2:
                        with st.form("signup_form"):
                            new_email = st.text_input("College Email", placeholder="faculty@ycce.edu")
                            new_password = st.text_input("Create Password", type="password")
                            confirm_password = st.text_input("Confirm Password", type="password")
                            if st.form_submit_button("Register", use_container_width=True):
                                if new_password == confirm_password and new_email:
                                    st.session_state.authenticated = True
                                    st.session_state.show_login = False
                                    st.rerun()
        else:
            # User Profile Section
            st.markdown("### 👤 Faculty Dashboard")
            # st.markdown("**Department:** Computer Science")
            st.markdown("**ID:** CS2023XYZ")

            st.markdown("---")

            # Quick Actions
            st.markdown("### 🚀 Quick Actions")
            if st.button("📢 Add Announcement",
                         use_container_width=True,
                         help="Post a new college announcement"):
                st.session_state.show_notification = True

            if st.button("📁 Upload Materials",
                         use_container_width=True,
                         help="Share study materials"):
                st.session_state.show_upload = True

            if st.button("📅 View Timetable",
                         use_container_width=True,
                         help="Check your class schedule"):
                st.session_state.show_timetable = True

            if st.button("🚪 Logout",
                         use_container_width=True,
                         type="secondary"):
                st.session_state.authenticated = False
                st.rerun()

# Middle Column - Main Chat Interface
with col2:
    with st.container(height=750, border=False):
        st.markdown("## 💬 SAHAYAK Assistant")
        st.caption("Your 24/7 college helpdesk powered by AI")

        # Chat History Display
        chat_container = st.container(height=500)
        with chat_container:
            for message in st.session_state.chat_history:
                if message.startswith("You:"):
                    st.markdown(f'<div class="user-message">{message}</div>', unsafe_allow_html=True)
                else:
                    st.markdown(f'<div class="bot-message">{message}</div>', unsafe_allow_html=True)

        # Chat Input Area
        with st.form("chat_form", border=False):
            user_input = st.text_area("Type your question:",
                                      placeholder="Ask about timetables, or college services...",
                                      height=100,
                                      key="chat_input")

            c1, c2 = st.columns([4, 1])
            with c1:
                if st.form_submit_button("✉️ Send Message",
                                         use_container_width=True,
                                         help="Get instant help from SAHAYAK"):
                    if user_input:
                        st.session_state.chat_history.append(f"You: {user_input}")

                        # Simulate intelligent response
                        model_response = getModelResponse(user_input)
                        bot_response = f"SAHAYAK: '{model_response}'"
                        # bot_response = f"SAHAYAK: I can help with that! For '{user_input}', please check the college portal or contact your department."
                        st.session_state.chat_history.append(bot_response)
                        st.rerun()
            with c2:
                if st.form_submit_button("🗑️ Clear Chat",
                                         type="secondary",
                                         use_container_width=True):
                    st.session_state.chat_history = []
                    st.rerun()

# Right Column - Notifications and Features
with col3:
    with st.container(height=750, border=False):
        st.markdown("## 🔔 Notifications")

        # Notification Display
        notif_container = st.container(height=300)
        with notif_container:
            if not st.session_state.notifications:
                st.info("No new notifications")
            else:
                for notification in st.session_state.notifications[-5:][::-1]:  # Show last 5
                    st.markdown(f"""
                    <div class="notification-card">
                        <div style="font-weight:600;">📌 New Announcement</div>
                        <div style="font-size:0.9em; margin-top:8px;">{notification}</div>
                        <div style="font-size:0.7em; color:#666; margin-top:8px;">
                            {datetime.datetime.now().strftime("%d %b %Y, %I:%M %p")}
                        </div>
                    </div>
                    """, unsafe_allow_html=True)

        if st.button("Clear All Notifications",
                     use_container_width=True,
                     type="secondary"):
            st.session_state.notifications = []
            st.rerun()

        st.markdown("---")

        # Dynamic Feature Panels (for authenticated users)
        if st.session_state.authenticated:
            # Notification Creation Panel
            if st.session_state.get('show_notification', False):
                with st.expander("📢 Create Announcement", expanded=True):
                    with st.form("announcement_form"):
                        announcement = st.text_area("Announcement text",
                                                    height=150,
                                                    placeholder="Type your official announcement...")
                        if st.form_submit_button("Post Announcement", use_container_width=True):
                            if announcement:
                                st.session_state.notifications.append(announcement)
                                st.session_state.show_notification = False
                                st.success("Announcement posted!")
                                st.rerun()

            # File Upload Panel
            if st.session_state.get('show_upload', False):
                with st.expander("📁 Upload Study Materials", expanded=True):
                    uploaded_file = st.file_uploader("Select file to share",
                                                     type=["xlsx", "xls"],
                                                     accept_multiple_files=False,
                                                     help="Timetables")
                    if uploaded_file:
                        if uploaded_file.type in [
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            "application/vnd.ms-excel"
                        ] or uploaded_file.name.endswith('.xlsx'):
                            st.success(f"✅ Uploaded: {uploaded_file.name}")
                            st.caption(
                                f"File type: {uploaded_file.type.split('/')[-1].upper()} | Size: {uploaded_file.size // 1024} KB")
                            # Saving the uploaded excel file
                            excel_file_path = ExcelDataParserJson.save_uploaded_file(uploaded_file,
                                                                                     save_dir="uploaded_files")
                            output_files = ExcelDataParserJson.excelToJsonConverter(excel_file_path)

                            timetableJsonPath = output_files["timetable"]["filepath"]
                            timetableJsonFilename = output_files["timetable"]["filename"]
                            subjectsJsonPath = output_files["subjects"]["filepath"]
                            subjectsJsonFilename = output_files["subjects"]["filename"]

                            # Converting Excel to timetable and subjects json files
                            processor = TimetableProcessor()
                            processor.load_json(timetableJsonPath, subjectsJsonPath)
                            st.session_state.structured_text = processor.generate_structured_text()
                            print("st.session_state.structured_text -", st.session_state.structured_text)
                            print("timetableJsonPath -", timetableJsonPath)
                            print("timetableJsonFilename -", timetableJsonFilename)
                            print("subjectsJsonPath -", subjectsJsonPath)
                            print("subjectsJsonFilename -", subjectsJsonFilename)
                            st.session_state.timetableTextFilepath = timetableJsonPath.replace(timetableJsonFilename, 'timetable_structured1.txt')
                            print("st.session_state.timetableTextFilepath -", st.session_state.timetableTextFilepath)
                            # Save to text file (or pass directly to your model)
                            with open(st.session_state.timetableTextFilepath, 'w') as f:
                                f.write(st.session_state.structured_text)
                                print("Timetable Text saved")

            # Timetable Panel
            if st.session_state.get('show_timetable', False):
                with st.expander("📅 Current Timetable", expanded=True):
                    # File uploader
                    uploaded_file = r"C:\Users\snehal\PycharmProjects\ChatbotRAG\data\UG CLASS CTECH TT_odd (24-25)_ 3RD-5TH-7TH SEM_TO STAFF_V3_9-08-2024-w.e.f. 12-08-2024.xls"

                    if uploaded_file:
                        try:
                            # Read Excel file
                            df = pd.read_excel(uploaded_file)
                            st.session_state.timetable_data = df
                        except Exception as e:
                            st.error(f"Error reading file: {e}")

                    # Display timetable
                    if st.session_state.timetable_data is not None:
                        st.markdown('<div class="timetable-table">', unsafe_allow_html=True)

                        # Download button
                        excel_buffer = BytesIO()
                        st.session_state.timetable_data.to_excel(excel_buffer, index=False)
                        st.download_button(
                            label="Download Timetable",
                            data=excel_buffer,
                            file_name="college_timetable.xlsx",
                            mime="application/vnd.ms-excel"
                        )
