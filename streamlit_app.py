import streamlit as st
import os
import base64
from pygoogledocs import (
    get_credentials,
    get_drive_service,
    get_docs_service,
    find_folder,
    create_folder,
    find_file,
    copy_document,
    Document,
    find_or_create_folder,
    MarkdownFormatter
)

# Configuration
SERVICE_ACCOUNT_FILE = '/Users/vishnubashyam/Documents/Projects/google_docs_packages/Docs/avian-presence-444016-i6-c1147821a882.json'
FOLDER_NAME = 'MCP_Shared'

def main():
    st.set_page_config(page_title="MCP Document Generator", layout="wide")
    
    st.title("MCP Document Generator")
    st.markdown("Create and manage MCP documents using PyGoogleDocs features.")
    
    # Setup authentication
    try:
        creds = get_credentials(SERVICE_ACCOUNT_FILE)
        drive_service = get_drive_service(creds)
        docs_service = get_docs_service(creds)
        
        # Create or find the main folder
        main_folder_id = find_folder(drive_service, FOLDER_NAME)
        if not main_folder_id:
            with st.spinner(f"Creating main folder '{FOLDER_NAME}'..."):
                main_folder = create_folder(drive_service, FOLDER_NAME)
                main_folder_id = main_folder['id']
            st.success(f"Created or connected to main folder: {FOLDER_NAME}")
        
        # Show tabs for different functionality
        tab1, tab2, tab3, tab4 = st.tabs([
            "Basic Document Creation", 
            "Markdown Document", 
            "Math Worksheet Creator",
            "Advanced Features"
        ])
        
        with tab1:
            basic_document_creator(drive_service, docs_service, main_folder_id)
        
        with tab2:
            markdown_document_creator(drive_service, docs_service, main_folder_id)
        
        with tab3:
            math_worksheet_creator(drive_service, docs_service, main_folder_id)
            
        with tab4:
            advanced_features_demo(drive_service, docs_service, main_folder_id)
    
    except Exception as e:
        st.error(f"Error initializing Google API services: {str(e)}")
        st.info("Please check your service account credentials and ensure the Google APIs are enabled.")

def basic_document_creator(drive_service, docs_service, main_folder_id):
    st.header("Basic Document Creator")
    st.markdown("Create a simple document with text and formatting.")
    
    with st.form("basic_document_form"):
        document_name = st.text_input("Document Name", "My Basic Document")
        
        col1, col2 = st.columns(2)
        with col1:
            title = st.text_input("Title", "Sample Document")
            subtitle = st.text_input("Subtitle", "Created with PyGoogleDocs")
        
        with col2:
            make_title_bold = st.checkbox("Bold Title", value=True)
            title_size = st.slider("Title Size (pt)", 10, 36, 18)
        
        content = st.text_area("Document Content", 
            """This is a sample document created using PyGoogleDocs.

You can add multiple paragraphs and the library will handle the formatting.

This demonstrates the basic text insertion and formatting capabilities.
""", height=150)
        
        submitted = st.form_submit_button("Create Document")
        
        if submitted:
            with st.spinner("Creating document... Please wait."):
                try:
                    # Create a blank document
                    file_metadata = {
                        'name': document_name,
                        'mimeType': 'application/vnd.google-apps.document',
                        'parents': [main_folder_id]
                    }
                    file = drive_service.files().create(body=file_metadata, fields='id').execute()
                    document_id = file['id']
                    
                    # Initialize the Document object
                    document = Document(docs_service, document_id)
                    
                    # Get the document to find the first tab ID
                    doc_data = document.fetch(include_tabs_content=True)
                    tab_id = doc_data['tabs'][0]['tabProperties']['tabId']
                    
                    # Insert title with formatting
                    document.insert_text(
                        {'index': 1, 'tabId': tab_id}, 
                        title + "\n", 
                        format_bold=make_title_bold, 
                        format_size=title_size
                    )
                    
                    # Insert subtitle
                    document.insert_text(
                        {'index': len(title) + 2, 'tabId': tab_id}, 
                        subtitle + "\n\n", 
                        format_italic=True
                    )
                    
                    # Insert content
                    document.insert_text(
                        {'index': len(title) + len(subtitle) + 4, 'tabId': tab_id}, 
                        content
                    )
                    
                    show_success_message(document_id)
                    
                except Exception as e:
                    st.error(f"Error creating document: {str(e)}")

def markdown_document_creator(drive_service, docs_service, main_folder_id):
    st.header("Markdown Document Creator")
    st.markdown("""
    Write your document content using Markdown syntax and it will be converted to Google Docs formatting.
    
    **Supported Markdown Features:**
    - **Bold** text: `**bold**`
    - *Italic* text: `*italic*`
    - # Headers (levels 1-6): `# Header`
    - Lists:
      * Unordered: `* item`
      * Ordered: `1. item`
    - [Links](https://example.com): `[text](url)`
    - Tables: `| Header | Header |`
    - Code: `` `code` ``
    """)
    
    with st.form("markdown_document_form"):
        document_name = st.text_input("Document Name", "My Markdown Document", key="md_doc_name")
        
        markdown_text = st.text_area("Markdown Content", 
        """# My Markdown Document

## Introduction
This document demonstrates the **markdown formatting capabilities** of PyGoogleDocs.

You can create *italic text* and **bold text** easily with markdown.

## Features
* Simple formatting
* Lists
* Tables
* Headers

## Example Code
Here's some `inline code` to demonstrate code formatting.

## Example Table
| Name | Value |
|------|-------|
| Item 1 | 100 |
| Item 2 | 200 |

## Links
[Visit Google](https://google.com)

1. First ordered item
2. Second ordered item
3. Third ordered item
""", height=400)
        
        submitted = st.form_submit_button("Create Markdown Document")
        
        if submitted:
            with st.spinner("Creating markdown document... Please wait."):
                try:
                    # Create a blank document
                    file_metadata = {
                        'name': document_name,
                        'mimeType': 'application/vnd.google-apps.document',
                        'parents': [main_folder_id]
                    }
                    file = drive_service.files().create(body=file_metadata, fields='id').execute()
                    document_id = file['id']
                    
                    # Initialize the Document object
                    document = Document(docs_service, document_id)
                    
                    # Get the document to find the first tab ID
                    doc_data = document.fetch(include_tabs_content=True)
                    tab_id = doc_data['tabs'][0]['tabProperties']['tabId']
                    
                    # Insert markdown content
                    document.insert_markdown(tab_id, 1, markdown_text)
                    
                    show_success_message(document_id)
                    
                except Exception as e:
                    st.error(f"Error creating markdown document: {str(e)}")

def math_worksheet_creator(drive_service, docs_service, main_folder_id):
    st.header("Math Worksheet Creator")
    st.markdown("Create customized math worksheets with problems and answer sheets.")
    
    with st.form("math_worksheet_form"):
        document_name = st.text_input("Document Name", "Algebra Practice Worksheet", key="math_doc_name")
        
        col1, col2 = st.columns(2)
        with col1:
            grade = st.text_input("Grade Level", "Grade 7", key="math_grade")
            unit = st.text_input("Unit", "Unit 1: Patterns and Algebraic Thinking", key="math_unit")
        
        with col2:
            create_answer_key = st.checkbox("Create Answer Key", value=True)
            num_problems = st.slider("Number of Problems", 1, 10, 5)
        
        directions = st.text_area("Directions", 
            "Solve the following problems. Show your work.", height=60, key="math_directions")
        
        # Default problems and answers
        default_problems = "\n".join([f"{i+1}. Solve for x: {3*i+2}x + {i} = {7*i+10}" for i in range(num_problems)])
        default_answers = "\n".join([f"{i+1}. x = {i+2}" for i in range(num_problems)])
        
        col1, col2 = st.columns(2)
        with col1:
            problems = st.text_area("Problems", default_problems, height=300, key="math_problems")
        
        with col2:
            if create_answer_key:
                answers = st.text_area("Answers", default_answers, height=300, key="math_answers")
        
        submitted = st.form_submit_button("Create Math Worksheet")
        
        if submitted:
            with st.spinner("Creating math worksheet... Please wait."):
                try:
                    # Create math worksheet folder if it doesn't exist
                    worksheet_folder = find_or_create_folder(drive_service, "Math Worksheets", main_folder_id)
                    worksheet_folder_id = worksheet_folder['id']
                    
                    # Create a blank document
                    file_metadata = {
                        'name': document_name,
                        'mimeType': 'application/vnd.google-apps.document',
                        'parents': [worksheet_folder_id]
                    }
                    file = drive_service.files().create(body=file_metadata, fields='id').execute()
                    document_id = file['id']
                    
                    # Initialize the Document object
                    document = Document(docs_service, document_id)
                    
                    # Format the problem list as a list of lines
                    problem_list = problems.strip().split('\n')
                    
                    # Create worksheet with header and problems
                    document.create_worksheet(
                        title=document_name, 
                        problems=problem_list
                    )
                    
                    # Create answer key if requested
                    if create_answer_key:
                        answer_list = answers.strip().split('\n')
                        problem_ids = [f"Problem {i+1}" for i in range(len(problem_list))]
                        
                        document.generate_answer_sheet(
                            title=f"Answer Key: {document_name}",
                            problems=problem_ids,
                            answers=answer_list
                        )
                    
                    show_success_message(document_id)
                    
                except Exception as e:
                    st.error(f"Error creating math worksheet: {str(e)}")

def advanced_features_demo(drive_service, docs_service, main_folder_id):
    st.header("Advanced Features Demo")
    
    # Tab navigation within the advanced features section
    adv_tab1, adv_tab2 = st.tabs(["Tables and Images", "Document Organization"])
    
    with adv_tab1:
        st.subheader("Tables and Images")
        
        with st.form("tables_images_form"):
            document_name = st.text_input("Document Name", "Advanced Features Demo", key="adv_doc_name")
            
            # Table section
            st.markdown("### Table Settings")
            col1, col2 = st.columns(2)
            with col1:
                rows = st.number_input("Number of Rows", 1, 10, 3)
                include_headers = st.checkbox("Include Headers", value=True)
            
            with col2:
                cols = st.number_input("Number of Columns", 1, 5, 2)
                headers = st.text_input("Headers (comma-separated)", "Name,Value")
            
            # Image section
            st.markdown("### Image Settings")
            image_url = st.text_input(
                "Image URL (must be publicly accessible)", 
                "https://upload.wikimedia.org/wikipedia/commons/thumb/5/53/Google_%22G%22_Logo.svg/150px-Google_%22G%22_Logo.svg.png"
            )
            
            submitted = st.form_submit_button("Create Document")
            
            if submitted:
                with st.spinner("Creating document with tables and images... Please wait."):
                    try:
                        # Create advanced folder if it doesn't exist
                        advanced_folder = find_or_create_folder(drive_service, "Advanced Examples", main_folder_id)
                        advanced_folder_id = advanced_folder['id']
                        
                        # Create a blank document
                        file_metadata = {
                            'name': document_name,
                            'mimeType': 'application/vnd.google-apps.document',
                            'parents': [advanced_folder_id]
                        }
                        file = drive_service.files().create(body=file_metadata, fields='id').execute()
                        document_id = file['id']
                        
                        # Initialize the Document object
                        document = Document(docs_service, document_id)
                        
                        # Add a title
                        document.create_header("Advanced Features Demo", level=1)
                        
                        # Add a subtitle for the table section
                        document.create_header("Table Example", level=2)
                        
                        # Parse headers
                        header_list = headers.split(',') if headers and include_headers else None
                        
                        # Add a table
                        document.create_table(rows=rows, cols=cols, headers=header_list)
                        
                        # Add a subtitle for the image section
                        document.create_header("Image Example", level=2)
                        
                        # Get the document to find the first tab ID
                        doc_data = document.fetch(include_tabs_content=True)
                        tab_id = doc_data['tabs'][0]['tabProperties']['tabId']
                        
                        # Add descriptive text at the end of the document
                        document.append_text(tab_id, "The image below is inserted using the insert_image method:\n\n")
                        
                        # Add an image
                        document.insert_image(uri=image_url)
                        
                        show_success_message(document_id)
                        
                    except Exception as e:
                        st.error(f"Error creating advanced document: {str(e)}")
    
    with adv_tab2:
        st.subheader("Document Organization")
        
        with st.form("document_organization_form"):
            folder_name = st.text_input("New Folder Name", "My Organized Documents")
            
            st.markdown("### Create Document in Folder")
            document_name = st.text_input("Document Name", "Organized Document", key="org_doc_name")
            content = st.text_area("Document Content", "This document is organized in a folder structure.", height=100)
            
            submitted = st.form_submit_button("Create Folder and Document")
            
            if submitted:
                with st.spinner("Creating folder and document... Please wait."):
                    try:
                        # Create the folder within the main folder
                        new_folder = create_folder(drive_service, folder_name, main_folder_id)
                        new_folder_id = new_folder['id']
                        
                        # Create a document in the new folder
                        file_metadata = {
                            'name': document_name,
                            'mimeType': 'application/vnd.google-apps.document',
                            'parents': [new_folder_id]
                        }
                        file = drive_service.files().create(body=file_metadata, fields='id').execute()
                        document_id = file['id']
                        
                        # Initialize the Document object
                        document = Document(docs_service, document_id)
                        
                        # Get the document to find the first tab ID
                        doc_data = document.fetch(include_tabs_content=True)
                        tab_id = doc_data['tabs'][0]['tabProperties']['tabId']
                        
                        # Insert content
                        document.insert_text({'index': 1, 'tabId': tab_id}, content)
                        
                        # Success message
                        st.success(f"✅ Created folder '{folder_name}' and document '{document_name}'")
                        doc_url = f"https://docs.google.com/document/d/{document_id}/edit"
                        st.markdown(f"[View document]({doc_url})", unsafe_allow_html=True)
                        
                    except Exception as e:
                        st.error(f"Error creating folder and document: {str(e)}")

def show_success_message(document_id):
    st.success(f"✅ Document created successfully!")
    doc_url = f"https://docs.google.com/document/d/{document_id}/edit"
    st.info(f"Document ID: {document_id}")
    st.markdown(f"[Click here to view and edit the document]({doc_url})", unsafe_allow_html=True)
    
    # Create a button that opens the document in a new tab
    js = f"""
    <script>
        function open_doc() {{
            window.open('{doc_url}', '_blank').focus();
        }}
    </script>
    <button
        style="background-color:#4CAF50;color:white;padding:10px 24px;border:none;border-radius:4px;cursor:pointer;"
        type="button"
        onclick="open_doc()"
    >
        Open Document in Google Docs
    </button>
    """
    st.components.v1.html(js, height=50)

if __name__ == "__main__":
    main()