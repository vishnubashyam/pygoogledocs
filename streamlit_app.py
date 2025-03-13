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
    find_or_create_folder
)

# Configuration
SERVICE_ACCOUNT_FILE = '/Users/vishnubashyam/Documents/Projects/Google Slides test/Slides/avian-presence-444016-i6-c1147821a882.json'
FOLDER_NAME = 'MCP_Shared'
TEMPLATE_NAME = 'Inquiry Activity Template'

# Main Streamlit app
def main():
    st.set_page_config(page_title="Inquiry Activity Generator", layout="centered")
    
    st.title("Inquiry Activity Generator")
    st.markdown("Create customized inquiry activities using Google Docs.")
    
    # Form for document information
    with st.form("doc_info_form"):
        st.subheader("Document Information")
        
        # Basic document information
        col1, col2 = st.columns(2)
        with col1:
            document_name = st.text_input("Document Name", "Algebra Patterns Inquiry 1.2")
        with col2:
            lesson_num_name = st.text_input("Lesson Number and Name", "Lesson 1.2: Exploring Patterns in Algebra")
        
        # Grade and unit information
        col1, col2 = st.columns(2)
        with col1:
            grade = st.text_input("Grade Level", "Grade 7")
        with col2:
            unit = st.text_input("Unit", "Unit 1: Patterns and Algebraic Thinking")
            
        # Directions
        directions = st.text_area("Directions", 
            "Work in groups to solve the following problems. Discuss your approach with your team members and be prepared to share your findings with the class.",
            height=100)
        
        # Worksheet content
        worksheet_contents = st.text_area("Worksheet Contents", 
            """
1. Pattern Recognition:
   Look at the sequence: 3, 7, 11, 15, 19, ...
   a) What pattern do you notice?
   b) What would be the next three numbers in this sequence?
   c) Write an expression for the nth term in this sequence.

2. Table Completion:
   Complete the table below and identify the relationship between x and y.
   
   | x | y |
   |---|---|
   | 1 | 3 |
   | 2 | 5 |
   | 3 | 7 |
   | 4 | ? |
   | 5 | ? |
   
   a) What is the pattern?
   b) Write an equation relating x and y.
   c) What would y equal when x = 10?

3. Problem Solving:
   A rectangular garden has a length that is 3 meters more than its width. If the perimeter of the garden is 22 meters:
   a) Write an equation to represent this situation.
   b) Solve for the dimensions of the garden.
   c) What is the area of the garden?

4. Visual Pattern:
   Consider the pattern of squares below:
   
   Pattern 1: □
   Pattern 2: □□
                □□
   Pattern 3: □□□
                □□□
                □□□
   
   a) How many squares would be in Pattern 4?
   b) How many squares would be in Pattern 10?
   c) Write a formula for the number of squares in Pattern n.
            """,
            height=400)
        
        
        submitted = st.form_submit_button("Generate Document")
        
        if submitted:
            template_fields = {
                "{Lesson_num_name}": lesson_num_name,
                "{Directions}": directions,
                "{Worksheet_contents}": worksheet_contents,
                "{Grade}": grade,
                "{Unit}": unit
            }
            
            # Show a spinner while generating the document
            with st.spinner("Generating document... Please wait."):
                document_id = generate_document(document_name, template_fields)
                
            # Success message
            if document_id:
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
            else:
                st.error("❌ Error creating document. Please check the logs.")

def generate_document(document_name, template_fields):
    try:
        # Authenticate using service account credentials
        creds = get_credentials(SERVICE_ACCOUNT_FILE)
        drive_service = get_drive_service(creds)
        docs_service = get_docs_service(creds)

        # Locate the main folder
        main_folder_id = find_folder(drive_service, FOLDER_NAME)
        if not main_folder_id:
            st.info(f"Main folder '{FOLDER_NAME}' not found. Creating folder...")
            main_folder = create_folder(drive_service, FOLDER_NAME)
            main_folder_id = main_folder['id']
        
        # Create or find the Inquiry Activities subfolder
        inquiry_folder = find_or_create_folder(drive_service, "Inquiry Activities", main_folder_id)
        inquiry_folder_id = inquiry_folder['id']
        
        # Check if a template document exists, if not create a blank one
        template_id = find_file(drive_service, TEMPLATE_NAME, main_folder_id)
        if not template_id:
            st.info(f"Template '{TEMPLATE_NAME}' not found. Creating a new template...")
            # Create a new blank document to serve as template
            file_metadata = {
                'name': TEMPLATE_NAME,
                'mimeType': 'application/vnd.google-apps.document',
                'parents': [main_folder_id]
            }
            file = drive_service.files().create(body=file_metadata, fields='id').execute()
            template_id = file['id']
            
            # Initialize the Document object for the template
            template_doc = Document(docs_service, template_id)
            
            # Add a title to the template
            template_doc.create_header("Inquiry Activity", level=1)
            
            # Add template placeholders
            content = """
{Lesson_num_name}

{Grade}
{Unit}

Directions:
{Directions}

{Worksheet_contents}
"""
            # Insert the template content
            template_doc.insert_text({'index': 1}, content)

        # Create a copy of the template document in the Inquiry Activities folder
        new_document = copy_document(drive_service, template_id, document_name, inquiry_folder_id)
        new_document_id = new_document['id']

        # Initialize the Document object
        document = Document(docs_service, new_document_id)

        # Replace all the template placeholders
        for placeholder, replacement in template_fields.items():
            document.replace_text(placeholder, replacement)
        
        # Return the document ID
        return new_document_id
        
    except Exception as e:
        st.error(f"Error: {str(e)}")
        return None

if __name__ == "__main__":
    main()