#!/usr/bin/env python
import os
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
NEW_DOCUMENT_NAME = 'Algebra Patterns Inquiry 1.2'

# Template field values
TEMPLATE_FIELDS = {
    "{Lesson_num_name}": "Lesson 1.2: Exploring Patterns in Algebra",
    "{Directions}": "Work in groups to solve the following problems. Discuss your approach with your team members and be prepared to share your findings with the class.",
    "{Worksheet_contents}": """
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
    "{Grade}": "Grade 7",
    "{Unit}": "Unit 1: Patterns and Algebraic Thinking"
}

# No answer key is needed

def main():
    # Authenticate using service account credentials
    creds = get_credentials(SERVICE_ACCOUNT_FILE)
    drive_service = get_drive_service(creds)
    docs_service = get_docs_service(creds)

    # Locate the main folder
    main_folder_id = find_folder(drive_service, FOLDER_NAME)
    if not main_folder_id:
        print(f"Folder '{FOLDER_NAME}' not found. Creating folder...")
        main_folder = create_folder(drive_service, FOLDER_NAME)
        main_folder_id = main_folder['id']
        print(f"Created main folder with ID: {main_folder_id}")
    
    # Create or find the Inquiry Activities subfolder
    inquiry_folder = find_or_create_folder(drive_service, "Inquiry Activities", main_folder_id)
    inquiry_folder_id = inquiry_folder['id']
    print(f"Using Inquiry Activities folder with ID: {inquiry_folder_id}")

    # Check if a template document exists, if not create a blank one
    template_id = find_file(drive_service, TEMPLATE_NAME, main_folder_id)
    if not template_id:
        print(f"Template '{TEMPLATE_NAME}' not found. Creating a new template...")
        # Create a new blank document to serve as template
        file_metadata = {
            'name': TEMPLATE_NAME,
            'mimeType': 'application/vnd.google-apps.document',
            'parents': [folder_id]
        }
        file = drive_service.files().create(body=file_metadata, fields='id').execute()
        template_id = file['id']
        print(f"Created template document with ID: {template_id}")
        
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
        
        print("Template document initialized with placeholders.")

    # Create a copy of the template document in the Inquiry Activities folder
    new_document = copy_document(drive_service, template_id, NEW_DOCUMENT_NAME, inquiry_folder_id)
    new_document_id = new_document['id']
    print(f"Created new document with ID: {new_document_id}")

    # Initialize the Document object
    document = Document(docs_service, new_document_id)

    # Replace all the template placeholders
    for placeholder, replacement in TEMPLATE_FIELDS.items():
        if placeholder == "{Lesson_num_name}":
            document.replace_text(placeholder, replacement, format_bold=True, format_size=16)
        elif placeholder == "{Worksheet_contents}":
            document.replace_text(placeholder, replacement)
        else:
            document.replace_text(placeholder, replacement)
    
    print(f'New document created with ID: {new_document_id}')
    print(f'Document saved to the "Inquiry Activities" folder ({inquiry_folder_id})')
    print('Note: These documents are only accessible to the service account.')

if __name__ == '__main__':
    main()