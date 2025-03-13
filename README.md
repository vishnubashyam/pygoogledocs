# PyGoogleDocs

A Python package for automating Google Docs specifically designed for creating mathematics worksheets.

## Features

- Create, copy, and modify Google Docs documents
- Generate math worksheets with numbered problems
- Create answer keys
- Support for mathematical equations
- Table creation and formatting
- Text styling and formatting
- Image insertion
- Shared folder management

## Requirements

- Python 3.9+
- Google API credentials (service account)
- Required Python packages:
  - google-api-python-client
  - google-auth
  - google-auth-httplib2
  - google-auth-oauthlib

## Usage

```python
from pygoogledocs import get_credentials, get_drive_service, get_docs_service, Document

# Authenticate
creds = get_credentials('path/to/credentials.json')
drive_service = get_drive_service(creds)
docs_service = get_docs_service(creds)

# Create a new document
file_metadata = {
    'name': 'Math Worksheet',
    'mimeType': 'application/vnd.google-apps.document'
}
doc = drive_service.files().create(body=file_metadata, fields='id').execute()
doc_id = doc['id']

# Initialize document
document = Document(docs_service, doc_id)

# Create a worksheet
problems = [
    "Solve for x: 3x + 7 = 22",
    "Simplify: 2(3x - 4) + 5x",
    "Factor: x² - 9"
]

document.create_worksheet("Algebra Skills", problems)
```

## Example

See `docs_demo.py` for a complete example that demonstrates:
- Finding or creating folders
- Creating template documents
- Generating worksheets with problems
- Creating answer keys

## Related Projects

This package is a companion to PyGoogleSlides, which provides similar automation capabilities for Google Slides presentations.