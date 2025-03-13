from google.oauth2 import service_account
from googleapiclient.discovery import build

DEFAULT_SCOPES = [
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/documents'
]

def get_credentials(service_account_file, scopes=None):
    """Return credentials using a service account key file."""
    if scopes is None:
        scopes = DEFAULT_SCOPES
    creds = service_account.Credentials.from_service_account_file(service_account_file, scopes=scopes)
    return creds

def get_docs_service(creds):
    """Return the Google Docs API service."""
    return build('docs', 'v1', credentials=creds)

def get_drive_service(creds):
    """Return the Google Drive API service."""
    return build('drive', 'v3', credentials=creds)