"""
PyGoogleDocs - A Python package for automating Google Docs for math worksheets.
"""

from .auth import get_credentials, get_drive_service, get_docs_service
from .drive import (
    find_folder, 
    create_folder, 
    find_file, 
    delete_file, 
    rename_file, 
    copy_document, 
    move_file,
    find_or_create_folder
)
from .document import Document
from .markdown import MarkdownFormatter

__version__ = "0.1.0"
__author__ = "Vishnu Bashyam"

__all__ = [
    'get_credentials',
    'get_drive_service',
    'get_docs_service',
    'find_folder',
    'create_folder',
    'find_file',
    'delete_file',
    'rename_file',
    'copy_document',
    'move_file',
    'find_or_create_folder',
    'Document',
    'MarkdownFormatter',
]