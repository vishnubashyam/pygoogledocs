def find_folder(drive_service, folder_name, parent_folder_id=None, return_all=False):
    """
    Find a folder by name, optionally within a parent folder.
    
    Args:
        drive_service: Google Drive API service instance
        folder_name: Name of the folder to find
        parent_folder_id: ID of the parent folder (optional)
        return_all: Whether to return all matching folders or just the first one (default: False)
        
    Returns:
        If return_all is False: Folder ID if found, None otherwise
        If return_all is True: List of folder dictionaries (id, name) if found, empty list otherwise
    """
    query = f"mimeType='application/vnd.google-apps.folder' and name='{folder_name}' and trashed=false"
    if parent_folder_id:
        query += f" and '{parent_folder_id}' in parents"
    
    results = drive_service.files().list(
        q=query,
        spaces='drive',
        fields='files(id, name)'
    ).execute()
    folders = results.get('files', [])
    
    if return_all:
        return folders
    else:
        return folders[0]['id'] if folders else None

def create_folder(drive_service, folder_name, parent_folder_id=None):
    """
    Create a new folder in Google Drive, optionally within a parent folder.
    If a folder with the same name already exists in the specified parent, it will return that folder instead.
    
    Args:
        drive_service: Google Drive API service instance
        folder_name: Name of the folder to create
        parent_folder_id: ID of the parent folder (optional)
        
    Returns:
        Dictionary containing folder metadata
    """
    # First check if folder already exists in the specified parent
    existing_folders = find_folder(drive_service, folder_name, parent_folder_id, return_all=True)
    
    if existing_folders:
        # If multiple folders exist with the same name, log a warning
        if len(existing_folders) > 1:
            print(f"Warning: Multiple folders named '{folder_name}' found in the same location.")
            print(f"Using the first folder found: {existing_folders[0]['id']}")
        
        # Return the first existing folder
        return drive_service.files().get(fileId=existing_folders[0]['id'], fields='id,name').execute()
    
    # Create new folder if none exists
    file_metadata = {
        'name': folder_name,
        'mimeType': 'application/vnd.google-apps.folder'
    }
    
    if parent_folder_id:
        file_metadata['parents'] = [parent_folder_id]
    
    return drive_service.files().create(body=file_metadata, fields='id,name').execute()

def find_file(drive_service, file_name, parent_folder_id=None):
    """
    Find a file by name, optionally within a parent folder.
    
    Args:
        drive_service: Google Drive API service instance
        file_name: Name of the file to find
        parent_folder_id: ID of the parent folder (optional)
        
    Returns:
        File ID if found, None otherwise
    """
    query = f"name='{file_name}' and trashed=false"
    if parent_folder_id:
        query += f" and '{parent_folder_id}' in parents"
    results = drive_service.files().list(
        q=query,
        spaces='drive',
        fields='files(id, name)'
    ).execute()
    files = results.get('files', [])
    return files[0]['id'] if files else None

def delete_file(drive_service, file_id):
    """
    Delete a file from Google Drive.
    
    Args:
        drive_service: Google Drive API service instance
        file_id: ID of the file to delete
        
    Returns:
        None
    """
    drive_service.files().delete(fileId=file_id).execute()

def rename_file(drive_service, file_id, new_name):
    """
    Rename a file in Google Drive.
    
    Args:
        drive_service: Google Drive API service instance
        file_id: ID of the file to rename
        new_name: New name for the file
        
    Returns:
        Updated file metadata
    """
    file_metadata = {'name': new_name}
    return drive_service.files().update(
        fileId=file_id,
        body=file_metadata,
        fields='id,name'
    ).execute()

def copy_document(drive_service, template_id, new_name, parents, overwrite=False):
    """
    Create a copy of a document, optionally overwriting an existing file with the same name.
    
    Args:
        drive_service: Google Drive API service instance
        template_id: ID of the template document to copy
        new_name: Name for the new document
        parents: ID or list of IDs of parent folders
        overwrite: Whether to overwrite an existing file with the same name (default: False)
        
    Returns:
        Dictionary containing created document metadata
    """
    # Convert parents to list if it's a single ID
    parent_folders = parents if isinstance(parents, list) else [parents]
    
    # Check if a file with the same name exists in any of the parent folders
    for parent_id in parent_folders:
        existing_file_id = find_file(drive_service, new_name, parent_id)
        if existing_file_id and overwrite:
            # Delete the existing file if overwrite is True
            delete_file(drive_service, existing_file_id)
    
    # Create the copy
    body = {
        'name': new_name,
        'parents': parent_folders
    }
    return drive_service.files().copy(fileId=template_id, body=body).execute()

def move_file(drive_service, file_id, new_parent_id, remove_parents=None):
    """
    Move a file to a different folder in Google Drive.
    
    Args:
        drive_service: Google Drive API service instance
        file_id: ID of the file to move
        new_parent_id: ID of the destination folder
        remove_parents: ID of the parent folder to remove (optional)
        
    Returns:
        Updated file metadata
    """
    # Build the request
    if remove_parents:
        return drive_service.files().update(
            fileId=file_id,
            addParents=new_parent_id,
            removeParents=remove_parents,
            fields='id,name,parents'
        ).execute()
    else:
        return drive_service.files().update(
            fileId=file_id,
            addParents=new_parent_id,
            fields='id,name,parents'
        ).execute()

def find_or_create_folder(drive_service, folder_name, parent_folder_id=None):
    """
    Find a folder by name or create it if it doesn't exist.
    This function handles the case of multiple folders with the same name
    by consolidating their contents into a single folder.
    
    Args:
        drive_service: Google Drive API service instance
        folder_name: Name of the folder to find or create
        parent_folder_id: ID of the parent folder (optional)
        
    Returns:
        Dictionary containing folder metadata
    """
    # Find all folders with this name in the parent
    existing_folders = find_folder(drive_service, folder_name, parent_folder_id, return_all=True)
    
    if not existing_folders:
        # No folder exists, create a new one
        file_metadata = {
            'name': folder_name,
            'mimeType': 'application/vnd.google-apps.folder'
        }
        if parent_folder_id:
            file_metadata['parents'] = [parent_folder_id]
        
        return drive_service.files().create(body=file_metadata, fields='id,name').execute()
    
    elif len(existing_folders) == 1:
        # Only one folder exists, return it
        return drive_service.files().get(fileId=existing_folders[0]['id'], fields='id,name').execute()
    
    else:
        # Multiple folders exist, consolidate them
        print(f"Found {len(existing_folders)} folders named '{folder_name}'. Consolidating...")
        
        # Use the first folder as the target
        target_folder = existing_folders[0]
        target_folder_id = target_folder['id']
        
        # Move all files from other folders to the target folder and delete the empty folders
        for folder in existing_folders[1:]:
            source_folder_id = folder['id']
            
            # Find all files in the source folder
            results = drive_service.files().list(
                q=f"'{source_folder_id}' in parents and trashed=false",
                spaces='drive',
                fields='files(id, name)'
            ).execute()
            
            files_to_move = results.get('files', [])
            
            # Move each file to the target folder
            for file_item in files_to_move:
                move_file(drive_service, file_item['id'], target_folder_id, source_folder_id)
                print(f"Moved file '{file_item['name']}' to the target folder")
            
            # Delete the now-empty source folder
            try:
                drive_service.files().delete(fileId=source_folder_id).execute()
                print(f"Deleted empty folder: {folder['name']} (ID: {source_folder_id})")
            except Exception as e:
                print(f"Could not delete folder {folder['name']} (ID: {source_folder_id}): {str(e)}")
        
        return drive_service.files().get(fileId=target_folder_id, fields='id,name').execute()