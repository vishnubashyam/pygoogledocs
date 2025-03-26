class Document:
    # Google Docs API bullet preset constants
    BULLET_DISC_CIRCLE_SQUARE = "BULLET_DISC_CIRCLE_SQUARE"
    BULLET_DIAMONDX_ARROW3D_SQUARE = "BULLET_DIAMONDX_ARROW3D_SQUARE"
    BULLET_CHECKBOX = "BULLET_CHECKBOX"
    BULLET_ARROW_DIAMOND_DISC = "BULLET_ARROW_DIAMOND_DISC"
    BULLET_STAR_CIRCLE_SQUARE = "BULLET_STAR_CIRCLE_SQUARE"
    BULLET_ARROW3D_CIRCLE_SQUARE = "BULLET_ARROW3D_CIRCLE_SQUARE"
    BULLET_LEFTTRIANGLE_DIAMOND_DISC = "BULLET_LEFTTRIANGLE_DIAMOND_DISC"
    BULLET_DIAMONDX_HOLLOWDIAMOND_SQUARE = "BULLET_DIAMONDX_HOLLOWDIAMOND_SQUARE"
    BULLET_DIAMOND_CIRCLE_SQUARE = "BULLET_DIAMOND_CIRCLE_SQUARE"
    NUMBERED_DECIMAL_ALPHA_ROMAN = "NUMBERED_DECIMAL_ALPHA_ROMAN"
    NUMBERED_DECIMAL_ALPHA_ROMAN_PARENS = "NUMBERED_DECIMAL_ALPHA_ROMAN_PARENS"
    NUMBERED_DECIMAL_NESTED = "NUMBERED_DECIMAL_NESTED"
    NUMBERED_UPPERALPHA_ALPHA_ROMAN = "NUMBERED_UPPERALPHA_ALPHA_ROMAN"
    NUMBERED_UPPERROMAN_UPPERALPHA_DECIMAL = "NUMBERED_UPPERROMAN_UPPERALPHA_DECIMAL"
    NUMBERED_ZERODECIMAL_ALPHA_ROMAN = "NUMBERED_ZERODECIMAL_ALPHA_ROMAN"
    
    def __init__(self, docs_service, document_id):
        self.docs_service = docs_service
        self.document_id = document_id
        self.last_index = None  # Track the last insertion point

    def fetch(self, include_tabs_content=True):
        """
        Fetch and return the document's full JSON structure.
        
        Args:
            include_tabs_content: Whether to include content from all tabs in the response
            
        Returns:
            dict: The document's JSON structure
        """
        doc = self.docs_service.documents().get(
            documentId=self.document_id,
            includeTabsContent=include_tabs_content
        ).execute()
        
        # Update last_index based on document content
        if include_tabs_content and 'tabs' in doc and doc['tabs']:
            tab = doc['tabs'][0]  # Get first tab
            if 'documentTab' in tab and 'body' in tab['documentTab']:
                body = tab['documentTab']['body']
                if 'content' in body and body['content']:
                    # Find the last element in the content array
                    last_element = body['content'][-1]
                    self.last_index = last_element.get('endIndex', 1)
        
        return doc

    def batch_update(self, requests):
        """
        Execute a batchUpdate request on the document.
        
        Args:
            requests: List of request objects to execute
            
        Returns:
            dict: The response from the API
        """
        body = {'requests': requests}
        response = self.docs_service.documents().batchUpdate(
            documentId=self.document_id,
            body=body
        ).execute()
        
        # After a batch update, refresh document to update last_index
        self.fetch()
        
        return response
        
    def get_end_index(self, tab_id=None):
        """
        Get the current end index of the document or specified tab.
        If document hasn't been fetched yet, fetch it first.
        
        Args:
            tab_id: Optional ID of the tab to get end index for
            
        Returns:
            int: The end index position
        """
        if self.last_index is None:
            self.fetch()
            
        # If still None after fetching, default to index 1
        if self.last_index is None:
            self.last_index = 1
            
        return self.last_index
        
    def append_text(self, tab_id, text, format_bold=False, format_italic=False, format_size=None, format_color=None):
        """
        Append text at the end of the document.
        
        Args:
            tab_id: ID of the tab to insert into
            text: Text to append
            format_bold: Whether to format the text as bold
            format_italic: Whether to format the text as italic
            format_size: Point size for the text (optional)
            format_color: Color for the text in RGB format (optional)
            
        Returns:
            dict: Response from the API
        """
        # Get the current end index
        end_index = self.get_end_index(tab_id)
        
        # Insert text at the end
        return self.insert_text(
            {'index': end_index - 1, 'tabId': tab_id},
            text,
            format_bold,
            format_italic,
            format_size,
            format_color
        )
        
    def insert_markdown(self, tab_id, index, markdown_text):
        """
        Insert text with markdown formatting at the specified location.
        
        Args:
            tab_id: ID of the tab to insert into
            index: Position in the document to insert the text
            markdown_text: Text with markdown formatting to insert
            
        Returns:
            dict: Response from the API
        """
        from .markdown import MarkdownFormatter
        
        formatter = MarkdownFormatter()
        requests = formatter.convert_to_doc_requests(markdown_text, index)
        
        # Add tab ID to all requests that have a location or range
        for request in requests:
            for key in request:
                if 'location' in request[key]:
                    request[key]['location']['tabId'] = tab_id
                if 'range' in request[key]:
                    request[key]['range']['tabId'] = tab_id
        
        return self.batch_update(requests)

    def replace_text(self, placeholder, replacement, format_bold=False, format_italic=False, format_size=None, format_color=None):
        """
        Replace all occurrences of placeholder text with replacement text.
        
        Args:
            placeholder (str): The text to search for and replace
            replacement (str): The text to replace with
            format_bold (bool): Whether to format the replaced text as bold
            format_italic (bool): Whether to format the replaced text as italic
            format_size (int): Point size for the text (optional)
            format_color (dict): Color for the text in RGB format: {'red': r, 'green': g, 'blue': b} (optional)
        
        Returns:
            dict: Response from the API
        """
        # Find all instances of the placeholder
        doc = self.fetch()
        
        # Build replacement request
        requests = [
            {
                'replaceAllText': {
                    'containsText': {
                        'text': placeholder,
                        'matchCase': True
                    },
                    'replaceText': replacement
                }
            }
        ]
        
        response = self.batch_update(requests)
        
        # If we need to update formatting and text was replaced
        if response.get('replies') and (format_bold or format_italic or format_size or format_color):
            # Get all the locations where text was replaced
            for reply in response.get('replies', []):
                if 'replaceAllText' in reply:
                    replace_info = reply['replaceAllText']
                    if replace_info.get('occurrencesChanged', 0) > 0:
                        # For each location, apply formatting
                        self._apply_formatting_to_replaced_text(doc, placeholder, replacement, 
                                                               format_bold, format_italic, 
                                                               format_size, format_color)
        
        return response
    
    def _apply_formatting_to_replaced_text(self, doc, placeholder, replacement, 
                                           format_bold, format_italic, format_size, format_color):
        """Apply formatting to the newly inserted text."""
        # We need to search the document again to find the ranges of the replaced text
        content = doc.get('body', {}).get('content', [])
        format_requests = []
        
        # Build the formatting request
        style = {}
        update_fields = []
        
        if format_bold:
            style['bold'] = True
            update_fields.append('bold')
        
        if format_italic:
            style['italic'] = True
            update_fields.append('italic')
        
        if format_size:
            style['fontSize'] = {'magnitude': format_size, 'unit': 'PT'}
            update_fields.append('fontSize')
        
        if format_color:
            style['foregroundColor'] = {'color': format_color}
            update_fields.append('foregroundColor')
        
        if update_fields:
            # Find all instances where the text was just replaced
            # This is an approximation since we can't directly get the locations
            # Use search API to find the locations precisely in a real implementation
            format_requests.append({
                'updateTextStyle': {
                    'range': {
                        'segmentId': '',  # Use the "name-only" segment for the main document
                        'startIndex': 1,  # Placeholder, would be determined from actual API response
                        'endIndex': len(replacement) + 1  # Placeholder
                    },
                    'textStyle': style,
                    'fields': ','.join(update_fields)
                }
            })
            
            self.batch_update(format_requests)

    def insert_text(self, location, text, format_bold=False, format_italic=False, format_size=None, format_color=None):
        """
        Insert text at a specific location in the document.
        
        Args:
            location (dict): Location object specifying where to insert text
            text (str): Text to insert
            format_bold (bool): Whether to format the inserted text as bold
            format_italic (bool): Whether to format the inserted text as italic
            format_size (int): Point size for the text (optional)
            format_color (dict): Color for the text in RGB format: {'red': r, 'green': g, 'blue': b} (optional)
        
        Returns:
            dict: Response from the API
        """
        # Insert text
        requests = [
            {
                'insertText': {
                    'location': location,
                    'text': text
                }
            }
        ]
        
        response = self.batch_update(requests)
        
        # Determine the range of the inserted text
        start_index = location.get('index', 0)
        end_index = start_index + len(text)
        
        # Apply formatting if needed
        if format_bold or format_italic or format_size or format_color:
            format_requests = []
            
            # Build the formatting request
            style = {}
            update_fields = []
            
            if format_bold:
                style['bold'] = True
                update_fields.append('bold')
            
            if format_italic:
                style['italic'] = True
                update_fields.append('italic')
            
            if format_size:
                style['fontSize'] = {'magnitude': format_size, 'unit': 'PT'}
                update_fields.append('fontSize')
            
            if format_color:
                style['foregroundColor'] = {'color': format_color}
                update_fields.append('foregroundColor')
            
            if update_fields:
                format_requests.append({
                    'updateTextStyle': {
                        'range': {
                            'segmentId': '',  # Use the "name-only" segment for the main document
                            'startIndex': start_index,
                            'endIndex': end_index
                        },
                        'textStyle': style,
                        'fields': ','.join(update_fields)
                    }
                })
                
                self.batch_update(format_requests)
        
        return response

    def create_header(self, text, level=1):
        """
        Create a new header at the end of the document.
        
        Args:
            text (str): Header text
            level (int): Header level (1-6)
        
        Returns:
            dict: Response from the API
        """
        doc = self.fetch()
        
        # Find the end of the document
        content = doc.get('body', {}).get('content', [])
        if content:
            last_element = content[-1]
            end_index = last_element.get('endIndex', 0)
        else:
            end_index = 1
        
        # Insert the header text
        insert_request = {
            'insertText': {
                'location': {
                    'index': end_index - 1
                },
                'text': text + '\n'
            }
        }
        
        # Apply header formatting
        style_requests = []
        
        if level == 1:
            style_name = 'HEADING_1'
        elif level == 2:
            style_name = 'HEADING_2'
        elif level == 3:
            style_name = 'HEADING_3'
        elif level == 4:
            style_name = 'HEADING_4'
        elif level == 5:
            style_name = 'HEADING_5'
        elif level == 6:
            style_name = 'HEADING_6'
        else:
            style_name = 'HEADING_1'  # Default to H1 if invalid level
        
        style_request = {
            'updateParagraphStyle': {
                'range': {
                    'startIndex': end_index - 1,
                    'endIndex': end_index + len(text)
                },
                'paragraphStyle': {
                    'namedStyleType': style_name
                },
                'fields': 'namedStyleType'
            }
        }
        
        style_requests.append(style_request)
        
        # Execute requests
        self.batch_update([insert_request])
        return self.batch_update(style_requests)

    def create_table(self, rows, cols, headers=None):
        """
        Create a table at the end of the document.
        
        Args:
            rows (int): Number of rows (excluding header row)
            cols (int): Number of columns
            headers (list): List of header values (optional)
        
        Returns:
            dict: Response from the API
        """
        doc = self.fetch()
        
        # Find the end of the document
        content = doc.get('body', {}).get('content', [])
        if content:
            last_element = content[-1]
            end_index = last_element.get('endIndex', 0)
        else:
            end_index = 1
        
        # Add a newline before the table if needed
        insert_newline = {
            'insertText': {
                'location': {
                    'index': end_index - 1
                },
                'text': '\n'
            }
        }
        
        # Create the table
        create_table = {
            'insertTable': {
                'location': {
                    'index': end_index  # +1 for the newline we just added
                },
                'rows': rows + (1 if headers else 0),  # Add a row for headers if provided
                'columns': cols
            }
        }
        
        # Insert the table
        response = self.batch_update([insert_newline, create_table])
        
        # Add headers if provided
        if headers and len(headers) <= cols:
            table_start_index = end_index  # This is where our table starts
            
            # We'd need to get the new structure of the document to find exact cell locations
            # This is a simplified approach and might need adjustment for the actual API
            header_requests = []
            
            for i, header in enumerate(headers):
                header_requests.append({
                    'insertText': {
                        'location': {
                            'index': table_start_index + (i * 2) + 2  # Approximate location
                        },
                        'text': header
                    }
                })
                
                # Make headers bold
                header_requests.append({
                    'updateTextStyle': {
                        'range': {
                            'startIndex': table_start_index + (i * 2) + 2,
                            'endIndex': table_start_index + (i * 2) + 2 + len(header)
                        },
                        'textStyle': {
                            'bold': True
                        },
                        'fields': 'bold'
                    }
                })
            
            self.batch_update(header_requests)
        
        return response

    def insert_image(self, uri, width=None, height=None):
        """
        Insert an image at the end of the document.
        
        Args:
            uri (str): URI of the image to insert
            width (int): Optional width in EMU (1/914400 of an inch)
            height (int): Optional height in EMU (1/914400 of an inch)
        
        Returns:
            dict: Response from the API
        """
        doc = self.fetch()
        
        # Find the end of the document
        content = doc.get('body', {}).get('content', [])
        if content:
            last_element = content[-1]
            end_index = last_element.get('endIndex', 0)
        else:
            end_index = 1
        
        # Create the image insertion request
        insert_image = {
            'insertInlineImage': {
                'location': {
                    'index': end_index - 1
                },
                'uri': uri
            }
        }
        
        # Add size if specified
        if width and height:
            insert_image['insertInlineImage']['objectSize'] = {
                'width': {
                    'magnitude': width,
                    'unit': 'EMU'
                },
                'height': {
                    'magnitude': height,
                    'unit': 'EMU'
                }
            }
        
        return self.batch_update([insert_image])

    def add_math_equation(self, latex):
        """
        Add a mathematical equation to the document using LaTeX notation.
        
        Args:
            latex (str): LaTeX notation for the math equation
        
        Returns:
            dict: Response from the API
        """
        doc = self.fetch()
        
        # Find the end of the document
        content = doc.get('body', {}).get('content', [])
        if content:
            last_element = content[-1]
            end_index = last_element.get('endIndex', 0)
        else:
            end_index = 1
        
        # Unfortunately, Google Docs API doesn't directly support adding LaTeX equations
        # As a workaround, we'll just insert the LaTeX code with a note
        insert_text = {
            'insertText': {
                'location': {
                    'index': end_index - 1
                },
                'text': f"[Math Equation: {latex}]\n"
            }
        }
        
        return self.batch_update([insert_text])

    def create_worksheet(self, title, problems, width=500, height=700):
        """
        Create a math worksheet with a title and a series of numbered problems.
        
        Args:
            title (str): Title of the worksheet
            problems (list): List of problem statements
            width (int): Width of the page in points
            height (int): Height of the page in points
        
        Returns:
            dict: Response from the API
        """
        # First fetch the document to get tab information
        doc = self.fetch()
        if 'tabs' not in doc or not doc['tabs']:
            raise ValueError("Document has no tabs")
            
        tab_id = doc['tabs'][0]['tabProperties']['tabId']
        
        # We'll execute operations one at a time to ensure document state is always current
        
        # Set page size
        self.batch_update([{
            'updateDocumentStyle': {
                'documentStyle': {
                    'pageSize': {
                        'width': {
                            'magnitude': width,
                            'unit': 'PT'
                        },
                        'height': {
                            'magnitude': height,
                            'unit': 'PT'
                        }
                    }
                },
                'fields': 'pageSize'
            }
        }])
        
        # Start with minimal content to ensure there's a paragraph
        response = self.batch_update([{
            'insertText': {
                'location': {
                    'index': 1,
                    'tabId': tab_id
                },
                'text': '\n'
            }
        }])
        
        # Add title
        self.append_text(
            tab_id, 
            title + '\n\n',
            format_bold=True
        )
        
        # Apply title style at the document beginning
        # Get the length of the title for proper styling
        title_length = len(title)
        self.batch_update([{
            'updateParagraphStyle': {
                'range': {
                    'startIndex': 1,
                    'endIndex': 1 + title_length,
                    'tabId': tab_id
                },
                'paragraphStyle': {
                    'namedStyleType': 'TITLE',
                    'alignment': 'CENTER'
                },
                'fields': 'namedStyleType,alignment'
            }
        }])
        
        # Add each problem one by one
        for i, problem in enumerate(problems, 1):
            problem_text = f"{i}. {problem}\n\n"
            
            # Append each problem to the end
            self.append_text(tab_id, problem_text)
            
        # Refresh document state
        self.fetch()
        
        return response

    def generate_answer_sheet(self, title, problems, answers):
        """
        Generate an answer sheet for a math worksheet.
        
        Args:
            title (str): Title of the answer sheet
            problems (list): List of problem numbers or identifiers
            answers (list): List of answers corresponding to the problems
        
        Returns:
            dict: Response from the API
        """
        if len(problems) != len(answers):
            raise ValueError("Number of problems must match number of answers")
        
        # First fetch the document to get tab information
        doc = self.fetch()
        if 'tabs' not in doc or not doc['tabs']:
            raise ValueError("Document has no tabs")
            
        tab_id = doc['tabs'][0]['tabProperties']['tabId']
        
        # We'll execute operations one by one to ensure document state is always current
        
        # Start with minimal content to ensure there's a paragraph
        self.batch_update([{
            'insertText': {
                'location': {
                    'index': 1,
                    'tabId': tab_id
                },
                'text': '\n'
            }
        }])
        
        # Add title
        self.append_text(
            tab_id, 
            title + '\n\n',
            format_bold=True
        )
        
        # Apply title style
        title_length = len(title)
        self.batch_update([{
            'updateParagraphStyle': {
                'range': {
                    'startIndex': 1,
                    'endIndex': 1 + title_length,
                    'tabId': tab_id
                },
                'paragraphStyle': {
                    'namedStyleType': 'TITLE',
                    'alignment': 'CENTER'
                },
                'fields': 'namedStyleType,alignment'
            }
        }])
        
        # Create formatted answer table using plain text for simplicity
        self.append_text(tab_id, "Problem\tAnswer\n")
        self.append_text(tab_id, "-------\t-------\n")
        
        # Add each problem-answer pair as a row
        for problem, answer in zip(problems, answers):
            row_text = f"{problem}\t{answer}\n"
            self.append_text(tab_id, row_text)
        
        # Refresh document state
        return self.fetch()