class Document:
    def __init__(self, docs_service, document_id):
        self.docs_service = docs_service
        self.document_id = document_id

    def fetch(self):
        """Fetch and return the document's full JSON structure."""
        return self.docs_service.documents().get(
            documentId=self.document_id
        ).execute()

    def batch_update(self, requests):
        """Execute a batchUpdate request on the document."""
        body = {'requests': requests}
        return self.docs_service.documents().batchUpdate(
            documentId=self.document_id,
            body=body
        ).execute()

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
        requests = []
        
        # Set page size
        requests.append({
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
        })
        
        # Add title
        doc = self.fetch()
        content = doc.get('body', {}).get('content', [])
        if content:
            last_element = content[-1]
            end_index = last_element.get('endIndex', 0)
        else:
            end_index = 1
        
        requests.append({
            'insertText': {
                'location': {
                    'index': end_index - 1
                },
                'text': title + '\n\n'
            }
        })
        
        # Format title
        title_end_index = end_index - 1 + len(title) + 2
        requests.append({
            'updateParagraphStyle': {
                'range': {
                    'startIndex': end_index - 1,
                    'endIndex': title_end_index - 2  # Exclude the newlines
                },
                'paragraphStyle': {
                    'namedStyleType': 'TITLE',
                    'alignment': 'CENTER'
                },
                'fields': 'namedStyleType,alignment'
            }
        })
        
        # Add problems
        current_index = title_end_index - 1
        for i, problem in enumerate(problems, 1):
            problem_text = f"{i}. {problem}\n\n"
            
            requests.append({
                'insertText': {
                    'location': {
                        'index': current_index
                    },
                    'text': problem_text
                }
            })
            
            current_index += len(problem_text)
        
        # Execute all requests
        return self.batch_update(requests)

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
        
        requests = []
        
        # Add title
        doc = self.fetch()
        content = doc.get('body', {}).get('content', [])
        if content:
            last_element = content[-1]
            end_index = last_element.get('endIndex', 0)
        else:
            end_index = 1
        
        requests.append({
            'insertText': {
                'location': {
                    'index': end_index - 1
                },
                'text': title + '\n\n'
            }
        })
        
        # Format title
        title_end_index = end_index - 1 + len(title) + 2
        requests.append({
            'updateParagraphStyle': {
                'range': {
                    'startIndex': end_index - 1,
                    'endIndex': title_end_index - 2  # Exclude the newlines
                },
                'paragraphStyle': {
                    'namedStyleType': 'TITLE',
                    'alignment': 'CENTER'
                },
                'fields': 'namedStyleType,alignment'
            }
        })
        
        # Create a two-column table for problem numbers and answers
        requests.append({
            'insertTable': {
                'location': {
                    'index': title_end_index - 1
                },
                'rows': len(problems) + 1,  # +1 for the header row
                'columns': 2
            }
        })
        
        # We would need to get the document structure again to find the exact cell positions
        # This is an approximation
        table_start_index = title_end_index - 1
        
        # Add headers
        headers = ["Problem", "Answer"]
        for i, header in enumerate(headers):
            cell_index = table_start_index + (i * 2) + 2  # Approximate cell location
            
            requests.append({
                'insertText': {
                    'location': {
                        'index': cell_index
                    },
                    'text': header
                }
            })
            
            # Make headers bold
            requests.append({
                'updateTextStyle': {
                    'range': {
                        'startIndex': cell_index,
                        'endIndex': cell_index + len(header)
                    },
                    'textStyle': {
                        'bold': True
                    },
                    'fields': 'bold'
                }
            })
        
        # Add problems and answers
        # This is a simplification - in a real implementation, you'd need to determine the exact cell positions
        row_offset = 20  # Approximate offset for each row
        for i, (problem, answer) in enumerate(zip(problems, answers), 1):
            problem_cell_index = table_start_index + 2 + (i * row_offset)
            answer_cell_index = problem_cell_index + 10  # Approximate offset for second column
            
            requests.append({
                'insertText': {
                    'location': {
                        'index': problem_cell_index
                    },
                    'text': str(problem)
                }
            })
            
            requests.append({
                'insertText': {
                    'location': {
                        'index': answer_cell_index
                    },
                    'text': str(answer)
                }
            })
        
        # Execute all requests
        return self.batch_update(requests)