"""
Markdown parsing and formatting for Google Docs API integration.

This module provides functionality to parse markdown text and convert it to
the appropriate Google Docs API requests for rich text formatting.
"""

import re
from typing import Dict, List, Tuple, Any, Optional, Union
import markdown  # pip install markdown
from bs4 import BeautifulSoup  # pip install beautifulsoup4


class MarkdownFormatter:
    """Converts markdown text to Google Docs API formatting requests."""
    
    def __init__(self):
        """Initialize the markdown formatter with regex patterns for markdown elements."""
        # Text formatting patterns
        self.bold_pattern = re.compile(r'\*\*(.*?)\*\*')
        self.italic_pattern = re.compile(r'\*(.*?)\*')
        self.code_pattern = re.compile(r'`(.*?)`')
        self.link_pattern = re.compile(r'\[(.*?)\]\((.*?)\)')
        
        # List patterns
        self.unordered_list_pattern = re.compile(r'^\s*[\*\-\+]\s+(.*?)$', re.MULTILINE)
        self.ordered_list_pattern = re.compile(r'^\s*(\d+)\.?\s+(.*?)$', re.MULTILINE)
        
        # Table pattern
        self.table_pattern = re.compile(r'^\|(.+)\|$', re.MULTILINE)
        self.table_separator_pattern = re.compile(r'^\|(\s*[-:]+[-|\s:]*)\|$', re.MULTILINE)
        
        # Header patterns
        self.header_pattern = re.compile(r'^(#{1,6})\s+(.*?)$', re.MULTILINE)

    def parse(self, text: str) -> List[Dict[str, Any]]:
        """
        Parse markdown text and return a list of Google Docs API requests.
        
        Args:
            text: Markdown text to parse
            
        Returns:
            List of Google Docs API requests to apply the formatting
        """
        requests = []
        
        # Track text spans and their formatting
        spans = self._identify_all_spans(text)
        
        # Convert identified spans to formatting requests
        # Implementation will depend on how you want to insert/replace text
        
        return requests
    
    def _identify_all_spans(self, text: str) -> List[Tuple[int, int, Dict[str, Any]]]:
        """
        Identify all spans of text that need formatting.
        
        Args:
            text: Text to analyze
            
        Returns:
            List of (start, end, style_dict) tuples
        """
        spans = []
        
        # Find all bold spans
        for match in self.bold_pattern.finditer(text):
            # The span is from the start to end, but we don't include the ** markers
            start = match.start() + 2
            end = match.end() - 2
            spans.append((start, end, {'bold': True}))
        
        # Find all italic spans
        for match in self.italic_pattern.finditer(text):
            # Skip if this is inside a bold pattern (which uses ** which would match *)
            # This is a simplified approach and might need refinement
            is_inside_bold = False
            for bold_match in self.bold_pattern.finditer(text):
                if bold_match.start() <= match.start() and bold_match.end() >= match.end():
                    is_inside_bold = True
                    break
            
            if not is_inside_bold:
                start = match.start() + 1
                end = match.end() - 1
                spans.append((start, end, {'italic': True}))
        
        # Find all code spans
        for match in self.code_pattern.finditer(text):
            start = match.start() + 1
            end = match.end() - 1
            # For code blocks, we might want to use monospace font and potentially a background color
            spans.append((start, end, {
                'weightedFontFamily': {'fontFamily': 'Courier New'},
                'backgroundColor': {'color': {'rgbColor': {'red': 0.95, 'green': 0.95, 'blue': 0.95}}}
            }))
        
        # Find all links
        for match in self.link_pattern.finditer(text):
            start = match.start()
            end = match.end()
            text_part = match.group(1)
            url_part = match.group(2)
            # For links, we'll need to handle this specially since it requires link object creation
            spans.append((start, end, {'link': {'url': url_part}}))
        
        # Sort by start position for processing
        spans.sort(key=lambda x: x[0])
        
        return spans

    def create_text_insertion_requests(self, text: str, index: int) -> Tuple[List[Dict[str, Any]], int]:
        """
        Create requests to insert markdown text at the specified index.
        
        Args:
            text: Markdown text to insert
            index: Position in the document to insert the text
            
        Returns:
            Tuple of (requests list, new index after insertion)
        """
        # Get the cleaned text without markdown syntax
        cleaned_text = self._remove_markdown_syntax(text)
        if not cleaned_text.strip():
            cleaned_text = " "  # Avoid empty text
        
        # Start with plain text insertion
        requests = [
            {
                'insertText': {
                    'location': {'index': index},
                    'text': cleaned_text
                }
            }
        ]
        
        # Process links separately for better link handling
        link_matches = list(self.link_pattern.finditer(text))
        for match in link_matches:
            text_part = match.group(1)
            url_part = match.group(2)
            
            # Find where this text appears in the cleaned text
            text_in_cleaned = cleaned_text.find(text_part)
            
            if text_in_cleaned >= 0:
                # Add link formatting
                requests.append({
                    'updateTextStyle': {
                        'range': {
                            'startIndex': index + text_in_cleaned,
                            'endIndex': index + text_in_cleaned + len(text_part)
                        },
                        'textStyle': {
                            'link': {'url': url_part}
                        },
                        'fields': 'link'
                    }
                })
        
        # Now add regular formatting requests (bold, italic, code)
        spans = self._identify_all_spans(text)
        
        for start, end, style in spans:
            # Skip link spans as we've already handled them
            if 'link' in style:
                continue
                
            # Convert the span positions from the markdown text to the cleaned text
            clean_start = start - self._count_preceding_syntax_chars(text, start)
            clean_end = end - self._count_preceding_syntax_chars(text, end)
            
            # Ensure ranges aren't empty and are valid
            if clean_start >= clean_end or clean_start < 0 or clean_end > len(cleaned_text):
                continue  # Skip invalid ranges
            
            # Prepare the fields list - only include valid and present fields
            style_fields = []
            for key in style:
                style_fields.append(key)
            
            # Skip if no valid fields
            if not style_fields:
                continue
                
            # Add a text style update request
            requests.append({
                'updateTextStyle': {
                    'range': {
                        'startIndex': index + clean_start,
                        'endIndex': index + clean_end
                    },
                    'textStyle': style,
                    'fields': ','.join(style_fields)
                }
            })
        
        return requests, index + len(cleaned_text)
    
    def _remove_markdown_syntax(self, text: str) -> str:
        """
        Remove markdown syntax from text.
        
        Args:
            text: Markdown text
            
        Returns:
            Text with markdown syntax removed, including headers, bold, italic,
            code markers, and list markers.
        """
        # Remove header markers (e.g., "# ")
        text = re.sub(r'^#{1,6}\s+', '', text, flags=re.MULTILINE)
        # Remove bold markers
        text = re.sub(r'\*\*(.*?)\*\*', r'\1', text)
        # Remove italic markers
        text = re.sub(r'\*(.*?)\*', r'\1', text)
        # Remove inline code markers
        text = re.sub(r'`(.*?)`', r'\1', text)
        # Convert links to just the text part
        text = re.sub(r'\[(.*?)\]\(.*?\)', r'\1', text)
        # Remove ordered list markers (e.g., "1. ", "2. ", etc.)
        text = re.sub(r'^\s*\d+\.\s+', '', text, flags=re.MULTILINE)
        # Remove unordered list markers (e.g., "* ", "- ", "+ ")
        text = re.sub(r'^\s*[\*\-\+]\s+', '', text, flags=re.MULTILINE)
        # (Add any other syntax removals as needed)
        return text
    
    def _count_preceding_syntax_chars(self, text: str, position: int) -> int:
        """
        Count syntax characters preceding the given position.
        
        Args:
            text: The text to analyze
            position: The position to check before
            
        Returns:
            Count of syntax characters
        """
        # Count markdown syntax characters that would be removed
        # This is an improved implementation to track position shifts
        
        # Characters to consider as syntax
        syntax_counts = 0
        
        # Count bold markers (**) before this position
        for match in self.bold_pattern.finditer(text):
            if match.start() < position:
                # If we're inside the bold section, only count the chars before our position
                if match.end() > position:
                    # Count the opening ** only
                    syntax_counts += 2
                else:
                    # Count both opening and closing ** (total 4 chars)
                    syntax_counts += 4
        
        # Count italic markers (*) before this position
        for match in self.italic_pattern.finditer(text):
            # Skip if this is inside a bold pattern which we've already counted
            is_inside_bold = False
            for bold_match in self.bold_pattern.finditer(text):
                if bold_match.start() <= match.start() and bold_match.end() >= match.end():
                    is_inside_bold = True
                    break
            
            if not is_inside_bold:
                if match.start() < position:
                    # If we're inside the italic section, only count the chars before our position
                    if match.end() > position:
                        # Count the opening * only
                        syntax_counts += 1
                    else:
                        # Count both opening and closing * (total 2 chars)
                        syntax_counts += 2
        
        # Count code markers (`) before this position
        for match in self.code_pattern.finditer(text):
            if match.start() < position:
                # If we're inside the code section, only count the chars before our position
                if match.end() > position:
                    # Count the opening ` only
                    syntax_counts += 1
                else:
                    # Count both opening and closing ` (total 2 chars)
                    syntax_counts += 2
        
        # Count link markers before this position
        for match in self.link_pattern.finditer(text):
            if match.start() < position:
                # Links have format [text](url)
                # We need to count [, ], (, and ) plus the URL that's not in the cleaned text
                text_part = match.group(1)
                url_part = match.group(2)
                
                if match.end() > position:
                    # We're inside the link, compute exact position
                    if match.start() + 1 + len(text_part) + 1 < position:
                        # We're after the text part, include the brackets and opening parens
                        syntax_counts += 3  # [, ], and (
                        # Count how much of the URL we've passed
                        url_chars_passed = position - (match.start() + 1 + len(text_part) + 2)
                        if url_chars_passed > 0:
                            syntax_counts += min(url_chars_passed, len(url_part))
                    else:
                        # We're in the text part, just count the opening bracket
                        syntax_counts += 1
                else:
                    # Count the entire syntax: [, ], (, ), and the URL
                    syntax_counts += 4 + len(url_part)
        
        return syntax_counts
    
    def convert_to_doc_requests(self, text: str, start_index: int) -> List[Dict[str, Any]]:
        """
        Convert Markdown text into a list of Google Docs API requests using a robust AST-based approach.
        We parse Markdown to HTML, then walk the HTML structure to create paragraphs, headings, lists,
        and inline formatting requests.

        :param text: The original Markdown text to convert
        :param start_index: The position in the document where we should begin inserting
        :return: A list of Google Docs API requests
        """
        # STEP 1: Convert Markdown to HTML via python-markdown
        html_string = markdown.markdown(text)  # e.g. <p>Some text</p>, <ul><li>...</li></ul>, etc.

        # STEP 2: Parse the HTML string into a BeautifulSoup DOM
        soup = BeautifulSoup(html_string, "html.parser")

        requests: List[Dict[str, Any]] = []
        insertion_index = start_index

        # We'll process block-level elements in order, so that the text is inserted into the doc in sequence.
        for block in soup.contents:
            # If it's just whitespace or a newline, ignore
            if not block.name and not block.strip():
                continue

            if block.name == 'p':
                # Basic paragraph
                block_requests, insertion_index = self._process_paragraph(block, insertion_index)
                requests.extend(block_requests)

            elif block.name and block.name in ['h1', 'h2', 'h3', 'h4', 'h5', 'h6']:
                # Headings
                block_requests, insertion_index = self._process_heading(block, insertion_index)
                requests.extend(block_requests)

            elif block.name == 'ul':
                # Unordered list
                block_requests, insertion_index = self._process_list(block, insertion_index, ordered=False)
                requests.extend(block_requests)

            elif block.name == 'ol':
                # Ordered list
                block_requests, insertion_index = self._process_list(block, insertion_index, ordered=True)
                requests.extend(block_requests)
            
            else:
                # Fallback: treat as paragraph
                block_requests, insertion_index = self._process_paragraph(block, insertion_index)
                requests.extend(block_requests)

        return requests

    def _process_paragraph(self, block, insertion_index: int) -> Tuple[List[Dict[str, Any]], int]:
        """
        Insert a paragraph node's text, then apply any inline formatting such as <strong>, <em>, <code>, <a>.
        """
        requests: List[Dict[str, Any]] = []
        paragraph_text = block.get_text()
        if not paragraph_text.strip():
            paragraph_text = "\n"  # Insert a blank line for empty paragraphs
        else:
            paragraph_text += "\n"  # Ensure newline at paragraph end

        # Insert the paragraph text
        requests.append({
            'insertText': {
                'location': {'index': insertion_index},
                'text': paragraph_text
            }
        })

        start_offset = insertion_index
        end_offset = insertion_index + len(paragraph_text)
        new_insertion_index = end_offset

        # Now apply inline formatting by iterating over the block's children
        # (e.g. <strong>, <em>, <code>, <a>)
        inline_requests = self._generate_inline_format_requests(block, start_offset)
        requests.extend(inline_requests)

        # If we want to maintain normal paragraph styling, we can do so here:
        # e.g., no bullet, no heading style, just normal text
        return requests, new_insertion_index

    def _process_heading(self, block, insertion_index: int) -> Tuple[List[Dict[str, Any]], int]:
        """
        Insert the heading text, then apply the relevant heading style (HEADING_1 through HEADING_6).
        """
        requests: List[Dict[str, Any]] = []
        heading_text = block.get_text() + "\n"  # ensure newline after heading
        requests.append({
            'insertText': {
                'location': {'index': insertion_index},
                'text': heading_text
            }
        })

        start_offset = insertion_index
        end_offset = insertion_index + len(heading_text)
        new_insertion_index = end_offset

        # Determine heading level from the tag name
        level = int(block.name[-1])  # for "h1" -> 1, "h2" -> 2, etc.
        if level < 1 or level > 6:
            level = 1

        # Apply the heading style
        requests.append({
            'updateParagraphStyle': {
                'range': {
                    'startIndex': start_offset,
                    'endIndex': end_offset
                },
                'paragraphStyle': {
                    'namedStyleType': f'HEADING_{level}'
                },
                'fields': 'namedStyleType'
            }
        })

        # Also apply inline formatting within the heading (e.g., <strong> inside an <h2>)
        inline_requests = self._generate_inline_format_requests(block, start_offset)
        requests.extend(inline_requests)

        return requests, new_insertion_index

    def _process_list(self, block, insertion_index: int, ordered: bool) -> Tuple[List[Dict[str, Any]], int]:
        """
        Insert a list (either <ul> or <ol>). Each <li> becomes a bullet or a numbered list item.
        We gather all paragraphs within each <li> and insert them at once, preserving inter-paragraph blank lines
        but not creating extra bullets.
        """
        requests: List[Dict[str, Any]] = []

        list_item_start_index = None
        list_item_end_index = None

        # Find all top-level <li> elements
        list_items = block.find_all('li', recursive=False)

        for li_idx, list_item in enumerate(list_items):
            # Gather all <p> (and possibly <br>) inside this <li>
            paragraphs = list_item.find_all('p', recursive=False)
            
            if paragraphs:
                # If there are multiple <p> tags, join them with blank lines or single newlines
                # so they remain in the same bullet. Here we use double newlines for spacing:
                combined_content = []
                for p in paragraphs:
                    # if there's text
                    txt = p.get_text(strip=False)
                    # If completely empty, maybe skip or preserve
                    if txt.strip() == '':
                        # If you want an actual blank line for spacing, do:
                        combined_content.append('\n')
                    else:
                        combined_content.append(txt)
                li_text = "\n\n".join(combined_content).rstrip() + "\n"
            else:
                # If no <p> tags, just use the text from the <li> itself
                li_text = list_item.get_text().rstrip() + "\n"

            # Insert text (all paragraphs as one chunk)
            requests.append({
                'insertText': {
                    'location': {'index': insertion_index},
                    'text': li_text
                }
            })

            if list_item_start_index is None:
                list_item_start_index = insertion_index
            
            insertion_index += len(li_text)
            list_item_end_index = insertion_index

            # Now apply inline formatting inside this <li> 
            # (bold, italic, links, etc.). Because we merged paragraphs, we call
            # _generate_inline_format_requests() on each paragraph or on list_item:
            inline_requests = self._generate_inline_format_requests(list_item, list_item_start_index)
            requests.extend(inline_requests)

        # Once we've inserted all items, apply bullet or numbered style
        if list_item_start_index is not None and list_item_end_index is not None:
            bullet_preset = 'BULLET_DISC_CIRCLE_SQUARE' if not ordered else 'NUMBERED_DECIMAL_ALPHA_ROMAN'
            requests.append({
                'createParagraphBullets': {
                    'range': {
                        'startIndex': list_item_start_index,
                        'endIndex': list_item_end_index
                    },
                    'bulletPreset': bullet_preset
                }
            })

        return requests, insertion_index

    def _generate_inline_format_requests(self, parent_tag, insertion_offset: int) -> List[Dict[str, Any]]:
        """
        Recursively walk an element's children to find <strong>, <em>, <a>, <code>, etc.
        Returns a list of 'updateTextStyle' requests (and possibly 'insertText' if needed).
        Each text snippet is found by its position in the parent's get_text().
        """
        requests: List[Dict[str, Any]] = []
        parent_text = parent_tag.get_text()

        # We'll keep track of the absolute offset of each snippet. We can use a running index or a method like find().
        # For robust handling of repeated substrings, you'd do a deeper approach, but here's a simple example:
        child_strings = []
        for elem in parent_tag.descendants:
            if elem.name in ['strong', 'b', 'em', 'i', 'code', 'a']:
                # Retrieve the text for this inline node
                snippet = elem.get_text()
                if not snippet:
                    continue
                # Find snippet occurrence in the parent text
                snippet_start = parent_text.find(snippet)
                if snippet_start == -1:
                    continue
                snippet_end = snippet_start + len(snippet)

                # Build style dict
                style = {}
                fields = []
                if elem.name in ['strong', 'b']:
                    style['bold'] = True
                    fields.append('bold')
                if elem.name in ['em', 'i']:
                    style['italic'] = True
                    fields.append('italic')
                if elem.name == 'code':
                    style['weightedFontFamily'] = {'fontFamily': 'Courier New'}
                    style['backgroundColor'] = {
                        'color': {
                            'rgbColor': {'red': 0.95, 'green': 0.95, 'blue': 0.95}
                        }
                    }
                    fields.extend(['weightedFontFamily', 'backgroundColor'])
                if elem.name == 'a':
                    # Get link from <a> tag
                    href = elem.get('href')
                    if href:
                        style['link'] = {'url': href}
                        fields.append('link')

                # Generate request
                requests.append({
                    'updateTextStyle': {
                        'range': {
                            'startIndex': insertion_offset + snippet_start,
                            'endIndex': insertion_offset + snippet_end
                        },
                        'textStyle': style,
                        'fields': ','.join(fields)
                    }
                })

        return requests

    def _apply_inline_formatting(
        self,
        requests: List[Dict[str, Any]],
        markdown_text: str,
        cleaned_text: str,
        current_index: int
    ) -> None:
        """
        Apply inline formatting (bold, italic, code, and link) to the inserted text.

        Args:
            requests: Request list to append to
            markdown_text: Original markdown text (with **, *, `, and [text](link))
            cleaned_text: The text after removing those markdown markers
            current_index: Where the inserted text begins in the document
        """

        # 1) Process inline code blocks first ( `some code` )
        #    We'll give them a monospace font and a light background.
        for code_match in self.code_pattern.finditer(markdown_text):
            code_text = code_match.group(1)  # the content inside backticks
            text_pos = cleaned_text.find(code_text)

            # If found in the cleaned text, apply "monospace" style
            if text_pos >= 0:
                requests.append({
                    'updateTextStyle': {
                        'range': {
                            'startIndex': current_index + text_pos,
                            'endIndex': current_index + text_pos + len(code_text)
                        },
                        'textStyle': {
                            'weightedFontFamily': {'fontFamily': 'Courier New'},
                            'backgroundColor': {
                                'color': {
                                    'rgbColor': {'red': 0.95, 'green': 0.95, 'blue': 0.95}
                                }
                            }
                        },
                        'fields': 'weightedFontFamily,backgroundColor'
                    }
                })

        # 2) Process bold text (**text**)
        for bold_match in self.bold_pattern.finditer(markdown_text):
            bold_text = bold_match.group(1)  # the content inside **
            # See if that substring is present in cleaned_text
            text_pos = cleaned_text.find(bold_text)
            if text_pos >= 0:
                # Apply bold formatting
                requests.append({
                    'updateTextStyle': {
                        'range': {
                            'startIndex': current_index + text_pos,
                            'endIndex': current_index + text_pos + len(bold_text)
                        },
                        'textStyle': {
                            'bold': True
                        },
                        'fields': 'bold'
                    }
                })

        # 3) Process italic text (*text*), skipping those inside bold
        for italic_match in self.italic_pattern.finditer(markdown_text):
            italic_text = italic_match.group(1)
            # Skip if it's inside a bold region
            is_inside_bold = False
            for bold_match in self.bold_pattern.finditer(markdown_text):
                if (bold_match.start() <= italic_match.start() and
                        bold_match.end() >= italic_match.end()):
                    is_inside_bold = True
                    break
            if not is_inside_bold:
                text_pos = cleaned_text.find(italic_text)
                if text_pos >= 0:
                    # Apply italic formatting
                    requests.append({
                        'updateTextStyle': {
                            'range': {
                                'startIndex': current_index + text_pos,
                                'endIndex': current_index + text_pos + len(italic_text)
                            },
                            'textStyle': {
                                'italic': True
                            },
                            'fields': 'italic'
                        }
                    })

        # 4) Process links ([text](url))
        for link_match in self.link_pattern.finditer(markdown_text):
            link_text = link_match.group(1)  # text inside [...]
            url = link_match.group(2)        # URL inside (...)
            
            # Find link_text in the cleaned text
            text_pos = cleaned_text.find(link_text)
            if text_pos >= 0:
                # Apply link style
                requests.append({
                    'updateTextStyle': {
                        'range': {
                            'startIndex': current_index + text_pos,
                            'endIndex': current_index + text_pos + len(link_text)
                        },
                        'textStyle': {
                            'link': {
                                'url': url
                            }
                        },
                        'fields': 'link'
                    }
                })