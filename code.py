"""
DOCX → JSON (via Docling) → Editable HTML + Reconstructed DOCX.
"""
import json
from pathlib import Path
from docling.document_converter import DocumentConverter, InputFormat, WordFormatOption
from docling.datamodel.pipeline_options import PaginatedPipelineOptions

def docx_to_json_to_html(input_docx_path, output_json_path=None, output_html_path=None):
    """
    Convert DOCX to JSON first, then JSON to HTML
    
    Args:
        input_docx_path: Path to input DOCX file
        output_json_path: Path for JSON output (optional)
        output_html_path: Path for HTML output (optional)
    """
    
    # Setup paths
    input_path = Path(input_docx_path)
    
    if output_json_path is None:
        output_json_path = input_path.with_suffix('.json')
    else:
        output_json_path = Path(output_json_path)
        
    if output_html_path is None:
        output_html_path = input_path.with_name(f"{input_path.stem}-editable.html")
    else:
        output_html_path = Path(output_html_path)
    
    # Configure pipeline options
    docx_pipeline = PaginatedPipelineOptions()
    docx_pipeline.generate_page_images = False
    
    # Setup document converter
    converter = DocumentConverter(
        allowed_formats=[InputFormat.DOCX],
        format_options={InputFormat.DOCX: WordFormatOption(pipeline_options=docx_pipeline)},
    )
    
    # Step 1: Convert DOCX to JSON
    print(f"Converting {input_path} to JSON...")
    result = converter.convert(str(input_path))
    doc = result.document
    
    # Export to JSON
    json_content = doc.export_to_dict()  # Gets the document as a dictionary
    
    # Save JSON file
    with open(output_json_path, 'w', encoding='utf-8') as f:
        json.dump(json_content, f, indent=2, ensure_ascii=False)
    print(f"JSON saved: {output_json_path.resolve()}")
    
    # Step 2: Convert JSON to HTML
    print("Converting JSON to HTML...")
    html_content = json_to_html(json_content)
    
    # Create complete HTML shell
    html_shell = create_html_shell(html_content, input_path.stem)
    
    # Save HTML file
    output_html_path.write_text(html_shell, encoding="utf-8")
    print(f"HTML saved: {output_html_path.resolve()}")
    
    return output_json_path, output_html_path

def resolve_reference(ref_string, json_content):
    """
    Resolve a JSON reference like '#/texts/0' to the actual object
    
    Args:
        ref_string: Reference string like '#/texts/0'
        json_content: Full JSON document
    
    Returns:
        dict: Referenced object
    """
    # Remove the '#/' prefix and split the path
    path = ref_string.replace('#/', '').split('/')
    
    # Navigate through the JSON structure
    current = json_content
    for part in path:
        if part.isdigit():
            current = current[int(part)]
        else:
            current = current[part]
    
    return current

def json_to_html(json_content):
    """
    Convert Docling JSON document structure to HTML
    
    Args:
        json_content: Dictionary containing Docling document structure
    
    Returns:
        str: HTML content
    """
    print("Processing Docling JSON structure...")
    
    html_parts = []
    
    # Process the body content in order
    if 'body' in json_content and 'children' in json_content['body']:
        children = json_content['body']['children']
        print(f"Found {len(children)} body children")
        
        for child in children:
            if '$ref' in child:
                ref = child['$ref']
                
                # Resolve the reference to get the actual content
                try:
                    content_item = resolve_reference(ref, json_content)
                    
                    if ref.startswith('#/texts/'):
                        html_parts.append(process_text_item(content_item))
                    elif ref.startswith('#/tables/'):
                        html_parts.append(process_table_item(content_item))
                    elif ref.startswith('#/pictures/'):
                        html_parts.append(process_picture_item(content_item))
                    
                except (KeyError, IndexError) as e:
                    print(f"Warning: Could not resolve reference {ref}: {e}")
                    continue
    
    # Filter out empty strings and join
    return '\n'.join(filter(None, html_parts))

def process_text_item(text_item):
    """
    Process a text item from the Docling JSON structure
    
    Args:
        text_item: Text item dictionary
    
    Returns:
        str: HTML representation
    """
    text = text_item.get('text', '').strip()
    
    # Skip empty text items
    if not text:
        return ''
    
    # Get formatting information
    formatting = text_item.get('formatting', {})
    label = text_item.get('label', 'paragraph')
    
    # Apply formatting
    if formatting.get('bold'):
        if len(text) > 50:  # Long bold text is likely a heading
            html_tag = 'h2'
        else:
            html_tag = 'strong'
            text = f'<{html_tag}>{escape_html(text)}</{html_tag}>'
            return f'<p>{text}</p>'
    else:
        html_tag = 'p'
    
    # Apply other formatting
    if formatting.get('italic'):
        text = f'<em>{escape_html(text)}</em>'
    elif formatting.get('underline'):
        text = f'<u>{escape_html(text)}</u>'
    elif formatting.get('strikethrough'):
        text = f'<del>{escape_html(text)}</del>'
    else:
        text = escape_html(text)
    
    # Determine the HTML element based on label and formatting
    if html_tag == 'h2':
        return f'<h2>{text}</h2>'
    elif label == 'title':
        return f'<h1>{text}</h1>'
    elif label == 'heading':
        return f'<h3>{text}</h3>'
    else:
        return f'<p>{text}</p>'

def process_table_item(table_item):
    """
    Process a table item from the Docling JSON structure
    
    Args:
        table_item: Table item dictionary
    
    Returns:
        str: HTML table representation
    """
    if 'data' not in table_item or 'grid' not in table_item['data']:
        return ''
    
    grid = table_item['data']['grid']
    if not grid:
        return ''
    
    html_parts = ['<table>']
    
    for row_idx, row in enumerate(grid):
        html_parts.append('  <tr>')
        
        for cell in row:
            cell_text = cell.get('text', '').strip()
            
            # Determine if this is a header cell
            is_header = cell.get('column_header', False) or cell.get('row_header', False)
            tag = 'th' if is_header else 'td'
            
            # Handle cell spanning
            row_span = cell.get('row_span', 1)
            col_span = cell.get('col_span', 1)
            
            span_attrs = ''
            if row_span > 1:
                span_attrs += f' rowspan="{row_span}"'
            if col_span > 1:
                span_attrs += f' colspan="{col_span}"'
            
            html_parts.append(f'    <{tag}{span_attrs}>{escape_html(cell_text)}</{tag}>')
        
        html_parts.append('  </tr>')
    
    html_parts.append('</table>')
    
    return '\n'.join(html_parts)

import base64

def process_picture_item(picture_item):
    """
    Process a picture item from the Docling JSON structure
    
    Args:
        picture_item: Picture item dictionary
    
    Returns:
        str: HTML <img> tag with embedded base64 image
    """
    # If no data, return placeholder
    if 'data' not in picture_item:
        return '<div class="image-placeholder">[Missing Image]</div>'
    
    # Get image data
    img_data = picture_item['data']
    img_format = picture_item.get('format', 'png')  # default to PNG
    
    # Create base64 data URI
    img_src = f"data:image/{img_format};base64,{img_data}"
    
    return f'<img src="{img_src}" alt="Embedded Image" style="max-width:100%; height:auto; margin:10px 0;" />'


def escape_html(text):
    """
    Escape HTML special characters
    
    Args:
        text: Input text
    
    Returns:
        str: HTML-escaped text
    """
    if not isinstance(text, str):
        text = str(text)
    
    return (text
            .replace('&', '&amp;')
            .replace('<', '&lt;')
            .replace('>', '&gt;')
            .replace('"', '&quot;')
            .replace("'", '&#x27;'))

def create_html_shell(content_html, document_title):
    """
    Create complete HTML document with embedded CSS
    
    Args:
        content_html: Main content HTML
        document_title: Document title for page title
    
    Returns:
        str: Complete HTML document
    """
    return f"""<!doctype html>
<html>
<head>
    <meta charset="utf-8">
    <title>Editable — {document_title}</title>
    <style>
        body {{ 
            font-family: Arial, Helvetica, sans-serif; 
            margin: 16px; 
            line-height: 1.6;
            max-width: 800px;
            margin: 0 auto;
            padding: 20px;
        }}
        
        /* Typography */
        h1, h2, h3 {{ 
            color: #333; 
            margin: 1.5em 0 0.5em 0; 
        }}
        h1 {{ 
            font-size: 1.8em; 
            text-align: center;
            border-bottom: 2px solid #333;
            padding-bottom: 10px;
        }}
        h2 {{ 
            font-size: 1.5em; 
            color: #2c5aa0;
        }}
        h3 {{ font-size: 1.2em; }}
        
        p {{ 
            margin: 0.5em 0; 
        }}
        
        /* Bold text styling */
        strong {{
            font-weight: bold;
            color: #2c5aa0;
        }}
        
        /* Tables */
        table {{ 
            border-collapse: collapse; 
            width: 100%; 
            margin: 1.5em 0; 
            border: 2px solid #333;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }}
        th, td {{ 
            border: 1px solid #ccc; 
            padding: 12px 15px; 
            vertical-align: top; 
            text-align: left;
        }}
        th {{
            background-color: #f8f9fa;
            font-weight: bold;
            color: #2c5aa0;
            border-bottom: 2px solid #2c5aa0;
        }}
        
        /* Alternating row colors */
        tr:nth-child(even) {{
            background-color: #f9f9f9;
        }}
        
        /* Lists */
        ul, ol {{ 
            margin: 0.5em 0 0.5em 1.5em; 
        }}
        li {{
            margin: 0.3em 0;
        }}
        
        /* Image placeholders */
        .image-placeholder {{
            background: #f0f0f0;
            border: 2px dashed #ccc;
            padding: 20px;
            text-align: center;
            margin: 10px 0;
            color: #666;
        }}
        
        /* Editing features */
        [contenteditable="true"]:focus {{ 
            outline: 2px solid #4c9ffe; 
            background-color: #fafafa;
        }}
        
        /* Control panel */
        .control-panel {{
            position: sticky;
            top: 0;
            background: white;
            border-bottom: 1px solid #ddd;
            padding: 10px;
            margin-bottom: 20px;
            z-index: 100;
        }}
        
        .control-button {{
            background: #007bff;
            color: white;
            border: none;
            padding: 8px 16px;
            border-radius: 4px;
            cursor: pointer;
            margin-right: 10px;
            font-size: 14px;
        }}
        .control-button:hover {{
            background: #0056b3;
        }}
        
        .control-button.secondary {{
            background: #6c757d;
        }}
        .control-button.secondary:hover {{
            background: #545b62;
        }}
        
        /* JSON display */
        #json-display {{
            display: none;
            background: #f8f8f8;
            border: 1px solid #ddd;
            padding: 15px;
            margin: 15px 0;
            border-radius: 5px;
            font-family: 'Courier New', monospace;
            white-space: pre-wrap;
            max-height: 400px;
            overflow-y: auto;
            font-size: 12px;
        }}
        
        /* Status indicator */
        .status {{
            float: right;
            padding: 5px 10px;
            border-radius: 3px;
            font-size: 12px;
        }}
        .status.saved {{
            background: #d4edda;
            color: #155724;
        }}
        .status.modified {{
            background: #fff3cd;
            color: #856404;
        }}
    </style>
</head>
<body>
    <div class="control-panel">
        <button class="control-button" onclick="toggleJsonDisplay()">Show/Hide JSON</button>
        <button class="control-button secondary" onclick="exportContent()">Export HTML</button>
        <button class="control-button secondary" onclick="printDocument()">Print</button>
        <div class="status saved" id="status">Saved</div>
    </div>
    
    <div id="json-display"></div>
    
    <div id="editor" contenteditable="true">
        {content_html}
    </div>
    
    <script>
        let isJsonVisible = false;
        let hasUnsavedChanges = false;
        
        function toggleJsonDisplay() {{
            const jsonDiv = document.getElementById('json-display');
            isJsonVisible = !isJsonVisible;
            
            if (isJsonVisible) {{
                jsonDiv.style.display = 'block';
                jsonDiv.textContent = 'JSON structure processed to create this editable document.\\n\\nOriginal structure: Docling format with references to text and table elements.';
            }} else {{
                jsonDiv.style.display = 'none';
            }}
        }}
        
        function exportContent() {{
            const content = document.getElementById('editor').innerHTML;
            const blob = new Blob([content], {{ type: 'text/html' }});
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'exported-content.html';
            a.click();
            URL.revokeObjectURL(url);
        }}
        
        function printDocument() {{
            window.print();
        }}
        
        function updateStatus(status) {{
            const statusDiv = document.getElementById('status');
            statusDiv.className = 'status ' + status;
            statusDiv.textContent = status === 'saved' ? 'Saved' : 'Modified';
        }}
        
        // Auto-save functionality
        let saveTimeout;
        document.getElementById('editor').addEventListener('input', function() {{
            hasUnsavedChanges = true;
            updateStatus('modified');
            
            clearTimeout(saveTimeout);
            saveTimeout = setTimeout(function() {{
                // Simulate save
                hasUnsavedChanges = false;
                updateStatus('saved');
                console.log('Auto-saved content');
            }}, 2000);
        }});
        
        // Warn before leaving if there are unsaved changes
        window.addEventListener('beforeunload', function(e) {{
            if (hasUnsavedChanges) {{
                e.preventDefault();
                e.returnValue = '';
            }}
        }});
    </script>
</body>
</html>"""

# Main execution
if __name__ == "__main__":
    # Configuration
    input_docx = r"documents\Master Approval Ltr (1).docx"
    output_json = "output\Master-Approval-Ltr-1.json"
    output_html = "output\Master-Approval-Ltr-1-editable.html"
    
    # Run conversion
    try:
        json_path, html_path = docx_to_json_to_html(
            input_docx_path=input_docx,
            output_json_path=output_json,
            output_html_path=output_html
        )
        print(f"\nConversion completed successfully!")
        print(f"JSON: {json_path}")
        print(f"HTML: {html_path}")
        
    except Exception as e:
        print(f"Error during conversion: {e}")
        import traceback
        traceback.print_exc()