from flask import Flask, request, jsonify
import base64
import io
import json
from mailmerge import MailMerge
import os

def create_app():
    app = Flask(__name__)

    @app.route('/', methods=['GET'])
    def home():
        return jsonify({
            'name': 'Document Mail Merge API',
            'version': '1.0.0',
            'status': 'running'
        })

    @app.route('/api/document-merge', methods=['POST'])
    def document_merge():
        """
        API endpoint for document mail merge using docx-mailmerge
        
        Expected request format:
        {
            "template": "base64EncodedWordDocumentTemplate", 
            "data": {
                "receivable_net": "1000.00",
                "date": "2024-12-18",
                "client_name": "Example Client"
            }
        }
        
        Response format:
        {
            "success": true,
            "document": "base64EncodedPopulatedDocument",
            "fields": ["receivable_net", "date", "client_name"]
        }
        """
        try:
            # Get request data
            request_data = request.get_json()
            
            # Validate required parameters
            if not request_data or 'template' not in request_data or 'data' not in request_data:
                return jsonify({
                    'success': False,
                    'error': 'Missing required parameters: template and data'
                }), 400
            
            template_base64 = request_data['template']
            data = request_data['data']
            
            # Validate data is a dictionary
            if not isinstance(data, dict):
                return jsonify({
                    'success': False,
                    'error': 'Data must be an object'
                }), 400
            
            # Validate all data values are strings
            for key, value in data.items():
                if not isinstance(value, str):
                    return jsonify({
                        'success': False,
                        'error': f"Data value for key '{key}' must be a string"
                    }), 400
            
            # Decode base64 template
            try:
                template_bytes = base64.b64decode(template_base64)
            except Exception:
                return jsonify({
                    'success': False,
                    'error': 'Invalid base64 encoded template'
                }), 400
            
            # Create in-memory document
            with io.BytesIO(template_bytes) as template_file:
                document = MailMerge(template_file)
                
                # Get available fields
                available_fields = document.get_merge_fields()
                
                # Merge data
                document.merge(**data)
                
                # Write to buffer
                output_buffer = io.BytesIO()
                document.write(output_buffer)
                
                # Get bytes and encode to base64
                output_buffer.seek(0)
                output_bytes = output_buffer.read()
                output_base64 = base64.b64encode(output_bytes).decode('utf-8')
                
                # Return success response
                return jsonify({
                    'success': True,
                    'document': output_base64,
                    'fields': list(available_fields)
                })
                
        except Exception as e:
            app.logger.error(f"Error processing document: {str(e)}")
            return jsonify({
                'success': False,
                'error': str(e)
            }), 500
    
    return app

# For local development
if __name__ == '__main__':
    app = create_app()
    port = int(os.environ.get('PORT', 8000)) 
    app.run(host='0.0.0.0', port=port, debug=False)