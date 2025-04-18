from flask import Flask, render_template, request, send_file
from converter import convert_txt_to_excel
import tempfile
import os

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert():
    if 'txtfile' not in request.files:
        return "No file uploaded", 400
    
    txt_file = request.files['txtfile']
    
    if txt_file.filename == '':
        return "No selected file", 400

    try:
        # Save uploaded file to a temp location
        with tempfile.NamedTemporaryFile(delete=False, suffix=".txt") as temp_input:
            txt_file.save(temp_input.name)
            output_path = convert_txt_to_excel(temp_input.name)
        
        return send_file(output_path, as_attachment=True)
    
    except Exception as e:
        return f"An error occurred: {str(e)}", 500

if __name__ == '__main__':
    app.run(debug=True)
