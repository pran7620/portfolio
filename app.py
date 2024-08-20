from flask import Flask, request, render_template_string, send_file
import pandas as pd

app = Flask(__name__)

# HTML template
template = '''
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Excel Cleaner</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            background-color: #f4f4f4;
            margin: 0;
            padding: 0;
            display: flex;
            justify-content: center;
            align-items: center;
            height: 100vh;
        }
        .container {
            background: #fff;
            border-radius: 8px;
            box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
            padding: 20px;
            width: 100%;
            max-width: 600px;
        }
        h1 {
            text-align: center;
            color: #333;
        }
        form {
            display: flex;
            flex-direction: column;
        }
        input[type="file"] {
            margin-bottom: 20px;
        }
        input[type="submit"] {
            background-color: #007bff;
            color: #fff;
            border: none;
            padding: 10px 20px;
            border-radius: 4px;
            cursor: pointer;
            font-size: 16px;
        }
        input[type="submit"]:hover {
            background-color: #0056b3;
        }
        .message {
            text-align: center;
            margin-top: 20px;
            color: #ff0000;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>Upload Excel File</h1>
        <form action="/upload" method="POST" enctype="multipart/form-data">
            <input type="file" name="file" accept=".xlsx" required>
            <input type="submit" value="Upload and Clean">
        </form>
        {% if message %}
        <div class="message">
            {{ message }}
        </div>
        {% endif %}
    </div>
</body>
</html>

'''

@app.route('/')
def index():
    return render_template_string(template)

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return 'No file part'
    
    file = request.files['file']
    
    if file.filename == '':
        return 'No selected file'
    
    if file:
        # Load the uploaded Excel file
        df = pd.read_excel(file, header=None)
        
        # Replace null, dash, and 0 values with blank
        cleaned_data = df.fillna('').replace('-', '').replace(0, '')
        
        # Save the cleaned data to a new Excel file
        output_filename = 'cleaned_data.xlsx'
        cleaned_data.to_excel(output_filename, index=False)
        
        # Send the cleaned file back to the user
        return send_file(output_filename, as_attachment=True)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
