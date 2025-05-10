from flask import Flask, render_template, request, jsonify
import pandas as pd
import base64
import io

app = Flask(__name__)

# Mock file data (replace with actual file storage or database in production)
gk_isXlsx = False
gk_xlsxFileLookup = {}
gk_fileData = {}

def filled_cell(cell):
    """Check if a cell is filled (not empty, null, or None)."""
    return cell != '' and cell is not None

def load_file_data(filename):
    """Process XLSX file data and return as CSV string."""
    if gk_isXlsx and filename in gk_xlsxFileLookup:
        try:
            # Decode base64 file data
            file_content = base64.b64decode(gk_fileData[filename])
            # Read Excel file using pandas
            df = pd.read_excel(io.BytesIO(file_content), engine='openpyxl')
            # Remove blank rows
            df = df.dropna(how='all')
            # Filter rows with at least one filled cell
            filtered_data = df[df.apply(lambda row: any(filled_cell(cell) for cell in row), axis=1)]
            
            # Heuristic to find header row (simplified)
            header_row_index = 0
            if len(filtered_data) > 1:
                for i in range(len(filtered_data)):
                    if filtered_data.iloc[i].count() >= filtered_data.iloc[i + 1].count():
                        header_row_index = i
                        break
            
            # Convert to CSV
            csv_data = filtered_data.iloc[header_row_index:].to_csv(index=False)
            return csv_data
        except Exception as e:
            print(f"Error processing file: {e}")
            return ""
    return gk_fileData.get(filename, "")

@app.route('/')
def index():
    """Render the main portfolio page."""
    return render_template('index.html')

@app.route('/send_message', methods=['POST'])
def send_message():
    """Handle contact form submission."""
    name = request.form.get('name')
    email = request.form.get('email')
    message = request.form.get('message')
    
    if name and email and message:
        # In a real application, save to database or send email
        return jsonify({'message': 'Message sent successfully!'})
    else:
        return jsonify({'error': 'Please fill in all fields.'}), 400

@app.route('/load_file/<filename>')
def load_file(filename):
    """Load and process file data."""
    data = load_file_data(filename)
    return jsonify({'data': data})

if __name__ == '__main__':
    app.run(debug=True)