# app.py
from flask import Flask, request, jsonify, send_file, send_from_directory
from flask_cors import CORS
import pandas as pd
import os
from datetime import datetime

app = Flask(__name__)
CORS(app)

# In-memory storage for results (in production, use a database)
results = []

def calculate_grade(marks):
    """Calculate grade based on marks"""
    if marks >= 90:
        return "A+"
    elif marks >= 80:
        return "A"
    elif marks >= 70:
        return "B"
    elif marks >= 60:
        return "C"
    elif marks >= 50:
        return "D"
    else:
        return "F"

# Serve the frontend
@app.route('/')
def serve_frontend():
    return send_from_directory('.', 'index.html')

@app.route('/<path:path>')
def serve_static(path):
    return send_from_directory('.', path)

@app.route('/calculate_grade', methods=['POST'])
def calculate_grade_endpoint():
    """Calculate grade for a student"""
    try:
        data = request.get_json()
        name = data.get('name', '').strip()
        marks = float(data.get('marks', 0))
        
        # Validate inputs
        if not name:
            return jsonify({'error': 'Student name is required'}), 400
        
        if marks < 0 or marks > 100:
            return jsonify({'error': 'Marks must be between 0 and 100'}), 400
        
        # Calculate grade
        grade = calculate_grade(marks)
        
        # Create result object
        result = {
            'name': name,
            'marks': marks,
            'grade': grade,
            'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        }
        
        # Store result
        results.append(result)
        
        return jsonify(result)
    
    except ValueError:
        return jsonify({'error': 'Invalid marks value'}), 400
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/export_excel', methods=['GET'])
def export_excel():
    """Export all results to Excel"""
    try:
        if not results:
            return jsonify({'error': 'No results to export'}), 400
        
        # Create DataFrame
        df = pd.DataFrame(results)
        
        # Create Excel file
        filename = f"student_grades_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        filepath = os.path.join('exports', filename)
        
        # Ensure exports directory exists
        os.makedirs('exports', exist_ok=True)
        
        # Save to Excel
        df.to_excel(filepath, index=False, sheet_name='Student Grades')
        
        return send_file(
            filepath,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/get_results', methods=['GET'])
def get_results():
    """Get all stored results"""
    return jsonify(results)

@app.route('/clear_results', methods=['DELETE'])
def clear_results():
    """Clear all stored results"""
    global results
    results = []
    return jsonify({'message': 'All results cleared'})

if __name__ == '__main__':
    # Create exports directory if it doesn't exist
    os.makedirs('exports', exist_ok=True)
    app.run(debug=True, port=5000)