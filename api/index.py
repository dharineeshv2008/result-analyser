import os
import io
import logging
import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file, jsonify
from werkzeug.utils import secure_filename
from dotenv import load_dotenv
from supabase import create_client, Client

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

load_dotenv()

app = Flask(__name__)
# Use a strong secret key for sessions
app.secret_key = os.environ.get('SECRET_KEY', 'result-analyser-production-9921')

# --- ENVIRONMENT VALIDATION ---
SUPABASE_URL = os.environ.get("SUPABASE_URL")
SUPABASE_KEY = os.environ.get("SUPABASE_KEY")

if not SUPABASE_URL or not SUPABASE_KEY:
    logger.error("CRITICAL: SUPABASE_URL or SUPABASE_KEY is not defined in environment variables.")
    # In a serverless environment, we want to fail gracefully but clearly
    supabase = None
else:
    try:
        supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)
        logger.info("Successfully initialized Supabase client.")
    except Exception as e:
        logger.error(f"CRITICAL: Failed to initialize Supabase client: {str(e)}")
        supabase = None

# --- CONSTANTS ---
ALLOWED_EXTENSIONS = {'xlsx'}
REQUIRED_COLUMNS = ['Name', 'Register Number / Roll No', 'Tamil', 'English', 'Maths', 'Science', 'Social Science']
SUBJECT_COLS = ['Tamil', 'English', 'Maths', 'Science', 'Social Science']

COLUMN_KEYWORDS = {
    'name': ['name', 'studentname', 'student_name'],
    'regno': ['rollno', 'regno', 'registerno', 'registrationnumber', 'regno'],
    'tamil': ['tamil'],
    'english': ['english', 'eng'],
    'maths': ['maths', 'math', 'mathematics'],
    'science': ['science', 'sci'],
    'social science': ['socialscience', 'social', 'ss', 'social_science']
}

# --- HELPERS ---
def normalize_column(name):
    if not isinstance(name, str):
        name = str(name)
    return ''.join(e for e in name.lower() if e.isalnum())

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def calculate_grade(total):
    if total >= 450: return 'A+'
    elif total >= 400: return 'A'
    elif total >= 350: return 'B'
    elif total >= 300: return 'C'
    elif total >= 250: return 'D'
    else: return 'F'

def handle_api_error(error_msg, status_code=400):
    logger.error(f"API Error: {error_msg}")
    return jsonify({"success": False, "error": str(error_msg)}), status_code

# --- ROUTES ---
@app.route('/')
def index():
    if 'user_id' in session:
        return redirect(url_for('history'))
    return redirect(url_for('login'))

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        email = request.form.get('email')
        password = request.form.get('password')
        if not email or not password:
            flash('Email and password are required', 'error')
            return render_template('login.html')
            
        try:
            if not supabase:
                raise ValueError("Database connection not available. Check environment variables.")
                
            response = supabase.auth.sign_in_with_password({"email": email, "password": password})
            session['user_id'] = response.user.id
            session['email'] = response.user.email
            return redirect(url_for('history'))
        except Exception as e:
            logger.error(f"Login failure: {str(e)}")
            flash(f'Login failed: {str(e)}', 'error')
            
    return render_template('login.html')

@app.route('/register', methods=['GET', 'POST'])
def register():
    if request.method == 'POST':
        email = request.form.get('email')
        password = request.form.get('password')
        try:
            if not supabase:
                raise ValueError("Database connection not available.")
            supabase.auth.sign_up({"email": email, "password": password})
            flash('Registration successful. Please login.', 'success')
            return redirect(url_for('login'))
        except Exception as e:
            flash(f'Registration failed: {str(e)}', 'error')
    return render_template('register.html')

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

@app.route('/dashboard')
def dashboard():
    if 'user_id' not in session:
        return redirect(url_for('login'))
    return render_template('dashboard.html', record=None)

@app.route('/process', methods=['POST'])
def process():
    """Processes an uploaded Excel file and saves results to Supabase."""
    logger.info("Received process request")
    
    if 'user_id' not in session:
        return handle_api_error("Unauthorized. Please login.", 401)

    if not supabase:
        return handle_api_error("Database connection missing.", 500)

    try:
        user_id = session['user_id']
        if 'file' not in request.files:
            return handle_api_error("No file part in request.")

        file = request.files['file']
        if file.filename == '':
            return handle_api_error("No selected file.")

        if not allowed_file(file.filename):
            return handle_api_error("Invalid file type. Only .xlsx files are supported.")

        # Read Excel directly from memory
        try:
            df = pd.read_excel(file.stream)
        except Exception as e:
            return handle_api_error(f"Failed to parse Excel file: {str(e)}")

        if df.empty:
            return handle_api_error("The uploaded file is empty.")

        # --- Intelligent Column Detection ---
        original_cols = df.columns.tolist()
        detected_mapping = {}
        std_names_ordered = sorted(COLUMN_KEYWORDS.keys(), key=lambda x: len(x), reverse=True)
        
        for orig_col in original_cols:
            norm_col = normalize_column(orig_col)
            for std_name in std_names_ordered:
                if std_name in detected_mapping: continue
                keywords = COLUMN_KEYWORDS[std_name]
                if any(normalize_column(k) in norm_col for k in keywords):
                    detected_mapping[std_name] = orig_col
                    break
        
        # Validation
        expected_standards = list(COLUMN_KEYWORDS.keys())
        missing_standards = [std for std in expected_standards if std not in detected_mapping]
        
        if missing_standards:
             display_names = {'name': 'Name', 'regno': 'Register Number', 'tamil': 'Tamil', 'english': 'English', 'maths': 'Maths', 'science': 'Science', 'social science': 'Social Science'}
             missing_display = [display_names.get(m, m) for m in missing_standards]
             return handle_api_error(f"Missing required columns: {', '.join(missing_display)}")

        # Rename and select
        df = df[[detected_mapping[std] for std in expected_standards]]
        df.columns = expected_standards
        
        # --- Data Validation ---
        if df.isnull().values.any():
            return handle_api_error("The file contains empty cells. Please fill all fields.")
        
        if df.duplicated(subset=['regno']).any():
            return handle_api_error("Duplicate Register Numbers found.")

        # Marks validation
        subject_cols_mapped = ['tamil', 'english', 'maths', 'science', 'social science']
        for col in subject_cols_mapped:
            df[col] = pd.to_numeric(df[col], errors='coerce')
            if df[col].isnull().any():
                return handle_api_error(f"Column '{col.title()}' contains non-numeric values.")
            if not df[col].between(0, 100).all():
                return handle_api_error(f"Marks in '{col.title()}' must be between 0 and 100.")

        # --- Calculations ---
        df['Total'] = df[subject_cols_mapped].sum(axis=1)
        df['Average'] = df['Total'] / 5
        df['Grade'] = df['Total'].apply(calculate_grade)

        # Database insertion logic
        record_data = {
            'filename': secure_filename(file.filename),
            'uploader_id': user_id,
            'student_count': len(df),
            'class_average': round(float(df['Average'].mean()), 2),
            'highest_score': int(df['Total'].max()),
            'fail_count': int((df['Grade'] == 'F').sum())
        }

        record_insert = supabase.table('records').insert(record_data).execute()
        if not record_insert.data:
            raise Exception("Failed to save batch record to database.")

        record_id = record_insert.data[0]['id']

        # Batch insert students
        student_records = []
        for _, row in df.iterrows():
            student_records.append({
                'record_id': record_id,
                'name': str(row['name']),
                'reg_no': str(row['regno']),
                'tamil': int(row['tamil']),
                'english': int(row['english']),
                'maths': int(row['maths']),
                'science': int(row['science']),
                'social_science': int(row['social science']),
                'total': int(row['Total']),
                'average': float(row['Average']),
                'grade': str(row['Grade'])
            })
        
        # Insert in chunks of 500 to avoid request size limits
        for i in range(0, len(student_records), 500):
            supabase.table('student_results').insert(student_records[i:i + 500]).execute()

        return jsonify({
            "success": True,
            "message": "Processed successfully",
            "redirect_url": url_for('view_record', record_id=record_id)
        })

    except Exception as e:
        logger.error(f"Processing crash: {str(e)}", exc_info=True)
        return handle_api_error(f"Server Error: {str(e)}", 500)

@app.route('/history')
def history():
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    try:
        if not supabase: raise ValueError("Database missing")
        records = supabase.table('records').select('*').order('created_at', desc=True).execute()
        return render_template('history.html', records=records.data)
    except Exception as e:
        flash(f"Error fetching history: {str(e)}", "error")
        return render_template('history.html', records=[])

@app.route('/view/<record_id>')
def view_record(record_id):
    if 'user_id' not in session:
        return redirect(url_for('login'))
        
    try:
        if not supabase: raise ValueError("Database missing")
        res_record = supabase.table('records').select('*').eq('id', record_id).execute()
        if not res_record.data:
            flash("Record not found", "error")
            return redirect(url_for('history'))

        record = res_record.data[0]
        students = supabase.table('student_results').select('*').eq('record_id', record_id).execute().data

        # Stats for charts
        subject_averages = {
            'Tamil': sum(s['tamil'] for s in students) / len(students) if students else 0,
            'English': sum(s['english'] for s in students) / len(students) if students else 0,
            'Maths': sum(s['maths'] for s in students) / len(students) if students else 0,
            'Science': sum(s['science'] for s in students) / len(students) if students else 0,
            'Social Science': sum(s['social_science'] for s in students) / len(students) if students else 0,
        }
        
        grades = [s['grade'] for s in students]
        grade_distribution = {g: grades.count(g) for g in ['A+', 'A', 'B', 'C', 'D', 'F']}

        record_bundle = {
            'info': record,
            'students': students,
            'charts': {
                'subjects': list(subject_averages.keys()),
                'averages': [round(v, 2) for v in subject_averages.values()],
                'grade_labels': list(grade_distribution.keys()),
                'grade_counts': list(grade_distribution.values())
            }
        }
        return render_template('dashboard.html', record=record_bundle)
    except Exception as e:
        flash(f"Error viewing record: {str(e)}", "error")
        return redirect(url_for('history'))

@app.route('/delete/<record_id>', methods=['POST'])
def delete_record(record_id):
    if 'user_id' not in session:
        return handle_api_error("Unauthorized", 401)
    
    try:
        if not supabase: raise ValueError("Database missing")
        supabase.table('records').delete().eq('id', record_id).execute()
        return jsonify({"success": True, "message": "Record deleted successfully"})
    except Exception as e:
        return handle_api_error(str(e))

@app.route('/export/excel/<record_id>')
def export_excel(record_id):
    if 'user_id' not in session:
        return redirect(url_for('login'))
        
    try:
        if not supabase: raise ValueError("Database missing")
        students = supabase.table('student_results').select('*').eq('record_id', record_id).execute().data
        if not students:
            flash('No data found.', 'error')
            return redirect(url_for('history'))
            
        df = pd.DataFrame(students)
        df = df[['name', 'reg_no', 'tamil', 'english', 'maths', 'science', 'social_science', 'total', 'average', 'grade']]
        df.columns = ['Name', 'Register Number', 'Tamil', 'English', 'Maths', 'Science', 'Social Science', 'Total', 'Average', 'Grade']
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
        output.seek(0)
        
        return send_file(output, as_attachment=True, download_name=f'results_{record_id[:8]}.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    except Exception as e:
        flash(f"Export failed: {str(e)}", "error")
        return redirect(url_for('history'))

# Standard export for Vercel
app = app

if __name__ == '__main__':
    app.run(debug=True)
