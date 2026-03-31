import os
import io
import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file
from werkzeug.utils import secure_filename
from dotenv import load_dotenv
from supabase import create_client, Client

load_dotenv()

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'default-dev-key')
app.config['UPLOAD_FOLDER'] = os.path.join(app.root_path, 'uploads')
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16 MB

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

SUPABASE_URL = os.environ.get("SUPABASE_URL")
SUPABASE_KEY = os.environ.get("SUPABASE_KEY")
supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)

ALLOWED_EXTENSIONS = {'xlsx'}
REQUIRED_COLUMNS = ['Name', 'Register Number / Roll No', 'Tamil', 'English', 'Maths', 'Science', 'Social Science']
SUBJECT_COLS = ['Tamil', 'English', 'Maths', 'Science', 'Social Science']

# Keyword groups for intelligent detection
COLUMN_KEYWORDS = {
    'name': ['name', 'studentname', 'student_name'],
    'regno': ['rollno', 'regno', 'registerno', 'registrationnumber', 'regno'],
    'tamil': ['tamil'],
    'english': ['english', 'eng'],
    'maths': ['maths', 'math', 'mathematics'],
    'science': ['science', 'sci'],
    'social science': ['socialscience', 'social', 'ss', 'social_science']
}

def normalize_column(name):
    """Lowercase and remove all non-alphanumeric characters."""
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
        try:
            response = supabase.auth.sign_in_with_password({"email": email, "password": password})
            session['user_id'] = response.user.id
            session['email'] = response.user.email
            return redirect(url_for('history'))
        except Exception as e:
            flash(f'Login failed: {str(e)}', 'error')
    return render_template('login.html')

@app.route('/register', methods=['GET', 'POST'])
def register():
    if request.method == 'POST':
        email = request.form.get('email')
        password = request.form.get('password')
        try:
            response = supabase.auth.sign_up({"email": email, "password": password})
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
    if 'user_id' not in session:
        return redirect(url_for('login'))

    user_id = session['user_id']
    if 'file' not in request.files:
        flash('No file part', 'error')
        return redirect(url_for('dashboard'))

    file = request.files['file']
    if file.filename == '':
        flash('No selected file', 'error')
        return redirect(url_for('dashboard'))

    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)

        try:
            df = pd.read_excel(filepath)
            
            # Intelligent Column Detection
            original_cols = df.columns.tolist()
            detected_mapping = {} # standard_name: original_column_name
            
            # Sort standard names by length (longest first) to catch 'social science' before 'science'
            std_names_ordered = sorted(COLUMN_KEYWORDS.keys(), key=lambda x: len(x), reverse=True)
            
            for orig_col in original_cols:
                norm_col = normalize_column(orig_col)
                for std_name in std_names_ordered:
                    if std_name in detected_mapping:
                        continue
                        
                    keywords = COLUMN_KEYWORDS[std_name]
                    # Check if any keyword matches as a substring of the column name
                    match_found = False
                    for k in keywords:
                        if normalize_column(k) in norm_col:
                            detected_mapping[std_name] = orig_col
                            match_found = True
                            break
                    if match_found:
                        break # mapped this column, move to next original column
            
            # Validation
            expected_standards = list(COLUMN_KEYWORDS.keys())
            missing_standards = [std for std in expected_standards if std not in detected_mapping]
            
            if missing_standards:
                display_names = {
                    'name': 'Name', 
                    'regno': 'Register Number / Roll No', 
                    'tamil': 'Tamil', 
                    'english': 'English', 
                    'maths': 'Maths', 
                    'science': 'Science', 
                    'social science': 'Social Science'
                }
                missing_display = [display_names.get(m, m) for m in missing_standards]
                
                # Check if subjects are missing for specific error message
                subjects = ['tamil', 'english', 'maths', 'science', 'social science']
                if any(s in missing_standards for s in subjects):
                    raise ValueError(f'Invalid column name detected. Expected subjects: Tamil, English, Maths, Science, Social Science')
                else:
                    raise ValueError(f'Missing required columns: {", ".join(missing_display)}')
            
            # Select and rename columns to standard format
            df = df[[detected_mapping[std] for std in expected_standards]]
            df.columns = expected_standards
            
            # 6. DATA VALIDATION
            if df.isnull().values.any():
                raise ValueError('The Excel file contains empty cells.')
            if df.duplicated(subset=['regno']).any():
                raise ValueError('Duplicate Register Numbers found.')
                
            subject_cols_mapped = ['tamil', 'english', 'maths', 'science', 'social science']
            for col in subject_cols_mapped:
                # Force conversion to numeric, coercing errors to NaN, which we can then catch
                df[col] = pd.to_numeric(df[col], errors='coerce')
                
                if df[col].isnull().any(): # Coerced NaN means it wasn't numeric
                     raise ValueError(f'Column "{col.title()}" contains non-numeric values.')
                if not df[col].between(0, 100).all():
                    raise ValueError(f'Marks out of range in "{col.title()}". Must be 0-100.')

            df['Total'] = df[subject_cols_mapped].sum(axis=1)
            df['Average'] = df['Total'] / 5
            df['Grade'] = df['Total'].apply(calculate_grade)

            student_count = len(df)
            class_average = float(df['Average'].mean())
            highest_score = int(df['Total'].max())
            fail_count = int((df['Grade'] == 'F').sum())

            # Insert into `records` table
            record_insert = supabase.table('records').insert({
                'filename': filename,
                'uploader_id': user_id,
                'student_count': student_count,
                'class_average': round(class_average, 2),
                'highest_score': highest_score,
                'fail_count': fail_count
            }).execute()

            if not record_insert.data:
                raise Exception("Failed to create record in database")

            record_id = record_insert.data[0]['id']

            # Insert into `student_results` in batches
            records_data = []
            for _, row in df.iterrows():
                records_data.append({
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
            
            # Sub-divide into chunks of 1000 due to REST constraints if needed, but 100 is fine directly.
            supabase.table('student_results').insert(records_data).execute()

            flash('Processing and upload successful!', 'success')
            return redirect(url_for('view_record', record_id=record_id))

        except Exception as e:
            flash(f'Error processing file: {str(e)}', 'error')
            return redirect(url_for('dashboard'))
        finally:
            if os.path.exists(filepath):
                os.remove(filepath)
    else:
        flash('Invalid file. Only .xlsx', 'error')
        return redirect(url_for('dashboard'))

@app.route('/history')
def history():
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    records = supabase.table('records').select('*').order('created_at', desc=True).execute()
    return render_template('history.html', records=records.data)

@app.route('/view/<record_id>')
def view_record(record_id):
    if 'user_id' not in session:
        return redirect(url_for('login'))
        
    res_record = supabase.table('records').select('*').eq('id', record_id).execute()
    if not res_record.data:
        flash("Record not found", "error")
        return redirect(url_for('history'))

    record = res_record.data[0]
    students = supabase.table('student_results').select('*').eq('record_id', record_id).execute().data

    # For charts
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

@app.route('/delete/<record_id>', methods=['POST'])
def delete_record(record_id):
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    supabase.table('records').delete().eq('id', record_id).execute()
    flash('Record deleted successfully', 'success')
    return redirect(url_for('history'))

@app.route('/export/excel/<record_id>')
def export_excel(record_id):
    if 'user_id' not in session:
        return redirect(url_for('login'))
        
    students = supabase.table('student_results').select('*').eq('record_id', record_id).execute().data
    if not students:
        flash('No data found.', 'error')
        return redirect(url_for('history'))
        
    df = pd.DataFrame(students)
    df = df[['name', 'reg_no', 'tamil', 'english', 'maths', 'science', 'social_science', 'total', 'average', 'grade']]
    df.columns = REQUIRED_COLUMNS + ['Total', 'Average', 'Grade']
    
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    buffer.seek(0)
    
    return send_file(buffer, as_attachment=True, download_name=f'results_{record_id[:8]}.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

if __name__ == '__main__':
    app.run(debug=True, port=5000)
