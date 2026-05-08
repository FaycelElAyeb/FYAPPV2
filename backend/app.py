import os
import traceback

from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
from dotenv import load_dotenv

from analyzer import AcademicAnalyzer


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FRONTEND_FOLDER = os.path.abspath(os.path.join(BASE_DIR, '..', 'frontend'))

UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')
REPORTS_FOLDER = os.path.join(BASE_DIR, 'reports')

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(REPORTS_FOLDER, exist_ok=True)

load_dotenv(os.path.join(BASE_DIR, '.env'))

app = Flask(__name__)

CORS(app, resources={
    r"/api/*": {
        "origins": "*"
    }
})

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024


def get_saved_file(base_name):
    """
    Find saved uploaded file as CSV, XLSX, or XLS.
    Example: gradebook.csv / gradebook.xlsx / gradebook.xls
    """
    for ext in ['.csv', '.xlsx', '.xls']:
        path = os.path.join(UPLOAD_FOLDER, f'{base_name}{ext}')
        if os.path.exists(path):
            return path
    return None


def get_extension(filename):
    filename = filename.lower()

    if filename.endswith('.csv'):
        return '.csv'

    if filename.endswith('.xls'):
        return '.xls'

    return '.xlsx'


@app.route('/')
def home():
    return send_file(os.path.join(FRONTEND_FOLDER, 'login.html'))


@app.route('/favicon.ico')
def favicon():
    return '', 204


@app.route('/api/health', methods=['GET'])
def health():
    return jsonify({
        'status': 'ok',
        'message': 'Academic Analytics System Running'
    })


@app.route('/api/analyze', methods=['POST'])
def analyze_api():
    try:
        if 'gradebook' not in request.files:
            return jsonify({'error': 'ملف Gradebook مفقود'}), 400

        if 'analytics' not in request.files:
            return jsonify({'error': 'ملف Analytics مفقود'}), 400

        gradebook_file = request.files['gradebook']
        analytics_file = request.files['analytics']

        if gradebook_file.filename == '':
            return jsonify({'error': 'لم يتم اختيار ملف Gradebook'}), 400

        if analytics_file.filename == '':
            return jsonify({'error': 'لم يتم اختيار ملف Analytics'}), 400

        gb_ext = get_extension(gradebook_file.filename)
        an_ext = get_extension(analytics_file.filename)

        gb_path = os.path.join(UPLOAD_FOLDER, f'gradebook{gb_ext}')
        an_path = os.path.join(UPLOAD_FOLDER, f'analytics{an_ext}')

        gradebook_file.save(gb_path)
        analytics_file.save(an_path)

        analyzer = AcademicAnalyzer(gb_path, an_path)
        report = analyzer.generate_full_report()

        return jsonify(report), 200

    except Exception as e:
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/api/download-report', methods=['POST'])
def download_report():
    try:
        gradebook_path = get_saved_file('gradebook')
        analytics_path = get_saved_file('analytics')

        if not gradebook_path or not analytics_path:
            return jsonify({
                'error': 'يرجى رفع ملفات Gradebook و Analytics أولاً'
            }), 400

        analyzer = AcademicAnalyzer(gradebook_path, analytics_path)
        report_path = analyzer.export_excel_report(REPORTS_FOLDER)

        return send_file(
            report_path,
            as_attachment=True,
            download_name='academic_analytics_report.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    except Exception as e:
        traceback.print_exc()
        return jsonify({'error': f'خطأ في تنزيل التقرير: {str(e)}'}), 500


@app.route('/api/send-email', methods=['POST'])
def send_email():
    try:
        data = request.get_json()

        required_fields = [
            'student_id',
            'student_name',
            'risk_level',
            'recommendations'
        ]

        if not data or any(field not in data for field in required_fields):
            return jsonify({'error': 'بيانات غير كاملة'}), 400

        student_id = str(data['student_id']).strip()
        student_name = str(data['student_name']).strip()
        risk_level = str(data['risk_level']).strip()
        recommendations = str(data['recommendations']).strip()

        sender_email = os.environ.get('MAIL_SENDER')
        sender_password = os.environ.get('MAIL_PASSWORD')
        smtp_host = os.environ.get('MAIL_HOST', 'smtp.gmail.com')
        smtp_port = os.environ.get('MAIL_PORT', '587')
        smtp_secure = os.environ.get('MAIL_SECURE', 'starttls')

        if not sender_email or not sender_password:
            return jsonify({
                'error': 'لم يتم إعداد بريد المرسل أو كلمة المرور في backend/.env'
            }), 500

        gradebook_path = get_saved_file('gradebook')
        analytics_path = get_saved_file('analytics')

        if not gradebook_path or not analytics_path:
            return jsonify({
                'error': 'يرجى رفع ملفات Gradebook و Analytics أولاً'
            }), 400

        analyzer = AcademicAnalyzer(gradebook_path, analytics_path)

        success, message = analyzer.send_email_notification(
            student_id,
            student_name,
            risk_level,
            recommendations,
            sender_email,
            sender_password,
            smtp_host,
            smtp_port,
            smtp_secure
        )

        if success:
            return jsonify({'message': message}), 200

        return jsonify({'error': message}), 500

    except Exception as e:
        traceback.print_exc()
        return jsonify({'error': f'خطأ في إرسال البريد: {str(e)}'}), 500


@app.route('/<path:path>')
def serve_frontend(path):
    if path.startswith('api'):
        return jsonify({'error': 'Not found'}), 404

    requested_file = os.path.abspath(os.path.join(FRONTEND_FOLDER, path))

    if not requested_file.startswith(FRONTEND_FOLDER):
        return jsonify({'error': 'Access denied'}), 403

    if os.path.exists(requested_file) and os.path.isfile(requested_file):
        return send_file(requested_file)

    return send_file(os.path.join(FRONTEND_FOLDER, 'login.html'))


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)