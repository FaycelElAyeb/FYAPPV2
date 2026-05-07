import pandas as pd
import numpy as np
from sklearn.ensemble import RandomForestClassifier
from sklearn.preprocessing import StandardScaler
import os
import re
from datetime import datetime
import xlsxwriter
import yagmail


class AcademicAnalyzer:
    def __init__(self, gradebook_path, analytics_path):
        self.gradebook_path = gradebook_path
        self.analytics_path = analytics_path
        self.gradebook_df = None
        self.analytics_df = None
        self.merged_df = None
        self._load_data()

    # ------------------------------------------------------------------ #
    #  DATA LOADING
    # ------------------------------------------------------------------ #
    def _load_excel_file(self, file_path, file_name):
        """محاولة قراءة ملف Excel بطرق متعددة"""
        engines = ['openpyxl', 'xlrd', None]
        last_error = None
        
        for engine in engines:
            try:
                if engine:
                    df = pd.read_excel(file_path, header=0, engine=engine)
                else:
                    df = pd.read_excel(file_path, header=0)
                return df
            except Exception as e:
                last_error = e
                continue
        
        raise Exception(f"خطأ في قراءة {file_name}: {str(last_error)}")
    
    def _load_data(self):
        self.gradebook_df = self._load_excel_file(self.gradebook_path, 'ملف Gradebook')
        self.analytics_df = self._load_excel_file(self.analytics_path, 'ملف Analytics')
        self._clean_gradebook()
        self._clean_analytics()
        self._merge_data()

    def _clean_gradebook(self):
        df = self.gradebook_df.copy()
        df.columns = [str(c).strip() for c in df.columns]
        # Try to find student-name column
        name_col = self._find_col(df, ['student', 'name', 'اسم', 'الطالب', 'username'])
        if name_col:
            df.rename(columns={name_col: 'student_name'}, inplace=True)
        else:
            df['student_name'] = df.iloc[:, 0]

        # Try to find student ID column
        id_col = self._find_col(df, ['id', 'student_id', 'رقم', 'معرف', 'الرقم الجامعي'])
        if id_col:
            df.rename(columns={id_col: 'student_id'}, inplace=True)
        else:
            # Assume student_name is the ID if it's numeric
            df['student_id'] = df['student_name'].astype(str).str.extract(r'(\d+)').fillna(df['student_name'])

        # Find total/final grade column
        total_col = self._find_col(df, ['total', 'final', 'grade', 'الكلي', 'الإجمالي', 'التقدير', 'overall'])
        if total_col:
            df.rename(columns={total_col: 'total_grade'}, inplace=True)
        else:
            # pick last numeric column
            num_cols = df.select_dtypes(include=[np.number]).columns.tolist()
            if num_cols:
                df.rename(columns={num_cols[-1]: 'total_grade'}, inplace=True)
            else:
                df['total_grade'] = np.nan

        df['total_grade'] = pd.to_numeric(df['total_grade'], errors='coerce')

        # Collect quiz/exam columns (numeric columns that are not total_grade)
        exam_cols = [c for c in df.select_dtypes(include=[np.number]).columns if c != 'total_grade']
        df['exam_avg'] = df[exam_cols].mean(axis=1) if exam_cols else np.nan
        df['exam_std'] = df[exam_cols].std(axis=1) if exam_cols else 0
        df['exam_min'] = df[exam_cols].min(axis=1) if exam_cols else np.nan
        df['exam_max'] = df[exam_cols].max(axis=1) if exam_cols else np.nan
        df['exam_count'] = len(exam_cols)
        df['exams_below_50'] = (df[exam_cols] < 50).sum(axis=1) if exam_cols else 0
        df['exam_cols'] = [exam_cols] * len(df)

        self.gradebook_df = df
        self.exam_cols = exam_cols

    def _clean_analytics(self):
        df = self.analytics_df.copy()
        df.columns = [str(c).strip() for c in df.columns]

        name_col = self._find_col(df, ['student', 'name', 'اسم', 'الطالب', 'username'])
        if name_col:
            df.rename(columns={name_col: 'student_name'}, inplace=True)
        else:
            df['student_name'] = df.iloc[:, 0]

        # Try to find student ID column
        id_col = self._find_col(df, ['id', 'student_id', 'رقم', 'معرف', 'الرقم الجامعي'])
        if id_col:
            df.rename(columns={id_col: 'student_id'}, inplace=True)
        else:
            # Assume student_name is the ID if it's numeric
            df['student_id'] = df['student_name'].astype(str).str.extract(r'(\d+)').fillna(df['student_name'])

        total_col = self._find_col(df, ['total', 'final', 'grade', 'الكلي', 'الإجمالي', 'التقدير', 'overall'])
        if total_col and total_col != 'student_name':
            df.rename(columns={total_col: 'total_grade_analytics'}, inplace=True)

        missed_col = self._find_col(df, ['missed', 'due', 'فائتة', 'متأخر', 'missing'])
        if missed_col:
            df.rename(columns={missed_col: 'missed_deadlines'}, inplace=True)
        else:
            df['missed_deadlines'] = 0

        hours_col = self._find_col(df, ['hours', 'time', 'ساعات', 'وقت', 'duration', 'spent'])
        if hours_col:
            df.rename(columns={hours_col: 'hours_spent'}, inplace=True)
        else:
            df['hours_spent'] = 0

        days_col = self._find_col(df, ['days', 'last', 'أيام', 'آخر', 'access', 'login'])
        if days_col:
            df.rename(columns={days_col: 'days_since_access'}, inplace=True)
        else:
            df['days_since_access'] = 0

        for col in ['missed_deadlines', 'hours_spent', 'days_since_access']:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        self.analytics_df = df

    def _merge_data(self):
        gb = self.gradebook_df[['student_name', 'student_id', 'total_grade', 'exam_avg',
                                 'exam_std', 'exam_min', 'exam_max',
                                 'exam_count', 'exams_below_50']].copy()
        an = self.analytics_df[['student_name', 'student_id', 'missed_deadlines',
                                  'hours_spent', 'days_since_access']].copy()
        if 'total_grade_analytics' in self.analytics_df.columns:
            an['total_grade_analytics'] = self.analytics_df['total_grade_analytics']

        self.merged_df = pd.merge(gb, an, on='student_name', how='outer')
        # If student_id_x and student_id_y differ, prefer the one from gradebook
        if 'student_id_x' in self.merged_df.columns and 'student_id_y' in self.merged_df.columns:
            self.merged_df['student_id'] = self.merged_df['student_id_x'].fillna(self.merged_df['student_id_y'])
            self.merged_df.drop(['student_id_x', 'student_id_y'], axis=1, inplace=True)
        elif 'student_id' not in self.merged_df.columns:
            self.merged_df['student_id'] = self.merged_df['student_name'].astype(str).str.extract(r'(\d+)').fillna(self.merged_df['student_name'])
        self.merged_df['total_grade'] = pd.to_numeric(self.merged_df['total_grade'], errors='coerce')

    # ------------------------------------------------------------------ #
    #  HELPERS
    # ------------------------------------------------------------------ #
    def _find_col(self, df, keywords):
        for col in df.columns:
            col_lower = str(col).lower()
            for kw in keywords:
                if kw.lower() in col_lower:
                    return col
        return None

    def _risk_level(self, score):
        if pd.isna(score):
            return 'غير محدد'
        if score >= 80:
            return 'منخفض'
        if score >= 60:
            return 'متوسط'
        if score >= 50:
            return 'مرتفع'
        return 'حرج'

    def _risk_color(self, level):
        colors = {'منخفض': '#27ae60', 'متوسط': '#f39c12',
                  'مرتفع': '#e67e22', 'حرج': '#e74c3c', 'غير محدد': '#95a5a6'}
        return colors.get(level, '#95a5a6')

    def _engagement_level(self, hours, days):
        if hours > 10 and days < 3:
            return 'ممتاز'
        if hours > 5 and days < 7:
            return 'جيد'
        if hours > 2 and days < 14:
            return 'متوسط'
        return 'ضعيف'

    def _performance_trend(self, row):
        std = row.get('exam_std', 0) or 0
        avg = row.get('exam_avg', 0) or 0
        mn  = row.get('exam_min', 0) or 0
        mx  = row.get('exam_max', 0) or 0
        if std == 0:
            return 'مستقر'
        cv = std / avg if avg > 0 else 0
        if cv < 0.15:
            return 'مستقر'
        if mx - mn > 30:
            return 'متذبذب'
        if avg >= 70:
            return 'تحسن'
        return 'تراجع'

    def _recommendations(self, row):
        recs = []
        grade = row.get('total_grade', np.nan)
        missed = row.get('missed_deadlines', 0) or 0
        hours = row.get('hours_spent', 0) or 0
        days = row.get('days_since_access', 0) or 0
        below50 = row.get('exams_below_50', 0) or 0

        if pd.isna(grade) or grade < 50:
            recs.append('⚠️ تدخل عاجل: جلسة دعم فردية مع الطالب')
        if missed > 3:
            recs.append('📅 متابعة الواجبات الفائتة وإعادة جدولتها')
        if days > 14:
            recs.append('📧 إرسال تنبيه فوري للطالب لإعادة الانخراط')
        if hours < 2:
            recs.append('⏱️ تشجيع الطالب على زيادة وقت الدراسة')
        if below50 > 2:
            recs.append('📚 مراجعة المفاهيم الأساسية للمقرر')
        if not recs:
            recs.append('✅ الطالب يسير بشكل جيد - استمر في المتابعة')
        return ' | '.join(recs)

    def send_email_notification(self, student_id, student_name, risk_level, recommendations, sender_email, sender_password, smtp_host, smtp_port, smtp_secure):
        """إرسال بريد إلكتروني للطالب باستخدام إعدادات SMTP المخصصة"""
        try:
            recipient_email = f"{student_id}@qu.edu.sa"
            subject = f"تنبيه أكاديمي - حالة أدائك في المقرر"
            
            body = f"""
عزيزي الطالب {student_name},

نحن نتابع أداءك في المقرر الحالي من خلال نظام التنبؤ المبكر بالتعثر الأكاديمي.

حالة الخطر الحالية: {risk_level}

التوصيات المقترحة:
{recommendations}

يرجى التواصل مع المدرس أو المشرف الأكاديمي للحصول على الدعم اللازم.

مع خالص التحية،
فريق الدعم الأكاديمي
جامعة القصيم
"""
            smtp_port_int = int(smtp_port)
            smtp_ssl = smtp_secure.lower() == 'ssl'
            smtp_starttls = smtp_secure.lower() == 'starttls'
            yag = yagmail.SMTP(
                user=sender_email,
                password=sender_password,
                host=smtp_host,
                port=smtp_port_int,
                smtp_ssl=smtp_ssl,
                smtp_starttls=smtp_starttls,
                timeout=30  # إضافة timeout لتجنب التعليق
            )
            yag.send(to=recipient_email, subject=subject, contents=body)
            return True, "تم إرسال البريد بنجاح"
        except Exception as e:
            error_text = str(e)
            help_msg = ''
            if 'getaddrinfo failed' in error_text or 'Name or service not known' in error_text:
                help_msg = '\n\n🔧 تحقق من اسم مخدم SMTP والاتصال بالإنترنت. جرب أزرار الخيارات السريعة: QU Email أو Gmail أو Outlook.'
            elif 'InvalidSecondFactor' in error_text or 'Application-specific password required' in error_text:
                help_msg = '\n\n🔐 يجب استخدام كلمة مرور تطبيق Gmail من: https://myaccount.google.com/apppasswords'
            elif 'Authentication failed' in error_text or 'Invalid login credentials' in error_text:
                help_msg = '\n\n⚠️ بيانات مدخل غير صحيحة. تحقق من البريد وكلمة المرور.'
            return False, f"خطأ في إرسال البريد: {error_text}{help_msg}"

    def _predict_at_risk(self):
        df = self.merged_df.copy()
        features = ['total_grade', 'exam_avg', 'exam_std', 'missed_deadlines',
                    'hours_spent', 'days_since_access', 'exams_below_50']
        X = df[features].fillna(df[features].median())
        # Rule-based risk score (0-100, higher = more at risk)
        risk_scores = []
        for _, row in X.iterrows():
            score = 0
            g = row['total_grade']
            if g < 50:   score += 40
            elif g < 60: score += 25
            elif g < 70: score += 15
            elif g < 80: score += 5
            if row['missed_deadlines'] > 5:  score += 20
            elif row['missed_deadlines'] > 2: score += 10
            if row['days_since_access'] > 21: score += 20
            elif row['days_since_access'] > 7: score += 10
            if row['hours_spent'] < 1:  score += 15
            elif row['hours_spent'] < 3: score += 7
            if row['exams_below_50'] > 2: score += 5
            risk_scores.append(min(score, 100))
        df['risk_score'] = risk_scores
        df['risk_level'] = df['total_grade'].apply(self._risk_level)
        df['at_risk'] = df['risk_score'] >= 40
        return df

    # ------------------------------------------------------------------ #
    #  MAIN REPORT
    # ------------------------------------------------------------------ #
    def generate_full_report(self):
        df = self._predict_at_risk()
        total_students = len(df)
        at_risk_count = int(df['at_risk'].sum())
        avg_grade = float(df['total_grade'].mean()) if not df['total_grade'].isna().all() else 0
        pass_rate = float((df['total_grade'] >= 60).sum() / total_students * 100) if total_students else 0

        # Risk distribution
        risk_dist = df['risk_level'].value_counts().to_dict()

        # Engagement distribution
        df['engagement'] = df.apply(
            lambda r: self._engagement_level(r['hours_spent'], r['days_since_access']), axis=1)
        engagement_dist = df['engagement'].value_counts().to_dict()

        # Performance trend distribution
        df['trend'] = df.apply(self._performance_trend, axis=1)
        trend_dist = df['trend'].value_counts().to_dict()

        # Grade distribution buckets
        grade_dist = {
            'ممتاز (90-100)': int((df['total_grade'] >= 90).sum()),
            'جيد جداً (80-89)': int(((df['total_grade'] >= 80) & (df['total_grade'] < 90)).sum()),
            'جيد (70-79)': int(((df['total_grade'] >= 70) & (df['total_grade'] < 80)).sum()),
            'مقبول (60-69)': int(((df['total_grade'] >= 60) & (df['total_grade'] < 70)).sum()),
            'راسب (<60)': int((df['total_grade'] < 60).sum()),
        }

        # Per-student data
        students = []
        for _, row in df.iterrows():
            students.append({
                'name': str(row['student_name']),
                'student_id': str(row['student_id']),
                'total_grade': round(float(row['total_grade']), 1) if not pd.isna(row['total_grade']) else None,
                'exam_avg': round(float(row['exam_avg']), 1) if not pd.isna(row.get('exam_avg', np.nan)) else None,
                'exam_std': round(float(row['exam_std']), 1) if not pd.isna(row.get('exam_std', np.nan)) else None,
                'exam_min': round(float(row['exam_min']), 1) if not pd.isna(row.get('exam_min', np.nan)) else None,
                'exam_max': round(float(row['exam_max']), 1) if not pd.isna(row.get('exam_max', np.nan)) else None,
                'missed_deadlines': int(row['missed_deadlines']),
                'hours_spent': round(float(row['hours_spent']), 1),
                'days_since_access': int(row['days_since_access']),
                'exams_below_50': int(row['exams_below_50']),
                'risk_score': int(row['risk_score']),
                'risk_level': str(row['risk_level']),
                'risk_color': self._risk_color(str(row['risk_level'])),
                'at_risk': bool(row['at_risk']),
                'engagement': str(row['engagement']),
                'trend': str(row['trend']),
                'recommendations': self._recommendations(row),
            })

        # Sort: at-risk first, then by risk_score desc
        students.sort(key=lambda x: (-x['risk_score'], x['name']))

        return {
    'summary': {
        'total': total_students,
        'atRisk': at_risk_count,
        'safe': total_students - at_risk_count,
        'avgGrade': round(avg_grade, 1),
        'passRate': round(pass_rate, 1),
        'avgHours': round(float(df['hours_spent'].mean()), 1),
        'avgMissed': round(float(df['missed_deadlines'].mean()), 1),
        'avgDays': round(float(df['days_since_access'].mean()), 1),
    },
    'riskDist': risk_dist,
    'engDist': engagement_dist,
    'trendDist': trend_dist,
    'gradeDist': grade_dist,
    'students': students,
}

    # ------------------------------------------------------------------ #
    #  EXCEL EXPORT
    # ------------------------------------------------------------------ #
    def export_excel_report(self, output_dir):
        report = self.generate_full_report()
        path = os.path.join(output_dir, 'academic_analytics_report.xlsx')

        wb = xlsxwriter.Workbook(path)

        # Formats
        title_fmt = wb.add_format({'bold': True, 'font_size': 16, 'align': 'center',
                                    'valign': 'vcenter', 'bg_color': '#1a237e',
                                    'font_color': 'white', 'border': 1})
        header_fmt = wb.add_format({'bold': True, 'font_size': 11, 'align': 'center',
                                     'valign': 'vcenter', 'bg_color': '#283593',
                                     'font_color': 'white', 'border': 1, 'text_wrap': True})
        cell_fmt = wb.add_format({'align': 'center', 'valign': 'vcenter',
                                   'border': 1, 'font_size': 10})
        red_fmt = wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1,
                                  'bg_color': '#ffebee', 'font_color': '#c62828', 'bold': True})
        orange_fmt = wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1,
                                     'bg_color': '#fff3e0', 'font_color': '#e65100', 'bold': True})
        green_fmt = wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1,
                                    'bg_color': '#e8f5e9', 'font_color': '#1b5e20', 'bold': True})
        summary_label_fmt = wb.add_format({'bold': True, 'font_size': 12, 'align': 'right',
                                            'valign': 'vcenter', 'bg_color': '#e8eaf6', 'border': 1})
        summary_val_fmt = wb.add_format({'bold': True, 'font_size': 14, 'align': 'center',
                                          'valign': 'vcenter', 'bg_color': '#ffffff', 'border': 1})

        # ---- Sheet 1: Summary ----
        ws1 = wb.add_worksheet('ملخص تنفيذي')
        ws1.set_column('A:F', 22)
        ws1.set_row(0, 50)
        ws1.merge_range('A1:F1', 'نظام التنبؤ المبكر بالتعثر الأكاديمي - تقرير شامل', title_fmt)

        s = report['summary']
        summary_data = [
            ('إجمالي الطلاب', s['total_students']),
            ('الطلاب في خطر', s['at_risk_count']),
            ('الطلاب بأمان', s['safe_count']),
            ('متوسط الدرجات', f"{s['avg_grade']}%"),
            ('نسبة النجاح', f"{s['pass_rate']}%"),
            ('متوسط ساعات الدراسة', s['avg_hours']),
            ('متوسط المهام الفائتة', s['avg_missed']),
            ('متوسط أيام الغياب', s['avg_days_access']),
        ]
        ws1.write('A3', 'مؤشرات الأداء العامة', header_fmt)
        ws1.merge_range('A3:F3', 'مؤشرات الأداء العامة', header_fmt)
        for i, (label, val) in enumerate(summary_data):
            row = 3 + i
            ws1.write(row, 0, label, summary_label_fmt)
            ws1.merge_range(row, 1, row, 5, val, summary_val_fmt)

        # Grade distribution
        ws1.merge_range(12, 0, 12, 5, 'توزيع الدرجات', header_fmt)
        for i, (k, v) in enumerate(report['grade_distribution'].items()):
            ws1.write(13 + i, 0, k, summary_label_fmt)
            ws1.merge_range(13 + i, 1, 13 + i, 5, v, summary_val_fmt)

        # ---- Sheet 2: Student Details ----
        ws2 = wb.add_worksheet('تفاصيل الطلاب')
        ws2.set_row(0, 40)
        headers = ['اسم الطالب', 'الدرجة الكلية', 'متوسط الاختبارات',
                   'أدنى درجة', 'أعلى درجة', 'المهام الفائتة',
                   'ساعات الدراسة', 'أيام منذ آخر دخول',
                   'مستوى الخطر', 'درجة الخطر', 'مستوى التفاعل',
                   'اتجاه الأداء', 'التوصيات']
        col_widths = [25, 14, 18, 12, 12, 16, 16, 20, 14, 14, 16, 16, 60]
        for i, (h, w) in enumerate(zip(headers, col_widths)):
            ws2.write(0, i, h, header_fmt)
            ws2.set_column(i, i, w)

        for r, st in enumerate(report['students']):
            row = r + 1
            ws2.set_row(row, 30)
            risk = st['risk_level']
            fmt = red_fmt if risk == 'حرج' else (orange_fmt if risk == 'مرتفع' else
                  (cell_fmt if risk == 'منخفض' else cell_fmt))
            vals = [
                st['name'],
                st['total_grade'] if st['total_grade'] is not None else 'غير متاح',
                st['exam_avg'] if st['exam_avg'] is not None else 'غير متاح',
                st['exam_min'] if st['exam_min'] is not None else 'غير متاح',
                st['exam_max'] if st['exam_max'] is not None else 'غير متاح',
                st['missed_deadlines'],
                st['hours_spent'],
                st['days_since_access'],
                st['risk_level'],
                st['risk_score'],
                st['engagement'],
                st['trend'],
                st['recommendations'],
            ]
            for c, v in enumerate(vals):
                if c == 8:
                    ws2.write(row, c, v, fmt)
                else:
                    ws2.write(row, c, v, cell_fmt)

        # ---- Sheet 3: At-Risk Students ----
        ws3 = wb.add_worksheet('الطلاب في خطر')
        ws3.set_row(0, 40)
        ws3.merge_range('A1:M1', 'قائمة الطلاب المتعثرين - تحتاج تدخلاً فورياً', title_fmt)
        for i, (h, w) in enumerate(zip(headers, col_widths)):
            ws3.write(1, i, h, header_fmt)
            ws3.set_column(i, i, w)
        at_risk = [s for s in report['students'] if s['at_risk']]
        for r, st in enumerate(at_risk):
            row = r + 2
            ws3.set_row(row, 30)
            vals = [
                st['name'], st['total_grade'], st['exam_avg'],
                st['exam_min'], st['exam_max'], st['missed_deadlines'],
                st['hours_spent'], st['days_since_access'],
                st['risk_level'], st['risk_score'],
                st['engagement'], st['trend'], st['recommendations'],
            ]
            for c, v in enumerate(vals):
                ws3.write(row, c, v if v is not None else 'غير متاح',
                          red_fmt if c == 8 else cell_fmt)

        wb.close()
        return path
