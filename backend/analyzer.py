import pandas as pd
import numpy as np
import os
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
    # DATA LOADING
    # ------------------------------------------------------------------ #

    def _load_file(self, file_path, file_name):
        try:
            if file_path.lower().endswith(".csv"):
                return pd.read_csv(file_path, encoding="utf-8-sig")

            return pd.read_excel(file_path, header=0)

        except Exception as e:
            raise Exception(f"خطأ في قراءة {file_name}: {str(e)}")

    def _load_data(self):
        self.gradebook_df = self._load_file(
            self.gradebook_path,
            'ملف Gradebook'
        )

        self.analytics_df = self._load_file(
            self.analytics_path,
            'ملف Analytics'
        )

        self._clean_gradebook()
        self._clean_analytics()
        self._merge_data()

    # ------------------------------------------------------------------ #
    # CLEAN GRADEBOOK
    # ------------------------------------------------------------------ #

    def _clean_gradebook(self):
        df = self.gradebook_df.copy()

        df.columns = [str(c).strip() for c in df.columns]

        # Names
        first_col = 'الاسم الأول'
        last_col = 'اسم العائلة'

        if first_col not in df.columns:
            first_col = 'الاسم الاول'

        if last_col not in df.columns:
            last_col = 'الاسم الأخير'

        df['student_name'] = (
            df[first_col].fillna('').astype(str).str.strip()
            + ' ' +
            df[last_col].fillna('').astype(str).str.strip()
        ).str.strip()

        # Student ID
        if 'معرف الطالب' in df.columns:
            df['student_id'] = (
                df['معرف الطالب']
                .astype(str)
                .str.strip()
            )

        elif 'اسم المستخدم' in df.columns:
            df['student_id'] = (
                df['اسم المستخدم']
                .astype(str)
                .str.strip()
            )

        else:
            raise Exception('لم يتم العثور على معرف الطالب')

        # Total grade
        total_col = None

        for col in df.columns:
            if 'التقدير الكلي' in str(col):
                total_col = col
                break

        if not total_col:
            for col in df.columns:
                if 'overall grade' in str(col).lower():
                    total_col = col
                    break

        if not total_col:
            raise Exception('لم يتم العثور على عمود الدرجة الكلية')

        df['total_grade'] = pd.to_numeric(
            df[total_col],
            errors='coerce'
        ).fillna(0)

        # Exam columns
        exam_cols = [
            c for c in df.columns
            if any(x in str(c).lower() for x in [
                'test',
                'quiz',
                'assignement',
                'assignment',
                'activity',
                'اختبار'
            ])
        ]

        for col in exam_cols:
            df[col] = pd.to_numeric(
                df[col],
                errors='coerce'
            ).fillna(0)

        if exam_cols:
            df['exam_avg'] = df[exam_cols].mean(axis=1)
            df['exam_std'] = df[exam_cols].std(axis=1).fillna(0)
            df['exam_min'] = df[exam_cols].min(axis=1)
            df['exam_max'] = df[exam_cols].max(axis=1)
            df['exams_below_50'] = (
                df[exam_cols] < 50
            ).sum(axis=1)

        else:
            df['exam_avg'] = 0
            df['exam_std'] = 0
            df['exam_min'] = 0
            df['exam_max'] = 0
            df['exams_below_50'] = 0

        df['exam_count'] = len(exam_cols)

        self.gradebook_df = df
        self.exam_cols = exam_cols

    # ------------------------------------------------------------------ #
    # CLEAN ANALYTICS
    # ------------------------------------------------------------------ #

    def _clean_analytics(self):
        df = self.analytics_df.copy()

        df.columns = [str(c).strip() for c in df.columns]

        first_col = 'الاسم الأول'
        last_col = 'الاسم الأخير'

        if first_col not in df.columns:
            first_col = 'الاسم الاول'

        if last_col not in df.columns:
            last_col = 'اسم العائلة'

        df['student_name'] = (
            df[first_col].fillna('').astype(str).str.strip()
            + ' ' +
            df[last_col].fillna('').astype(str).str.strip()
        ).str.strip()

        # Student ID
        if 'معرف الطالب' in df.columns:
            df['student_id'] = (
                df['معرف الطالب']
                .astype(str)
                .str.strip()
            )

        elif 'اسم المستخدم' in df.columns:
            df['student_id'] = (
                df['اسم المستخدم']
                .astype(str)
                .str.strip()
            )

        else:
            raise Exception(
                'لم يتم العثور على معرف الطالب في ملف Analytics'
            )

        missed_col = None
        hours_col = None
        days_col = None

        for col in df.columns:
            c = str(col)

            if 'فائتة' in c or 'missed' in c.lower():
                missed_col = col

            if 'ساعات' in c or 'hours' in c.lower():
                hours_col = col

            if 'آخر وصول' in c or 'days' in c.lower():
                days_col = col

        df['missed_deadlines'] = (
            pd.to_numeric(df[missed_col], errors='coerce').fillna(0)
            if missed_col else 0
        )

        df['hours_spent'] = (
            pd.to_numeric(df[hours_col], errors='coerce').fillna(0)
            if hours_col else 0
        )

        df['days_since_access'] = (
            pd.to_numeric(df[days_col], errors='coerce').fillna(0)
            if days_col else 0
        )

        self.analytics_df = df

    # ------------------------------------------------------------------ #
    # MERGE
    # ------------------------------------------------------------------ #

    def _merge_data(self):

        gb = self.gradebook_df[[
            'student_name',
            'student_id',
            'total_grade',
            'exam_avg',
            'exam_std',
            'exam_min',
            'exam_max',
            'exam_count',
            'exams_below_50'
        ]].copy()

        an = self.analytics_df[[
            'student_name',
            'student_id',
            'missed_deadlines',
            'hours_spent',
            'days_since_access'
        ]].copy()

        self.merged_df = pd.merge(
            gb,
            an,
            on='student_id',
            how='inner',
            suffixes=('_gb', '_an')
        )

        self.merged_df['student_name'] = (
            self.merged_df['student_name_gb']
            .fillna(self.merged_df['student_name_an'])
        )

        self.merged_df.drop(
            ['student_name_gb', 'student_name_an'],
            axis=1,
            inplace=True
        )

        numeric_cols = [
            'total_grade',
            'exam_avg',
            'exam_std',
            'exam_min',
            'exam_max',
            'exam_count',
            'exams_below_50',
            'missed_deadlines',
            'hours_spent',
            'days_since_access'
        ]

        for col in numeric_cols:
            self.merged_df[col] = pd.to_numeric(
                self.merged_df[col],
                errors='coerce'
            ).fillna(0)

    # ------------------------------------------------------------------ #
    # HELPERS
    # ------------------------------------------------------------------ #

    def _risk_color(self, level):

        colors = {
            'منخفض': '#27ae60',
            'متوسط': '#f39c12',
            'مرتفع': '#e67e22',
            'حرج': '#e74c3c'
        }

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

        std = row.get('exam_std', 0)
        avg = row.get('exam_avg', 0)
        mn = row.get('exam_min', 0)
        mx = row.get('exam_max', 0)

        if avg == 0:
            return 'غير محدد'

        if std == 0:
            return 'مستقر'

        cv = std / avg

        if cv < 0.15:
            return 'مستقر'

        if mx - mn > 30:
            return 'متذبذب'

        if avg >= 70:
            return 'تحسن'

        return 'تراجع'

    def _recommendations(self, row):

        recs = []

        grade = row.get('total_grade', 0)
        missed = row.get('missed_deadlines', 0)
        hours = row.get('hours_spent', 0)
        days = row.get('days_since_access', 0)
        below50 = row.get('exams_below_50', 0)

        if grade < 50:
            recs.append(
                '⚠️ تدخل عاجل: جلسة دعم فردية مع الطالب'
            )

        if missed > 3:
            recs.append(
                '📅 متابعة الواجبات الفائتة وإعادة جدولتها'
            )

        if days > 14:
            recs.append(
                '📧 إرسال تنبيه فوري للطالب لإعادة الانخراط'
            )

        if hours < 2:
            recs.append(
                '⏱️ تشجيع الطالب على زيادة وقت الدراسة'
            )

        if below50 > 2:
            recs.append(
                '📚 مراجعة المفاهيم الأساسية للمقرر'
            )

        if not recs:
            recs.append(
                '✅ الطالب يسير بشكل جيد - استمر في المتابعة'
            )

        return ' | '.join(recs)

    # ------------------------------------------------------------------ #
    # EMAIL
    # ------------------------------------------------------------------ #

    def send_email_notification(
        self,
        student_id,
        student_name,
        risk_level,
        recommendations,
        sender_email,
        sender_password,
        smtp_host,
        smtp_port,
        smtp_secure
    ):

        try:

            recipient_email = f"{student_id}@qu.edu.sa"

            subject = "تنبيه أكاديمي - حالة أدائك في المقرر"

            body = f"""
عزيزي الطالب {student_name},

حالة الخطر الحالية: {risk_level}

التوصيات:
{recommendations}

يرجى التواصل مع المدرس أو المشرف الأكاديمي.

مع خالص التحية،
جامعة القصيم
"""

            smtp_port_int = int(smtp_port)

            yag = yagmail.SMTP(
                user=sender_email,
                password=sender_password,
                host=smtp_host,
                port=smtp_port_int,
                smtp_ssl=smtp_secure.lower() == 'ssl',
                smtp_starttls=smtp_secure.lower() == 'starttls',
                timeout=30
            )

            yag.send(
                to=recipient_email,
                subject=subject,
                contents=body
            )

            return True, "تم إرسال البريد بنجاح"

        except Exception as e:
            return False, f"خطأ في إرسال البريد: {str(e)}"

    # ------------------------------------------------------------------ #
    # RISK PREDICTION
    # ------------------------------------------------------------------ #

    def _predict_at_risk(self):

        df = self.merged_df.copy()

        features = [
            'total_grade',
            'exam_avg',
            'exam_std',
            'missed_deadlines',
            'hours_spent',
            'days_since_access',
            'exams_below_50'
        ]

        X = df[features].fillna(
            df[features].median()
        )

        risk_scores = []

        for _, row in X.iterrows():

            score = 0

            g = row['total_grade']

            if g < 50:
                score += 40

            elif g < 60:
                score += 25

            elif g < 70:
                score += 15

            elif g < 80:
                score += 5

            if row['missed_deadlines'] > 5:
                score += 20

            elif row['missed_deadlines'] > 2:
                score += 10

            if row['days_since_access'] > 21:
                score += 20

            elif row['days_since_access'] > 7:
                score += 10

            if row['hours_spent'] < 1:
                score += 15

            elif row['hours_spent'] < 3:
                score += 7

            if row['exams_below_50'] > 2:
                score += 5

            risk_scores.append(min(score, 100))

        df['risk_score'] = risk_scores

        df['risk_level'] = df['risk_score'].apply(
            lambda x:
                'حرج' if x >= 70 else
                'مرتفع' if x >= 50 else
                'متوسط' if x >= 30 else
                'منخفض'
        )

        df['at_risk'] = df['risk_score'] >= 40

        return df

    # ------------------------------------------------------------------ #
    # MAIN REPORT
    # ------------------------------------------------------------------ #

    def generate_full_report(self):

        df = self._predict_at_risk()

        total_students = len(df)

        at_risk_count = int(
            df['at_risk'].sum()
        )

        avg_grade = (
            float(df['total_grade'].mean())
            if total_students else 0
        )

        pass_rate = (
            float(
                (df['total_grade'] >= 60).sum()
                / total_students * 100
            )
            if total_students else 0
        )

        df['engagement'] = df.apply(
            lambda r: self._engagement_level(
                r['hours_spent'],
                r['days_since_access']
            ),
            axis=1
        )

        df['trend'] = df.apply(
            self._performance_trend,
            axis=1
        )

        risk_dist = df['risk_level'].value_counts().to_dict()

        engagement_dist = (
            df['engagement']
            .value_counts()
            .to_dict()
        )

        trend_dist = (
            df['trend']
            .value_counts()
            .to_dict()
        )

        grade_dist = {
            'ممتاز (90-100)': int(
                (df['total_grade'] >= 90).sum()
            ),

            'جيد جداً (80-89)': int(
                (
                    (df['total_grade'] >= 80)
                    &
                    (df['total_grade'] < 90)
                ).sum()
            ),

            'جيد (70-79)': int(
                (
                    (df['total_grade'] >= 70)
                    &
                    (df['total_grade'] < 80)
                ).sum()
            ),

            'مقبول (60-69)': int(
                (
                    (df['total_grade'] >= 60)
                    &
                    (df['total_grade'] < 70)
                ).sum()
            ),

            'راسب (<60)': int(
                (df['total_grade'] < 60).sum()
            ),
        }

        students = []

        for _, row in df.iterrows():

            row = row.fillna(0)

            students.append({
                'name': str(row['student_name']),
                'student_id': str(row['student_id']),
                'total_grade': round(float(row['total_grade']), 1),
                'exam_avg': round(float(row['exam_avg']), 1),
                'exam_std': round(float(row['exam_std']), 1),
                'exam_min': round(float(row['exam_min']), 1),
                'exam_max': round(float(row['exam_max']), 1),
                'missed_deadlines': int(row['missed_deadlines']),
                'hours_spent': round(float(row['hours_spent']), 1),
                'days_since_access': int(row['days_since_access']),
                'exams_below_50': int(row['exams_below_50']),
                'risk_score': int(row['risk_score']),
                'risk_level': str(row['risk_level']),
                'risk_color': self._risk_color(
                    str(row['risk_level'])
                ),
                'at_risk': bool(row['at_risk']),
                'engagement': str(row['engagement']),
                'trend': str(row['trend']),
                'recommendations': self._recommendations(row),
            })

        students.sort(
            key=lambda x: (
                -x['risk_score'],
                x['name']
            )
        )

        return {
            'summary': {
                'total': total_students,
                'atRisk': at_risk_count,
                'safe': total_students - at_risk_count,
                'avgGrade': round(avg_grade, 1),
                'passRate': round(pass_rate, 1),
                'avgHours': round(
                    float(df['hours_spent'].mean()), 1
                ),
                'avgMissed': round(
                    float(df['missed_deadlines'].mean()), 1
                ),
                'avgDays': round(
                    float(df['days_since_access'].mean()), 1
                ),
            },

            'riskDist': risk_dist,
            'engDist': engagement_dist,
            'trendDist': trend_dist,
            'gradeDist': grade_dist,
            'students': students,
        }

    # ------------------------------------------------------------------ #
    # EXCEL EXPORT
    # ------------------------------------------------------------------ #

    def export_excel_report(self, output_dir):

        report = self.generate_full_report()

        path = os.path.join(
            output_dir,
            'academic_analytics_report.xlsx'
        )

        wb = xlsxwriter.Workbook(path)

        title_fmt = wb.add_format({
            'bold': True,
            'font_size': 16,
            'align': 'center',
            'valign': 'vcenter',
            'bg_color': '#1a237e',
            'font_color': 'white',
            'border': 1
        })

        header_fmt = wb.add_format({
            'bold': True,
            'font_size': 11,
            'align': 'center',
            'valign': 'vcenter',
            'bg_color': '#283593',
            'font_color': 'white',
            'border': 1,
            'text_wrap': True
        })

        cell_fmt = wb.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'border': 1,
            'font_size': 10
        })

        summary_label_fmt = wb.add_format({
            'bold': True,
            'font_size': 12,
            'align': 'right',
            'valign': 'vcenter',
            'bg_color': '#e8eaf6',
            'border': 1
        })

        summary_val_fmt = wb.add_format({
            'bold': True,
            'font_size': 14,
            'align': 'center',
            'valign': 'vcenter',
            'bg_color': '#ffffff',
            'border': 1
        })

        # SUMMARY SHEET
        ws1 = wb.add_worksheet('ملخص تنفيذي')

        ws1.set_column('A:F', 22)

        ws1.merge_range(
            'A1:F1',
            'نظام التنبؤ المبكر بالتعثر الأكاديمي - تقرير شامل',
            title_fmt
        )

        s = report['summary']

        summary_data = [
            ('إجمالي الطلاب', s['total']),
            ('الطلاب في خطر', s['atRisk']),
            ('الطلاب بأمان', s['safe']),
            ('متوسط الدرجات', f"{s['avgGrade']}%"),
            ('نسبة النجاح', f"{s['passRate']}%"),
            ('متوسط ساعات الدراسة', s['avgHours']),
            ('متوسط المهام الفائتة', s['avgMissed']),
            ('متوسط أيام الغياب', s['avgDays']),
        ]

        ws1.merge_range(
            'A3:F3',
            'مؤشرات الأداء العامة',
            header_fmt
        )

        for i, (label, val) in enumerate(summary_data):

            row = 3 + i

            ws1.write(
                row,
                0,
                label,
                summary_label_fmt
            )

            ws1.merge_range(
                row,
                1,
                row,
                5,
                val,
                summary_val_fmt
            )

        # GRADE DISTRIBUTION
        ws1.merge_range(
            12,
            0,
            12,
            5,
            'توزيع الدرجات',
            header_fmt
        )

        for i, (k, v) in enumerate(
            report['gradeDist'].items()
        ):

            ws1.write(
                13 + i,
                0,
                k,
                summary_label_fmt
            )

            ws1.merge_range(
                13 + i,
                1,
                13 + i,
                5,
                v,
                summary_val_fmt
            )

        wb.close()

        return path