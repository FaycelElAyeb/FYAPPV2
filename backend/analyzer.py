import os
import re

import numpy as np
import pandas as pd
import xlsxwriter
import yagmail


class AcademicAnalyzer:
    def __init__(self, gradebook_path, analytics_path):
        self.gradebook_path = gradebook_path
        self.analytics_path = analytics_path

        self.gradebook_df = None
        self.analytics_df = None
        self.merged_df = None

        self.total_grade_col = None
        self.total_grade_max = 100
        self.exam_cols = []

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
            "ملف Gradebook"
        )

        self.analytics_df = self._load_file(
            self.analytics_path,
            "ملف Analytics"
        )

        self._clean_gradebook()
        self._clean_analytics()
        self._merge_data()

    # ------------------------------------------------------------------ #
    # GENERAL HELPERS
    # ------------------------------------------------------------------ #

    def _clean_columns(self, df):
        df.columns = [str(c).strip() for c in df.columns]
        return df

    def _find_col(self, df, keywords):
        for col in df.columns:
            col_text = str(col).strip().lower()
            for kw in keywords:
                if kw.lower() in col_text:
                    return col
        return None

    def _to_number(self, value):
        if value is None:
            return np.nan

        if pd.isna(value):
            return np.nan

        if isinstance(value, (int, float, np.integer, np.floating)):
            return float(value)

        text = str(value).strip()

        if text == "" or text == "-":
            return np.nan

        text = text.replace("%", "")
        text = re.sub(r"[^0-9.\-]", "", text)

        if text == "":
            return np.nan

        try:
            return float(text)
        except Exception:
            return np.nan

    def _extract_max_points(self, column_name, default=100):
        """
        Extract max points from Blackboard column names like:
        التقدير الكلي [إجمالي النقاط: حتى 40النتيجة]
        Test1 [إجمالي النقاط: 0.5النتيجة]
        """
        text = str(column_name)

        nums = re.findall(r"\d+(?:\.\d+)?", text)

        candidates = []

        for n in nums:
            try:
                value = float(n)

                # Ignore Blackboard internal IDs such as 1956597
                if 0 < value <= 200:
                    candidates.append(value)

            except Exception:
                continue

        if candidates:
            return candidates[0]

        return default

    def _normalize_score(self, value, max_points):
        n = self._to_number(value)

        if pd.isna(n):
            return np.nan

        if max_points and max_points > 0:
            return round((n / max_points) * 100, 1)

        return round(n, 1)

    def _build_name(self, row, first_col, last_col):
        first = str(row.get(first_col, "")).strip()
        last = str(row.get(last_col, "")).strip()
        name = f"{first} {last}".strip()
        return name if name else "غير متوفر"

    # ------------------------------------------------------------------ #
    # CLEAN GRADEBOOK
    # ------------------------------------------------------------------ #

    def _clean_gradebook(self):
        df = self.gradebook_df.copy()
        df = self._clean_columns(df)

        first_col = "الاسم الأول" if "الاسم الأول" in df.columns else "الاسم الاول"
        last_col = "اسم العائلة" if "اسم العائلة" in df.columns else "الاسم الأخير"

        if first_col not in df.columns or last_col not in df.columns:
            raise Exception("لم يتم العثور على أعمدة الاسم في ملف Gradebook")

        df["student_name"] = df.apply(
            lambda r: self._build_name(r, first_col, last_col),
            axis=1
        )

        if "معرف الطالب" in df.columns:
            df["student_id"] = df["معرف الطالب"].astype(str).str.strip()
        elif "اسم المستخدم" in df.columns:
            df["student_id"] = df["اسم المستخدم"].astype(str).str.strip()
        else:
            raise Exception("لم يتم العثور على معرف الطالب في ملف Gradebook")

        # Total grade column
        total_col = None

        for col in df.columns:
            c = str(col).lower()
            if "التقدير الكلي" in str(col) or "overall grade" in c:
                total_col = col
                break

        if not total_col:
            raise Exception("لم يتم العثور على عمود الدرجة الكلية في Gradebook")

        self.total_grade_col = total_col
        self.total_grade_max = self._extract_max_points(total_col, default=100)

        df["total_grade_raw"] = df[total_col].apply(self._to_number)
        df["total_grade"] = df[total_col].apply(
            lambda v: self._normalize_score(v, self.total_grade_max)
        )

        # Exam/activity columns
        exam_cols = []

        for col in df.columns:
            c = str(col).lower()

            if col == total_col:
                continue

            if any(x in c for x in [
                "test",
                "quiz",
                "assignement",
                "assignment",
                "activity",
                "اختبار"
            ]):
                exam_cols.append(col)

        self.exam_cols = exam_cols

        for col in exam_cols:
            max_points = self._extract_max_points(col, default=100)
            df[f"__norm__{col}"] = df[col].apply(
                lambda v: self._normalize_score(v, max_points)
            )

        norm_cols = [f"__norm__{c}" for c in exam_cols]

        if norm_cols:
            df["exam_avg"] = df[norm_cols].mean(axis=1, skipna=True)
            df["exam_std"] = df[norm_cols].std(axis=1, ddof=0, skipna=True)
            df["exam_min"] = df[norm_cols].min(axis=1, skipna=True)
            df["exam_max"] = df[norm_cols].max(axis=1, skipna=True)
            df["exams_below_50"] = (df[norm_cols] < 50).sum(axis=1)
        else:
            df["exam_avg"] = np.nan
            df["exam_std"] = 0
            df["exam_min"] = np.nan
            df["exam_max"] = np.nan
            df["exams_below_50"] = 0

        df["exam_count"] = len(exam_cols)

        self.gradebook_df = df

    # ------------------------------------------------------------------ #
    # CLEAN ANALYTICS
    # ------------------------------------------------------------------ #

    def _clean_analytics(self):
        df = self.analytics_df.copy()
        df = self._clean_columns(df)

        first_col = "الاسم الأول" if "الاسم الأول" in df.columns else "الاسم الاول"
        last_col = "الاسم الأخير" if "الاسم الأخير" in df.columns else "اسم العائلة"

        if first_col not in df.columns or last_col not in df.columns:
            raise Exception("لم يتم العثور على أعمدة الاسم في ملف Analytics")

        df["student_name"] = df.apply(
            lambda r: self._build_name(r, first_col, last_col),
            axis=1
        )

        if "معرف الطالب" in df.columns:
            df["student_id"] = df["معرف الطالب"].astype(str).str.strip()
        elif "اسم المستخدم" in df.columns:
            df["student_id"] = df["اسم المستخدم"].astype(str).str.strip()
        else:
            raise Exception("لم يتم العثور على معرف الطالب في ملف Analytics")

        total_col = self._find_col(df, [
            "التقدير الكلي",
            "overall grade",
            "total",
            "grade"
        ])

        missed_col = self._find_col(df, [
            "تواريخ الاستحقاق الفائتة",
            "استحقاق",
            "فائتة",
            "missed",
            "missing",
            "overdue"
        ])

        hours_col = self._find_col(df, [
            "عدد الساعات في المقرر الدراسي",
            "ساعات",
            "hours"
        ])

        last_access_col = self._find_col(df, [
            "تاريخ آخر وصول",
            "last access",
            "last login"
        ])

        days_col = self._find_col(df, [
            "عدد الأيام منذ آخر وصول",
            "days since",
            "days"
        ])

        df["total_grade_analytics"] = (
            df[total_col].apply(self._to_number)
            if total_col else np.nan
        )

        df["missed_deadlines"] = (
            df[missed_col].apply(self._to_number).fillna(0)
            if missed_col else 0
        )

        df["hours_spent"] = (
            df[hours_col].apply(self._to_number).fillna(0)
            if hours_col else 0
        )

        df["days_since_access"] = (
            df[days_col].apply(self._to_number).fillna(0)
            if days_col else 0
        )

        df["last_access"] = (
            df[last_access_col].astype(str).fillna("-")
            if last_access_col else "-"
        )

        self.analytics_df = df

    # ------------------------------------------------------------------ #
    # MERGE - IMPORTANT: ANALYTICS IS MASTER
    # ------------------------------------------------------------------ #

    def _merge_data(self):
        students = []

        gb_map = {}

        for _, row in self.gradebook_df.iterrows():
            sid = str(row.get("student_id", "")).strip()

            if sid:
                gb_map[sid] = row

        for _, an in self.analytics_df.iterrows():
            sid = str(an.get("student_id", "")).strip()
            gb = gb_map.get(sid)

            name = str(an.get("student_name", "")).strip()

            if not name and gb is not None:
                name = str(gb.get("student_name", "")).strip()

            if not name:
                name = "غير متوفر"

            # Grade: prefer normalized Gradebook, fallback Analytics
            total_grade = np.nan

            if gb is not None:
                total_grade = self._to_number(gb.get("total_grade"))

            if pd.isna(total_grade):
                total_grade = self._to_number(an.get("total_grade_analytics"))

            if pd.isna(total_grade):
                total_grade = np.nan

            # Exam stats from Gradebook
            exam_avg = np.nan
            exam_std = 0
            exam_min = np.nan
            exam_max = np.nan
            below50 = 0

            if gb is not None and self.exam_cols:
                scores = []

                for col in self.exam_cols:
                    norm_col = f"__norm__{col}"
                    v = self._to_number(gb.get(norm_col))

                    if pd.notna(v) and 0 <= v <= 100:
                        scores.append(v)

                if scores:
                    exam_avg = round(float(np.mean(scores)), 1)
                    exam_std = round(float(np.std(scores)), 1)
                    exam_min = round(float(np.min(scores)), 1)
                    exam_max = round(float(np.max(scores)), 1)
                    below50 = int(np.sum(np.array(scores) < 50))

            if pd.isna(exam_avg) and pd.notna(total_grade):
                exam_avg = total_grade
                exam_min = total_grade
                exam_max = total_grade
                exam_std = 0

            missed = self._to_number(an.get("missed_deadlines"))
            hours = self._to_number(an.get("hours_spent"))
            days = self._to_number(an.get("days_since_access"))

            students.append({
                "student_id": sid,
                "student_name": name,
                "total_grade": total_grade,
                "exam_avg": exam_avg,
                "exam_std": exam_std,
                "exam_min": exam_min,
                "exam_max": exam_max,
                "exams_below_50": below50,
                "missed_deadlines": 0 if pd.isna(missed) else missed,
                "hours_spent": 0 if pd.isna(hours) else hours,
                "days_since_access": 0 if pd.isna(days) else days,
                "last_access": str(an.get("last_access", "-")),
            })

        self.merged_df = pd.DataFrame(students)

    # ------------------------------------------------------------------ #
    # ANALYTICS LOGIC - SAME AS OLD APP.JS
    # ------------------------------------------------------------------ #

    def _calc_risk(self, grade, missed, hours, days, below50):
        score = 0

        g = 0 if pd.isna(grade) else grade

        if g < 50:
            score += 40
        elif g < 60:
            score += 30
        elif g < 70:
            score += 20
        elif g < 80:
            score += 10
        elif g < 90:
            score += 5

        if missed >= 10:
            score += 35
        elif missed >= 7:
            score += 25
        elif missed >= 5:
            score += 20
        elif missed >= 3:
            score += 12
        elif missed >= 1:
            score += 5

        if days > 21:
            score += 20
        elif days > 14:
            score += 15
        elif days > 7:
            score += 10

        if hours < 1:
            score += 15
        elif hours < 3:
            score += 7
        elif hours < 5:
            score += 3

        if below50 > 3:
            score += 10
        elif below50 > 1:
            score += 5

        return min(int(score), 100)

    def _risk_level_by_grade(self, grade):
        if pd.isna(grade) or grade <= 0:
            return "غير محدد"

        if grade >= 80:
            return "منخفض"

        if grade >= 60:
            return "متوسط"

        if grade >= 50:
            return "مرتفع"

        return "حرج"

    def _risk_color(self, level):
        colors = {
            "منخفض": "#27ae60",
            "متوسط": "#f39c12",
            "مرتفع": "#e67e22",
            "حرج": "#e74c3c",
            "غير محدد": "#95a5a6",
        }

        return colors.get(level, "#95a5a6")

    def _engagement_level(self, hours, days):
        if hours > 10 and days < 3:
            return "ممتاز"

        if hours > 5 and days < 7:
            return "جيد"

        if hours > 2 and days < 14:
            return "متوسط"

        return "ضعيف"

    def _performance_trend(self, row):
        sd = row.get("exam_std", 0) or 0
        av = row.get("exam_avg", 0) or 0
        mn = row.get("exam_min", 0) or 0
        mx = row.get("exam_max", 0) or 0

        if sd == 0:
            return "مستقر"

        cv = sd / av if av > 0 else 0

        if cv < 0.15:
            return "مستقر"

        if mx - mn > 30:
            return "متذبذب"

        if av >= 70:
            return "تحسن"

        return "تراجع"

    def _recommendations(self, row):
        recs = []

        grade = row.get("total_grade", np.nan)
        missed = row.get("missed_deadlines", 0) or 0
        hours = row.get("hours_spent", 0) or 0
        days = row.get("days_since_access", 0) or 0
        below50 = row.get("exams_below_50", 0) or 0

        if pd.isna(grade) or grade < 50:
            recs.append("⚠️ تدخل عاجل: جلسة دعم فردية")

        if missed > 3:
            recs.append("📅 متابعة الواجبات الفائتة")

        if days > 14:
            recs.append("📧 إرسال تنبيه للطالب")

        if hours < 2:
            recs.append("⏱️ زيادة وقت الدراسة")

        if below50 > 2:
            recs.append("📚 مراجعة المفاهيم الأساسية")

        if not recs:
            recs.append("✅ الطالب يسير بشكل جيد")

        return " | ".join(recs)

    def _predict_at_risk(self):
        df = self.merged_df.copy()

        risk_scores = []
        risk_levels = []
        at_risk_flags = []

        for _, row in df.iterrows():
            grade = row.get("total_grade", np.nan)
            missed = row.get("missed_deadlines", 0)
            hours = row.get("hours_spent", 0)
            days = row.get("days_since_access", 0)
            below50 = row.get("exams_below_50", 0)

            risk_score = self._calc_risk(
                grade,
                missed,
                hours,
                days,
                below50
            )

            risk_level = self._risk_level_by_grade(grade)

            at_risk = (
                risk_score >= 30 or
                (pd.notna(grade) and grade < 60)
            )

            risk_scores.append(risk_score)
            risk_levels.append(risk_level)
            at_risk_flags.append(at_risk)

        df["risk_score"] = risk_scores
        df["risk_level"] = risk_levels
        df["at_risk"] = at_risk_flags

        return df

    # ------------------------------------------------------------------ #
    # MAIN REPORT
    # ------------------------------------------------------------------ #

    def generate_full_report(self):
        df = self._predict_at_risk()

        total_students = len(df)
        at_risk_count = int(df["at_risk"].sum()) if total_students else 0

        valid_grades = df["total_grade"].dropna()
        valid_grades = valid_grades[valid_grades > 0]

        avg_grade = float(valid_grades.mean()) if len(valid_grades) else 0

        pass_rate = (
            float((valid_grades >= 60).sum() / len(valid_grades) * 100)
            if len(valid_grades) else 0
        )

        df["engagement"] = df.apply(
            lambda r: self._engagement_level(
                r["hours_spent"],
                r["days_since_access"]
            ),
            axis=1
        )

        df["trend"] = df.apply(
            self._performance_trend,
            axis=1
        )

        risk_dist = df["risk_level"].value_counts().to_dict()
        engagement_dist = df["engagement"].value_counts().to_dict()
        trend_dist = df["trend"].value_counts().to_dict()

        grade_dist = {
            "ممتاز (90-100)": int((valid_grades >= 90).sum()),
            "جيد جداً (80-89)": int(((valid_grades >= 80) & (valid_grades < 90)).sum()),
            "جيد (70-79)": int(((valid_grades >= 70) & (valid_grades < 80)).sum()),
            "مقبول (60-69)": int(((valid_grades >= 60) & (valid_grades < 70)).sum()),
            "راسب (<60)": int((valid_grades < 60).sum()),
        }

        students = []

        for _, row in df.iterrows():
            row = row.fillna(0)

            students.append({
                "name": str(row["student_name"]),
                "student_id": str(row["student_id"]),
                "total_grade": round(float(row["total_grade"]), 1),
                "exam_avg": round(float(row["exam_avg"]), 1),
                "exam_std": round(float(row["exam_std"]), 1),
                "exam_min": round(float(row["exam_min"]), 1),
                "exam_max": round(float(row["exam_max"]), 1),
                "missed_deadlines": int(row["missed_deadlines"]),
                "hours_spent": round(float(row["hours_spent"]), 2),
                "days_since_access": int(row["days_since_access"]),
                "last_access": str(row.get("last_access", "-")),
                "exams_below_50": int(row["exams_below_50"]),
                "risk_score": int(row["risk_score"]),
                "risk_level": str(row["risk_level"]),
                "risk_color": self._risk_color(str(row["risk_level"])),
                "at_risk": bool(row["at_risk"]),
                "engagement": str(row["engagement"]),
                "trend": str(row["trend"]),
                "recommendations": self._recommendations(row),
            })

        students.sort(
            key=lambda x: (
                -x["risk_score"],
                x["name"]
            )
        )

        return {
            "summary": {
                "total": total_students,
                "atRisk": at_risk_count,
                "safe": total_students - at_risk_count,
                "avgGrade": round(avg_grade, 1),
                "passRate": round(pass_rate, 1),
                "avgHours": round(float(df["hours_spent"].mean()), 1) if total_students else 0,
                "avgMissed": round(float(df["missed_deadlines"].mean()), 1) if total_students else 0,
                "avgDays": round(float(df["days_since_access"].mean()), 1) if total_students else 0,
            },
            "riskDist": risk_dist,
            "engDist": engagement_dist,
            "trendDist": trend_dist,
            "gradeDist": grade_dist,
            "students": students,
        }

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

السلام عليكم ورحمة الله وبركاته،

نحن نتابع أداءك في المقرر الحالي من خلال نظام التنبؤ المبكر بالتعثر الأكاديمي.

حالة الخطر الحالية: {risk_level}

التوصيات:
{recommendations}

يرجى التواصل مع المدرس أو المشرف الأكاديمي للحصول على الدعم اللازم.

مع خالص التحية،
جامعة القصيم
"""

            yag = yagmail.SMTP(
                user=sender_email,
                password=sender_password,
                host=smtp_host,
                port=int(smtp_port),
                smtp_ssl=str(smtp_secure).lower() == "ssl",
                smtp_starttls=str(smtp_secure).lower() == "starttls",
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
    # EXCEL EXPORT
    # ------------------------------------------------------------------ #

    def export_excel_report(self, output_dir):
        os.makedirs(output_dir, exist_ok=True)

        report = self.generate_full_report()

        path = os.path.join(
            output_dir,
            "academic_analytics_report.xlsx"
        )

        wb = xlsxwriter.Workbook(path)

        title_fmt = wb.add_format({
            "bold": True,
            "font_size": 16,
            "align": "center",
            "valign": "vcenter",
            "bg_color": "#1a237e",
            "font_color": "white",
            "border": 1
        })

        header_fmt = wb.add_format({
            "bold": True,
            "font_size": 11,
            "align": "center",
            "valign": "vcenter",
            "bg_color": "#283593",
            "font_color": "white",
            "border": 1,
            "text_wrap": True
        })

        cell_fmt = wb.add_format({
            "align": "center",
            "valign": "vcenter",
            "border": 1,
            "font_size": 10
        })

        summary_label_fmt = wb.add_format({
            "bold": True,
            "font_size": 12,
            "align": "right",
            "valign": "vcenter",
            "bg_color": "#e8eaf6",
            "border": 1
        })

        summary_val_fmt = wb.add_format({
            "bold": True,
            "font_size": 14,
            "align": "center",
            "valign": "vcenter",
            "bg_color": "#ffffff",
            "border": 1
        })

        # Sheet 1: Summary
        ws1 = wb.add_worksheet("ملخص تنفيذي")
        ws1.set_column("A:F", 22)

        ws1.merge_range(
            "A1:F1",
            "نظام التنبؤ المبكر بالتعثر الأكاديمي - تقرير شامل",
            title_fmt
        )

        s = report["summary"]

        summary_data = [
            ("إجمالي الطلاب", s["total"]),
            ("الطلاب في خطر", s["atRisk"]),
            ("الطلاب بأمان", s["safe"]),
            ("متوسط الدرجات", f"{s['avgGrade']}%"),
            ("نسبة النجاح", f"{s['passRate']}%"),
            ("متوسط ساعات الدراسة", s["avgHours"]),
            ("متوسط المهام الفائتة", s["avgMissed"]),
            ("متوسط أيام الغياب", s["avgDays"]),
        ]

        ws1.merge_range("A3:F3", "مؤشرات الأداء العامة", header_fmt)

        for i, (label, val) in enumerate(summary_data):
            row = 3 + i
            ws1.write(row, 0, label, summary_label_fmt)
            ws1.merge_range(row, 1, row, 5, val, summary_val_fmt)

        ws1.merge_range(12, 0, 12, 5, "توزيع الدرجات", header_fmt)

        for i, (k, v) in enumerate(report["gradeDist"].items()):
            ws1.write(13 + i, 0, k, summary_label_fmt)
            ws1.merge_range(13 + i, 1, 13 + i, 5, v, summary_val_fmt)

        # Sheet 2: Student details
        ws2 = wb.add_worksheet("تفاصيل الطلاب")

        headers = [
            "اسم الطالب",
            "معرف الطالب",
            "الدرجة الكلية",
            "متوسط الاختبارات",
            "أدنى درجة",
            "أعلى درجة",
            "المهام الفائتة",
            "ساعات الدراسة",
            "تاريخ آخر وصول",
            "أيام منذ آخر دخول",
            "مستوى الخطر",
            "درجة الخطر",
            "مستوى التفاعل",
            "اتجاه الأداء",
            "التوصيات"
        ]

        widths = [25, 16, 14, 18, 12, 12, 16, 16, 22, 20, 14, 14, 16, 16, 60]

        for i, (h, w) in enumerate(zip(headers, widths)):
            ws2.write(0, i, h, header_fmt)
            ws2.set_column(i, i, w)

        for r, st in enumerate(report["students"], start=1):
            vals = [
                st["name"],
                st["student_id"],
                st["total_grade"],
                st["exam_avg"],
                st["exam_min"],
                st["exam_max"],
                st["missed_deadlines"],
                st["hours_spent"],
                st["last_access"],
                st["days_since_access"],
                st["risk_level"],
                st["risk_score"],
                st["engagement"],
                st["trend"],
                st["recommendations"],
            ]

            for c, v in enumerate(vals):
                ws2.write(r, c, v, cell_fmt)

        # Sheet 3: At-risk students
        ws3 = wb.add_worksheet("الطلاب في خطر")

        for i, (h, w) in enumerate(zip(headers, widths)):
            ws3.write(0, i, h, header_fmt)
            ws3.set_column(i, i, w)

        at_risk = [s for s in report["students"] if s["at_risk"]]

        for r, st in enumerate(at_risk, start=1):
            vals = [
                st["name"],
                st["student_id"],
                st["total_grade"],
                st["exam_avg"],
                st["exam_min"],
                st["exam_max"],
                st["missed_deadlines"],
                st["hours_spent"],
                st["last_access"],
                st["days_since_access"],
                st["risk_level"],
                st["risk_score"],
                st["engagement"],
                st["trend"],
                st["recommendations"],
            ]

            for c, v in enumerate(vals):
                ws3.write(r, c, v, cell_fmt)


        wb.close()

        return path