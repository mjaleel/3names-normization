import streamlit as st
import pandas as pd
import re
from rapidfuzz import fuzz
from openpyxl.styles import PatternFill
from io import BytesIO

# إعدادات الصفحة
st.set_page_config(page_title="نظام مطابقة الرواتب والـ IBAN", layout="wide")

st.title("📊 نظام مطابقة بيانات العقاري والبصرة")
st.write("قم برفع ملفات الـ Excel لمطابقة الأسماء وجلب الـ IBAN تلقائياً.")

# --- الدوال الأساسية (نفس منطقك الأصلي) ---
def normalize_name(name):
    if pd.isnull(name): return ""
    name = str(name).strip()
    name = name.replace("ه","ة").replace("أ","ا").replace("إ","ا").replace("آ","ا")
    name = name.replace("ى","ي").replace("ئ","ي")
    name = re.sub(r'(عبد)([^\s])', r'\1 \2', name)
    return " ".join(name.split()).lower()

def get_first_three_words(name):
    if pd.isnull(name) or name == "": return ""
    words = str(name).split()
    return " ".join(words[:3]) if len(words) >= 3 else " ".join(words)

# --- واجهة رفع الملفات ---
col1, col2 = st.columns(2)
with col1:
    file_basra = st.file_uploader("رفع ملف (بصرة شهر 4)", type=['xlsx'])
with col2:
    file_aqari = st.file_uploader("رفع ملف (امانات العقاري)", type=['xlsx'])

if file_basra and file_aqari:
    if st.button("بدء عملية المطابقة وتوليد الملف"):
        with st.spinner("جاري المعالجة والمطابقة..."):
            # قراءة الملفات
            df1 = pd.read_excel(file_basra)
            df2 = pd.read_excel(file_aqari)

            # المعالجة
            df1["norm_name"]   = df1["الاسم"].apply(normalize_name)
            df1["three_word"]  = df1["norm_name"].apply(get_first_three_words)
            df1["norm_school"] = df1["القسم"].apply(normalize_name)

            df2["norm_name"]   = df2["اسم المنتسب"].apply(normalize_name)
            df2["three_word"]  = df2["norm_name"].apply(get_first_three_words)
            df2["norm_school"] = df2["المدرسة"].apply(normalize_name)

            matches = []
            for _, db_row in df2.iterrows():
                db_three  = db_row["three_word"]
                db_school = db_row["norm_school"]
                candidates = df1[df1["three_word"] == db_three]

                if len(candidates) == 0:
                    matches.append({
                        "اسم المنتسب (قاعدة)": db_row["اسم المنتسب"],
                        "المدرسة (قاعدة)":     db_row["المدرسة"],
                        "الاسم المطابق (ملف)": "",
                        "القسم (ملف)":         "",
                        "نسبة تطابق المدرسة":  "",
                        "القسط الثابت":        db_row["القسط الثابت"],
                        "اسم المصرف":          db_row["اسم المصرف"],
                        "IBAN":                 "",
                        "ملاحظة":              "❌ لا يوجد اسم مطابق"
                    })
                    continue

                best_score, best_row = 0, None
                for _, c_row in candidates.iterrows():
                    sc = fuzz.ratio(db_school, c_row["norm_school"])
                    if sc > best_score:
                        best_score, best_row = sc, c_row

                school_ok = (best_score >= 85 or db_school in best_row["norm_school"] or best_row["norm_school"] in db_school)
                note = "✅ اسم + مدرسة" if school_ok else "⚠️ اسم فقط — مدرسة مختلفة"
                iban = best_row["Iban"] if (school_ok and "Iban" in best_row) else ""

                matches.append({
                    "اسم المنتسب (قاعدة)": db_row["اسم المنتسب"],
                    "المدرسة (قاعدة)":     db_row["المدرسة"],
                    "الاسم المطابق (ملف)": best_row["الاسم"],
                    "القسم (ملف)":         best_row["القسم"],
                    "نسبة تطابق المدرسة":  f"{round(best_score)}%",
                    "القسط الثابت":        db_row["القسط الثابت"],
                    "اسم المصرف":          db_row["اسم المصرف"],
                    "IBAN":                 iban,
                    "ملاحظة":              note
                })

            result_df = pd.DataFrame(matches)

            # التصدير لملف Excel ملون في الذاكرة
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                result_df.to_excel(writer, index=False, sheet_name="نتائج المطابقة")
                ws = writer.sheets["نتائج المطابقة"]

                green_fill  = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
                yellow_fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
                red_fill    = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

                for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                    val = str(row[8].value) # عمود الملاحظة
                    if "✅" in val:
                        for cell in row: cell.fill = green_fill
                    elif "⚠️" in val:
                        for cell in row: cell.fill = yellow_fill
                    elif "❌" in val:
                        for cell in row: cell.fill = red_fill

            st.success("تمت عملية المطابقة بنجاح!")
            
            # عرض إحصائيات سريعة
            st.dataframe(result_df)

            # زر التحميل
            st.download_button(
                label="📥 تحميل ملف النتائج الملون",
                data=output.getvalue(),
                file_name="نتائج_المطابقة_النهائية.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
