import streamlit as st
import pandas as pd
import re
from rapidfuzz import fuzz
from openpyxl.styles import PatternFill
from io import BytesIO

# --- الجزء الأول: الدوال الأصلية (بدون أي تغيير) ---
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

# --- واجهة Streamlit ---
st.set_page_config(page_title="المطابق الدقيق", layout="wide")
st.title("🎯 نظام المطابقة الذكي (نسخة الدقة القصوى)")

# رفع الملفات
col_files1, col_files2 = st.columns(2)
with col_files1:
    file1 = st.file_uploader("رفع ملف (بصرة شهر 4)", type=['xlsx'])
with col_files2:
    file2 = st.file_uploader("رفع ملف (امانات العقاري)", type=['xlsx'])

if file1 and file2:
    df1 = pd.read_excel(file1)
    df2 = pd.read_excel(file2)

    st.markdown("---")
    st.subheader("🛠️ تحديد الأعمدة بدقة")
    st.info("قم بتعريف الأعمدة لكي يعرف الكود أين يبحث، مهما اختلفت مسمياتها في ملفاتك.")

    c1, c2 = st.columns(2)
    with c1:
        st.write("**ملف بصرة (المصدر):**")
        name_col_1 = st.selectbox("اختر عمود (الاسم) في ملف بصرة", df1.columns, index=0)
        school_col_1 = st.selectbox("اختر عمود (القسم/المدرسة) في ملف بصرة", df1.columns, index=1)
        iban_col_name = st.selectbox("اختر عمود (IBAN) في ملف بصرة", df1.columns)

    with c2:
        st.write("**ملف العقاري (الهدف):**")
        name_col_2 = st.selectbox("اختر عمود (الاسم) في ملف العقاري", df2.columns, index=0)
        school_col_2 = st.selectbox("اختر عمود (المدرسة) في ملف العقاري", df2.columns, index=1)
        # الأعمدة الإضافية التي تريد بقاءها في النتيجة
        extra_cols_2 = st.multiselect("أعمدة إضافية من العقاري تريد الاحتفاظ بها", [c for c in df2.columns if c not in [name_col_2, school_col_2]])

    if st.button("▶️ تشغيل عملية المطابقة"):
        with st.spinner("جاري المطابقة..."):
            
            # تنفيذ نفس منطق التحضير الأصلي
            df1["_norm_name"]   = df1[name_col_1].apply(normalize_name)
            df1["_three_word"]  = df1["_norm_name"].apply(get_first_three_words)
            df1["_norm_school"] = df1[school_col_1].apply(normalize_name)

            df2["_norm_name"]   = df2[name_col_2].apply(normalize_name)
            df2["_three_word"]  = df2["_norm_name"].apply(get_first_three_words)
            df2["_norm_school"] = df2[school_col_2].apply(normalize_name)

            matches = []
            
            # حلقة المطابقة (نفس منطق كودك 100%)
            for _, db_row in df2.iterrows():
                db_three  = db_row["_three_word"]
                db_school = db_row["_norm_school"]
                candidates = df1[df1["_three_word"] == db_three]

                if len(candidates) == 0:
                    entry = {
                        "اسم المنتسب (قاعدة)": db_row[name_col_2],
                        "المدرسة (قاعدة)":     db_row[school_col_2],
                        "الاسم المطابق (ملف)": "",
                        "القسم (ملف)":         "",
                        "نسبة تطابق المدرسة":  "",
                        "IBAN":                 "",
                        "ملاحظة":              "❌ لا يوجد اسم مطابق"
                    }
                    for ec in extra_cols_2: entry[ec] = db_row[ec]
                    matches.append(entry)
                    continue

                best_score, best_row = 0, None
                for _, c_row in candidates.iterrows():
                    sc = fuzz.ratio(db_school, c_row["_norm_school"])
                    if sc > best_score:
                        best_score, best_row = sc, c_row

                school_ok = (best_score >= 85 or db_school in best_row["_norm_school"] or best_row["_norm_school"] in db_school)
                note = "✅ اسم + مدرسة" if school_ok else "⚠️ اسم فقط — مدرسة مختلفة"
                iban = best_row[iban_col_name] if school_ok else ""

                entry = {
                    "اسم المنتسب (قاعدة)": db_row[name_col_2],
                    "المدرسة (قاعدة)":     db_row[school_col_2],
                    "الاسم المطابق (ملف)": best_row[name_col_1],
                    "القسم (ملف)":         best_row[school_col_1],
                    "نسبة تطابق المدرسة":  f"{round(best_score)}%",
                    "IBAN":                 iban,
                    "ملاحظة":              note
                }
                for ec in extra_cols_2: entry[ec] = db_row[ec]
                matches.append(entry)

            result_df = pd.DataFrame(matches)

            # التصدير الملون (نفس التنسيق والألوان)
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                result_df.to_excel(writer, index=False, sheet_name="نتائج المطابقة")
                ws = writer.sheets["نتائج المطابقة"]

                green_fill  = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
                yellow_fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
                red_fill    = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

                # التلوين بناءً على عمود الملاحظة
                note_col_idx = list(result_df.columns).index("ملاحظة") + 1
                for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                    val = str(row[note_col_idx-1].value)
                    if "✅" in val:
                        for cell in row: cell.fill = green_fill
                    elif "⚠️" in val:
                        for cell in row: cell.fill = yellow_fill
                    elif "❌" in val:
                        for cell in row: cell.fill = red_fill

            st.success("تم الانتهاء من المطابقة!")
            st.download_button("📥 تحميل ملف النتائج الملون", output.getvalue(), "Matched_Results.xlsx")
