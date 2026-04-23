import streamlit as st
import pandas as pd
import re
from rapidfuzz import fuzz
from openpyxl.styles import PatternFill
from io import BytesIO

# --- نفس الدوال الأساسية بدقة 100% ---
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

# --- واجهة المستخدم ---
st.title("🔄 نظام المطابقة الذكي الديناميكي")

col1, col2 = st.columns(2)
with col1:
    file_target = st.file_uploader("رفع الملف المطلوب تعبئته (مثلاً: العقاري)", type=['xlsx'])
with col2:
    file_source = st.file_uploader("رفع ملف المصدر (مثلاً: بصره)", type=['xlsx'])

if file_target and file_source:
    df_target = pd.read_excel(file_target)
    df_source = pd.read_excel(file_source)

    st.sidebar.header("⚙️ إعدادات الأعمدة")
    
    # اختيار الأعمدة من ملف الهدف (العقاري)
    name_col_t = st.sidebar.selectbox("عمود الاسم (الهدف)", df_target.columns)
    school_col_t = st.sidebar.selectbox("عمود المدرسة/القسم (الهدف)", df_target.columns)
    
    # اختيار الأعمدة من ملف المصدر (بصره)
    name_col_s = st.sidebar.selectbox("عمود الاسم (المصدر)", df_source.columns)
    school_col_s = st.sidebar.selectbox("عمود المدرسة/القسم (المصدر)", df_source.columns)
    
    # اختيار الأعمدة المراد جلبها ديناميكياً
    fetch_cols = st.sidebar.multiselect("الأعمدة المراد جلبها عند التطابق (مثلاً: Iban)", df_source.columns)

    if st.button("🚀 بدء المطابقة"):
        # 1. تحضير ملف المصدر (Normalization)
        df_source["_norm_name"] = df_source[name_col_s].apply(normalize_name)
        df_source["_three_word"] = df_source["_norm_name"].apply(get_first_three_words)
        df_source["_norm_school"] = df_source[school_col_s].apply(normalize_name)

        matches = []
        
        # 2. حلقة المطابقة (نفس المنطق الأصلي)
        for _, row_t in df_target.iterrows():
            target_name_norm = normalize_name(row_t[name_col_t])
            target_three = get_first_three_words(target_name_norm)
            target_school_norm = normalize_name(row_t[school_col_t])
            
            # البحث عن المرشحين بناءً على أول 3 كلمات
            candidates = df_source[df_source["_three_word"] == target_three]

            if len(candidates) == 0:
                res = row_t.to_dict()
                res.update({"الملاحظة": "❌ لا يوجد اسم مطابق", "نسبة تطابق المدرسة": ""})
                for c in fetch_cols: res[c] = ""
                matches.append(res)
                continue

            best_score, best_row = 0, None
            for _, row_s in candidates.iterrows():
                sc = fuzz.ratio(target_school_norm, row_s["_norm_school"])
                if sc > best_score:
                    best_score, best_row = sc, row_s

            school_ok = (best_score >= 85 or target_school_norm in best_row["_norm_school"] or best_row["_norm_school"] in target_school_norm)
            note = "✅ اسم + مدرسة" if school_ok else "⚠️ اسم فقط — مدرسة مختلفة"
            
            res = row_t.to_dict()
            res.update({
                "الاسم المطابق (ملف المصدر)": best_row[name_col_s],
                "القسم (ملف المصدر)": best_row[school_col_s],
                "نسبة تطابق المدرسة": f"{round(best_score)}%",
                "الملاحظة": note
            })
            
            # جلب الأعمدة المختارة ديناميكياً
            for c in fetch_cols:
                res[c] = best_row[c] if school_ok else ""
            
            matches.append(res)

        result_df = pd.DataFrame(matches)

        # 3. التصدير الملون (نفس التنسيق الأصلي)
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            result_df.to_excel(writer, index=False, sheet_name="نتائج المطابقة")
            ws = writer.sheets["نتائج المطابقة"]

            green_fill  = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
            yellow_fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
            red_fill    = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

            # تلوين الصفوف بناءً على عمود "الملاحظة"
            for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                val = str(row[result_df.columns.get_loc("الملاحظة")].value)
                if "✅" in val:
                    for cell in row: cell.fill = green_fill
                elif "⚠️" in val:
                    for cell in row: cell.fill = yellow_fill
                elif "❌" in val:
                    for cell in row: cell.fill = red_fill

        st.success("تمت المطابقة بنجاح!")
        st.download_button("📥 تحميل ملف النتائج الملون", output.getvalue(), "Matched_Results.xlsx")
