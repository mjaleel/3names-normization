import streamlit as st
import pandas as pd
import re
from rapidfuzz import fuzz
from openpyxl.styles import PatternFill
from io import BytesIO

# إعدادات الصفحة
st.set_page_config(page_title="المطابق الذكي الديناميكي", layout="wide")

st.title("🔄 نظام المطابقة المرن")
st.info("ارفع الملفات، اختر الأعمدة، واترك الباقي على النظام.")

# --- دوال المعالجة ---
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

# --- رفع الملفات ---
col1, col2 = st.columns(2)
with col1:
    file_a = st.file_uploader("رفع ملف (أ) - الملف المطلوب تعبئته", type=['xlsx'])
with col2:
    file_b = st.file_uploader("رفع ملف (ب) - ملف المصدر (بصرة مثلاً)", type=['xlsx'])

if file_a and file_b:
    df_a = pd.read_excel(file_a)
    df_b = pd.read_excel(file_b)

    st.divider()
    
    # --- اختيار الأعمدة ديناميكياً ---
    st.subheader("⚙️ إعدادات المطابقة")
    c1, c2, c3 = st.columns(3)
    
    with c1:
        name_col_a = st.selectbox("عمود الاسم في ملف (أ)", df_a.columns)
        school_col_a = st.selectbox("عمود المدرسة/القسم في ملف (أ)", df_a.columns)
    
    with c2:
        name_col_b = st.selectbox("عمود الاسم في ملف (ب)", df_b.columns)
        school_col_b = st.selectbox("عمود المدرسة/القسم في ملف (ب)", df_b.columns)
    
    with c3:
        target_cols_b = st.multiselect("الأعمدة المراد جلبها من ملف (ب)", [c for c in df_b.columns if c != name_col_b])

    if st.button("🚀 بدء المطابقة الذكية"):
        with st.spinner("جاري تحليل البيانات..."):
            # تجهيز البيانات
            df_b["_norm_name"] = df_b[name_col_b].apply(normalize_name)
            df_b["_three_word"] = df_b["_norm_name"].apply(get_first_three_words)
            df_b["_norm_school"] = df_b[school_col_b].apply(normalize_name)

            results = []

            for _, row_a in df_a.iterrows():
                name_a = normalize_name(row_a[name_col_a])
                three_a = get_first_three_words(name_a)
                school_a = normalize_name(row_a[school_col_a])
                
                # البحث الأولي بالأسماء الثلاثية
                candidates = df_b[df_b["_three_word"] == three_a]
                
                match_found = False
                best_row_b = None
                best_score = 0
                
                if not candidates.empty:
                    for _, row_b in candidates.iterrows():
                        score = fuzz.ratio(school_a, row_b["_norm_school"])
                        if score > best_score:
                            best_score = score
                            best_row_b = row_b
                    
                    if best_score >= 85 or school_a in best_row_b["_norm_school"] or best_row_b["_norm_school"] in school_a:
                        note = "✅ تطابق تام"
                        match_found = True
                    else:
                        note = "⚠️ الاسم مطابق - المدرسة مختلفة"
                else:
                    note = "❌ لا يوجد تطابق"

                # بناء السجل الناتج
                res_row = row_a.to_dict()
                res_row["الملاحظة"] = note
                res_row["نسبة تشابه المدرسة"] = f"{round(best_score)}%"
                
                for col in target_cols_b:
                    res_row[f"جلب_{col}"] = best_row_b[col] if (match_found and best_row_b is not None) else ""
                
                results.append(res_row)

            result_df = pd.DataFrame(results)

            # --- التصدير والتلوين ---
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                result_df.to_excel(writer, index=False, sheet_name="النتائج")
                ws = writer.sheets["النتائج"]
                
                # تعريف الألوان
                fills = {
                    "✅": PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"),
                    "⚠️": PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid"),
                    "❌": PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
                }

                # البحث عن عمود الملاحظة للتلوين
                note_idx = result_df.columns.get_loc("الملاحظة") + 1
                for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                    note_val = str(row[note_idx-1].value)
                    for char, fill in fills.items():
                        if char in note_val:
                            for cell in row: cell.fill = fill

            st.success("اكتملت المطابقة!")
            st.dataframe(result_df.head(100)) # عرض عينة
            
            st.download_button(
                label="📥 تحميل النتائج (Excel ملون)",
                data=output.getvalue(),
                file_name="Dynamic_Match_Results.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
