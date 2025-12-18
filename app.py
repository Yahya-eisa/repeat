import streamlit as st
import pandas as pd
from io import BytesIO
import datetime
import pytz

st.set_page_config(page_title="🔍 مراجعة الاوردرات المكررة", layout="wide")
st.title("🔍 مراجعة الاوردرات المكررة")
st.markdown("ارفع الملف علشان تطلع الاوردرات المكررة 🔥")

uploaded_file = st.file_uploader("📤 ارفع ملف Excel", type=["xlsx"])

if uploaded_file:
    # قراءة الملف
    df = pd.read_excel(uploaded_file, engine="openpyxl", dtype=str)
    
    # البحث عن أعمدة كود الأوردر ورقم التليفون
    code_col = None
    phone_col = None
    
    for col in df.columns:
        col_lower = str(col).lower()
        if 'كود' in col_lower or 'رقم' in col_lower and 'عشوائي' in col_lower:
            code_col = col
        elif 'موبايل' in col_lower or 'تليفون' in col_lower or 'هاتف' in col_lower:
            phone_col = col
    
    if code_col and phone_col:
        # تنظيف البيانات
        df_clean = df[[code_col, phone_col]].copy()
        df_clean = df_clean.dropna(subset=[code_col, phone_col])
        df_clean[phone_col] = df_clean[phone_col].astype(str).str.strip()
        df_clean[code_col] = df_clean[code_col].astype(str).str.strip()
        
        # إزالة المكررات (نفس الكود ونفس التليفون)
        df_clean = df_clean.drop_duplicates()
        
        # البحث عن التليفونات المكررة
        phone_counts = df_clean[phone_col].value_counts()
        duplicated_phones = phone_counts[phone_counts > 1].index.tolist()
        
        if duplicated_phones:
            # استخراج الأكواد المكررة
            duplicates_df = df_clean[df_clean[phone_col].isin(duplicated_phones)].copy()
            duplicates_df = duplicates_df.sort_values(phone_col)
            
            # إضافة عمود عدد الأكواد لكل تليفون
            duplicates_df['عدد الأكواد'] = duplicates_df.groupby(phone_col)[phone_col].transform('count')
            
            st.error(f"⚠️ تم العثور على {len(duplicated_phones)} رقم تليفون مكرر!")
            st.warning(f"📊 إجمالي الأكواد المكررة: {len(duplicates_df)}")
            
            # عرض النتيجة
            st.dataframe(duplicates_df, use_container_width=True, hide_index=True)
            
            # تحميل الملف
            buffer = BytesIO()
            duplicates_df.to_excel(buffer, sheet_name='التليفونات المكررة', index=False, engine='openpyxl')
            buffer.seek(0)
            
            tz = pytz.timezone('Africa/Cairo')
            today = datetime.datetime.now(tz).strftime("%Y-%m-%d")
            file_name = f"التليفونات المكررة - {today}.xlsx"
            
            st.download_button(
                label="⬇️ تحميل التليفونات المكررة",
                data=buffer.getvalue(),
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            # ملخص إحصائي
            st.divider()
            st.subheader("📈 ملخص إحصائي")
            
            summary_df = duplicates_df.groupby(phone_col)[code_col].agg(['count', lambda x: ', '.join(x)]).reset_index()
            summary_df.columns = ['رقم التليفون', 'عدد الأكواد', 'الأكواد']
            summary_df = summary_df.sort_values('عدد الأكواد', ascending=False)
            
            st.dataframe(summary_df, use_container_width=True, hide_index=True)
            
        else:
            st.success("✅ مفيش تليفونات مكررة! كل رقم تليفون له كود واحد بس")
    
    else:
        st.error("❌ مش لاقي عمود كود الأوردر أو رقم التليفون في الملف!")
        st.info(f"الأعمدة الموجودة: {', '.join(df.columns.tolist())}")



