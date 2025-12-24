import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Jolly Jupiter IT Department", layout="wide")

st.title("中文組做卷管理系統")

# Sidebar with step-by-step templates
st.sidebar.title("操作步驟")
step = st.sidebar.radio(
    "請選擇步驟",
    [
        "1. 做卷有效資料",
        "2. 匯入做卷老師資料",
        "3. 計算老師佣金",
        "4. 其他"
    ]
)

if step == "1. 做卷有效資料":
    st.header("上傳報表 (JJCustomer Report)")
    uploaded_file = st.file_uploader("請上傳 JJCustomer 報表 (xls/xlsx)", type=["xls", "xlsx"])
    if uploaded_file:
        try:
            # Read with header at row 6 (index 5)
            df = pd.read_excel(uploaded_file, header=5, dtype=str)
        except Exception as e:
            st.error(f"讀取檔案時發生錯誤: {e}")
            st.stop()

        # Filter by 班別
        class_types = [
            "etup 測考卷 - 高小",
            "etlp 測考卷 - 初小",
            "etlp 測考卷 - 初小 - 1小時",
            "etup 測考卷 - 高小 - 1小時"
        ]
        class_col = [col for col in df.columns if "班別" in str(col)]
        if not class_col:
            st.error("找不到班別欄位，請檢查檔案格式。")
            st.stop()
        class_col = class_col[0]

        df_filtered = df[df[class_col].astype(str).str.contains('|'.join(class_types), na=False)]

        # Filter by 學生出席狀況
        att_col = [col for col in df.columns if "學生出席狀況" in str(col)]
        if not att_col:
            st.error("找不到學生出席狀況欄位，請檢查檔案格式。")
            st.stop()
        att_col = att_col[0]
        df_filtered = df_filtered[df_filtered[att_col] == "出席"]

        # Find relevant columns for duplicate checking
        id_col = [col for col in df.columns if "學生編號" in str(col)][0]
        name_col = [col for col in df.columns if "學栍姓名" in str(col) or "學生姓名" in str(col)][0]
        date_col = [col for col in df.columns if "上課日期" in str(col)][0]
        time_col = [col for col in df.columns if "時間" in str(col)][0]

        # Special duplicate logic: group and keep non-請假 if present
        teacher_status_col = [col for col in df.columns if "老師出席狀況" in str(col)]
        teacher_status_col = teacher_status_col[0] if teacher_status_col else None

        group_cols = [id_col, name_col, date_col, class_col, time_col]
        if teacher_status_col:
            def pick_row(group):
                # If any row is not 請假, keep the first such row; else keep the first row
                not_leave = group[group[teacher_status_col] != "請假"]
                return not_leave.iloc[0] if not_leave.shape[0] > 0 else group.iloc[0]
            df_valid = df_filtered.groupby(group_cols, as_index=False).apply(pick_row).reset_index(drop=True)
        else:
            df_valid = df_filtered.drop_duplicates(subset=group_cols, keep='first')

        # Find duplicates (rows that would have been dropped)
        merged = df_filtered.merge(df_valid[group_cols], on=group_cols, how='left', indicator=True)
        df_duplicates = df_filtered[merged['_merge'] == 'left_only']

        st.success(f"有效資料共 {len(df_valid)} 筆，重複資料共 {len(df_duplicates)} 筆。")
        st.subheader("有效資料")
        st.dataframe(df_valid)
        st.subheader("重複資料")
        st.dataframe(df_duplicates)

        # Download buttons
        def to_excel(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            return output.getvalue()

        st.download_button(
            label="下載有效資料 Excel",
            data=to_excel(df_valid),
            file_name="valid_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.download_button(
            label="下載重複資料 Excel",
            data=to_excel(df_duplicates),
            file_name="duplicate_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

elif step == "2. 匯入做卷老師資料":
    st.header("匯入做卷老師資料")
    uploaded_teacher_file = st.file_uploader(
        "請上傳做卷老師分配表 (Excel/CSV/圖片)", 
        type=["xls", "xlsx", "csv", "png", "jpg", "jpeg"]
    )
    if uploaded_teacher_file:
        file_type = uploaded_teacher_file.type
        if "excel" in file_type or uploaded_teacher_file.name.endswith((".xls", ".xlsx")):
            try:
                df_teacher = pd.read_excel(uploaded_teacher_file, dtype=str)
                st.success("Excel 檔案已上傳！")
                st.subheader("做卷老師分配表 預覽")
                st.dataframe(df_teacher)
            except Exception as e:
                st.error(f"讀取 Excel 檔案時發生錯誤: {e}")
        elif "csv" in file_type or uploaded_teacher_file.name.endswith(".csv"):
            try:
                df_teacher = pd.read_csv(uploaded_teacher_file, dtype=str)
                st.success("CSV 檔案已上傳！")
                st.subheader("做卷老師分配表 預覽")
                st.dataframe(df_teacher)
            except Exception as e:
                st.error(f"讀取 CSV 檔案時發生錯誤: {e}")
        elif any(uploaded_teacher_file.name.lower().endswith(ext) for ext in [".png", ".jpg", ".jpeg"]):
            st.success("圖片檔案已上傳！")
            st.subheader("做卷老師分配表 圖片預覽")
            st.image(uploaded_teacher_file)
        else:
            st.warning("不支援的檔案格式，請上傳 Excel、CSV 或圖片檔案。")

elif step == "3. 計算老師佣金":
    st.header("計算老師佣金")
    st.info("此步驟尚未實作，請稍候。")

else:
    st.header("其他功能")
    st.info("此步驟尚未實作，請稍候。")
