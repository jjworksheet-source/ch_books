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

# 用於全局暫存有效資料
if 'valid_data' not in st.session_state:
    st.session_state['valid_data'] = None

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
        df_duplicates = merged.loc[merged['_merge'] == 'left_only', df_filtered.columns]

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

        # Save valid_data to session_state for step 2
        st.session_state['valid_data'] = df_valid

elif step == "2. 匯入做卷老師資料":
    st.header("自動統計出卷老師份數")
    df_valid = st.session_state.get('valid_data', None)
    if df_valid is None:
        st.warning("請先在步驟一上傳並產生有效資料。")
    else:
        # 年級+卷對應老師規則
        cb_list = [
            "P1女拔_", "P1男拔_", "P1男拔_ 1小時", "P5女拔_", "P5男拔_", "P5男拔_ 1小時", "P6女拔_", "P6男拔_"
        ]
        kt_list = [
            "P1保羅_", "P1喇沙_", "P2保羅_", "P2喇沙_", "P3保羅_", "P3喇沙_", "P4保羅_", "P4喇沙_", "P5保羅_", "P5喇沙_", "P6喇沙_"
        ]
        mc_list = [
            "P2女拔_", "P2男拔_", "P2男拔_ 1小時", "P3女拔_", "P3男拔_", "P3男拔_ 1小時", "P4女拔_", "P4男拔_", "P4男拔_ 1小時"
        ]

        # 取得年級、學校、時間欄位
        grade_col = [col for col in df_valid.columns if "年級" in str(col)][0]
        school_col = [col for col in df_valid.columns if "學校" in str(col)][0]
        time_col = [col for col in df_valid.columns if "時間" in str(col)][0]

        # 產生年級+卷
        def get_grade卷(row):
            base = f"{row[grade_col]}{row[school_col]}_"
            if "1小時" in str(row[time_col]):
                return f"{base} 1小時"
            else:
                return base

        df_valid['年級+卷'] = df_valid.apply(get_grade卷, axis=1)

        # 統計每個年級+卷的學生數
        group_counts = df_valid.groupby('年級+卷').size().reset_index(name='人數')

        # 建立最終表格
        all卷 = sorted(set(cb_list + kt_list + mc_list))
        result = pd.DataFrame({'年級+卷': all卷})
        result['cb'] = 0
        result['kt'] = 0
        result['mc'] = 0

        # 填入各老師人數
        for _, row in group_counts.iterrows():
            g卷 = row['年級+卷']
            n = row['人數']
            if g卷 in cb_list:
                result.loc[result['年級+卷'] == g卷, 'cb'] = n
            if g卷 in kt_list:
                result.loc[result['年級+卷'] == g卷, 'kt'] = n
            if g卷 in mc_list:
                result.loc[result['年級+卷'] == g卷, 'mc'] = n

        result['總和'] = result[['cb', 'kt', 'mc']].sum(axis=1)
        result = result[['年級+卷', 'cb', 'kt', 'mc', '總和']]

        st.subheader("出卷老師的做卷人數統計表")
        st.dataframe(result)

        # 下載按鈕
        def to_excel(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            return output.getvalue()

        st.download_button(
            label="下載出卷老師統計表 Excel",
            data=to_excel(result),
            file_name="teacher_assignment_summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

elif step == "3. 計算老師佣金":
    st.header("計算老師佣金")
    st.info("此步驟尚未實作，請稍候。")

else:
    st.header("其他功能")
    st.info("此步驟尚未實作，請稍候。")
