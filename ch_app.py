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

        # 新增「年級_卷」欄位
        grade_col = [col for col in df_valid.columns if "年級" in str(col)]
        school_col = [col for col in df_valid.columns if "學校" in str(col)]
        if not grade_col or not school_col:
            st.error("找不到年級或學校欄位，請檢查檔案格式。")
            st.stop()
        grade_col = grade_col[0]
        school_col = school_col[0]

        def extract_school_short(s):
            # 去掉第一個底線，取第一個中文字（如 _喇沙_喇沙小學 -> 喇沙）
            if pd.isna(s):
                return ""
            s = str(s)
            if s.startswith("_"):
                s = s[1:]
            # 取第一個中文字（遇到底線或非中文字就停）
            result = ""
            for ch in s:
                if '\u4e00' <= ch <= '\u9fff':
                    result += ch
                elif ch == "_":
                    break
            return result

        def make_grade_juan(row):
            grade = str(row[grade_col]).strip() if not pd.isna(row[grade_col]) else ""
            school = extract_school_short(row[school_col])
            juan = f"{grade}{school}_"
            # 檢查班別是否有1小時
            class_val = str(row[class_col]) if not pd.isna(row[class_col]) else ""
            if "1小時" in class_val:
                juan += "1小時"
            return juan

        df_valid["年級_卷"] = df_valid.apply(make_grade_juan, axis=1)

        # 新增「出卷老師」欄位
        cb_list = [
            "P1女拔_", "P1男拔_", "P1男拔_1小時", "P5女拔_", "P5男拔_", "P5男拔_1小時", "P6女拔_", "P6男拔_"
        ]
        kt_list = [
            "P1保羅_", "P1喇沙_", "P2保羅_", "P2喇沙_", "P3保羅_", "P3喇沙_", "P4保羅_", "P4喇沙_", "P5保羅_", "P5喇沙_", "P6喇沙_"
        ]
        mc_list = [
            "P2女拔_", "P2男拔_", "P2男拔_1小時", "P3女拔_", "P3男拔_", "P3男拔_1小時", "P4女拔_", "P4男拔_", "P4男拔_1小時"
        ]

        def get_teacher(juan):
            if juan in cb_list:
                return "cb"
            elif juan in kt_list:
                return "kt"
            elif juan in mc_list:
                return "mc"
            else:
                return ""

        df_valid["出卷老師"] = df_valid["年級_卷"].apply(get_teacher)

        # 調整欄位順序：將「年級_卷」和「出卷老師」插入在「班別」和「上課日期」之間
        columns = list(df_valid.columns)
        class_idx = columns.index(class_col)
        date_idx = columns.index(date_col)
        # 移除新欄位，準備插入
        columns.remove("年級_卷")
        columns.remove("出卷老師")
        # 插入順序：班別後面插入年級_卷，再插入出卷老師
        columns.insert(class_idx + 1, "年級_卷")
        columns.insert(class_idx + 2, "出卷老師")
        # 重新排序
        df_valid = df_valid[columns]

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
        cb_list = [
            "P1女拔_", "P1男拔_", "P1男拔_1小時", "P5女拔_", "P5男拔_", "P5男拔_1小時", "P6女拔_", "P6男拔_"
        ]
        kt_list = [
            "P1保羅_", "P1喇沙_", "P2保羅_", "P2喇沙_", "P3保羅_", "P3喇沙_", "P4保羅_", "P4喇沙_", "P5保羅_", "P5喇沙_", "P6喇沙_"
        ]
        mc_list = [
            "P2女拔_", "P2男拔_", "P2男拔_1小時", "P3女拔_", "P3男拔_", "P3男拔_1小時", "P4女拔_", "P4男拔_", "P4男拔_1小時"
        ]

        grade_col = [col for col in df_valid.columns if "年級" in str(col)][0]
        school_col = [col for col in df_valid.columns if "學校" in str(col)][0]
        class_col = [col for col in df_valid.columns if "班別" in str(col)][0]
        time_col = [col for col in df_valid.columns if "時間" in str(col)][0]

        # 自動萃取學校簡稱
        def extract_short(s):
            if pd.isna(s):
                return ""
            if "男拔" in s:
                return "男拔"
            if "女拔" in s:
                return "女拔"
            if "保羅" in s:
                return "保羅"
            if "喇沙" in s:
                return "喇沙"
            if "英華" in s:
                return "英華"
            if "聖若瑟" in s:
                return "聖若瑟"
            if "真光" in s:
                return "真光"
            return s[:2]  # fallback

        df_valid['學校簡稱'] = df_valid[school_col].apply(extract_short)

        # 用班別欄位判斷 1小時
        def get_grade卷(row):
            base = f"{str(row[grade_col]).strip()}{str(row['學校簡稱']).strip()}_"
            if "1小時" in str(row[class_col]):
                return f"{base}1小時"
            else:
                return base

        df_valid['年級+卷'] = df_valid.apply(get_grade卷, axis=1)

        # Debug 輸出
        st.write("有效資料產生的年級+卷：", list(df_valid['年級+卷'].unique()))
        st.write("cb_list:", cb_list)
        st.write("kt_list:", kt_list)
        st.write("mc_list:", mc_list)

        group_counts = df_valid.groupby('年級+卷').size().reset_index(name='人數')

        all卷 = sorted(set(cb_list + kt_list + mc_list))
        result = pd.DataFrame({'年級+卷': all卷})
        result['cb'] = 0
        result['kt'] = 0
        result['mc'] = 0

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
