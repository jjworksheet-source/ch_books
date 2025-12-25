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
        "2. 出卷老師資料",
        "3. 分校做卷情況",
        "4. 其他"
    ]
)

# cb/kt/mc list
cb_list = [
    "P1女拔_", "P1男拔_", "P1男拔_1小時", "P5女拔_", "P5男拔_", "P5男拔_1小時", "P6女拔_", "P6男拔_", "P6男拔_1小時"
]
kt_list = [
    "P1保羅_", "P1喇沙_", "P2保羅_", "P2喇沙_", "P3保羅_", "P3喇沙_", "P4保羅_", "P4喇沙_", "P5保羅_", "P5喇沙_", "P6保羅_", "P6喇沙_"
]
mc_list = [
    "P2女拔_", "P2男拔_", "P2男拔_1小時", "P3女拔_", "P3男拔_", "P3男拔_1小時", "P4女拔_", "P4男拔_", "P4男拔_1小時"
]
all_juan_list = cb_list + kt_list + mc_list

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

        # 將「年級_卷」和「出卷老師」移到最後
        columns = [col for col in df_valid.columns if col not in ["年級_卷", "出卷老師"]]
        columns += ["年級_卷", "出卷老師"]
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

elif step == "2. 出卷老師資料":
    st.header("出卷老師資料")
    df_valid = st.session_state.get('valid_data', None)
    if df_valid is None:
        st.warning("請先在步驟一上傳並產生有效資料。")
    else:
        # 只統計三個 list 的年級_卷
        juan_types = [j for j in cb_list + kt_list + mc_list if j in df_valid["年級_卷"].unique()]
        rows = []
        for juan in juan_types:
            price = 25 if "1小時" in juan else 32
            cb_count = df_valid[(df_valid["年級_卷"] == juan) & (df_valid["出卷老師"] == "cb")].shape[0]
            kt_count = df_valid[(df_valid["年級_卷"] == juan) & (df_valid["出卷老師"] == "kt")].shape[0]
            mc_count = df_valid[(df_valid["年級_卷"] == juan) & (df_valid["出卷老師"] == "mc")].shape[0]
            cb_commission = cb_count * price
            kt_commission = kt_count * price
            mc_commission = mc_count * price
            row = {
                "年級+卷": juan,
                "單價": price,
                "cb": cb_count,
                "cb 佣金": cb_commission,
                "kt": kt_count,
                "kt 佣金": kt_commission,
                "mc": mc_count,
                "mc 佣金": mc_commission,
                "總和": cb_count + kt_count + mc_count,
                "佣金總和": cb_commission + kt_commission + mc_commission
            }
            rows.append(row)
        result = pd.DataFrame(rows)

        # 加總列
        total_row = {
            "年級+卷": "總和",
            "單價": "-",
            "cb": result["cb"].sum(),
            "cb 佣金": result["cb 佣金"].sum(),
            "kt": result["kt"].sum(),
            "kt 佣金": result["kt 佣金"].sum(),
            "mc": result["mc"].sum(),
            "mc 佣金": result["mc 佣金"].sum(),
            "總和": result["總和"].sum(),
            "佣金總和": result["佣金總和"].sum()
        }
        result = pd.concat([result, pd.DataFrame([total_row])], ignore_index=True)

        # 儲存總金額到 session_state
        st.session_state['step2_total'] = total_row["佣金總和"]

        # 顯示
        st.subheader("出卷老師的做卷人數及佣金統計表")
        st.dataframe(result)

        # 下載
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

elif step == "3. 分校做卷情況":
    st.header("分校做卷情況")
    df_valid = st.session_state.get('valid_data', None)
    if df_valid is None:
        st.warning("請先在步驟一上傳並產生有效資料。")
    else:
        # 分校清單
        branch_list = ["IRM", "KLN", "NFC", "NPC", "PEC", "SMC", "TKO", "WCC", "WNC"]
        # 檢查分校欄位
        branch_col = [col for col in df_valid.columns if "分校" in str(col)]
        if not branch_col:
            st.error("找不到分校欄位，請檢查檔案格式。")
        else:
            branch_col = branch_col[0]
            # 只統計三個 list 的年級_卷
            juan_types = [j for j in cb_list + kt_list + mc_list if j in df_valid["年級_卷"].unique()]
            rows = []
            for juan in juan_types:
                price = 25 if "1小時" in juan else 32
                row = {"年級+卷": juan, "單價": price}
                total_students = 0
                for branch in branch_list:
                    s_count = df_valid[(df_valid["年級_卷"] == juan) & (df_valid[branch_col] == branch)].shape[0]
                    row[f"{branch}_S"] = s_count
                    row[f"{branch}_P"] = s_count * price
                    total_students += s_count
                row["總和"] = total_students
                row["總和_P"] = total_students * price
                rows.append(row)
            result = pd.DataFrame(rows)

            # 加總列
            total_row = {"年級+卷": "總和", "單價": "-"}
            for branch in branch_list:
                total_row[f"{branch}_S"] = result[f"{branch}_S"].sum()
                total_row[f"{branch}_P"] = result[f"{branch}_P"].sum()
            total_row["總和"] = result["總和"].sum()
            total_row["總和_P"] = result["總和_P"].sum()
            result = pd.concat([result, pd.DataFrame([total_row])], ignore_index=True)

            # 指定欄位順序
            columns = ["年級+卷", "單價"]
            for branch in branch_list:
                columns += [f"{branch}_S", f"{branch}_P"]
            columns += ["總和", "總和_P"]
            result = result[columns]

            # 取得 step2 總金額
            step2_total = st.session_state.get('step2_total', None)
            step3_total = total_row["總和_P"]

            # 顯示
            st.subheader("分校做卷情況統計表")
            st.dataframe(result)

            # 自動比對總金額
            if step2_total is not None:
                if step2_total == step3_total:
                    st.success(f"總金額一致：{step2_total} 元")
                else:
                    st.error(f"總金額不一致！Step 2：{step2_total} 元，Step 3：{step3_total} 元，請檢查資料！")
            else:
                st.info("尚未產生 Step 2 總金額，請先執行 Step 2。")

            # 下載
            def to_excel(df):
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False)
                return output.getvalue()
            st.download_button(
                label="下載分校做卷情況統計表 Excel",
                data=to_excel(result),
                file_name="branch_assignment_summary.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

else:
    st.header("其他功能")
    st.info("此步驟尚未實作，請稍候。")
