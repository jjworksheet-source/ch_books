import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Jolly Jupiter IT Department", layout="wide")

st.markdown("""
    <style>
    .main {background-color: #f5f7fa;}
    .block-container {padding-top: 2rem;}
    h1 {color: #2c3e50;}
    </style>
""", unsafe_allow_html=True)

st.title("Jolly Jupiter IT Department")
st.subheader("中文組做卷管理系統")

menu = st.sidebar.selectbox(
    "功能選單",
    ("上傳報表", "報銷管理", "（預留）印刷物料預測")
)

if menu == "上傳報表":
    st.header("upload jjcustomer report")
    uploaded_file = st.file_uploader("上傳 Excel 報表", type=["xls", "xlsx"])
    if uploaded_file:
        try:
            # 1. 正確讀取第6列為欄位名稱
            df = pd.read_excel(uploaded_file, header=5)
            st.success("檔案上傳成功！")

            # 2. 篩選班別
            class_list = [
                "etup 測考卷 - 高小",
                "etlp 測考卷 - 初小",
                "etlp 測考卷 - 初小 - 1小時",
                "etup 測考卷 - 高小 - 1小時"
            ]
            class_col = [col for col in df.columns if "班別" in str(col)]
            if class_col:
                class_col = class_col[0]
            else:
                st.error("找不到班別欄位，請檢查檔案格式。")
                st.write("所有欄位名稱：", list(df.columns))
                st.stop()
            df_filtered = df[df[class_col].astype(str).str.strip().isin(class_list)]

            # 3. 只保留出席
            attend_col = [col for col in df.columns if "出席狀況" in str(col)]
            if attend_col:
                attend_col = attend_col[0]
            else:
                st.error("找不到學生出席狀況欄位，請檢查檔案格式。")
                st.write("所有欄位名稱：", list(df.columns))
                st.stop()
            df_valid = df_filtered[df_filtered[attend_col].astype(str).str.strip() == "出席"]

            # 4. 檢查重複（加上時間欄位，並處理老師請假/代課問題）
            id_col = [col for col in df.columns if "學生編號" in str(col)]
            name_col = [col for col in df.columns if "學栍姓名" in str(col)]
            date_col = [col for col in df.columns if "上課日期" in str(col)]
            time_col = [col for col in df.columns if "時間" in str(col)]
            teacher_status_col = [col for col in df.columns if "老師出席狀況" in str(col)]
            if not (id_col and name_col and date_col and time_col and teacher_status_col):
                st.error("找不到學生編號、學栍姓名、上課日期、時間或老師出席狀況欄位，請檢查檔案格式。")
                st.write("所有欄位名稱：", list(df.columns))
                st.stop()
            id_col = id_col[0]
            name_col = name_col[0]
            date_col = date_col[0]
            time_col = time_col[0]
            teacher_status_col = teacher_status_col[0]

            group_cols = [id_col, name_col, date_col, class_col, time_col]

            def pick_row(group):
                # 先找不是請假的
                not_leave = group[group[teacher_status_col] != "請假"]
                if not_leave.shape[0] > 0:
                    return not_leave.iloc[0]
                else:
                    return group.iloc[0]

            # 先去除老師請假重複
            df_valid_nodup = df_valid.groupby(group_cols, as_index=False).apply(pick_row).reset_index(drop=True)

            # 再找真正重複（理論上已經沒有，但保險起見）
            df_duplicates = df_valid_nodup[df_valid_nodup.duplicated(subset=group_cols, keep=False)]

            st.write("## 篩選後有效資料")
            st.dataframe(df_valid_nodup)

            # 下載有效資料
            towrite = io.BytesIO()
            df_valid_nodup.to_excel(towrite, index=False)
            towrite.seek(0)
            st.download_button(
                label="下載有效資料 Excel",
                data=towrite,
                file_name="valid_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            if not df_duplicates.empty:
                st.warning(f"發現 {len(df_duplicates)} 筆重複資料（同一學生編號、姓名、上課日期、班別、時間）如下：")
                st.dataframe(df_duplicates)
                # 下載重複資料
                towrite_dup = io.BytesIO()
                df_duplicates.to_excel(towrite_dup, index=False)
                towrite_dup.seek(0)
                st.download_button(
                    label="下載重複資料 Excel",
                    data=towrite_dup,
                    file_name="duplicate_data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.success("沒有發現重複資料！")

        except Exception as e:
            st.error(f"檔案處理時發生錯誤：{e}")

elif menu == "報銷管理":
    st.header("報銷管理模組")
    st.write("請上傳報銷單據，填寫相關資訊，並可查詢報銷紀錄。")
else:
    st.header("（預留）印刷物料預測")
    st.info("此功能尚未開放，敬請期待。")

st.markdown("---")
st.caption("© 2025 Jolly Jupiter IT Department")
