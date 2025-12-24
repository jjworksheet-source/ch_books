import streamlit as st

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
elif menu == "報銷管理":
    st.header("報銷管理模組")
    st.write("請上傳報銷單據，填寫相關資訊，並可查詢報銷紀錄。")
else:
    st.header("（預留）印刷物料預測")
    st.info("此功能尚未開放，敬請期待。")

st.markdown("---")
st.caption("© 2025 Jolly Jupiter IT Department")
