import streamlit as st
import pandas as pd
from cleaner import clean_dataframe, create_excel_file

st.set_page_config(
    page_title="Excel 資料清洗工具",
    page_icon="🧹",
    layout="wide"
)

st.title("🧹 Excel 資料清洗工具")
st.markdown(
    """
    上傳 Excel 檔案後，系統會自動：

    1. 檢查 NULL、空值、N/A
    2. 檢查日期、電話、電郵、IP 格式
    3. 檢查數值異常
    4. 檢查大小寫混亂
    5. 將資料分為「完整」與「不完整」
    6. 產生統計報表
    7. 下載包含三個 Sheet 的 Excel 檔案
    """
)

uploaded_file = st.file_uploader(
    "請上傳 Excel 檔案",
    type=["xlsx", "xls"]
)

if uploaded_file is not None:

    try:
        raw_df = pd.read_excel(uploaded_file, sheet_name=0)
    except Exception as e:
        st.error("無法讀取 Excel 檔案，詳細錯誤如下：")
        st.exception(e)
        st.stop()

    st.success(f"成功讀取檔案：{uploaded_file.name}")

    st.subheader("原始資料預覽")
    st.dataframe(raw_df.head(20), use_container_width=True)

    complete_df = None
    incomplete_df = None
    summary_df = None
    excel_bytes = None

    try:
        with st.spinner("正在清洗資料..."):

            complete_df, incomplete_df, summary_df, issue_map = clean_dataframe(
                raw_df
            )

            excel_bytes = create_excel_file(
                complete_df,
                incomplete_df,
                summary_df
            )

    except Exception as e:
        st.error("清洗資料時發生錯誤，詳細錯誤如下：")
        st.exception(e)

        st.subheader("除錯資訊")
        st.write("上傳檔案名稱：", uploaded_file.name)
        st.write("資料列數：", raw_df.shape[0])
        st.write("資料欄數：", raw_df.shape[1])
        st.write("欄位名稱：")
        st.write(list(raw_df.columns))

        st.stop()

    st.subheader("📊 清洗結果統計")

    col1, col2, col3 = st.columns(3)

    col1.metric("總紀錄數", len(raw_df))
    col2.metric("完整資料", len(complete_df))
    col3.metric("不完整資料", len(incomplete_df))

    st.subheader("📋 詳細統計")
    st.dataframe(summary_df, use_container_width=True)

    st.subheader("✅ 完整資料")
    st.dataframe(complete_df, use_container_width=True)

    st.subheader("⚠️ 不完整資料")
    st.dataframe(incomplete_df, use_container_width=True)

    st.subheader("⬇️ 下載整理後的 Excel")

    st.download_button(
        label="下載 Cleaned_Org_Data.xlsx",
        data=excel_bytes,
        file_name="Cleaned_Org_Data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("請先上傳 Excel 檔案。")
