import streamlit as st
import pandas as pd

st.title("事務機使用量批次統計工具")
st.write("請上傳多個 CSV 或 Excel 報表，系統會自動依據 UserID 幫您加總特定欄位的數據。")

# 允許上傳多個檔案
uploaded_files = st.file_uploader("選擇檔案 (支援 CSV 或 Excel)", type=['csv', 'xlsx'], accept_multiple_files=True)

# 我們想要統計的目標欄位
target_cols = [
    'UserID', 
    '複印(黑白)累積頁數', 
    '複印(彩色)累積頁數', 
    '列印(黑白)累積頁數', 
    '列印(彩色)累積頁數'
]

if uploaded_files:
    all_data = []
    
    # 逐一讀取上傳的檔案
    for file in uploaded_files:
        try:
            if file.name.endswith('.csv'):
                df = pd.read_csv(file)
            else:
                df = pd.read_excel(file)
                
            available_cols = [col for col in target_cols if col in df.columns]
            df = df[available_cols]
            all_data.append(df)
            
        except Exception as e:
            st.error(f"讀取檔案 {file.name} 時發生錯誤: {e}")

    if all_data:
        combined_df = pd.concat(all_data, ignore_index=True)
        
        numeric_cols = [col for col in target_cols if col != 'UserID' and col in combined_df.columns]
        for col in numeric_cols:
            combined_df[col] = pd.to_numeric(combined_df[col], errors='coerce').fillna(0)
            
        summary_df = combined_df.groupby('UserID', as_index=False).sum()
        
        st.success("統計完成！")
        st.write("預覽合併後的結果：")
        st.dataframe(summary_df)
        
        csv = summary_df.to_csv(index=False).encode('utf-8-sig')
        st.download_button(
            label="下載統計結果 (CSV)",
            data=csv,
            file_name="彙整統計表.csv",
            mime="text/csv",
        )