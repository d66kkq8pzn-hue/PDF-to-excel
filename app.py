import streamlit as st
import pandas as pd
import pdf2image
import pytesseract
import re

st.title("事務機使用量批次統計工具 (極致過濾版) 🚀")
st.write("支援 CSV、Excel 與 **掃描版 PDF**。系統會自動辨識報表類型並歸類印量。")

uploaded_files = st.file_uploader("選擇檔案", type=['csv', 'xlsx', 'pdf'], accept_multiple_files=True)

# 最終彙整的標準表頭
target_cols = [
    'UserID', 
    '複印累計黑白頁數', '複印累計彩色頁數', 
    '列印累計黑白頁數', '列印累計彩色頁數'
]

if uploaded_files:
    all_data = []
    
    for file in uploaded_files:
        try:
            # --- 1. 處理 CSV ---
            if file.name.lower().endswith('.csv'):
                df = pd.read_csv(file)
                
            # --- 2. 處理 Excel ---
            elif file.name.lower().endswith('.xlsx'):
                df = pd.read_excel(file)
                
            # --- 3. 處理 掃描版 PDF (OCR 辨識) ---
            elif file.name.lower().endswith('.pdf'):
                st.info(f"正在使用 AI 辨識掃描檔：{file.name} (請稍候...)")
                
                images = pdf2image.convert_from_bytes(file.read())
                pdf_extracted_rows = []
                
                for img in images:
                    text = pytesseract.image_to_string(img, lang='eng+chi_tra')
                    lines = text.split('\n')
                    
                    is_copy = True if "複印" in text else False
                    is_print = True if "列印" in text else False
                    is_color_machine = True if "彩色" in text else False
                    
                    for line in lines:
                        tokens = [t for t in re.split(r'\s+|\|', line.strip()) if t]
                        if not tokens: continue
                        
                        first_token = tokens[0]
                        numbers = [int(t) for t in tokens if t.isdigit()]
                        
                        # ==========================================
                        # 【最新強化】：嚴格的黑名單過濾與去雜質
                        # ==========================================
                        
                        # 1. 清除第一組字串中的特殊符號 (如引號、逗號等)，只保留英數字
                        clean_token = re.sub(r'[^a-zA-Z0-9]', '', first_token)
                        
                        # 2. 建立嚴格的黑名單 (全部轉小寫比對，避免大小寫問題)
                        # 包含你截圖中出現的 Apeos, No, Ko, 還有表頭常見字眼
                        ignore_words_lower = ['no', 'ko', 'ce', 'apeos', '報表', '總計', '頁數', '張數', '最終', 'job', 'owner']
                        
                        # 如果清除雜質後字串是空的，或者存在於黑名單中，或者這行連一個數字都沒有 -> 直接跳過
                        if not clean_token or clean_token.lower() in ignore_words_lower or not numbers:
                            continue
                            
                        # 3. 檢查長度與中文 (UserID 不應該太長，也不該有中文)
                        user_id = clean_token
                        if len(user_id) > 15 or bool(re.search(r'[\u4e00-\u9fff]', user_id)):
                           if len(tokens) > 1:
                               # 如果第一個抓錯了，嘗試抓第二個 (同樣要清理雜質)
                               second_clean = re.sub(r'[^a-zA-Z0-9]', '', tokens[1])
                               if second_clean and second_clean.lower() not in ignore_words_lower:
                                   user_id = second_clean
                               else:
                                   continue
                           else:
                               continue

                        # ==========================================
                        # 數字抓取邏輯 (與上一版相同)
                        # ==========================================
                        copy_bw, copy_color, print_bw, print_color = 0, 0, 0, 0
                        
                        try:
                            if is_color_machine:
                                if len(numbers) >= 2:
                                    bw_pages = numbers[-3] if len(numbers) >= 3 else numbers[-2]
                                    color_pages = numbers[-2] if len(numbers) >= 3 else numbers[-1]
                                    if bw_pages == 9999999: bw_pages = 0
                                    if color_pages == 9999999: color_pages = 0
                                    
                                    if is_copy:
                                        copy_bw = bw_pages
                                        copy_color = color_pages
                                    else:
                                        print_bw = bw_pages
                                        print_color = color_pages
                            else:
                                total_pages = numbers[-1] if len(numbers) >= 1 else 0
                                if total_pages == 9999999: total_pages = 0
                                
                                if is_copy:
                                    copy_bw = total_pages
                                else:
                                    print_bw = total_pages
                                    
                        except IndexError:
                            continue
                            
                        row_data = {
                            'UserID': user_id,
                            '複印累計黑白頁數': copy_bw,
                            '複印累計彩色頁數': copy_color,
                            '列印累計黑白頁數': print_bw,
                            '列印累計彩色頁數': print_color
                        }
                        pdf_extracted_rows.append(row_data)
                
                if pdf_extracted_rows:
                    df = pd.DataFrame(pdf_extracted_rows)
                    df = df.groupby('UserID', as_index=False).sum()
                else:
                    st.warning(f"檔案 {file.name} 辨識完成，但沒有找到符合格式的數據。")
                    continue

            # --- 資料清理與合併 ---
            rename_mapping = {
                '複印(黑白)累積頁數': '複印累計黑白頁數',
                '複印(彩色)累積頁數': '複印累計彩色頁數',
                '列印(黑白)累積頁數': '列印累計黑白頁數',
                '列印(彩色)累積頁數': '列印累計彩色頁數'
            }
            df = df.rename(columns=rename_mapping)
            
            for col in target_cols:
                if col not in df.columns:
                    df[col] = 0
                    
            df = df[target_cols]
            all_data.append(df)
            
        except Exception as e:
            st.error(f"處理檔案 {file.name} 時發生錯誤: {e}")

    # --- 最終統計 ---
    if all_data:
        combined_df = pd.concat(all_data, ignore_index=True)
        
        numeric_cols = [col for col in target_cols if col != 'UserID']
        for col in numeric_cols:
            combined_df[col] = pd.to_numeric(combined_df[col], errors='coerce').fillna(0)
            
        summary_df = combined_df.groupby('UserID', as_index=False).sum()
        
        st.success("🎉 統計完成！")
        st.dataframe(summary_df)
        
        csv = summary_df.to_csv(index=False).encode('utf-8-sig')
        st.download_button("📥 下載統計結果 (CSV)", data=csv, file_name="事務機彙整統計表_整合版.csv", mime="text/csv")
