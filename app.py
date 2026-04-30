import streamlit as st
import pandas as pd
import pdf2image
import pytesseract
import re

st.title("事務機使用量批次統計工具 (智能影像辨識版) 🚀")
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
                    
                    # 狀態標記：判斷目前頁面是複印還是列印
                    is_copy = True if "複印" in text else False
                    is_print = True if "列印" in text else False
                    
                    # 狀態標記：判斷是單色機(只有總印次)還是彩色機(有分黑白彩色)
                    is_color_machine = True if "彩色" in text else False
                    
                    for line in lines:
                        # 尋找以 UserID 開頭的資料列 (假設 UserID 是英文或數字，長度至少2，後面跟著數字)
                        # 使用更寬鬆的尋找方式：找出一行裡面的所有英數字組合
                        tokens = [t for t in re.split(r'\s+|\|', line.strip()) if t]
                        
                        if not tokens: continue
                        
                        # 檢查這行是不是我們要的資料 (避開表頭、總計等文字)
                        # 邏輯：第一個 token 是 UserID，且這行至少包含 1 個數字
                        first_token = tokens[0]
                        numbers = [int(t) for t in tokens if t.isdigit()]
                        
                        # 排除不可能是 UserID 的關鍵字
                        ignore_words = ['No.', '報表', '總計', '頁數', '張數', 'KO', 'CE']
                        if first_token in ignore_words or not numbers:
                            continue
                            
                        # 如果第一個字串太長或包含中文，可能抓錯了，嘗試往後找真正的 UserID
                        user_id = first_token
                        if len(user_id) > 15 or bool(re.search(r'[\u4e00-\u9fff]', user_id)):
                           if len(tokens) > 1 and re.match(r'^[A-Za-z0-9]+$', tokens[1]):
                               user_id = tokens[1]
                           else:
                               continue # 找不到合理的 UserID 就跳過這行

                        # 初始化這筆資料的數字
                        copy_bw, copy_color, print_bw, print_color = 0, 0, 0, 0
                        
                        # 取出這行最後面的數字作為印量 (因為排版問題，印量通常在最後面)
                        # 如果是彩色機，最後幾個數字通常是 黑白, 彩色, (可能有的彩色A3), 累計張數
                        # 如果是單色機，最後幾個數字通常是 總印次, 累計張數
                        
                        try:
                            if is_color_machine:
                                # 彩色機邏輯：倒數第3個數字(或倒數第2個)通常是黑白累計，下一個是彩色累計
                                if len(numbers) >= 2:
                                    # 簡單啟發式：找最大的幾個數字，假設倒數第三、第二個為黑白/彩色頁數 (依據您提供的報表)
                                    bw_pages = numbers[-3] if len(numbers) >= 3 else numbers[-2]
                                    color_pages = numbers[-2] if len(numbers) >= 3 else numbers[-1]
                                    # 如果數值異常大(例如把卡號9999999抓進來了)，需要過濾
                                    if bw_pages == 9999999: bw_pages = 0
                                    if color_pages == 9999999: color_pages = 0
                                    
                                    if is_copy:
                                        copy_bw = bw_pages
                                        copy_color = color_pages
                                    else:
                                        print_bw = bw_pages
                                        print_color = color_pages
                            else:
                                # 單色機邏輯：只有總印次，全部歸在黑白
                                total_pages = numbers[-1] if len(numbers) >= 1 else 0
                                if total_pages == 9999999: total_pages = 0
                                
                                if is_copy:
                                    copy_bw = total_pages
                                else:
                                    print_bw = total_pages
                                    
                        except IndexError:
                            continue # 如果抓數字發生錯誤就跳過這行
                            
                        # 整理這筆資料
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
                    # 處理同一個 UserID 跨頁出現的情況 (先在單檔內加總)
                    df = df.groupby('UserID', as_index=False).sum()
                else:
                    st.warning(f"檔案 {file.name} 辨識完成，但沒有找到符合格式的數據。")
                    continue

            # --- 資料清理與合併 ---
            # 將 CSV/Excel 的舊表頭對應到新的標準表頭 (如果有的話)
            # 這裡確保不管是哪種來源，最後合併的表頭都是一致的
            rename_mapping = {
                '複印(黑白)累積頁數': '複印累計黑白頁數',
                '複印(彩色)累積頁數': '複印累計彩色頁數',
                '列印(黑白)累積頁數': '列印累計黑白頁數',
                '列印(彩色)累積頁數': '列印累計彩色頁數'
            }
            df = df.rename(columns=rename_mapping)
            
            # 補齊可能缺少的目標欄位
            for col in target_cols:
                if col not in df.columns:
                    df[col] = 0
                    
            df = df[target_cols] # 只保留標準欄位
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
