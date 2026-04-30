import streamlit as st
import pandas as pd
import pdf2image
import pytesseract
import re

st.title("事務機使用量批次統計工具 (錨點定位終極版) 🚀")
st.write("支援 CSV、Excel 與 **掃描版 PDF**。AI 將精準反推抓取任何格式的 UserID。")

uploaded_files = st.file_uploader("選擇檔案", type=['csv', 'xlsx', 'pdf'], accept_multiple_files=True)

target_cols = [
    'UserID', 
    '複印累計黑白頁數', '複印累計彩色頁數', 
    '列印累計黑白頁數', '列印累計彩色頁數'
]

if uploaded_files:
    all_data = []
    
    for file in uploaded_files:
        try:
            if file.name.lower().endswith('.csv'):
                df = pd.read_csv(file)
                
            elif file.name.lower().endswith('.xlsx'):
                df = pd.read_excel(file)
                
            elif file.name.lower().endswith('.pdf'):
                st.info(f"正在使用 AI 辨識掃描檔：{file.name} (請稍候...)")
                
                images = pdf2image.convert_from_bytes(file.read(), dpi=200)
                pdf_extracted_rows = []
                
                for img in images:
                    text = pytesseract.image_to_string(img, lang='eng+chi_tra')
                    lines = text.split('\n')
                    
                    is_copy = True if "複印" in text else False
                    is_print = True if "列印" in text else False
                    is_color_machine = True if "彩色" in text else False
                    
                    for line in lines:
                        # 將直線符號替換為空白，再切分字串
                        clean_line = line.replace('|', ' ')
                        tokens = [t.strip() for t in clean_line.split() if t.strip()]
                        if not tokens: continue
                        
                        numbers = [int(t) for t in tokens if t.isdigit()]
                        if not numbers: continue
                        
                        # 黑名單防線：整行若包含總計等字眼直接跳過
                        if any(bad in line.lower() for bad in ['no job owner', '總計', '報表列印日期']):
                            continue
                            
                        # ==========================================
                        # 【錨點定位法】：尋找 UserID
                        # ==========================================
                        real_user_id = None
                        
                        # 定義系統限制碼的特徵 (作為錨點)
                        limit_codes = ['0', '9999999', '99999999999999', '(禁止)', '禁止']
                        
                        # 從左到右掃描，尋找「第一個出現的限制碼」
                        anchor_index = -1
                        for i, token in enumerate(tokens):
                            clean_t = re.sub(r'[^a-zA-Z0-9]', '', token)
                            if clean_t in limit_codes or token in limit_codes:
                                anchor_index = i
                                break
                                
                        # 如果找到了限制碼，而且它前面還有字串
                        if anchor_index > 0:
                            # UserID 應該是限制碼的「前一個字串」
                            candidate_index = anchor_index - 1
                            candidate_token = re.sub(r'[^a-zA-Z0-9]', '', tokens[candidate_index])
                            
                            # 檢查這個候選人是不是剛好是 4 碼的「序號」(例如 0001)
                            # 如果是序號，代表 OCR 排版錯位，真正的 UserID 還要再往前一個
                            if len(candidate_token) == 4 and candidate_token.isdigit() and candidate_token.startswith('0'):
                                if candidate_index > 0:
                                    candidate_index -= 1
                                    candidate_token = re.sub(r'[^a-zA-Z0-9]', '', tokens[candidate_index])
                                else:
                                    candidate_token = ""
                            
                            # 如果清出來的候選人有內容，且不是黑名單字眼，這就是 UserID！
                            ignore_words = ['no', 'ko', 'ce', 'apeos']
                            if candidate_token and candidate_token.lower() not in ignore_words:
                                # 再次確保它不是限制碼
                                if candidate_token not in limit_codes:
                                    real_user_id = candidate_token

                        if not real_user_id:
                            continue

                        # ==========================================
                        # 數字抓取邏輯
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
                            'UserID': real_user_id,
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
        st.download_button("📥 下載統計結果 (CSV)", data=csv, file_name="事務機彙整統計表_完成版.csv", mime="text/csv")
