import streamlit as st
import pandas as pd
import pdf2image
import pytesseract
import re

st.title("事務機使用量批次統計工具 (空間座標定位版) 🚀")
st.write("支援 CSV、Excel 與 **掃描版 PDF**。系統將精準對齊表頭座標抓取資料。")

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
                st.info(f"正在使用 AI 空間定位辨識掃描檔：{file.name} (請稍候...)")
                
                images = pdf2image.convert_from_bytes(file.read())
                pdf_extracted_rows = []
                
                for img in images:
                    # 使用 image_to_data 獲取每個單字的實體座標 (x, y, w, h)
                    ocr_data = pytesseract.image_to_data(img, lang='eng+chi_tra', output_type=pytesseract.Output.DICT)
                    
                    # 1. 尋找「使用者ID」表頭的 X 軸座標範圍
                    id_col_left = -1
                    id_col_right = -1
                    
                    # 順便判斷報表類型
                    text_full = " ".join([str(t) for t in ocr_data['text'] if str(t).strip()])
                    is_copy = True if "複印" in text_full else False
                    is_print = True if "列印" in text_full else False
                    is_color_machine = True if "彩色" in text_full else False
                    
                    for i in range(len(ocr_data['text'])):
                        word = str(ocr_data['text'][i]).strip()
                        if "使用者ID" in word or word == "ID":
                            id_col_left = ocr_data['left'][i]
                            # 增加一點寬容度 (向左右擴展)，因為下方文字可能稍微凸出
                            id_col_right = id_col_left + ocr_data['width'][i] + 30 
                            id_col_left -= 30
                            break
                    
                    # 如果找不到表頭，代表這頁可能沒有資料，跳過
                    if id_col_left == -1:
                        continue

                    # 2. 將散落的單字依據 Y 軸重新組合成「行」 (Line)
                    # Y 座標相近的單字視為同一行 (容差 15 像素)
                    lines = {}
                    for i in range(len(ocr_data['text'])):
                        word = str(ocr_data['text'][i]).strip()
                        if not word: continue
                        
                        y = ocr_data['top'][i]
                        x = ocr_data['left'][i]
                        w = ocr_data['width'][i]
                        
                        # 尋找是否已有相近的 Y 座標行
                        found_line = False
                        for line_y in lines.keys():
                            if abs(y - line_y) < 15:
                                lines[line_y].append({'text': word, 'x': x, 'w': w})
                                found_line = True
                                break
                                
                        if not found_line:
                            lines[y] = [{'text': word, 'x': x, 'w': w}]
                            
                    # 3. 逐行解析資料
                    # 將行依據 Y 座標由上到下排序
                    sorted_lines = sorted(lines.items(), key=lambda item: item[0])
                    
                    for y, words in sorted_lines:
                        # 將同行單字依據 X 座標由左到右排序
                        words = sorted(words, key=lambda w: w['x'])
                        
                        # 提取整行文字，方便後續判斷黑名單
                        line_text = " ".join([w['text'] for w in words])
                        
                        # 黑名單檢查
                        ignore_words_lower = ['no.', 'ko', 'ce', 'apeos', '報表', '總計', '頁數', '張數', '最終', 'job', 'owner']
                        if any(bad_word in line_text.lower() for bad_word in ignore_words_lower):
                            continue
                            
                        # 提取這行的所有數字
                        numbers = [int(w['text']) for w in words if w['text'].isdigit()]
                        if not numbers:
                            continue

                        # 【核心定位邏輯】：尋找 X 座標落在「使用者ID」直行下方的文字
                        real_user_id = None
                        
                        for w in words:
                            word_center_x = w['x'] + (w['w'] / 2)
                            # 如果這個單字的中心點，落在「使用者ID」的左右邊界內
                            if id_col_left <= word_center_x <= id_col_right:
                                # 清除雜質
                                clean_t = re.sub(r'[^a-zA-Z0-9]', '', w['text'])
                                if clean_t and clean_t not in ['0', '9999999', '99999999999999']: # 避開卡號等無效數字
                                    real_user_id = clean_t
                                    break
                                    
                        if not real_user_id:
                            continue

                        # ==========================================
                        # 數字抓取邏輯 (取得最後面的印量數字)
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
        st.download_button("📥 下載統計結果 (CSV)", data=csv, file_name="事務機彙整統計表_精準對齊版.csv", mime="text/csv")
