from flask import Flask, request, send_file, render_template, session, jsonify, redirect, url_for
import pandas as pd
import openpyxl
import json
import datetime
import re
import io
import traceback

app = Flask(__name__)
app.secret_key = "plexure_automation_master_key" # 用於保持登入狀態

# ==========================================
# Step 1: 輔助函數 (邏輯處理)
# ==========================================

def extract_main_numbers(text):
    text = str(text)
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    matches = re.findall(r"(?m)^\s*(\d+)", text)
    return matches

def is_buy1get1(text):
    if not text: return False
    text = str(text)
    return "買一送一" in text or bool(re.search(r"買\s*[\d一二三四五六七八九十]+\s*送\s*[\d一二三四五六七八九十]+", text))

def clean_empty_text(val):
    if pd.isna(val): return ""
    val_str = str(val).strip()
    if val_str.lower() in ["", "nan", "none", "00:00:00", "0", "0.0"]: return ""
    return val_str

def split_product_codes(r_text, promo_text):
    codes = extract_main_numbers(r_text)
    if not codes: return "", ""
    promo_str = str(promo_text)
    if is_buy1get1(promo_str):
        result = "|".join(codes)
        return result, result
    return "|".join(codes), (codes[-1] if len(codes) > 0 else "")

def split_time(text):
    text = str(text).strip()
    if text.upper() == "NA" or text == "" or text.lower() == "nan": return "Nan", "Nan"
    if "-" in text:
        parts = text.split("-")
        s = parts[0].strip()
        e = parts[1].strip() if len(parts) > 1 else "Nan"
        return (s if s else "Nan"), (e if e else "Nan")
    return text, "Nan"

def get_extended_template(text):
    if pd.isna(text): return ""
    text = str(text)
    if re.search(r"\+.*(?:特價|限時優惠)", text):
        return "TW Price Deal - Two Product Sets - Reduced Price"
    if re.search(r"滿.*(?:送|折)|單點.*折", text):
        return "MPA TW Discount Off Product(Percentage) Pre tax False"
    if re.search(r"買一送一|買.*送|單點.*送", text):
        return "TW Buy One Get One Or Another Discounted(Percentage)"
    if re.search(r"現折|單點.*特價", text):
        return "MPA TW Discount Off Product($Amount) Pre tax False"
    return ""

# ==========================================
# Step 2: 核心轉換函式 (解決 NameError)
# ==========================================

def perform_transformation(data_stream, model_stream):
    wb_source = openpyxl.load_workbook(data_stream, data_only=True)
    ws_source = wb_source.worksheets[0]
    wb_target = openpyxl.load_workbook(model_stream)
    ws_target = wb_target.worksheets[0]

    # 定義 Excel 固定對應地址 (E9 到 E76)
    target_cells_addr = [
        "E9", "E19", "D20", "E24", "E26", "E28", "E30", "E32", "E34", "E36",
        "E38", "E40", "E42", "E47", "E52", "E54", "E56", "E58", "E60", "E62",
        "E64", "E66", "D69", "E71", "E73", "E75", "E77"
    ]
    
    template_height = 76 # 基礎樣板高度 (到 E76)
    next_row = 3
    d2_url = ws_target.cell(row=2, column=4).value

    for data_row in range(2, ws_source.max_row + 1):
        # 0. 複製基礎樣板 (複製 1 到 76 行的內容)
        if data_row > 2:
            for r_offset in range(template_height):
                for c in range(1, ws_target.max_column + 1):
                    ws_target.cell(row=next_row + r_offset, column=c).value = \
                        ws_target.cell(row=3 + r_offset, column=c).value

        # 1. 填寫固定欄位 (E9 ~ E76)
        for i, addr in enumerate(target_cells_addr):
            source_col = i + 1
            source_val = str(ws_source.cell(row=data_row, column=source_col).value or "").strip()
            orig_cell = wb_target.worksheets[0][addr]
            target_r = orig_cell.row + (next_row - 3)
            target_cell = ws_target.cell(row=target_r, column=orig_cell.column)
            target_cell.number_format = "@"

            if addr == "D20":
                ws_target.cell(row=target_r, column=3).value = "executeScript"
                target_cell.value = f"var targetText = '{source_val}'; var $select = window.jQuery('#ExtendedDataTemplateSelector'); var $opt = $select.find('option').filter(function() {{ return window.jQuery(this).text().trim() === targetText; }}); if($opt.length > 0) {{ $select.val($opt.val()).trigger('change'); }}"
            elif addr == "D69":
                ws_target.cell(row=target_r, column=3).value = "executeScript"
                target_cell.value = f"(function(){{var t='{source_val}';var s=window.jQuery('#OfferDetails_CategoryId');var v=s.find('option').filter(function(){{return window.jQuery(this).text().trim().indexOf(t)>-1;}}).val();if(v){{s.val(v).trigger('change').trigger('change.select2');}}window.jQuery('.select2-result-label').filter(function(){{return window.jQuery(this).text().trim().indexOf(t)>-1;}}).click();}})();"
            elif addr == "E76":
                ws_target.cell(row=target_r, column=3).value = "type"
                ws_target.cell(row=target_r, column=4).value = "id=OfferDetails_TermsAndConditionsTranslated_zh_"
                if not source_val or source_val.lower() == "nan":
                    source_val = "每券限兌換一次。每筆交易可以同時使用多張不同品項之回饋券或優惠券。"
                target_cell.value = source_val
            else:
                target_cell.value = source_val

        # 當前動態指令起始行 (在 E76 之後)
        dynamic_row = next_row + template_height

        # 2. 處理 AddSelection (對應 VBA 的 AB, AC, AD 欄位)
        # 根據 CSV：AB=28, AC=29, AD=30
        selection_cols = [28, 29, 30] 
        for i, col_idx in enumerate(selection_cols):
            cell_text = str(ws_source.cell(row=data_row, column=col_idx).value or "").strip()
            if cell_text:
                tag_list = [t.strip() for t in cell_text.split(",") if t.strip()]
                if tag_list:
                    for tag in tag_list:
                        ws_target.cell(row=dynamic_row, column=3).value = "addSelection"
                        ws_target.cell(row=dynamic_row, column=4).value = f"id=allTags{i+1}"
                        ws_target.cell(row=dynamic_row, column=5).value = f"label={tag}"
                        dynamic_row += 1
                    # 補上點擊 Add 按鈕
                    ws_target.cell(row=dynamic_row, column=3).value = "click"
                    ws_target.cell(row=dynamic_row, column=4).value = f"id=btnAdd{i+1}"
                    ws_target.cell(row=dynamic_row, column=5).value = f"id=btnAdd{i+1}"
                    dynamic_row += 1

        # 3. 處理 Stores 門市 (對應 VBA 的 AE 欄位)
        # 根據 CSV：AE=31
        stores_raw = str(ws_source.cell(row=data_row, column=31).value or "").strip()
        if stores_raw:
            store_ids = [s.strip() for s in stores_raw.split("#") if s.strip()]
            if store_ids:
                # 點選「非全門市」選項 (依據 VBA 邏輯)
                ws_target.cell(row=dynamic_row, column=3).value = "click"
                ws_target.cell(row=dynamic_row, column=4).value = "xpath=//input[@id='OfferStores_isAvailableAllStores' and @name='OfferStores.isAvailableAllStores' and @value='False']"
                dynamic_row += 1
                
                for sid in store_ids:
                    target_xpath = f"xpath=//input[@name='OfferStores.venues' and @value='{sid}']"
                    # Type Store ID
                    ws_target.cell(row=dynamic_row, column=3).value = "type"
                    ws_target.cell(row=dynamic_row, column=4).value = target_xpath
                    ws_target.cell(row=dynamic_row, column=5).value = sid
                    dynamic_row += 1
                    # Click Checkbox
                    ws_target.cell(row=dynamic_row, column=3).value = "click"
                    ws_target.cell(row=dynamic_row, column=4).value = target_xpath
                    dynamic_row += 1

        # 4. 結尾：儲存按鈕
        ws_target.cell(row=dynamic_row, column=3).value = "click"
        ws_target.cell(row=dynamic_row, column=4).value = "id=btnSave2"
        ws_target.cell(row=dynamic_row, column=5).value = "id=btnSave2"
        
        # 更新 next_row 供下一筆資料使用
        next_row = dynamic_row + 1

    # --- 產出 JSON 部分 ---
    plexure_json = {"Name": int(datetime.datetime.now().strftime("%Y%m%d")), "CreationDate": 45951, "Commands": []}
    plexure_json["Commands"].append({"Command": "open", "Target": str(d2_url), "Value": ""})
    
    curr_r = 3
    current_template_name = ""
    while curr_r <= ws_target.max_row:
        cmd = str(ws_target.cell(row=curr_r, column=3).value or "").strip()
        target = str(ws_target.cell(row=curr_r, column=4).value or "").strip()
        val = str(ws_target.cell(row=curr_r, column=5).value or "").strip()

        if "ExtendedDataTemplateSelector" in target and cmd == "executeScript":
            match = re.search(r"var targetText = '(.*?)';", target)
            if match: current_template_name = match.group(1)

        if cmd or target:
            # 圖片按鈕滾動邏輯
            if "Promo_en_saveButton" in target or "Promo_zh_saveButton" in target:
                plexure_json["Commands"].append({
                    "Command": "executeScript",
                    "Target": f"document.getElementById('{target.replace('id=', '')}').scrollIntoView();"
                })
                plexure_json["Commands"].append({"Command": "pause", "Target": "3000"})

            new_cmd = {"Command": cmd, "Target": target}
            if val: new_cmd["Value"] = val
            plexure_json["Commands"].append(new_cmd)

            # 特定模組注入邏輯
            if "ExtendedDataDynamicFields_8__Value" in target and cmd == "type":
                if current_template_name == "MPA TW Discount Off Product(Percentage) Pre tax False":
                    plexure_json["Commands"].append({"Command": "click", "Target": "id=OfferPlacement_ExtendedDataFields_ExtendedDataDynamicFields_9__Value"})
                    plexure_json["Commands"].append({"Command": "type", "Target": "id=OfferPlacement_ExtendedDataFields_ExtendedDataDynamicFields_9__Value", "Value": "0"})

            if "btnSave2" in target: 
                plexure_json["Commands"].append({"Command": "pause", "Target": "1000"})
                # 修復 bfcache 報錯問題
                plexure_json["Commands"].append({"Command": "selectWindow", "Target": "tab=0"})
        curr_r += 1

    return json.dumps(plexure_json, ensure_ascii=False, indent=2)
# ==========================================
# Step 3: Flask 路由 (含安全性)
# ==========================================

@app.route('/')
def index():
    # 如果有登入，標籤會引導你來這裡，然後被這行踢去 /step1
    if session.get('logged_in'): 
        return redirect(url_for('step1'))
    return render_template('index.html')

def login_page():
    if session.get('logged_in'): return redirect(url_for('step1'))
    return render_template('index.html')

@app.route('/login', methods=['POST'])
def login():
    data = request.get_json()
    if data.get('username') == "admin" and data.get('password') == "12345":
        session['logged_in'] = True
        return jsonify({"status": "success"})
    return jsonify({"status": "error", "message": "帳號或密碼錯誤"})

@app.route('/step1')
def step1():
    if not session.get('logged_in'): return redirect(url_for('login_page'))
    return render_template('step1.html')

@app.route('/step2')
def step2():
    if not session.get('logged_in'): return redirect(url_for('login_page'))
    return render_template('step2.html')

@app.route('/transform', methods=['POST'])
def transform():
    if not session.get('logged_in'): return jsonify({"status": "error", "message": "請登入"}), 401
    try:
        file = request.files.get('file')
        image_base_path = request.form.get('imagePath', "").strip()
        if image_base_path and not image_base_path.endswith(('/', '\\')): image_base_path += "/"

        df = pd.read_excel(file, header=None, skiprows=2, engine='openpyxl')
        df = df[df.iloc[:, 11].notna()].reset_index(drop=True)
        out = pd.DataFrame()

        def get_img(val):
            c = clean_empty_text(val)
            return f"{image_base_path}{c}" if c else ""

        # Excel 欄位對應
        out["Internal Name"] = df.iloc[:, 11]
        out["Base Weight"] = df.iloc[:, 14]
        out["Extended Data Templates"] = df.iloc[:, 11].apply(get_extended_template)
        out["Promotion Name(EN)"] = df.iloc[:, 25].apply(clean_empty_text)
        out["Promotion Short Description(EN)"] = df.iloc[:, 27].apply(clean_empty_text)
        out["Promotion Long Description(EN)"] = df.iloc[:, 29].apply(clean_empty_text)
        out["Promotion Name(ZH)"] = df.iloc[:, 24].apply(clean_empty_text)
        out["Promotion Short Description(ZH)"] = df.iloc[:, 26].apply(clean_empty_text)
        out["Promotion Long Description(ZH)"] = df.iloc[:, 28].apply(clean_empty_text)
        res = df.apply(lambda r: split_product_codes(r.iloc[17], r.iloc[11]), axis=1)
        out["Product Code Buy"] = [r[0] for r in res]
        out["Product Code Discounted"] = [r[1] for r in res]
        out["Percentage"] = "1%"
        out["Promotional Image En"] = df.iloc[:, 49].apply(get_img)
        out["Promotional Image Zh"] = df.iloc[:, 50].apply(get_img)
        out["Title EN"] = df.iloc[:, 31].apply(clean_empty_text)
        out["Title CH"] = df.iloc[:, 30].apply(clean_empty_text)
        out["Start Date"] = pd.to_datetime(df.iloc[:, 9], errors="coerce").dt.strftime("%Y/%m/%d")
        out["Daily Start"] = "12:00 AM"
        out["End Date"] = pd.to_datetime(df.iloc[:, 10], errors="coerce").dt.strftime("%Y/%m/%d")
        out["Daily End"] = "11:59 PM"
        time_split = df.iloc[:, 34].apply(lambda x: pd.Series(split_time(x)))
        out["Daily Start Split"] = time_split[0]
        out["Daily End Split"] = time_split[1]
        out["Category"] = df.iloc[:, 38].apply(clean_empty_text) 
        out["Desc EN"] = df.iloc[:, 40].apply(clean_empty_text)
        out["Terms EN"] = df.iloc[:, 42].apply(clean_empty_text)
        out["Desc ZH"] = df.iloc[:, 39].apply(clean_empty_text)
        out["Terms ZH"] = df.iloc[:, 41].apply(clean_empty_text)
        out["addSelection1"] = df.iloc[:, 44].apply(clean_empty_text)
        out["addSelection2"] = df.iloc[:, 45].apply(clean_empty_text)
        out["addSelection3"] = df.iloc[:, 46].apply(clean_empty_text)
        out["stores"] = df.iloc[:, 47].apply(clean_empty_text)
        
        out = out[~out["Internal Name"].astype(str).str.contains("系統排序", na=False)]
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer: out.to_excel(writer, index=False)
        output.seek(0)
        return send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name='data.xlsx')
    except Exception as e:
        traceback.print_exc()
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route('/generate_result', methods=['POST'])
def generate_result():
    if not session.get('logged_in'):
        return redirect(url_for('index'))
    
    try:
        data_file = request.files.get('dataFile')
        model_file = request.files.get('modelFile')

        # --- 新增後端防禦檢查 ---
        if not data_file or data_file.filename == '':
            return "錯誤：未上傳資料表 (data.xlsx)", 400
        if not model_file or model_file.filename == '':
            return "錯誤：未上傳 UI 樣板 (modle.xlsm)", 400

        # 執行轉換
        json_result = perform_transformation(data_file, model_file)
        
        return render_template('result.html', json_content=json_result)
        
    except Exception as e:
        traceback.print_exc()
        # 這裡會捕捉到 Zip 相關錯誤，並顯示比較友善的訊息
        return f"發生錯誤：{str(e)}。請確認上傳的是正確的 Excel 檔案格式。", 500

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True, port=5000)