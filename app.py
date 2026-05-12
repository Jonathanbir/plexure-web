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

    # 定義 Excel 對應地址 (A-AA 欄對應地址)
    target_cells_addr = [
        "E9", "E19", "D20", "E24", "E26", "E28", "E30", "E32", "E34", "E36",
        "E38", "E40", "E42", "E47", "E52", "E54", "E56", "E58", "E60", "E62",
        "E64", "E66", "D69", "E71", "E73", "E75", "E77"
    ]
    
    template_height = 76 # 維持 76 確保不蓋掉 E77 內容
    next_row = 3
    d2_url = ws_target.cell(row=2, column=4).value

    # 1. 填寫工作表樣板
    for data_row in range(2, ws_source.max_row + 1):
        if data_row > 2:
            for r_offset in range(template_height):
                for c in range(1, ws_target.max_column + 1):
                    ws_target.cell(row=next_row + r_offset, column=c).value = \
                        ws_target.cell(row=3 + r_offset, column=c).value

        for i, addr in enumerate(target_cells_addr):
            source_col = i + 1
            source_val = str(ws_source.cell(row=data_row, column=source_col).value or "").strip()
            orig_cell = wb_target.worksheets[0][addr]
            target_r = orig_cell.row + (next_row - 3)
            target_cell = ws_target.cell(row=target_r, column=orig_cell.column)
            target_cell.number_format = "@"

            # 模組選擇與分類選擇的 JS 注入
            if addr == "D20":
                ws_target.cell(row=target_r, column=3).value = "executeScript"
                target_cell.value = f"var targetText = '{source_val}'; var $select = window.jQuery('#ExtendedDataTemplateSelector'); var $opt = $select.find('option').filter(function() {{ return window.jQuery(this).text().trim() === targetText; }}); if($opt.length > 0) {{ $select.val($opt.val()).trigger('change'); }}"
            elif addr == "D69":
                    ws_target.cell(row=target_r, column=3).value = "executeScript"
                    # 移除註解以防止壓縮成一行時發生錯誤
                    target_cell.value = f"""
                    (function() {{
                        var targetText = '{source_val}';
                        var $select = window.jQuery('#OfferDetails_CategoryId');
                        var val = $select.find('option').filter(function() {{ 
                            return window.jQuery(this).text().trim().indexOf(targetText) > -1; 
                        }}).val();
                        
                        if (val) {{
                            $select.val(val).trigger('change').trigger('change.select2');
                        }}
                        
                        window.jQuery('.select2-result-label').filter(function() {{
                            return window.jQuery(this).text().trim().indexOf(targetText) > -1;
                        }}).click();
                    }})();
                    """.replace('\n', ' ').replace('\r', '').strip()
            elif addr == "E77": # 中文條款強制改 type 並設定預設值
                ws_target.cell(row=target_r, column=3).value = "type"
                ws_target.cell(row=target_r, column=4).value = "id=OfferDetails_TermsAndConditionsTranslated_zh_"
                if not source_val or source_val.lower() == "nan":
                    source_val = "每券限兌換一次。每筆交易可以同時使用多張不同品項之回饋券或優惠券。"
                target_cell.value = source_val
            else:
                target_cell.value = source_val

        # 設定儲存按鈕
        footer_row = next_row + template_height - 1
        ws_target.cell(row=footer_row, column=3).value = "click"
        ws_target.cell(row=footer_row, column=4).value = "id=btnSave2"
        ws_target.cell(row=footer_row, column=5).value = "id=btnSave2"
        next_row = footer_row + 1

    # 2. 產出 JSON 結構
    plexure_json = {"Name": int(datetime.datetime.now().strftime("%Y%m%d")), "CreationDate": 45951, "Commands": []}
    plexure_json["Commands"].append({"Command": "open", "Target": str(d2_url), "Value": ""})
    
    curr_r = 3
    current_template_name = ""
    while curr_r <= ws_target.max_row:
        cmd = str(ws_target.cell(row=curr_r, column=3).value or "").strip()
        target = str(ws_target.cell(row=curr_r, column=4).value or "").strip()
        val = str(ws_target.cell(row=curr_r, column=5).value or "").strip()

        # 抓取當前模組名稱以供後續判斷
        if "ExtendedDataTemplateSelector" in target and cmd == "executeScript":
            match = re.search(r"var targetText = '(.*?)';", target)
            if match: current_template_name = match.group(1)

        if cmd or target:
            # 針對圖片上傳按鈕增加滾動防呆
            if "Promo_en_saveButton" in target or "Promo_zh_saveButton" in target:
                plexure_json["Commands"].append({
                    "Command": "executeScript",
                    "Target": f"document.getElementById('{target.replace('id=', '')}').scrollIntoView();"
                })
                plexure_json["Commands"].append({"Command": "pause", "Target": "3000"})

            new_cmd = {"Command": cmd, "Target": target}
            if val: new_cmd["Value"] = val
            plexure_json["Commands"].append(new_cmd)

            # MPA Percentage 模組自動補欄位 9
            if "ExtendedDataDynamicFields_8__Value" in target and cmd == "type":
                if current_template_name == "MPA TW Discount Off Product(Percentage) Pre tax False":
                    plexure_json["Commands"].append({"Command": "click", "Target": "id=OfferPlacement_ExtendedDataFields_ExtendedDataDynamicFields_9__Value"})
                    plexure_json["Commands"].append({"Command": "type", "Target": "id=OfferPlacement_ExtendedDataFields_ExtendedDataDynamicFields_9__Value", "Value": "0"})

            if "btnSave2" in target: plexure_json["Commands"].append({"Command": "pause", "Target": "1000"})
        curr_r += 1

    return json.dumps(plexure_json, ensure_ascii=False, indent=2)

# ==========================================
# Step 3: Flask 路由 (含安全性)
# ==========================================

@app.route('/')
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
    if not session.get('logged_in'): return redirect(url_for('login_page'))
    try:
        data_file = request.files['dataFile']
        model_file = request.files['modelFile']
        json_result = perform_transformation(data_file, model_file)
        return render_template('result.html', json_content=json_result)
    except Exception as e:
        traceback.print_exc()
        return f"發生錯誤：{str(e)}", 500

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login_page'))

if __name__ == '__main__':
    app.run(debug=True, port=5000)