from flask import Flask, request, send_file, render_template
import pandas as pd
import openpyxl
import json
import datetime
import re
import io

app = Flask(__name__)

# =============================
# Step 1: 轉換邏輯輔助函數
# =============================

def extract_main_numbers(text):
    text = str(text)
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    matches = re.findall(r"(?m)^\s*(\d+)", text)
    return matches

def is_buy1get1(text):
    if not text: return False
    text = str(text)
    if re.search(r"買\s*[\d一二三四五六七八九十]+\s*送\s*[\d一二三四五六七八九十]+", text):
        return True
    return "買一送一" in text

def is_gift_case(text):
    text = str(text)
    return "送" in text and not is_buy1get1(text) and "滿$" not in text

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

def clean_empty_text(val):
    if pd.isna(val): return ""
    val_str = str(val).strip()
    if val_str.lower() in ["", "nan", "none", "00:00:00", "0", "0.0"]:
        return ""
    return val_str

def split_product_codes(r_text, promo_text):
    codes = extract_main_numbers(r_text)
    if not codes: return "", ""
    promo_str = str(promo_text)
    if is_buy1get1(promo_str):
        result = "|".join(codes)
        return result, result
    if is_gift_case(promo_str) and len(codes) >= 2:
        return "|".join(codes[:-1]), codes[-1]
    result = "|".join(codes)
    return result, result

def split_time(text):
    text = str(text).strip()
    if text.upper() == "NA" or text == "" or text.lower() == "nan": 
        return "Nan", "Nan"
    if "-" in text:
        parts = text.split("-")
        start = parts[0].strip()
        end = parts[1].strip() if len(parts) > 1 else "Nan"
        return (start if start else "Nan"), (end if end else "Nan")
    return text, "Nan"

# =============================
# Step 2: JSON 轉換輔助函數
# =============================

def perform_transformation(data_stream, model_stream):
    wb_source = openpyxl.load_workbook(data_stream, data_only=True)
    ws_source = wb_source.worksheets[0]
    wb_target = openpyxl.load_workbook(model_stream)
    ws_target = wb_target.worksheets[0]

    target_cells_addr = [
        "E9", "E19", "D20", "E24", "E26", "E28", "E30", "E32", "E34", "E36",
        "E38", "E40", "E42", "E47", "E52", "E54",
        "E56", "E58", "E60", "E62",
        "E64", "E66", "D69", "E71", "E73",
        "E75", "E77" 
    ]
    
    template_height = 76 
    next_row = 3
    d2_url = ws_target.cell(row=2, column=4).value

    # 1. 填寫 Excel 樣板邏輯
    for data_row in range(2, ws_source.max_row + 1):
        if data_row > 2:
            for r_offset in range(template_height):
                for c in range(1, ws_target.max_column + 1):
                    ws_target.cell(row=next_row + r_offset, column=c).value = \
                        ws_target.cell(row=3 + r_offset, column=c).value

        for i, addr in enumerate(target_cells_addr):
            source_col = i + 1
            source_val = str(ws_source.cell(row=data_row, column=source_col).value or "").strip()
            orig_cell = ws_target[addr]
            target_r = orig_cell.row + (next_row - 3)
            target_cell = ws_target.cell(row=target_r, column=orig_cell.column)
            target_cell.number_format = "@"

            if addr == "D20":
                ws_target.cell(row=target_r, column=3).value = "executeScript"
                target_cell.value = f"var targetText = '{source_val}'; var $select = window.jQuery('#ExtendedDataTemplateSelector'); var $opt = $select.find('option').filter(function() {{ return window.jQuery(this).text().trim() === targetText; }}); if($opt.length > 0) {{ $select.val($opt.val()).trigger('change'); }}"
            elif addr == "D69":
                ws_target.cell(row=target_r, column=3).value = "executeScript"
                target_cell.value = f"var targetText = '{source_val}'; var $select = window.jQuery('#OfferSetup_CategoryId'); var val = $select.find('option').filter(function() {{ return window.jQuery(this).text().trim() === targetText; }}).val(); $select.val(val).trigger('change');"
            elif addr == "E77":
                ws_target.cell(row=target_r, column=3).value = "type"
                ws_target.cell(row=target_r, column=4).value = "id=OfferDetails_TermsAndConditionsTranslated_zh_"
                if not source_val or source_val.lower() == "nan":
                    source_val = "每券限兌換一次。每筆交易可以同時使用多張不同品項之回饋券或優惠券。"
                target_cell.value = source_val
            else:
                target_cell.value = source_val

        footer_row = next_row + template_height - 1
        ws_target.cell(row=footer_row, column=3).value = "click"
        ws_target.cell(row=footer_row, column=4).value = "id=btnSave2"
        ws_target.cell(row=footer_row, column=5).value = "id=btnSave2"
        next_row = footer_row + 1

    # 2. 轉換為 JSON 邏輯
    plexure_json = {
        "Name": int(datetime.datetime.now().strftime("%Y%m%d")),
        "CreationDate": 45951,
        "Commands": []
    }

    def inject_start_flow(cmds, url):
        cmds.append({"Command": "open", "Target": str(url), "Value": ""})

    curr_r = 3
    inject_start_flow(plexure_json["Commands"], d2_url)

    # 用來記錄當前這組優惠使用的是哪個模組
    current_template_name = ""

    while curr_r <= ws_target.max_row:
        cmd = str(ws_target.cell(row=curr_r, column=3).value or "").strip()
        target = str(ws_target.cell(row=curr_r, column=4).value or "").strip()
        val = str(ws_target.cell(row=curr_r, column=5).value or "").strip()

        if curr_r == 3 and cmd == "open":
            curr_r += 6
            continue
        if cmd == "open" and curr_r > 3:
            inject_start_flow(plexure_json["Commands"], target)
            curr_r += 6
            continue

        # --- 【核心修正點：追蹤模組名稱】 ---
        # 檢查是否為 D20 產出的模板選擇腳本
        if "ExtendedDataTemplateSelector" in target and cmd == "executeScript":
            match = re.search(r"var targetText = '(.*?)';", target)
            if match:
                current_template_name = match.group(1)

        if cmd or target:
            new_cmd = {"Command": cmd, "Target": target}
            if val: new_cmd["Value"] = val
            plexure_json["Commands"].append(new_cmd)

            # --- 【核心修正點：注入特殊邏輯】 ---
            # 當遇到 DynamicFields_8__Value 且為 type 指令時
            if "ExtendedDataDynamicFields_8__Value" in target and cmd == "type":
                # 判斷模組是否為指定的那一個
                special_template = "MPA TW Discount Off Product(Percentage) Pre tax False"
                if current_template_name == special_template:
                    # 自動追加 click 9 與 type 9
                    plexure_json["Commands"].append({
                        "Command": "click",
                        "Target": "id=OfferPlacement_ExtendedDataFields_ExtendedDataDynamicFields_9__Value"
                    })
                    plexure_json["Commands"].append({
                        "Command": "type",
                        "Target": "id=OfferPlacement_ExtendedDataFields_ExtendedDataDynamicFields_9__Value",
                        "Value": "0"
                    })

            if "btnSave2" in target:
                plexure_json["Commands"].append({"Command": "pause", "Target": "1000", "Value": ""})
        curr_r += 1

    return json.dumps(plexure_json, ensure_ascii=False, indent=2)

# =============================
# Step 3: Flask Routes
# =============================

@app.route('/')
def index(): return render_template('index.html')

@app.route('/step2')
def step2(): return render_template('template.html')

@app.route('/transform', methods=['POST'])
def transform():
    file = request.files['file']
    image_base_path = request.form.get('imagePath', "").strip()
    if image_base_path and not image_base_path.endswith(('/', '\\')):
        image_base_path += "/"

    df = pd.read_excel(file, header=None, skiprows=2, engine='openpyxl')
    df = df[df.iloc[:, 11].notna()].reset_index(drop=True)
    out = pd.DataFrame()

    def get_full_image_path(val):
        cleaned = clean_empty_text(val)
        if cleaned == "": return ""
        return f"{image_base_path}{cleaned}"

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

    out["Promotional Image En"] = df.iloc[:, 49].apply(get_full_image_path)
    out["Promotional Image Zh"] = df.iloc[:, 50].apply(get_full_image_path)
    
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
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        out.to_excel(writer, index=False)
    output.seek(0)
    return send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name='data.xlsx')

@app.route('/transform_to_json', methods=['POST'])
def transform_to_json():
    data_file = request.files['dataFile']
    model_file = request.files['modelFile']
    json_result = perform_transformation(data_file, model_file)
    return send_file(io.BytesIO(json_result.encode('utf-8')), mimetype='application/json', as_attachment=True, download_name='plexure.json')

if __name__ == '__main__':
    app.run(debug=True, port=5000)