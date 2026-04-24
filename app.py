from flask import Flask, request, send_file, render_template
import pandas as pd
import openpyxl
import json
import datetime
import re
import io

app = Flask(__name__)

# =============================
# Step 1: 轉換邏輯輔助函數 (Excel 處理)
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
    return "" if val_str == "00:00:00" or val_str == "" else val

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
    if text.upper() == "NA" or text == "": return "NA", "NA"
    if "-" in text:
        parts = text.split("-")
        return parts[0].strip(), parts[1].strip()
    return text, ""

# =============================
# Step 2: JSON 轉換輔助函數 (JSON 生成)
# =============================

def perform_transformation(data_stream, model_stream):
    """執行將資料表與模組表結合並轉換為 JSON 的邏輯"""
    # 讀取上傳的資料流
    wb_source = openpyxl.load_workbook(data_stream, data_only=True)
    ws_source = wb_source.worksheets[0]
    
    wb_target = openpyxl.load_workbook(model_stream)
    ws_target = wb_target.worksheets[0]

    # 設定目標儲存格位址
    target_cells_addr = ["E9", "E19", "D22", "E24", "E26", "E28", "E30", "E32", "E34", "E36",
                         "E38", "E40", "E42", "E47", "E52", "E54", "E56", "E58", "E60", "E62", 
                         "E64", "E66", "D69", "E71", "E73", "E75", "E77"]
    template_height = 75
    next_row = 3
    
    # 從模組表的 D2 取得基礎網址
    d2_url = ws_target.cell(row=2, column=4).value

    # 1. 在記憶體中進行模組填充 (不存檔)
    for data_row in range(2, ws_source.max_row + 1):
        for i, addr in enumerate(target_cells_addr):
            source_val = str(ws_source.cell(row=data_row, column=i + 1).value or "")
            orig_cell = ws_target[addr]
            target_r = orig_cell.row + (next_row - 3)
            ws_target.cell(row=target_r, column=orig_cell.column).value = source_val
        
        # 插入結束按鈕 id=btnSave2
        footer_row = next_row + template_height - 1
        ws_target.cell(row=footer_row, column=3).value = "click"
        ws_target.cell(row=footer_row, column=4).value = "id=btnSave2"
        ws_target.cell(row=footer_row, column=5).value = "id=btnSave2"
        next_row = footer_row + 1

    # 2. 轉換為 JSON 結構
    plexure_json = {
        "Name": int(datetime.datetime.now().strftime("%Y%m%d")),
        "CreationDate": 45951,
        "Commands": []
    }

    def inject_start_flow(cmds, url):
        """注入起始標準流程 (共 6 個動作)"""
        cmds.append({"Command": "open", "Target": str(url), "Value": ""})

    curr_r = 3
    # 核心修正：強制先注入第一筆資料的啟動流程
    inject_start_flow(plexure_json["Commands"], d2_url)

    while curr_r <= ws_target.max_row:
        cmd = str(ws_target.cell(row=curr_r, column=3).value or "").strip()
        target = str(ws_target.cell(row=curr_r, column=4).value or "").strip()
        val = str(ws_target.cell(row=curr_r, column=5).value or "").strip()

        # 如果是第一筆模組的起始區域，跳過範本自帶的重複起始指令 (通常是前 6 行)
        if curr_r == 3 and cmd == "open":
            curr_r += 6 # 跳過 open, pause, wait, click, pause, wait
            continue

        # 如果遇到後續模組的 open
        if cmd == "open" and curr_r > 3:
            inject_start_flow(plexure_json["Commands"], target)
            curr_r += 6
            continue

        if cmd or target:
            # 建立指令物件，若 Value 為空則不放入 (符合您的 JSON 範例)
            new_cmd = {"Command": cmd, "Target": target}
            if val:
                new_cmd["Value"] = val
            
            plexure_json["Commands"].append(new_cmd)
            
            if "btnSave2" in target:
                plexure_json["Commands"].append({"Command": "pause", "Target": "1000", "Value": ""})

        curr_r += 1

    return json.dumps(plexure_json, ensure_ascii=False, indent=2)

# =============================
# Flask 路由設定
# =============================

@app.route('/')
def index():
    """首頁：處理第一步轉檔 (index.html)"""
    return render_template('index.html')

@app.route('/step2')
def step2():
    """第二步：處理結合模組轉 JSON (template.html)"""
    return render_template('template.html')

@app.route('/transform', methods=['POST'])
def transform():
    """Step 1: 處理原始檔轉 data.xlsx"""
    if 'file' not in request.files:
        return "請上傳檔案", 400
    
    file = request.files['file']
    
    try:
        # 使用 pandas 處理 Step 1 轉換邏輯
        df = pd.read_excel(file, header=None, skiprows=2, engine='openpyxl')
        df = df[df.iloc[:, 11].notna()].reset_index(drop=True)
        
        out = pd.DataFrame()
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
        out["Product CodeProduct Code to Buy"] = [r[0] for r in res]
        out["Product Code Discounted"] = [r[1] for r in res]
        out["Percentage of Total"] = ""
        out["Promotional Image En"] = ""
        out["Promotional Image Zh"] = ""
        out["Title EN"] = df.iloc[:, 31].apply(clean_empty_text)
        out["Title CH"] = df.iloc[:, 30].apply(clean_empty_text)
        out["Start Date and Time"] = pd.to_datetime(df.iloc[:, 9], errors="coerce").dt.strftime("%Y/%m/%d")
        out["End Date and Time"] = pd.to_datetime(df.iloc[:, 10], errors="coerce").dt.strftime("%Y/%m/%d")
        out["Daily Start Time"] = "12:00 AM"
        out["Daily End Time"] = "11:59 PM"
        
        time_split = df.iloc[:, 34].apply(lambda x: pd.Series(split_time(x)))
        out["Daily Start Time (Split)"] = time_split[0]
        out["Daily End Time (Split)"] = time_split[1]
        out["Number of days after activation"] = 3
        out["Specify expiry time"] = "11:59 PM"
        out["Category"] = df.iloc[:, 38].apply(clean_empty_text)
        out["Description (Max. 500 chars)(EN)"] = df.iloc[:, 40].apply(clean_empty_text)
        out["Terms (Max. 4000 chars)(EN)"] = df.iloc[:, 42].apply(clean_empty_text)
        out["Description (Max. 500 chars)(CH)"] = df.iloc[:, 39].apply(clean_empty_text)
        out["Terms (Max. 4000 chars)(CH)"] = df.iloc[:, 41].apply(clean_empty_text)
        out["addSelection1"] = df.iloc[:, 44].apply(clean_empty_text)
        out["addSelection2"] = df.iloc[:, 45].apply(clean_empty_text)
        out["addSelection3"] = df.iloc[:, 46].apply(clean_empty_text)
        out["stores"] = df.iloc[:, 47].apply(clean_empty_text)
        
        out = out[~out["Internal Name"].astype(str).str.contains("系統排序", na=False)]
        
        # 產出 Excel 下載流
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            out.to_excel(writer, index=False)
        output.seek(0)
        
        return send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                         as_attachment=True, download_name='data.xlsx')
    except Exception as e:
        return f"Step 1 轉換失敗: {str(e)}", 500

@app.route('/transform_to_json', methods=['POST'])
def transform_to_json():
    """Step 2: 處理 data.xlsx + 模組表轉 JSON"""
    if 'dataFile' not in request.files or 'modelFile' not in request.files:
        return "請確保上傳了資料表與模組表", 400
        
    data_file = request.files['dataFile']
    model_file = request.files['modelFile']
    
    try:
        # 呼叫 Step 2 處理核心
        json_result = perform_transformation(data_file, model_file)
        
        return send_file(
            io.BytesIO(json_result.encode('utf-8')),
            mimetype='application/json',
            as_attachment=True,
            download_name='plexure.json'
        )
    except Exception as e:
        # 詳細錯誤捕捉
        return f"JSON 轉換失敗: {str(e)}", 500

if __name__ == '__main__':
    # 統一運行於 5000 埠
    app.run(debug=True, port=5000)