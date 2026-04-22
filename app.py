from flask import Flask, request, send_file, render_template
import pandas as pd
import re
import io

app = Flask(__name__)

# =============================
# 轉換邏輯輔助函數
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
# Flask 路由設定
# =============================

@app.route('/')
def index():
    # 這裡會去讀取 templates/index.html
    return render_template('index.html')


@app.route('/step2')
def step2():
    # 這裡會去讀取 templates/template.html
    return render_template('template.html')

@app.route('/transform', methods=['GET', 'POST']) # 修改這裡，同時允許 GET 和 POST
def transform():
    if 'file' not in request.files:
        return "請上傳檔案", 400
    
    file = request.files['file']
    
    try:
        # 直接讀取上傳的檔案物件，跳過前兩列
        df = pd.read_excel(file, header=None, skiprows=2, engine='openpyxl')
        
        # 移除 L 欄（索引11）空白的列
        df = df[df.iloc[:, 11].notna()].reset_index(drop=True)
        
        out = pd.DataFrame()
        
        # 核心欄位對應與清理
        out["Internal Name"] = df.iloc[:, 11]
        out["Base Weight"] = df.iloc[:, 14]
        out["Extended Data Templates"] = df.iloc[:, 11].apply(get_extended_template)
        
        out["Promotion Name(EN)"] = df.iloc[:, 25].apply(clean_empty_text)
        out["Promotion Short Description(EN)"] = df.iloc[:, 27].apply(clean_empty_text)
        out["Promotion Long Description(EN)"] = df.iloc[:, 29].apply(clean_empty_text)
        
        out["Promotion Name(ZH)"] = df.iloc[:, 24].apply(clean_empty_text)
        out["Promotion Short Description(ZH)"] = df.iloc[:, 26].apply(clean_empty_text)
        out["Promotion Long Description(ZH)"] = df.iloc[:, 28].apply(clean_empty_text)
        
        # Product Code 拆分
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
        
        # AI 時間拆分
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
        
        # 二次過濾「系統排序」列
        out = out[~out["Internal Name"].astype(str).str.contains("系統排序", na=False)]
        
        # 產出 Excel 下載流
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            out.to_excel(writer, index=False)
        output.seek(0)
        
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name='data.xlsx'
        )
    except Exception as e:
        return f"轉換過程中發生錯誤: {str(e)}", 500

if __name__ == '__main__':
    app.run(debug=True, port=5000)