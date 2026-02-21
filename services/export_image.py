import io
import os
import re
from PIL import Image, ImageDraw, ImageFont, ImageOps

# ==========================================
# 1. 您的完美座標字典 (已修正寬度限制)
# ==========================================
COORDS_DICT = {
    '案名': (398, 410, 800), 
    '地址': (400, 529, 800),
    '社區名稱': (388, 1647, 400),
    '承辦人及電話': (1570, 2891, 800),
    
    # --- 其他標準欄位 ---
    '物件類型': (128, 194), '案件類型': (1768, 117),
    '登記總建坪': (555, 650), '主建物坪數': (420, 778),
    '附屬建物坪數': (420, 912), '公設坪數': (420, 1046), '不含車位坪數': (420, 1178), '車位坪數': (420, 1306),
    '地上層': (528, 1439), '地下層': (766, 1439), '位於樓層': (964, 1439), '建築完成日': (483, 1548, 800),
    '公設比': (888, 1647), '管理費': (388, 1745), '繳納方式': (688, 1745),
    '學校': (388, 1844), '市場': (388, 1945), '公園': (388, 2046), '面臨路寬': (388, 2148), '面道路': (688, 2148),
    '售價': (1600, 450, 600), '房屋單價': (1313, 661), '車位價格': (1809, 662), '格局': (1400, 830, 600 ),
    '土地面積': (1313, 1045), '權利範圍': (1900, 1059, 120), '座向': (1313, 1184), '貸款設定': (1900, 1182, 120),
    '車位形式': (1313, 1319), '車位樓層': (1637, 1306, 120), '汽車編號': (1888, 1321, 120), '機車位樓層': (1635, 1431, 120),
    '機車編號': (1888, 1431, 120), '建物KEY': (1313, 1548), '使用現況': (1313, 1647), '有無警衛': (1313, 1747),
    '總戶數': (1921, 1747), '同層戶數': (1313, 1845), '電梯數': (1921, 1845), '房地合一': (1313, 1945),
    '瓦斯': (1313, 2046), '冒泡位置圖': (122, 2289), '物件特色描述': (1064, 2314), 
    '委託契約書編號': (500, 270)
}

# ==========================================
# 2. 客製化字體與顏色樣式表
# ==========================================
FIELD_STYLES = {
    '案名': {'size': 16},
    '地址': {'size': 16},
    '車位形式': {'color': (255, 0, 0)},       
    '社區名稱': {'color': (255, 0, 0)},       
    '繳納方式': {'size': 16, 'color': (255, 0, 0)}, 
    '同層戶數': {'color': (255, 0, 0)},       
    '售價': {'size': 28, 'color': (255, 0, 0), 'bold': True},   
    '登記總建坪': {'size': 22, 'color': (255, 0, 0), 'bold': True}, 
    '主建物坪數': {'size': 16, 'color': (0, 0, 0)}, 
    '公設坪數': {'size': 16},
    '附屬建物坪數': {'size': 16},
    '不含車位坪數': {'size': 16},
    '車位坪數': {'size': 16},
    '格局': {'size': 22, 'bold': True}, # 🔥 已改為 28 級粗體
    '權利範圍': {'size': 9},
    '汽車編號': {'size': 16},
    '車位樓層': {'size': 16},
    '機車編號': {'size': 16},
    '機車位樓層': {'size': 16},
    '委託契約書編號': {'size': 16},
    '管理費': {'size': 16},
    '土地面積': {'size': 16},
}

# ==========================================
# 3. 雙字體智慧抓取模組
# ==========================================
def format_layout(val):
    if not val: return ""
    parts = val.split('/')
    labels = ['房', '廳', '衛', '陽台']
    result = ""
    for i, part in enumerate(parts):
        if i < len(labels): result += f"{part}{labels[i]}"
    return result

def format_tw_date(val):
    if not val: return ""
    match = re.match(r'^(\d{2,3})[/-](\d{1,2})[/-](\d{1,2})$', val.strip())
    if match:
        y, m, d = match.groups()
        return f"民國{y}年{m}月{d}號"
    return val

# 🔥 新增：根據是否為粗體，自動抓取不同的字體檔案
def get_font_path(is_bold=False):
    if is_bold:
        font_paths = [
            "NotoSansTC-Bold.ttf",             # 您的雲端粗體
            "C:\\Windows\\Fonts\\msjhbd.ttc",  # 本機正黑體粗體 (備用)
        ]
    else:
        font_paths = [
            "NotoSansTC-Regular.ttf",          # 您的雲端常規體
            "C:\\Windows\\Fonts\\msjh.ttc",    # 本機正黑體 (備用)
        ]
        
    for path in font_paths:
        if os.path.exists(path):
            return path
    return None

def get_safe_text_width(draw, text, font):
    try:
        bbox = draw.multiline_textbbox((0, 0), text, font=font)
        return bbox[2] - bbox[0]
    except AttributeError:
        lines = text.split('\n')
        max_w = 0
        for line in lines:
            try: w = draw.textlength(line, font=font)
            except: w = draw.textsize(line, font=font)[0]
            if w > max_w: max_w = w
        return max_w

# 🔥 已更新：直接在畫字時判斷要用哪個字體檔
def draw_text_auto_shrink(draw, text, xy, max_width, excel_pt_size, fill=(0,0,0), is_bold=False):
    actual_font_path = get_font_path(is_bold)
    
    if not actual_font_path:
        font = ImageFont.load_default()
        draw.text(xy, text, fill=fill, font=font)
        return

    ratio = 3.5
    font_size = int(excel_pt_size * ratio)
    min_font_size = int(8.5 * ratio)  
    font = ImageFont.truetype(actual_font_path, font_size)
    
    while font_size > min_font_size:
        if get_safe_text_width(draw, text, font) <= max_width:
            break
        font_size -= 2
        font = ImageFont.truetype(actual_font_path, font_size)
        
    final_lines = []
    for existing_line in text.split('\n'):
        current_line = ""
        for char in existing_line:
            test_line = current_line + char
            if get_safe_text_width(draw, test_line, font) <= max_width:
                current_line = test_line
            else:
                if current_line: final_lines.append(current_line)
                current_line = char
        if current_line: final_lines.append(current_line)
        
    final_text = "\n".join(final_lines)
        
    line_count = len(final_lines)
    spacing = 8  
    y_offset = int((excel_pt_size * ratio) // 2)
    
    if line_count > 1:
        extra_height = (line_count - 1) * (font_size + spacing)
        y_offset += extra_height // 2

    # 如果我們成功載入了實體的粗體字檔 (檔名有 bold)，就不需要額外的筆觸加粗了
    needs_stroke = False
    if is_bold:
        path_lower = actual_font_path.lower()
        if "bold" not in path_lower and "bd" not in path_lower:
            needs_stroke = True

    stroke_width = 1 if needs_stroke else 0
    draw.text((xy[0], xy[1] - y_offset), final_text, fill=fill, font=font, spacing=spacing, stroke_width=stroke_width, stroke_fill=fill)

# ==========================================
# 4. 核心繪圖邏輯
# ==========================================
def generate_jpg_from_template(template_img_path, final_data, map_image_file, desc_input):
    if not os.path.exists(template_img_path):
        raise FileNotFoundError(f"找不到底圖：{template_img_path}，請確保您的空白底圖放在正確位置。")
        
    img = Image.open(template_img_path).convert("RGB")
    draw = ImageDraw.Draw(img)
    
    # 特色描述固定使用一般字體 (is_bold=False)
    font_path_regular = get_font_path(is_bold=False)
    
    desc_input = str(desc_input).upper() if desc_input else ""
    
    for field_name, coords in COORDS_DICT.items():
        if len(coords) == 3:
            x, y, max_w = coords
        else:
            x, y = coords
            max_w = 250  
            
        if field_name == '冒泡位置圖' and map_image_file:
            try:
                map_img = Image.open(map_image_file).convert("RGB")
                # 🔥 自動處理手機上傳照片橫直顛倒的問題
                map_img = ImageOps.exif_transpose(map_img)
                map_img = ImageOps.fit(map_img, (906, 542), method=Image.Resampling.LANCZOS, centering=(0.5, 0.5))
                img.paste(map_img, (x, y))
            except Exception as e:
                print(f"貼上冒泡圖失敗: {e}")
            continue
            
        # =======================================================
        # 🔥 特殊處理：物件特色描述 (真實邊界框黑科技)
        # =======================================================
        if field_name == '物件特色描述' and desc_input:
            desc_font_pt = 12 
            
            for k, v in final_data.items():
                val_str = str(v.get('value', '')) if isinstance(v, dict) else str(v)
                match = re.search(r'(\d+)\s*級', val_str)
                if match:
                    desc_font_pt = int(match.group(1))
                    break
            
            desc_px_size = int(desc_font_pt * 3.5)
            font_desc = ImageFont.truetype(font_path_regular, desc_px_size) if font_path_regular else ImageFont.load_default()
            
            box_max_w = 2029 - x       
            box_max_h = 2824 - (y-20)  
            spacing = 15               
            
            final_lines = []
            for paragraph in desc_input.split('\n'):
                current_line = ""
                for char in paragraph:
                    test_line = current_line + char
                    if get_safe_text_width(draw, test_line, font_desc) <= box_max_w:
                        current_line = test_line
                    else:
                        if current_line: final_lines.append(current_line)
                        current_line = char
                if current_line: final_lines.append(current_line)
                
            line_height = desc_px_size + spacing
            max_allowed_lines = int(box_max_h / line_height)
            
            if len(final_lines) > max_allowed_lines:
                final_lines = final_lines[:max_allowed_lines]
                if len(final_lines[-1]) > 1:
                    final_lines[-1] = final_lines[-1][:-1] + "..."
                    
            wrapped_desc = "\n".join(final_lines)
            draw.text((x, y - 20), wrapped_desc, fill=(0, 0, 0), font=font_desc, spacing=spacing)
            continue
            
        val_to_draw = ""
        c_type = ""
        suffix = ""
        
        if field_name in final_data:
            val = final_data[field_name].get('value', '')
            c_type = final_data[field_name].get('type', '')
            suffix = final_data[field_name].get('suffix', '')
            if val: val_to_draw = str(val)
        else:
            alt_name = field_name.replace("建物坪數", "建坪數") 
            for data_key, content in final_data.items():
                if f"*{field_name}*" in data_key or field_name in data_key or alt_name in data_key:
                    val = content.get('value', '')
                    c_type = content.get('type', '')
                    suffix = content.get('suffix', '')
                    if val: val_to_draw = str(val)
                    break
                
        if val_to_draw:
            val_to_draw = str(val_to_draw).upper()
            
            if c_type == 'layout': val_to_draw = format_layout(val_to_draw)
            elif c_type == 'date_tw': val_to_draw = format_tw_date(val_to_draw)
            else: val_to_draw = f"{val_to_draw}{suffix}"
            
            style = FIELD_STYLES.get(field_name, {})
            current_size = style.get('size', 12)  
            current_color = style.get('color', (0, 0, 0)) 
            is_bold = style.get('bold', False)
            
            # 🔥 圖片版專屬：案名與地址字數超過 17 字自動換行，並縮小字體
            text_len = len(val_to_draw)
            if field_name == '案名' and text_len > 21:
                current_size = 12
                # 案名：每 15 個字強制切斷換行
                val_to_draw = '\n'.join([val_to_draw[i:i+17] for i in range(0, text_len, 17)])
                
            elif field_name == '地址' and text_len > 22:
                current_size = 10
                # 地址：每 15 個字強制切斷換行
                val_to_draw = '\n'.join([val_to_draw[i:i+22] for i in range(0, text_len, 22)])
            
            draw_text_auto_shrink(draw, str(val_to_draw), (x, y), max_w, current_size, current_color, is_bold)
            
    img_byte_arr = io.BytesIO()
    img.save(img_byte_arr, format='JPEG', quality=95)
    img_byte_arr.seek(0)
    
    return img_byte_arr