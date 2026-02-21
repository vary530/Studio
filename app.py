import streamlit as st
import openpyxl
from openpyxl.styles import Font, Alignment
from openpyxl.drawing.image import Image as XLImage
from PIL import Image, ImageDraw, ImageFont, ImageOps
import io
import re
import unicodedata
import os
import base64

# ==========================================
# 🌟 0. 頁面初始化
# ==========================================
st.set_page_config(
    page_title="Studio",
    page_icon="icon.png",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ==========================================
# 🌟 1. 終極暴力 CSS：鎖定螢幕、模糊背景、獵殺圖示
# ==========================================
bg_path = "background.png" 
bg_b64 = ""
if os.path.exists(bg_path):
    with open(bg_path, "rb") as f:
        bg_b64 = base64.b64encode(f.read()).decode()

st.markdown(f"""
<style>
    /* 👉 1. 終極鎖死畫面，拒絕 iOS 橡皮筋回彈與白邊 */
    html, body {{
        position: fixed !important;
        overflow: hidden !important; 
        width: 100vw !important;
        height: 100vh !important;
        height: 100dvh !important; /* 🔥 新增：針對 iOS 動態高度鎖死 */
        background-color: #000000 !important; /* 確保底色是純黑 */
        touch-action: none !important; /* 沒收瀏覽器滑動權限 */
        -webkit-overflow-scrolling: auto !important;
    }}
    
    /* 👉 2. 強制所有分頁內部容器撐滿 100% 螢幕，沒內容也要撐起來！ */
    [data-testid="stAppViewContainer"], .stApp {{
        background-color: transparent !important;
        overflow-y: auto !important; 
        height: 100vh !important;
        height: 100dvh !important; /* 🔥 新增：強制拉高 */
        min-height: 100dvh !important; /* 🔥 新增：讓矮個子分頁也能撐滿螢幕 */
        -webkit-overflow-scrolling: touch !important;
    }}

    /* 👉 3. 完美模糊背景：使用偽元素放在最底層 */
    [data-testid="stAppViewContainer"]::before {{
        content: "";
        position: fixed;
        top: 0; left: 0; right: 0; bottom: 0;
        background-image: url("data:image/png;base64,{bg_b64}");
        background-size: cover;
        background-position: center;
        background-repeat: no-repeat;
        filter: blur(8px) brightness(0.7); 
        -webkit-filter: blur(8px) brightness(0.7);
        transform: scale(1.1); 
        z-index: -1;
    }}

    /* 👉 4. 隱藏頂部選單與邊距調整 (避開上面那條白白的) */
    [data-testid="stHeader"] {{
        display: none !important;
    }}
    .block-container {{
        /* 🔥 把 top 加大，把內容往下推，避開 iOS 的上方狀態列 */
        padding-top: 4rem !important; 
        padding-bottom: 3rem !important;
        padding-left: 1.5rem !important;
        padding-right: 1.5rem !important;
        max-width: 100% !important;
        min-height: 100vh !important; /* 🔥 新增：強制內容區塊也撐高 */
        z-index: 1; 
    }}

    /* 👉 5. 終極無差別獵殺右下角圖示 */
    [data-testid="stViewerBadge"], 
    [class^="viewerBadge"], 
    div[class*="viewerBadge"], 
    iframe[title*="streamlit"],
    #MainMenu, footer {{
        display: none !important;
        opacity: 0 !important;
        pointer-events: none !important;
        visibility: hidden !important;
    }}
</style>
""", unsafe_allow_html=True)
# 嘗試匯入 PDF 模組
try:
    from services.pdf_scanner import extract_text_from_pdf
    PDF_MODULE_ACTIVE = True
except ImportError:
    PDF_MODULE_ACTIVE = False

# ==========================================
# 2. 全域變數與設定 (接下來接你原本後面的程式碼...)
# ==========================================

# 字體縮小邏輯
FONT_LOGIC = {10: (27, 11), 11: (23, 10), 12: (23, 9), 14: (18, 8), 16: (17, 7)}

# 選項欄位設定
OPTION_FIELDS = [
    ("""*物件類型*□店面□商辦□別墅□透天□電梯大樓□華廈□套房□公寓□廠房□農舍□其他______""", 'option_mark', None),
    ("""*車位形式*□坡道平面□坡道機械□升降平面□升降機械□庭院□平移機械□獨立車庫□搭式車位□無""", 'option_replace', None),
    ("""*座向*□座東朝西□座西朝東□座南朝北□座北朝南□座東南朝西北□座東北朝西南□座西南朝東北□座西北朝東南""", 'option_replace', None),
    ("""*有無警衛*□有 □無""", 'option_mark', None),
    ("""*瓦斯*□天然瓦斯  □桶裝""", 'option_mark', None),
    ("""*繳納方式*□月繳□年繳□季繳""", 'option_replace', {"custom_label": "管理費繳納方式"}),
    ("""*使用現況*□空屋  □自住  □出租""", 'option_mark', None),
    ("""*建物KEY*□公司 □警衛室 □洽開發""", 'option_mark', None),
    ("""*面道路*□雙向道□單向道□無尾巷""", 'option_mark', None),
    ("""*房地合一*□有 □無""", 'option_mark', None),
    ("""*案件類型*□1 □2 □3""", 'option_mark', {'mapping': {'1': '1.中人', '2': '2.開底合一', '3': '3.專任委託'}}),
]

# 基本資料欄位設定
BASIC_FIELDS_CONFIG = [
    {"label": "案名", "key": "案名", "type": "text", "note": ""},
    {"label": "售價", "key": "售價", "type": "number", "note": "萬"},
    {"label": "社區名稱", "key": "社區名稱", "type": "text", "note": ""},
    {"label": "地址", "key": "地址", "type": "text", "note": ""},
    {"label": "總樓層", "key": "地上層", "type": "text", "note": ""},  
    {"label": "位於樓層", "key": "位於樓層", "type": "text", "note": ""},
    {"label": "建築完成日(例:82/12/22)", "key": "建築完成日", "type": "date_tw", "note": ""},
    {"label": "登記總建坪", "key": "登記總建坪", "type": "number", "note": ""},
    {"label": "主建物坪數", "key": "主建物坪數", "type": "number", "note": ""},
    {"label": "附屬建坪數", "key": "附屬建坪數", "type": "number", "note": ""},
    {"label": "公設坪數", "key": "公設坪數", "type": "number", "note": ""},
    {"label": "車位坪數", "key": "車位坪數", "type": "number", "note": ""},
    {"label": "不含車位坪數", "key": "不含車位坪數", "type": "number_overwrite", "note": ""},
    {"label": "汽車編號", "key": "汽車編號", "type": "text", "note": ""},
    {"label": "汽車位樓層", "key": "車位樓層", "type": "text", "note": ""}, 
    {"label": "格局 (例:3/2/2)", "key": "格局", "type": "layout", "note": ""},
    {"label": "管理費", "key": "管理費", "type": "number", "note": "元"},
    {"label": "總戶數", "key": "總戶數", "type": "number", "note": ""},
    {"label": "同層戶數", "key": "同層戶數", "type": "number", "note": "戶"},
    {"label": "電梯數", "key": "電梯數", "type": "number", "note": "梯"},
    {"label": "地下共幾層", "key": "地下層", "type": "text", "note": ""},
    {"label": "學校", "key": "學校", "type": "text", "note": ""},
    {"label": "市場", "key": "市場", "type": "text", "note": ""},
    {"label": "公園", "key": "公園", "type": "text", "note": ""},
    {"label": "房屋單價 (填0自動算)", "key": "房屋單價", "type": "formula_price", "note": "萬"},
    {"label": "公設比 (填0自動算)", "key": "公設比", "type": "formula_ratio", "note": "%"},
    {"label": "土地面積", "key": "土地面積", "type": "number", "note": "坪"},
    {"label": "權利範圍", "key": "權利範圍", "type": "text", "note": ""},
    {"label": "車位價格", "key": "車位價格", "type": "number", "note": "萬"},
    {"label": "面臨路寬", "key": "面臨路寬", "type": "number", "note": "米"},
    {"label": "機車樓層", "key": "機車位樓層", "type": "text_overwrite", "note": ""}, 
    {"label": "機車編號", "key": "機車編號", "type": "text", "note": ""},
    {"label": "貸款設定", "key": "貸款設定", "type": "number", "note": "萬"},
    {"label": "委託書編號", "key": "委託契約書編號", "type": "text", "note": ""}, 
    {"label": "承辦人電話", "key": "承辦人及電話", "type": "text_area", "note": ""}, 
]

# ==========================================
# 2. 輔助函式
# ==========================================

def clear_all_options_callback():
    for item in OPTION_FIELDS:
        key = item[0]
        if key in st.session_state:
            st.session_state[key] = None

def smart_format(val):
    if not val: return ""
    s_val = str(val)
    try:
        if float(s_val.replace(',','')) == 0: return "0"
    except: pass
    s_val = re.sub(r'\.0+(?=\D|$)', '', s_val) 
    s_val = re.sub(r'(\.\d*[1-9])0+(?=\D|$)', r'\1', s_val) 
    return s_val

def get_visual_width(char):
    ew = unicodedata.east_asian_width(char)
    if ew in ['F', 'W', 'A']: return 1.0
    return 0.5

def split_text_by_visual_width(text, max_len):
    if not text: return ""
    result_lines = []
    original_paragraphs = text.split('\n')
    for paragraph in original_paragraphs:
        current_line, current_width = "", 0.0
        for char in paragraph:
            char_w = get_visual_width(char)
            if current_width + char_w > max_len:
                result_lines.append(current_line)
                current_line = char
                current_width = char_w
            else:
                current_line += char
                current_width += char_w
        result_lines.append(current_line)
    return '\n'.join(result_lines)

def parse_option_placeholder(placeholder):
    clean_ph = placeholder.replace('"""', '')
    match = re.match(r'\*(.*?)\*(.*)', clean_ph)
    if match:
        title, raw_options = match.group(1), match.group(2)
        options = [opt.strip() for opt in raw_options.split('□') if opt.strip()]
        clean_options = [opt.replace('______', '') if '______' in opt else opt for opt in options]
        return title, clean_options, raw_options
    return clean_ph, [], ""

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

def try_float(val):
    try:
        if isinstance(val, str): val = val.replace(',', '') 
        return float(val)
    except: return 0.0

def crop_and_resize_image(image_file):
    try:
        img = Image.open(image_file)
        img_w, img_h = img.size
        # 這是 Excel 畫面上要顯示的長寬 (維持不變，保證不跑版)
        target_w, target_h = 329, 197
        
        target_ratio = target_w / target_h
        current_ratio = img_w / img_h
        
        if current_ratio > target_ratio:
            new_width = int(img_h * target_ratio)
            new_height = img_h
            left, top = (img_w - new_width) // 2, 0
            right, bottom = left + new_width, img_h
        else:
            new_width = img_w
            new_height = int(img_w / target_ratio)
            left, top = 0, (img_h - new_height) // 2
            right, bottom = img_w, top + new_height
            
        img_cropped = img.crop((left, top, right, bottom))
        
        # 【關鍵修復】：把實際塞入 Excel 的圖片像素放大 4 倍 (1312x788)，解決 PDF 模糊問題！
        scale_factor = 4
        hi_res_w = target_w * scale_factor
        hi_res_h = target_h * scale_factor
        img_resized = img_cropped.resize((hi_res_w, hi_res_h), Image.Resampling.LANCZOS)
        
        img_byte_arr = io.BytesIO()
        img_resized.save(img_byte_arr, format='PNG')
        img_byte_arr.seek(0)
        
        # 回傳高畫質圖片，但最後兩個參數(顯示大小)依然告訴 Excel 只要顯示 328x197
        return img_byte_arr, target_w, target_h
    except Exception as e:
        st.error(f"圖片處理失敗: {e}")
        return None, 0, 0

# ==========================================
# 3. CSS 樣式核心 (V65 RWD 手機優化版)
# ==========================================
def set_bg_hack(main_bg):
    bg_css = ""
    try:
        with open(main_bg, "rb") as f:
            encoded_string = base64.b64encode(f.read()).decode()
        
        bg_css = f"""
        .stApp {{
            background: transparent;
        }}
        .stApp::before {{
            content: "";
            position: fixed;
            top: 0; left: 0; width: 100vw; height: 100vh;
            background-image: url(data:image/png;base64,{encoded_string});
            background-size: cover; 
            background-position: center; 
            background-repeat: no-repeat;
            background-attachment: fixed;
            filter: blur(15px);
            -webkit-filter: blur(15px);
            z-index: -1;
            transform: scale(1.1);
        }}
        """
    except FileNotFoundError:
        bg_css = ".stApp { background-color: #1a1a1a; }"

    st.markdown(
         f"""
         <style>
         {bg_css}

         /* === PWA 優化 === */
         ::-webkit-scrollbar {{ display: none; }}
         body {{ overscroll-behavior-y: none; }}
         header[data-testid="stHeader"] {{ display: none; }}
         footer {{ display: none; }}
         
         /* 主區塊間距 */
         .main .block-container {{
             padding-top: 0rem !important; 
             padding-bottom: 3rem !important;
             max-width: 100%;
             padding-left: 1rem !important; /* RWD 手機版優化 */
             padding-right: 1rem !important; /* RWD 手機版優化 */
         }}

         /* === 全域字體強制白色 === */
         h1, h2, h3, h4, h5, h6, p, label, span, div, li, .stMarkdown {{
             color: #ffffff !important;
             font-family: "Microsoft JhengHei", "Segoe UI", sans-serif;
         }}

         /* === 強制黃色字體 class === */
         #yellow-target {{
             color: #FFD700 !important;
         }}

         /* === PDF 掃描結果區塊 === */
         .info-card {{
             background-color: rgba(20, 20, 20, 0.6);
             border: 1px solid rgba(255, 255, 255, 0.15);
             border-radius: 12px;
             padding: 20px;
             margin-bottom: 15px;
         }}
         
         .info-row {{
             display: flex;
             align-items: baseline;
             margin-bottom: 3px;
             line-height: 1.5;
         }}

         .text-label {{ 
             color: #bbbbbb !important; 
             font-weight: 500; 
             margin-right: 8px;
             white-space: nowrap;
         }}
         
         .text-label-white {{
             color: #ffffff !important;
             font-weight: 500;
             margin-right: 8px;
             white-space: nowrap;
             font-size: 18px;
         }}
         
         .text-value {{ 
             color: #ffffff !important; 
             font-weight: bold; 
             font-size: 15px;
         }}
         
         .text-yellow {{
             color: #FFD700 !important;
             font-size: 22px;
             font-weight: 800;
             text-shadow: 0 0 10px rgba(255, 215, 0, 0.2);
         }}
         
         .section-title {{
             font-size: 16px;
             font-weight: bold;
             color: #ffffff !important;
             margin-top: 12px;
             margin-bottom: 2px;
         }}
         
         .sub-text {{
             font-size: 13px;
             color: #cccccc !important;
             margin-left: 0px;
             margin-bottom: 1px;
             font-family: monospace; 
         }}

         /* === 輸入元件美化 === */
         .stTextInput input, .stTextArea textarea, .stNumberInput input, .stSelectbox div[data-baseweb="select"] {{
             background-color: rgba(255, 255, 255, 0.1) !important;
             color: #ffffff !important;
             border: 1px solid rgba(255, 255, 255, 0.3) !important;
             border-radius: 8px;
         }}
         ul[data-baseweb="menu"] {{ background-color: #222222 !important; }}

         /* === 按鈕美化 === */
         .stButton button {{
             background-color: rgba(255, 255, 255, 0.15) !important;
             color: white !important;
             border: 1px solid rgba(255, 255, 255, 0.5) !important;
             border-radius: 20px;
             font-weight: bold;
         }}
         
         a[data-testid="stLinkButton"] {{
             background-color: rgba(66, 133, 244, 0.6) !important;
             color: white !important;
             border-radius: 8px;
             justify-content: center;
             font-weight: bold;
         }}

         /* === Tab 分頁籤美化 (移除紅條) === */
         div[data-testid="stTabs"] button {{
            background-color: transparent !important;
            color: #888888 !important;
            font-size: 16px;
            font-weight: bold;
         }}
         
         div[data-testid="stTabs"] button[aria-selected="true"] {{
            color: #ffffff !important;
            border-bottom: 2px solid #ffffff !important;
         }}
         
         div[data-baseweb="tab-highlight"] {{
             background-color: #ffffff !important;
         }}
         
         /* === 上傳元件文字替換 (RWD 手機優化版) === */
         
         /* 隱藏原生按鈕，但保留點擊功能 */
         div[data-testid="stFileUploader"] > section > button {{
             font-size: 0 !important;
             border: none !important;
             background: transparent !important;
             width: 100% !important;
             height: 100% !important;
             position: absolute;
             top: 0; left: 0;
             z-index: 2;
             display: block !important;
         }}
         
         /* 上傳方格容器樣式 */
         div[data-testid="stFileUploader"] section {{
             background-color: rgba(255, 255, 255, 0.15) !important;
             border: 1px dashed rgba(255, 255, 255, 0.4) !important;
             border-radius: 10px;
             padding: 0px !important;
             min-height: 80px !important; /* 加高，方便手指點擊 */
             position: relative;
             display: flex; 
             align-items: center; 
             justify-content: center;
             width: 100% !important; /* RWD 填滿寬度 */
         }}
         
         /* 隱藏拖曳文字 */
         div[data-testid="stFileUploader"] span {{ display: none !important; }}

         /* 文字樣式 (絕對置中) */
         .st-key-scanner_upload div[data-testid="stFileUploader"] section::after {{
             content: "建物謄本PDF上傳";
             color: white; 
             font-size: 15px; 
             font-weight: bold;
             position: absolute; 
             top: 50%; left: 50%;
             transform: translate(-50%, -50%);
             pointer-events: none; 
             z-index: 1;
             white-space: nowrap; /* 禁止換行 */
         }}
         
         .st-key-map_uploader div[data-testid="stFileUploader"] section::after {{
             content: "冒泡位置照片上傳";
             color: white; 
             font-size: 15px; 
             font-weight: bold;
             position: absolute; 
             top: 50%; left: 50%;
             transform: translate(-50%, -50%);
             pointer-events: none; 
             z-index: 1;
             white-space: nowrap; /* 禁止換行 */
         }}

         /* Expander & Toast */
         .streamlit-expanderHeader {{ background-color: rgba(255,255,255,0.1) !important; border-radius: 8px; }}
         div[data-testid="stToast"] {{ background-color: rgba(50,50,50,0.95) !important; color: white !important; border-radius: 10px; }}
         </style>
         """,
         unsafe_allow_html=True
     )

# 套用背景
set_bg_hack('background.png')

# 標題
st.markdown('<h2 style="text-shadow: 2px 2px 5px black; text-align:center; margin-bottom: 10px; margin-top: -60px;">物調整合Studio</h2>', unsafe_allow_html=True)

# --- 自動建立輸出資料夾 ---
if not os.path.exists("outputs"):
    os.makedirs("outputs")

# --- 分頁建構 ---
tab1, tab2, tab3 = st.tabs(["PDF 智慧掃描", "不動產物調產出", " 歷史產出紀錄"])

# ==========================================
# 分頁 1: PDF 掃描
# ==========================================
with tab1:
    def render_row(label, value):
        val_str = smart_format(value)
        return f'<div class="info-row"><span class="text-label">{label}</span><span class="text-value">{val_str}</span></div>'

    st.subheader("PDF 智慧掃描")
    
    if PDF_MODULE_ACTIVE:
        uploaded_pdf = st.file_uploader("上傳建物謄本 PDF", type=['pdf'], key="scanner_upload")
        
        if uploaded_pdf and st.button("開始掃描", type="primary"):
            with st.spinner("掃描中..."):
                data, raw_text = extract_text_from_pdf(uploaded_pdf)
                
                if "error" in data:
                    st.error(f"解析失敗: {data['error']}")
                else:
                    # 1. 日期處理
                    raw_date = data.get('completion_date', '')
                    clean_date = raw_date.split('（')[0].strip() if '（' in raw_date else raw_date
                    
                    # 2. 車位編號抓取 (精準抓取所有數字)
                    car_numbers = []
                    p_items = data.get('parking_items', [])
                    for item in p_items:
                        id_match = re.search(r'編號(.*?)[號：]', item)
                        if id_match:
                            raw_id_content = id_match.group(1)
                            found_digits = re.findall(r'\d+', raw_id_content)
                            if found_digits:
                                car_numbers.append(found_digits[-1]) 
                    
                    extracted_car_no = " ".join(car_numbers)

                    # 3. 存入 Session
                    processed_data = data.copy()
                    processed_data['clean_completion_date'] = clean_date
                    processed_data['extracted_car_no'] = extracted_car_no
                    st.session_state['pdf_cached_data'] = processed_data
                    
                    st.success("解析成功！資料已暫存。")

        # === 顯示掃描結果 (Persistent) ===
        if 'pdf_cached_data' in st.session_state and st.session_state['pdf_cached_data']:
            data = st.session_state['pdf_cached_data']
            st.markdown("<br>", unsafe_allow_html=True)
            
            # 卡片1: 基本資料
            basic_parts = []
            basic_parts.append('<div class="info-card">')
            basic_parts.append(render_row("所有權人：", data.get('owner_display', '')))
            basic_parts.append(render_row("取得日期：", data.get('acquisition_date_display', '')))
            basic_parts.append(render_row("登記原因：", data.get('reg_reason', '')))
            basic_parts.append(render_row("主要用途：", data.get('usage', '')))
            basic_parts.append(render_row("主要建材：", data.get('material', '')))
            basic_parts.append(render_row("建物門牌：", data.get('address', '')))
            basic_parts.append(render_row("建築完成日：", data.get('completion_date', ''))) 
            basic_parts.append(render_row("總樓層：", data.get('total_floors', '')))
            basic_parts.append(render_row("位於樓層：", data.get('layer', '')))
            basic_parts.append('</div>')
            st.markdown("".join(basic_parts), unsafe_allow_html=True)

            # 卡片 2 (坪數)
            area_parts = []
            area_parts.append('<div class="info-card">')
            
            # 登記總坪 (標題白色，數值黃色)
            area_parts.append('<div style="margin-bottom: 20px; border-bottom: 1px solid rgba(255,255,255,0.2); padding-bottom: 10px;">')
            area_parts.append(f'<span class="text-label-white">登記總坪數：</span> <span class="text-yellow">{smart_format(data.get("reg_total"))}坪</span>')
            area_parts.append(f'<span style="font-size:14px; color:#cccccc; margin-left:10px;">(含車位{smart_format(data.get("parking_total"))}坪)</span>')
            area_parts.append('</div>')
            
            # 主建物
            area_parts.append(f'<div class="section-title">主建物：{smart_format(data.get("main_total"))}坪</div>')
            for item in data.get('main_items', []):
                parts = item.split('：')
                if len(parts) == 2: item = f"{parts[0]}：{smart_format(parts[1].replace('坪', ''))}坪"
                area_parts.append(f"<div class='sub-text'>{item}</div>")
            
            # 附屬建物
            area_parts.append(f'<div class="section-title">附屬建物：{smart_format(data.get("annex_total"))}坪</div>')
            for item in data.get('annex_items', []):
                parts = item.split('：')
                if len(parts) == 2: item = f"{parts[0]}：{smart_format(parts[1].replace('坪', ''))}坪"
                area_parts.append(f"<div class='sub-text'>{item}</div>")
            
            # 公設
            area_parts.append(f'<div class="section-title">公設坪數：{smart_format(data.get("net_common_str"))}坪</div>')
            area_parts.append(f"<div class='sub-text'>-共有總坪數：{smart_format(data.get('gross_common_str'))}坪(含車位:{smart_format(data.get('parking_total'))}坪)</div>")
            for item in data.get('common_items', []):
                parts = item.split('：')
                if len(parts) == 2: item = f"{parts[0]}：{smart_format(parts[1].replace('坪', ''))}坪"
                area_parts.append(f"<div class='sub-text'>{item}</div>")
            
            # 車位
            area_parts.append(f'<div class="section-title">車位坪數：{smart_format(data.get("parking_total"))}坪</div>')
            p_items = data.get('parking_items', [])
            if p_items:
                for item in p_items:
                    parts = item.split('：')
                    if len(parts) == 2: item = f"{parts[0]}：{smart_format(parts[1].replace('坪', ''))}坪"
                    area_parts.append(f"<div class='sub-text'>{item}</div>")
            else:
                area_parts.append("<div class='sub-text'>(無車位資料)</div>")
            
            # 不含車位
            area_parts.append('<div style="margin-top:20px; border-top:1px solid rgba(255,255,255,0.2); padding-top:10px;">')
            area_parts.append(f'<span class="section-title">不含車位總坪數：{smart_format(data.get("net_reg_total"))}坪</span>')
            area_parts.append('</div>')
            
            area_parts.append('</div>')
            st.markdown("".join(area_parts), unsafe_allow_html=True)

            # 3. 他項權利 (ID 強制變黃)
            rights_count = int(data.get('rights_count', 0))
            
            st.markdown(f"""
            <div class="info-card" style="padding: 15px;">
                <div style="font-weight:bold; font-size:16px; color:white;">相關他項權利登記次序：共{rights_count}筆</div>
            </div>
            """, unsafe_allow_html=True)

            if rights_count > 0:
                with st.expander("點擊查看他項權利明細"):
                    rights_content = data.get('rights_display', [])
                    for r in rights_content:
                        r_html = r.strip().replace("\n", "<br>")
                        st.markdown(f"""
                        <div style='color:#cccccc; font-size:14px; margin-bottom:15px; border-bottom:1px dashed #555; padding-bottom:10px; line-height:1.6;'>
                            {r_html}
                        </div>""", unsafe_allow_html=True)
                    
                    # 使用 ID selector 強制變黃
                    st.markdown(f"""
                    <div style="margin-top: 10px; text-align: right; padding-top: 5px;">
                        <span id="yellow-target" style="font-size: 18px; font-weight: bold;">總擔保債權總金額：{data.get('total_rights_money')}</span>
                    </div>""", unsafe_allow_html=True)
    else:
        if not PDF_MODULE_ACTIVE:
            st.error("找不到 `services/pdf_scanner.py`。")

# ==========================================
# 分頁 2: 不動產物調產出
# ==========================================
with tab2:
    if 'pdf_cached_data' in st.session_state and st.session_state['pdf_cached_data']:
        st.info("暫存區有一筆掃描資料，是否匯入？")
        if st.button("匯入掃描資料", type="primary"):
            data = st.session_state['pdf_cached_data']
            
            st.session_state["地址"] = data.get("address", "")
            st.session_state["建築完成日"] = data.get("clean_completion_date", "")
            st.session_state["地上層"] = data.get("total_floors", "")
            st.session_state["位於樓層"] = data.get("layer", "")
            st.session_state["登記總建坪"] = smart_format(data.get("reg_total", ""))
            st.session_state["主建物坪數"] = smart_format(data.get("main_total", ""))
            st.session_state["附屬建坪數"] = smart_format(data.get("annex_total", ""))
            st.session_state["公設坪數"] = smart_format(data.get("gross_common_str", "")) 
            st.session_state["車位坪數"] = smart_format(data.get("parking_total", ""))
            st.session_state["不含車位坪數"] = smart_format(data.get("net_reg_total", ""))
            st.session_state["汽車編號"] = data.get("extracted_car_no", "")
            
            st.toast("資料匯入完成！")
            st.rerun()

    # 填表介面
    st.markdown("### 基本資料")
    input_vars = {} 
    parsed_data = {}

    for i in range(0, len(BASIC_FIELDS_CONFIG), 3):
        cols = st.columns(3)
        for j in range(3):
            if i + j < len(BASIC_FIELDS_CONFIG):
                item = BASIC_FIELDS_CONFIG[i + j]
                with cols[j]:
                    
                    # === 修正重點：初始化 Session State，而不是用 value 塞入 ===
                    if item['key'] not in st.session_state:
                        st.session_state[item['key']] = ""

                    if item['key'] == "地址":
                        addr_col1, addr_col2 = st.columns([0.7, 0.3])
                        with addr_col1:
                            # 把 value=default_val 刪除了，因為有了 key=item['key'] 系統就會自動帶入
                            val = st.text_input(label=item['label'], placeholder=item['label'], label_visibility="collapsed", key=item['key'])
                        with addr_col2:
                            map_url = f"https://www.google.com/maps/search/?api=1&query={val}" if val else "https://www.google.com/maps"
                            st.link_button("查看地址位置", map_url, use_container_width=True)
                    else:
                        if item.get('type') == 'text_area':
                            # 把 value=default_val 刪除了
                            val = st.text_area(label=item['label'], placeholder=item['label'], label_visibility="collapsed", key=item['key'], height=68)
                        else:
                            # 把 value=default_val 刪除了
                            val = st.text_input(label=item['label'], placeholder=item['label'], label_visibility="collapsed", key=item['key'])
                    
                    input_vars[item['key']] = val
                    suffix = item['note']
                    if "(" in suffix or "例" in suffix: suffix = ""
                    parsed_data[item['key']] = {'value': val, 'type': item['type'], 'suffix': suffix}

    st.markdown("---")
    st.markdown("### 選項勾選")
    options_data = {}

    for i in range(0, len(OPTION_FIELDS), 2):
        cols_opt = st.columns(2)
        for j in range(2):
            if i + j < len(OPTION_FIELDS):
                item = OPTION_FIELDS[i + j]
                placeholder, dtype, extra = item[0], item[1], item[2]
                with cols_opt[j]:
                    raw_label = placeholder.replace('"""', '')
                    title_match = re.match(r'\*(.*?)\*', raw_label)
                    display_label = title_match.group(1) if title_match else raw_label
                    if extra and isinstance(extra, dict) and 'custom_label' in extra:
                        display_label = extra['custom_label']
                    _, options, _ = parse_option_placeholder(placeholder)
                    display_options = [extra['mapping'].get(opt, opt) if extra and 'mapping' in extra else opt for opt in options]
                    selected_idx = None
                    if len(options) <= 3:
                        st.markdown(f"**{display_label}**")
                        selected_idx = st.radio(display_label, range(len(options)), format_func=lambda x: display_options[x], horizontal=True, index=None, key=placeholder, label_visibility="collapsed")
                    else:
                        selected_idx = st.selectbox(display_label, range(len(options)), format_func=lambda x: display_options[x], index=None, placeholder=f"請選擇...", key=placeholder)
                    val = options[selected_idx] if selected_idx is not None else None
                    options_data[placeholder] = {'type': dtype, 'value': val, 'raw_full_str': parse_option_placeholder(placeholder)[2]}

    st.write("") 
    if st.button("清空所有選項", type="secondary", on_click=clear_all_options_callback): pass

    st.markdown("---")
    st.markdown("### 冒泡位置圖")
    map_image_file = st.file_uploader("請上傳圖片 (自動調整為 328x197 像素)", type=['png', 'jpg', 'jpeg'], key="map_uploader")
    if map_image_file: st.caption("圖片已上傳，將自動調整為 328x197 像素並插入 Excel")

    st.markdown("---")
    st.markdown("### 物件特色描述")

    if 'font_size' not in st.session_state: st.session_state['font_size'] = 12
    font_size = st.radio("選擇字體大小：", options=[10, 11, 12, 14, 16], horizontal=True, index=2)
    st.session_state['font_size'] = font_size
    max_char_width, max_line = FONT_LOGIC[font_size]
    st.info(f"設定：字體 {font_size} 級 | 每行 {max_char_width} 字 | 上限 {max_line} 行")

    desc_input = st.text_area("特色內容輸入", height=150, key="desc_input_area", label_visibility="collapsed", placeholder="請在此輸入特色描述... (輸入後點擊空白處即可預覽)")

    formatted_text = ""
    line_count = 0
    if desc_input:
        formatted_text = split_text_by_visual_width(desc_input, max_char_width)
        line_count = len(formatted_text.split('\n'))
        progress = min(line_count / max_line, 1.0)
        if line_count > max_line: st.markdown(f"**已達 {line_count} 行 (上限 {max_line})，超出範圍！**")
        else: st.progress(progress)

    st.markdown(f"""<div style="margin-top: 5px; background-color: rgba(0,0,0,0.5); border: 2px dashed #ccc; padding: 10px; width: 100%; color: white; font-family: 'DFKai-SB', 'KaiTi', 'BiauKai', monospace; font-size: {font_size}pt; line-height: 1.25; white-space: pre; overflow-x: auto;">{formatted_text if formatted_text else "(預覽區域：點擊產生物調表顯示)"}</div>""", unsafe_allow_html=True)

    st.markdown("---")
    
    # 建立雙排按鈕
    col_btn1, col_btn2 = st.columns(2)
    with col_btn1:
        submit_btn = st.button("產生物調表 (Excel)", type="primary", use_container_width=True)
    with col_btn2:
        img_btn = st.button("產生物調表圖片 (JPG)", type="secondary", use_container_width=True)

    # ==========================
    # 觸發功能 1：產生完美 JPG
    # ==========================
    if img_btn:
        # 計算數值 (與下方 Excel 邏輯相同)
        price = try_float(input_vars.get("售價", 0))
        area = try_float(input_vars.get("不含車位坪數", 0))
        public = try_float(input_vars.get("公設坪數", 0))
        
        # (保留您原本計算單價與公設比的邏輯)
        if "房屋單價" in parsed_data and parsed_data["房屋單價"]['value'] == '0' and area > 0:
            parsed_data["房屋單價"]['value'] = f"{(price/area):.2f}"
        if "公設比" in parsed_data and parsed_data["公設比"]['value'] == '0' and area > 0:
            parsed_data["公設比"]['value'] = f"{(public/area)*100:.1f}"
            
        # ==========================================
        # 🌟 修正 1：先建立 final_data 箱子
        # ==========================================
        final_data = {**parsed_data, **options_data}
        
        # ==========================================
        # 🌟 修正 2：【終極完美解法】直接抓取畫面上計算好的變數！
        # ==========================================
        # 放棄去記憶體尋找字串！既然上方已經把 font_size, max_char_width, max_line 算出來了，
        # 我們直接現場把它們組裝成 export_image.py 看得懂的句子！
        layout_str = f"字體 {font_size} 級 | 每行 {max_char_width} 字 | 上限 {max_line} 行"
        
        # 把這句完美組裝的話塞進去給圖片程式讀取
        final_data['自訂排版'] = {'value': layout_str, 'type': 'text', 'suffix': ''}
        
        try:
            from services.export_image import generate_jpg_from_template
            
            # 請確認這張圖片的檔名與您的檔案一致！
            blank_image_path = "blank_template.jpg" 
            
            with st.spinner("正在為您繪製高畫質圖片..."):
                img_bytes = generate_jpg_from_template(blank_image_path, final_data, map_image_file, desc_input)
                
            safe_name = str(input_vars.get("案名", "未命名")).strip()

# --- 新增：偷偷備份一份到 outputs 資料夾 ---
            with open(f"outputs/{safe_name}_物調卡.jpg", "wb") as f:
                f.write(img_bytes.getvalue())

            # 1. 成功提示
            st.success("🎉 物調表圖卡繪製成功！")
            
            # 2. 直接在畫面上顯示圖片 (變數換成您專屬的 img_bytes)
            st.image(img_bytes, use_column_width=True)
            
            # 3. 給手機版用戶的溫馨提示
            st.info("📱 手機版用戶：請直接「長按上方圖片」 ➜ 選擇「儲存到相簿」或「分享」即可！")
            
        except Exception as e:
            st.error(f"產生圖片時發生錯誤: {e}")
    # ==========================
    # 觸發功能 2：產生物調表 (Excel)
    # ==========================
    if submit_btn:
        try:
            template_path = "template.xlsx"
            # ...(保留您原本後面寫入 Excel 的全部程式碼)
            wb = openpyxl.load_workbook(template_path)
            ws = wb.active 
            
            contract_id = st.session_state.get('委託契約書編號', "")
            if not contract_id: contract_id = input_vars.get("委託契約書編號", "")
            case_name = st.session_state.get('案名', "")
            if not case_name: case_name = input_vars.get("案名", "")

            contract_id = str(contract_id).strip().upper()
            case_name = str(case_name).strip()
            
            safe_id = re.sub(r'[\\/*?:"<>|]', '', contract_id)
            safe_name = re.sub(r'[\\/*?:"<>|]', '', case_name)
            download_filename = f"{safe_id}{safe_name}.xlsx" if safe_id or safe_name else "物調表_完成.xlsx"

            price = try_float(input_vars.get("售價", 0))
            area = try_float(input_vars.get("不含車位坪數", 0))
            public = try_float(input_vars.get("公設坪數", 0))
            
            if "房屋單價" in parsed_data:
                u_val = parsed_data["房屋單價"]['value']
                c_type = parsed_data["房屋單價"]['type']
                if u_val == '0':
                    if area > 0: parsed_data["房屋單價"]['value'] = f"{(price/area):.2f}"
            
            if "公設比" in parsed_data:
                u_val = parsed_data["公設比"]['value']
                c_type = parsed_data["公設比"]['type']
                if u_val == '0':
                    if area > 0: parsed_data["公設比"]['value'] = f"{(public/area)*100:.1f}"

            for k in ["委託契約書編號", "車位樓層", "機車位樓層"]:
                if k in parsed_data and parsed_data[k]['value']:
                    parsed_data[k]['value'] = str(parsed_data[k]['value']).upper()

            final_data = {**parsed_data, **options_data}
            
            img_to_insert = None
            if map_image_file:
                img_buffer, w, h = crop_and_resize_image(map_image_file)
                if img_buffer:
                    img_to_insert = XLImage(img_buffer)
                    img_to_insert.width, img_to_insert.height = w, h
            
            final_data['"""物件特色描述"""'] = {'type': 'area_text_advanced', 'value': desc_input, 'max_char': max_char_width, 'font_size': font_size}
            auto_shrink_keys = ["汽車編號", "車位樓層", "機車編號", "機車位樓層"]

            for row in ws.iter_rows():
                for cell in row:
                    if cell.value and isinstance(cell.value, str):
                        if '"""冒泡位置圖"""' in cell.value:
                            cell.value = "" 
                            if img_to_insert: ws.add_image(img_to_insert, cell.coordinate)
                            continue

                        if '"""' in cell.value:
                            original_text, is_overwritten = cell.value, False
                            
                            for k, content in final_data.items():
                                ph = f'"""{k}"""' if '*' not in k and '"""' not in k else k
                                    
                                if ph in original_text:
                                    val = content['value']
                                    c_type = content.get('type', 'text')
                                    
                                    if val:
                                        # V60: 字體自動縮小邏輯
                                        v_len = sum(get_visual_width(c) for c in str(val))
                                        curr_font_name = cell.font.name if cell.font else 'KaiTi'
                                        
                                        if k == "案名" and v_len >= 15:
                                            cell.font = Font(name=curr_font_name, size=14)
                                        elif k == "地址" and v_len >= 15:
                                            cell.font = Font(name=curr_font_name, size=12)
                                        elif k in ["學校", "市場", "公園"] and v_len >= 15:
                                            cell.font = Font(name=curr_font_name, size=10)
                                    
                                        if k in auto_shrink_keys and val and len(str(val)) > 3:
                                            cell.font = Font(name=cell.font.name if cell.font else 'KaiTi', size=10)

                                    if c_type in ['text_overwrite', 'number_overwrite']:
                                        suffix = content.get('suffix', '')
                                        cell.value = f"{val}{suffix}" if val else ""
                                        is_overwritten = True
                                        break
                                    elif c_type == 'option_mark':
                                        new_text = content['raw_full_str'].replace(f"□{val}", f"■{val}") if val else content['raw_full_str']
                                        original_text = original_text.replace(ph, new_text)
                                    elif c_type == 'option_replace':
                                        original_text = original_text.replace(ph, val if val else "")
                                    elif c_type == 'area_text_advanced':
                                        if val:
                                            formatted = split_text_by_visual_width(val, content['max_char'])
                                            original_text = original_text.replace(ph, formatted)
                                            cell.alignment = Alignment(wrap_text=True, vertical='top')
                                            cell.font = Font(name=cell.font.name if cell.font else 'KaiTi', size=content['font_size'])
                                        else: original_text = original_text.replace(ph, "")
                                    elif c_type == 'text_area':
                                        suffix = content.get('suffix', '')
                                        if val:
                                            original_text = original_text.replace(ph, f"{val}{suffix}")
                                            cell.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')
                                        else: original_text = original_text.replace(ph, "")
                                    else:
                                        suffix = content.get('suffix', '')
                                        if val:
                                            if c_type == 'layout': val = format_layout(val)
                                            elif c_type == 'date_tw': val = format_tw_date(val)
                                            replacement = f"{val}{suffix}"
                                        else: replacement = ""
                                        
                                        if ph == '"""不含車位坪數"""':
                                            original_text = original_text.replace(ph, replacement).replace("不含", "")
                                        else:
                                            original_text = original_text.replace(ph, replacement)
                            
                            if not is_overwritten: cell.value = original_text.replace('"""', '')

            output = io.BytesIO()
            wb.save(output)
            output.seek(0)
# --- 新增：偷偷備份一份到 outputs 資料夾 ---
            with open(f"outputs/{download_filename}", "wb") as f:
                f.write(output.getvalue())

            st.success(f"Excel 檔案製作完成！檔名：{download_filename}")
            
            # 👇 --- 替換開始：改用防卡死的 HTML 下載按鈕 --- 👇
            import base64
            
            # 將 Excel 檔案轉成 Base64 編碼
            b64 = base64.b64encode(output.getvalue()).decode()
            
            # 製作專屬下載連結 (target="_blank" 是防卡死的關鍵)
            download_html = f"""
            <a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" 
               download="{download_filename}" 
               target="_blank" 
               style="display: block; width: 100%; padding: 0.5rem 1rem; background-color: #FF4B4B; color: white; text-align: center; text-decoration: none; border-radius: 0.5rem; font-weight: 600; margin-top: 1rem;">
               點擊下載 Excel 底稿
            </a>
            """
            
            st.markdown(download_html, unsafe_allow_html=True)
            st.caption("提示：如果點擊後進入全螢幕預覽，請按左上角的「完成(Done)」或將手指放在螢幕【最左側邊緣】向右滑動返回！")
            # 👆 --- 替換結束 --- 👆
            
        except FileNotFoundError: st.error("找不到 `template.xlsx`。")
        except Exception as e: st.error(f"發生錯誤: {e}")

       # ==========================================
# 分頁 3: 歷史產出紀錄
# ==========================================
with tab3:
    st.markdown("###  本機歷史產出紀錄")
    
    
    output_dir = "outputs"
    if os.path.exists(output_dir):
        files = os.listdir(output_dir)
        if not files:
            st.info("目前還沒有任何產出紀錄喔！趕快去產出一張物調卡吧！")
        else:
            # 取得檔案並依照修改時間排序 (最新的在最上面)
            file_details = []
            for f in files:
                file_path = os.path.join(output_dir, f)
                file_time = os.path.getmtime(file_path)
                file_details.append((f, file_path, file_time))
            
            file_details.sort(key=lambda x: x[2], reverse=True)
            
            for f_name, f_path, f_time in file_details:
                # 判斷是不是圖片
                is_img = f_name.lower().endswith(('.jpg', '.png', '.jpeg'))
                
                col_file, col_dl, col_del = st.columns([0.6, 0.2, 0.2])
                
                with col_file:
                    icon = "" if is_img else ""
                    st.markdown(f"{icon} **{f_name}**")
                    
                with col_dl:
                    with open(f_path, "rb") as file_data:
                        st.download_button(
                            label=" 下載",
                            data=file_data.read(),
                            file_name=f_name,
                            mime="image/jpeg" if is_img else "application/octet-stream",
                            key=f"dl_{f_name}",
                            use_container_width=True
                        )
                        
                with col_del:
                    if st.button(" 刪除", key=f"del_{f_name}", use_container_width=True):
                        try:
                            os.remove(f_path) 
                            st.rerun()        
                        except Exception as e:
                            st.error(f"刪除失敗: {e}")
                            
                # 🔥 終極破解法：如果是圖片，產生一個預覽區塊讓手機可以「長按分享」
                if is_img:
                    with st.expander("點擊預覽圖片 (長按圖片即可跳出手機功能列)"):
                        st.image(f_path, use_container_width=True)
                            

                st.markdown("<hr style='margin: 0.5em 0; border-color: rgba(255,255,255,0.1);'>", unsafe_allow_html=True)









