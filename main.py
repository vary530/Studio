import streamlit as st
import base64
from services.pdf_scanner import extract_text_from_pdf

st.set_page_config(page_title="建物謄本掃描", layout="centered")

def set_bg_hack(main_bg):
    main_bg_ext = "png"
    st.markdown(
         f"""
         <style>
         /* === 背景模糊特效層 === */
         .stApp {{
             background: none;
         }}
         .stApp::before {{
             content: "";
             position: fixed;
             top: 0; 
             left: 0; 
             width: 100vw; 
             height: 100vh;
             background: url(data:image/{main_bg_ext};base64,{base64.b64encode(open(main_bg, "rb").read()).decode()});
             background-size: cover; 
             background-position: center; 
             background-repeat: no-repeat; 
             background-attachment: fixed;
             filter: blur(10px);
             z-index: -1;
             transform: scale(1.05);
         }}

         /* === 強制內容層浮出 === */
         .main .block-container {{
             z-index: 1;
             position: relative;
         }}

         /* === 資訊卡樣式 (毛玻璃) === */
         .glass-container {{
             background-color: rgba(0, 0, 0, 0.6); 
             padding: 25px; 
             border-radius: 16px; 
             margin-bottom: 20px;
             border: 1px solid rgba(255, 255, 255, 0.15); 
             box-shadow: 0 8px 32px 0 rgba(0, 0, 0, 0.37);
             backdrop-filter: blur(4px); 
             -webkit-backdrop-filter: blur(4px); 
             color: #ffffff;
         }}

         /* 字體樣式 */
         .text-label {{ color: #cccccc; font-size: 16px; font-weight: 500; }}
         .text-value {{ color: #ffffff; font-size: 16px; font-weight: 400; margin-left: 5px; }}
         .text-highlight {{ color: #FFD700; font-size: 24px; font-weight: 700; }}
         .text-medium {{ color: #ffffff; font-size: 18px; font-weight: 600; }}
         .text-detail {{ color: #aaaaaa; font-size: 14px; margin-left: 15px; line-height: 1.6; }}

         /* === 上傳區塊美化 === */
         div[data-testid="stFileUploader"] section {{
             background-color: rgba(0, 0, 0, 0.5);
             border: none;
             border-radius: 16px;
             padding: 20px;
             box-shadow: 0 4px 12px rgba(0,0,0,0.3);
         }}
         
         div[data-testid="stFileUploader"] section > div > div > span {{
             color: #ddd !important;
             font-size: 16px !important;
         }}
         
         div[data-testid="stFileUploader"] button {{
             background-color: transparent !important;
             border: 1px solid rgba(255, 255, 255, 0.8) !important;
             border-radius: 20px !important;
             padding: 5px 20px !important;
             color: transparent !important;
             position: relative;
             transition: all 0.3s ease;
         }}
         
         div[data-testid="stFileUploader"] button::after {{
             content: "上傳檔案";
             color: white;
             font-size: 14px;
             position: absolute;
             left: 50%;
             top: 50%;
             transform: translate(-50%, -50%);
             font-weight: 500;
         }}

         div[data-testid="stFileUploader"] button:hover {{
             background-color: rgba(255, 255, 255, 0.1) !important;
             border-color: #ffffff !important;
             box-shadow: 0 0 8px rgba(255, 255, 255, 0.3);
         }}

         /* === 折疊選單美化 === */
         .stExpander {{ background-color: transparent; border: none; }}
         .streamlit-expanderHeader {{ 
             background-color: rgba(0, 0, 0, 0.6) !important; 
             border: 1px solid rgba(255, 255, 255, 0.1) !important;
             border-radius: 10px; 
             color: white !important; 
             font-weight: 600; 
         }}
         div[data-testid="stExpanderDetails"] {{ 
             background-color: rgba(0, 0, 0, 0.5); 
             border: 1px solid rgba(255, 255, 255, 0.1);
             border-top: none;
             border-radius: 0 0 10px 10px;
             color: #dddddd; 
             padding: 15px;
             margin-top: -5px;
         }}
         </style>
         """,
         unsafe_allow_html=True
     )

try:
    set_bg_hack('background.png') 
except:
    st.markdown("""<style>.stApp { background-color: #2c3e50; }</style>""", unsafe_allow_html=True)

st.title("建物謄本掃描")

menu = st.sidebar.selectbox("選擇功能", ["自動掃描 PDF", "製作物調表 (開發中)"])

def render_row(label, value):
    return f'<div><span class="text-label">{label}</span><span class="text-value">{value}</span></div>'

def render_detail(text):
    return f'<div class="text-detail">{text}</div>'

if menu == "自動掃描 PDF":
    uploaded_file = st.file_uploader("請上傳建物謄本 PDF", type="pdf")
    
    if uploaded_file is not None:
        if st.button("開始掃描"):
            with st.spinner("掃描中..."):
                data, raw_text = extract_text_from_pdf(uploaded_file)
                
                if "error" in data:
                    st.error(f"發生錯誤: {data['error']}")
                else:
                    # 1. 基本資料
                    basic_parts = []
                    basic_parts.append('<div class="glass-container">')
                    basic_parts.append(render_row("所有權人：", data.get('owner_display')))
                    basic_parts.append(render_row("取得日期：", data.get('acquisition_date_display')))
                    basic_parts.append(render_row("登記原因：", data.get('reg_reason')))
                    basic_parts.append(render_row("主要用途：", data.get('usage')))
                    basic_parts.append(render_row("主要建材：", data.get('material')))
                    basic_parts.append(render_row("建物門牌：", data.get('address')))
                    basic_parts.append(render_row("建築完成日：", data.get('completion_date')))
                    basic_parts.append(render_row("總樓層：", data.get('total_floors')))
                    basic_parts.append(render_row("位於樓層：", data.get('layer')))
                    basic_parts.append('</div>')
                    st.markdown("".join(basic_parts), unsafe_allow_html=True)

                    # 2. 坪數細節
                    area_parts = []
                    area_parts.append('<div class="glass-container">')
                    
                    area_parts.append('<div style="margin-bottom: 15px;">')
                    area_parts.append('<span class="text-label">登記總坪數：</span>')
                    area_parts.append(f'<span class="text-highlight">{data.get("reg_total")}坪</span>')
                    area_parts.append(f'<span class="text-value" style="font-size: 14px;">(車位{data.get("parking_total")}坪)</span>')
                    area_parts.append('</div>')
                    
                    area_parts.append(f'<div style="margin-top: 15px;"><div class="text-medium">主建物：{data.get("main_total")}坪</div></div>')
                    for item in data.get('main_items', []):
                        area_parts.append(render_detail(item))
                        
                    area_parts.append(f'<div style="margin-top: 15px;"><div class="text-medium">附屬建物：{data.get("annex_total")}坪</div></div>')
                    for item in data.get('annex_items', []):
                        area_parts.append(render_detail(item))
                        
                    area_parts.append(f'<div style="margin-top: 15px;"><div class="text-medium">公設坪數：{data.get("net_common_str")}坪</div></div>')
                    area_parts.append(render_detail(f'-共有總坪數：{data.get("gross_common_str")}坪(含車位:{data.get("parking_total")}坪)'))
                    for item in data.get('common_items', []):
                        area_parts.append(render_detail(item))
                        
                    area_parts.append(f'<div style="margin-top: 15px;"><div class="text-medium">車位坪數：{data.get("parking_total")}坪</div></div>')
                    p_items = data.get('parking_items', [])
                    if p_items:
                        for item in p_items:
                            area_parts.append(render_detail(item))
                    else:
                        area_parts.append(render_detail('(無車位資料)'))
                    
                    area_parts.append('<div style="margin-top: 25px; border-top: 1px solid rgba(255,255,255,0.3); padding-top: 15px;">')
                    area_parts.append(f'<span class="text-medium">不含車位總坪數：{data.get("net_reg_total")}坪</span>')
                    area_parts.append('</div>')
                    
                    area_parts.append('</div>')
                    st.markdown("".join(area_parts), unsafe_allow_html=True)

                    # 3. 他項權利
                    rights_count = int(data.get('rights_count', 0))
                    
                    st.markdown(f"""
                    <div class="glass-container" style="padding: 15px; margin-bottom: 5px;">
                        <div class="text-medium">相關他項權利登記次序：共{rights_count}筆</div>
                    </div>""", unsafe_allow_html=True)

                    if rights_count > 0:
                        with st.expander("點擊查看他項權利明細"):
                            rights_content = data.get('rights_display', [])
                            for r in rights_content:
                                r_clean = r.strip().replace("\n", "<br>")
                                st.markdown(f"""
                                <div style="border-bottom: 1px solid rgba(255,255,255,0.1); padding: 10px 0; color: #ddd; line-height: 1.6;">
                                    {r_clean}
                                </div>""", unsafe_allow_html=True)
                            
                            st.markdown(f"""
                            <div style="margin-top: 15px; text-align: right; padding-top: 10px; border-top: 1px solid rgba(255,255,255,0.3);">
                                <span class="text-highlight" style="font-size: 20px;">總擔保債權總金額：{data.get('total_rights_money')}</span>
                            </div>""", unsafe_allow_html=True)
                    
                    # === 除錯模式 ===
                    st.markdown("<br>", unsafe_allow_html=True)
                    with st.expander("🛠️ 除錯模式：查看原始文字 (Debug)"):
                        st.text_area("PDF 原始提取內容", raw_text, height=300)

elif menu == "製作物調表 (開發中)":
    st.info("此功能尚未開發。")