# ==========================================
    # 🌟 二合一終極按鈕：一鍵產出 Excel 與 JPG
    # ==========================================
    if st.button("🚀 輸出物調表", type="primary", use_container_width=True):
        with st.spinner("努力產生物調表與圖片中..."):
            try:
                # --- 1. 共用的前置資料準備 ---
                price = try_float(input_vars.get("售價", 0))
                area = try_float(input_vars.get("不含車位坪數", 0))
                public = try_float(input_vars.get("公設坪數", 0))

                # 計算單價與公設比
                if "房屋單價" in parsed_data and parsed_data["房屋單價"]['value'] == '0' and area > 0:
                    parsed_data["房屋單價"]['value'] = f"{(price/area):.2f}"
                if "公設比" in parsed_data and parsed_data["公設比"]['value'] == '0' and area > 0:
                    parsed_data["公設比"]['value'] = f"{(public/area)*100:.1f}"

                # 特殊欄位轉大寫
                for k in ["委託契約書編號", "車位樓層", "機車位樓層"]:
                    if k in parsed_data and parsed_data[k]['value']:
                        parsed_data[k]['value'] = str(parsed_data[k]['value']).upper()

                # 準備檔名
                contract_id = st.session_state.get('委託契約書編號', "")
                if not contract_id: contract_id = input_vars.get("委託契約書編號", "")
                case_name = st.session_state.get('案名', "")
                if not case_name: case_name = input_vars.get("案名", "")

                contract_id = str(contract_id).strip().upper()
                case_name = str(case_name).strip()
                safe_id = re.sub(r'[\\/*?:"<>|]', '', contract_id)
                safe_name = re.sub(r'[\\/*?:"<>|]', '', case_name)
                
                download_filename = f"{safe_id}{safe_name}.xlsx" if safe_id or safe_name else "物調表_完成.xlsx"
                safe_img_name = safe_name if safe_name else "未命名"

                # 建立基礎 final_data
                final_data = {**parsed_data, **options_data}

                # --- 2. 產生 JPG 圖片 ---
                from services.export_image import generate_jpg_from_template
                layout_str = f"字體 {font_size} 級 | 每行 {max_char_width} 字 | 上限 {max_line} 行"
                final_data_jpg = final_data.copy()
                final_data_jpg['自訂排版'] = {'value': layout_str, 'type': 'text', 'suffix': ''}
                
                blank_image_path = "blank_template.jpg" 
                img_bytes = generate_jpg_from_template(blank_image_path, final_data_jpg, map_image_file, desc_input)
                
                # 備份 JPG 到 outputs
                if not os.path.exists("outputs"): os.makedirs("outputs")
                with open(f"outputs/{safe_img_name}_物調卡.jpg", "wb") as f:
                    f.write(img_bytes.getvalue())

                # --- 3. 產生 Excel ---
                template_path = "template.xlsx"
                wb = openpyxl.load_workbook(template_path)
                ws = wb.active 
                
                img_to_insert = None
                if map_image_file:
                    img_buffer, w, h = crop_and_resize_image(map_image_file)
                    if img_buffer:
                        img_to_insert = XLImage(img_buffer)
                        img_to_insert.width, img_to_insert.height = w, h
                
                final_data_excel = final_data.copy()
                final_data_excel['"""物件特色描述"""'] = {'type': 'area_text_advanced', 'value': desc_input, 'max_char': max_char_width, 'font_size': font_size}
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
                                for k, content in final_data_excel.items():
                                    ph = f'"""{k}"""' if '*' not in k and '"""' not in k else k
                                    if ph in original_text:
                                        val = content['value']
                                        c_type = content.get('type', 'text')
                                        if val:
                                            # V60: 字體自動縮小邏輯
                                            v_len = sum(get_visual_width(c) for c in str(val))
                                            curr_font_name = cell.font.name if cell.font else 'KaiTi'
                                            if k == "案名" and v_len >= 15: cell.font = Font(name=curr_font_name, size=14)
                                            elif k == "地址" and v_len >= 15: cell.font = Font(name=curr_font_name, size=12)
                                            elif k in ["學校", "市場", "公園"] and v_len >= 15: cell.font = Font(name=curr_font_name, size=10)
                                            
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
                
                # 備份 Excel 到 outputs
                with open(f"outputs/{download_filename}", "wb") as f:
                    f.write(output.getvalue())

                # --- 4. 畫面完美輸出 (合併顯示區) ---
                st.success("🎉 物調表 Excel 與圖卡皆產生成功！")
                
                # (A) 防卡死的 Excel HTML 下載按鈕
                import base64
                b64 = base64.b64encode(output.getvalue()).decode()
                download_html = f"""
                <a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" 
                   download="{download_filename}" 
                   target="_blank" 
                   style="display: block; width: 100%; padding: 0.8rem 1rem; background-color: #FF4B4B; color: white; text-align: center; text-decoration: none; border-radius: 0.5rem; font-weight: bold; margin-bottom: 1rem; font-size: 16px;">
                   📥 點擊下載 Excel 底稿
                </a>
                """
                st.markdown(download_html, unsafe_allow_html=True)
                st.caption("💡 提示：若點擊後全螢幕卡住，請按左上角完成，或將手指放在螢幕【最左側邊緣】向右滑動返回！")
                
                st.markdown("<hr style='border-color: rgba(255,255,255,0.2); margin: 1.5rem 0;'>", unsafe_allow_html=True)
                
                # (B) 圖片直接預覽與長按提示
                st.info("📱 手機版用戶：請直接「長按下方圖片」 ➜ 選擇「儲存到照片」或「分享」即可！")
                st.image(img_bytes, use_container_width=True)
                
            except Exception as e:
                st.error(f"發生錯誤: {e}")

    # 加入隱形墊子，防止按鈕黏在螢幕最底部 (解決分頁2按鈕黏底問題)
    st.markdown("<div style='height: 3rem;'></div>", unsafe_allow_html=True)
