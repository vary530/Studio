import pdfplumber
import re
from datetime import datetime
import math
import unicodedata

# --- 基礎工具函式 ---

def full_to_half(s):
    """V21 強力版：全形轉半形"""
    if not s: return ""
    return unicodedata.normalize('NFKC', s)

def m2_to_ping(m2_str):
    """平方公尺 -> 坪 (四捨五入至小數點第3位)"""
    try:
        clean_num = re.sub(r'[^\d.]', '', str(m2_str))
        if not clean_num: return 0.0
        m2 = float(clean_num)
        ping = m2 * 0.3025
        return int(ping * 1000 + 0.5) / 1000.0
    except:
        return 0.0

def format_money(amount_str):
    """將金額轉為萬單位"""
    try:
        clean = re.sub(r'[^\d]', '', amount_str)
        if not clean: return "0"
        val = int(clean)
        if val >= 10000:
            wan = val / 10000
            if wan.is_integer():
                return f"{int(wan)}萬"
            else:
                return f"{wan}萬"
        return str(val)
    except:
        return amount_str

def identify_owner_type(id_no):
    """身份識別核心邏輯"""
    id_clean = full_to_half(id_no).strip().upper()
    if not id_clean: return ""
    if re.match(r'^\d{8}$', id_clean): return "(法人)"
    if re.match(r'^[A-Z][12][0-9\*]{8}$', id_clean):
        gender_code = id_clean[1]
        if gender_code == '1': return "(本國男)"
        if gender_code == '2': return "(本國女)"
    if re.match(r'^[A-Z][89][0-9\*]{8}$', id_clean):
        gender_code = id_clean[1]
        if gender_code == '8': return "(外籍人士:男)"
        if gender_code == '9': return "(外籍人士:女)"
    if re.match(r'^[A-Z]{2}[0-9\*]{8}$', id_clean):
        return "(外籍人士)"
    return ""

def parse_chinese_number(cn_str):
    """V25 增強版：中文數字轉阿拉伯數字"""
    if not cn_str: return (False, cn_str)
    s = cn_str.replace("層", "")
    if s.isdigit(): return (True, s)

    cn_map = {
        '0': 0, '０': 0, '零': 0, '〇': 0, 'o': 0, 'O': 0,
        '1': 1, '１': 1, '一': 1,
        '2': 2, '２': 2, '二': 2,
        '3': 3, '３': 3, '三': 3,
        '4': 4, '４': 4, '四': 4,
        '5': 5, '５': 5, '五': 5,
        '6': 6, '６': 6, '六': 6,
        '7': 7, '７': 7, '七': 7,
        '8': 8, '８': 8, '八': 8,
        '9': 9, '９': 9, '九': 9,
        '十': 10, '百': 100
    }
    
    has_unit = any(c in ['十', '百'] for c in s)
    if not has_unit:
        direct_str = ""
        is_direct = True
        for char in s:
            if char in cn_map and cn_map[char] < 10:
                direct_str += str(cn_map[char])
            else:
                is_direct = False
                break
        if is_direct and direct_str:
            return (True, direct_str)

    has_num = any(c in cn_map for c in s)
    if not has_num: return (False, cn_str)

    total = 0
    current_unit = 0
    for char in s:
        if char in cn_map:
            val = cn_map[char]
            if val == 10:
                if current_unit == 0: current_unit = 1
                total += current_unit * 10
                current_unit = 0
            elif val == 100:
                if current_unit == 0: current_unit = 1
                total += current_unit * 100
                current_unit = 0
            else:
                current_unit = val
    total += current_unit
    if total == 0 and s not in ['0', '零', '０']: return (False, cn_str)
    return (True, str(total))

def format_tw_date(raw_date):
    """民國日期去零"""
    if not raw_date: return "未抓取"
    clean = re.sub(r'(?<=\D)0(?=\d)', '', raw_date)
    clean = re.sub(r'民國0+', '民國', clean)
    return clean

def calculate_age(date_str):
    """計算屋齡或持有年限"""
    try:
        match = re.search(r'(\d{2,3})年(\d{1,2})月(\d{1,2})日', date_str)
        if match:
            y, m, d = map(int, match.groups())
            ad_year = y + 1911
            target_date = datetime(ad_year, m, d)
            diff = (datetime.now() - target_date).days / 365.0
            return f"{diff:.1f}"
    except:
        pass
    return "??"

def clean_address(raw_addr):
    """地址截斷"""
    s = full_to_half(raw_addr)
    match = re.search(r'(.*(?:號|樓|室|層|邸|莊|衖))(.*)', s)
    if match:
        main_addr = match.group(1)
        suffix = match.group(2)
        if not re.match(r'^[之\-\d]', suffix):
             s = main_addr
    return s

def clean_common_name(raw_name):
    """V31 強力清洗"""
    clean = re.sub(r'^(平方公尺|共有部分資料|日共有部分資料|共有部分|分資料|建物電傳資訊|資料|電傳|號)+', '', raw_name)
    clean = clean.lstrip('號')
    stop_keywords = ['主要用途', '權利範圍', '使用執照', '建築基地', '其他登記']
    for keyword in stop_keywords:
        if keyword in clean:
            clean = clean.split(keyword)[0]
    return clean.strip()

def fix_sticky_numbers(text):
    """V25 關鍵修復：解決數字沾黏"""
    def replacer(match):
        prefix = match.group(1) # 分之
        digits = match.group(2) # 1517570.06
        for i in range(1, len(digits)-3): 
            num_part = digits[:i]
            area_part = digits[i:]
            try:
                val = float(area_part)
                if val < 200000:
                    return f"{prefix}{num_part} {area_part}"
            except:
                continue
        return match.group(0)

    return re.sub(r'(分之)(\d+\.\d+)', replacer, text)

# --- 核心解析邏輯 ---

def parse_pdf_logic(raw_text):
    result = {}
    clean_text = re.sub(r'\s+', '', raw_text)
    clean_text = fix_sticky_numbers(clean_text)
    
    parts = re.split(r'(建物所有權部|建物他項權利部)', clean_text)
    section_marking = parts[0]
    section_ownership = ""
    section_rights = ""
    
    current_header = ""
    for part in parts:
        if "建物所有權部" in part:
            current_header = "ownership"
            continue
        elif "建物他項權利部" in part:
            current_header = "rights"
            continue
        if current_header == "ownership":
            section_ownership += part
        elif current_header == "rights":
            section_rights += part
        else:
            if part != parts[0]: section_marking += part

    # 1. 標示部
    usage_match = re.search(r'主要用途(.*?)(?=主要建材)', section_marking)
    result['usage'] = usage_match.group(1) if usage_match else "未抓取"

    mat_match = re.search(r'主要建材(.*?)(?=層數)', section_marking)
    result['material'] = mat_match.group(1) if mat_match else "未抓取"

    # V33 修正：先抓取「地區」與「段名」
    district_match = re.search(r'(?:建物電傳資訊|建物標示部).*?(\w+[縣市]\w+[鄉鎮市區])', section_marking)
    district = district_match.group(1) if district_match else ""
    
    section_name = ""
    if district:
        try:
            sec_match = re.search(f'{re.escape(district)}(\D+?)(?=\d)', section_marking)
            if sec_match:
                section_name = sec_match.group(1)
        except:
            pass

    # 地址抓取與清洗
    addr_match = re.search(r'建物門牌(.*?)(?=建物坐落地號|主要用途|權利範圍|由.*?變更)', section_marking)
    street_raw = addr_match.group(1) if addr_match else ""
    cleaned_street = clean_address(street_raw)
    
    # V33 關鍵：移除地址結尾的段名
    clean_sec_name = full_to_half(section_name)
    if clean_sec_name and cleaned_street.endswith(clean_sec_name):
        cleaned_street = cleaned_street[:-len(clean_sec_name)]
        
    result['address'] = f"{district}{cleaned_street}"

    date_match = re.search(r'建築完成日期(民國\d+年\d+月\d+日)', section_marking)
    if date_match:
        d_clean = format_tw_date(date_match.group(1))
        age = calculate_age(d_clean)
        result['completion_date'] = f"{d_clean}（屋齡：{age}年）"
    else:
        result['completion_date'] = "未抓取"

    result['total_floors'] = "??"
    floors_match = re.search(r'層數(\S+?)層', section_marking)
    if floors_match: result['total_floors'] = full_to_half(floors_match.group(1)).lstrip('0')

    result['layer'] = "??"
    layer_match = re.search(r'層次(\S+?)層次面積', section_marking)
    if layer_match: 
        is_num, val = parse_chinese_number(layer_match.group(1))
        result['layer'] = val if is_num else layer_match.group(1)

    # 2. 所有權部
    owner_match = re.search(r'所有權人(.*?)(?=統一編號)', section_ownership)
    owner_name = owner_match.group(1) if owner_match else "未抓取"
    
    id_match = re.search(r'統一編號([A-Z0-9\*]+)', section_ownership)
    id_no = id_match.group(1) if id_match else ""
    owner_type = identify_owner_type(id_no)
    result['owner_display'] = f"{owner_name}{owner_type}"

    reg_date_match = re.search(r'登記日期(民國\d+年\d+月\d+日)', section_ownership)
    if reg_date_match:
        reg_d_clean = format_tw_date(reg_date_match.group(1))
        holding_years = calculate_age(reg_d_clean)
        result['acquisition_date_display'] = f"{reg_d_clean}（持有：{holding_years}年）"
    else:
        result['acquisition_date_display'] = "未抓取"

    reason_match = re.search(r'登記原因(\S+?)原因發生日期', section_ownership)
    result['reg_reason'] = reason_match.group(1) if reason_match else "未抓取"

    # 3. 他項權利部
    rights_list = []
    total_rights_value = 0 

    right_entries = re.split(r'登記次序\d{4}-\d{3}', section_rights)
    count = 0
    for entry in right_entries:
        if not entry.strip(): continue
        
        r_owner_match = re.search(r'權利人(.*?)(?=權利人統編|地址)', entry)
        if not r_owner_match: continue
        
        r_name = r_owner_match.group(1)
        r_id_match = re.search(r'權利人統編([A-Z0-9\*]+)', entry)
        r_gender = identify_owner_type(r_id_match.group(1)) if r_id_match else ""
        
        r_date_match = re.search(r'登記日期(民國\d+年\d+月\d+日)', entry)
        r_date_str = f" ({format_tw_date(r_date_match.group(1))})" if r_date_match else ""

        r_type_match = re.search(r'權利種類(.*?)(?=收件年期|字號)', entry)
        r_type = r_type_match.group(1) if r_type_match else "未知種類"
        
        r_reason_match = re.search(r'登記原因(.*?)(?=權利人)', entry)
        r_reason = r_reason_match.group(1) if r_reason_match else ""
        
        r_amount_match = re.search(r'擔保債權總金額新台幣(\d+)元', entry)
        if r_amount_match:
            amount_val = int(r_amount_match.group(1))
            total_rights_value += amount_val
            amt_formatted = format_money(r_amount_match.group(1))
            r_amount_display = f"新台幣{r_amount_match.group(1)}元（{amt_formatted}）"
        else:
            r_amount_display = "未抓取"

        count += 1
        rights_list.append(f"""
{count}.權利人：{r_name}{r_gender}{r_date_str}
   權利種類：{r_type}（登記原因設定：{r_reason}）
   擔保債權總金額：{r_amount_display}
""")

    result['rights_display'] = rights_list
    result['rights_count'] = str(count)
    
    if total_rights_value > 0:
        total_wan = format_money(str(total_rights_value))
        result['total_rights_money'] = f"新台幣{total_rights_value}元（{total_wan}）"
    else:
        result['total_rights_money'] = "無"

    # 4. 面積計算
    main_items = []
    annex_items = []
    common_items = []
    parking_items = []
    
    total_main = 0.0
    total_annex = 0.0
    total_common_gross = 0.0 
    parking_exact_ping_accumulator = 0.0

    # V26: 抓取主建物建號
    main_build_no = ""
    main_build_match = re.search(r'建物標示部.*?(\d{3,}-\d{3})建號', clean_text)
    if main_build_match:
        main_build_no = main_build_match.group(1)

    # 主建物
    main_total_match = re.search(r'總面積([\d\.]+)平方公尺', clean_text)
    if main_total_match:
        ping = m2_to_ping(main_total_match.group(1))
        total_main = ping
        result['main_total'] = f"{ping:.3f}"
        
        old_matches = re.findall(r'層次(.+?)層次面積([\d\.]+)平方公尺', section_marking)
        old_format_found = False
        if old_matches:
            old_format_found = True
            for name, m2_str in old_matches:
                if "總" in name: continue
                is_num, clean_name = parse_chinese_number(name)
                item_ping = m2_to_ping(m2_str)
                if is_num:
                    main_items.append(f"-{clean_name}層：{item_ping}坪")
                else:
                    clean_name = name.replace("層次", "").replace("面積", "")
                    main_items.append(f"-{clean_name}：{item_ping}坪")

        if not old_format_found:
            try:
                start_keywords = r'(?:層次.*?|總面積[\d\.]+平方公尺)'
                end_keywords = r'(?:附屬建物用途|共有部分資料|共有部分|建物分標示)'
                main_block_match = re.search(f'{start_keywords}(.*?){end_keywords}', clean_text)
                if main_block_match:
                    sub_items = re.findall(r'([^\d:：]+?)(?:層次)?面積[:：]?([\d\.]+)平方公尺', main_block_match.group(1))
                    for name, m2_str in sub_items:
                        if "總" in name or "合計" in name: continue
                        is_num, clean_name = parse_chinese_number(name)
                        item_ping = m2_to_ping(m2_str)
                        if is_num:
                            main_items.append(f"-{clean_name}層：{item_ping}坪")
                        else:
                            clean_name = name.replace("層次", "").replace("面積", "")
                            main_items.append(f"-{clean_name}：{item_ping}坪")
            except: pass
    else:
        result['main_total'] = "0.000"

    # 附屬建物
    annex_matches = re.findall(r'附屬建物用途(\S+?)面積([\d\.]+)平方公尺', clean_text)
    for name, m2_str in annex_matches:
        ping = m2_to_ping(m2_str)
        annex_items.append(f"-{name}：{ping}坪")
        total_annex += ping
    result['annex_total'] = f"{total_annex:.3f}"

    # 共有部分 & 車位
    common_definitions = []
    
    # V31: 強力 regex
    def_iter = re.finditer(r'([\u4e00-\u9fa5]+段[\d-]+建號).*?([\d\.]+)(?:平方公尺|平)', clean_text)
    
    for match in def_iter:
        raw_name = match.group(1)
        area_str = match.group(2)
        clean_name = clean_common_name(raw_name)
        
        # V34 修正：使用標題的「段名」來淨化共有部分名稱
        if section_name and section_name in clean_name:
            idx = clean_name.find(section_name)
            clean_name = clean_name[idx:]

        if main_build_no and main_build_no in raw_name: continue
        if "電傳" in clean_name: continue

        common_definitions.append({
            'start_index': match.start(), 
            'name': clean_name, 
            'area': float(area_str)
        })

    scope_matches = re.finditer(r'權利範圍(\d+)分之(\d+)', clean_text)
    for match in scope_matches:
        den = float(match.group(1))
        num = float(match.group(2))
        scope_start = match.start()
        
        valid_defs = [d for d in common_definitions if d['start_index'] <= scope_start]
        if not valid_defs: continue
        
        target_def = valid_defs[-1]
        current_m2 = target_def['area']
        
        raw_ping = (current_m2 * (num / den)) * 0.3025
        display_ping = int(raw_ping * 1000 + 0.5) / 1000.0
        
        raw_context = clean_text[max(0, scope_start-40):scope_start]
        context = full_to_half(raw_context)
        
        is_parking = False
        p_num = ""
        
        # V32/V34 車位判斷
        if "含停車位" in context or "編號" in context:
            p_match_full = re.findall(r'編號\s*([^\s權]+?)(?:號)?(?:權利範圍)', context)
            if not p_match_full:
                p_match_full = re.findall(r'編號\s*([0-9A-Za-z\u4e00-\u9fa5]+)(?:號)?', context)

            if p_match_full:
                is_parking = True
                p_num = p_match_full[-1].rstrip('號') # V34 修正

        if is_parking:
            parking_items.append(f"-編號{p_num}號：{display_ping:.3f}坪")
            parking_exact_ping_accumulator += raw_ping
        elif "所有權" in context or "建築基地" in context:
            pass 
        else:
            if "共有部分" in context or "建號" in context:
                common_items.append(f"-共有部分({target_def['name']})：{display_ping:.3f}坪")
                total_common_gross += display_ping

    # === 計算邏輯 ===
    
    final_parking_total = int(parking_exact_ping_accumulator * 1000 + 0.5) / 1000.0
    
    gross_common = total_common_gross
    net_common = max(0.0, gross_common - final_parking_total)
    reg_total = total_main + total_annex + gross_common
    net_reg_total = max(0.0, reg_total - final_parking_total)

    result['reg_total'] = f"{reg_total:.3f}"
    result['main_items'] = main_items
    result['annex_items'] = annex_items
    
    result['net_common_str'] = f"{net_common:.3f}"
    result['gross_common_str'] = f"{gross_common:.3f}"
    result['common_items'] = common_items
    
    result['parking_total'] = f"{final_parking_total:.3f}"
    result['parking_items'] = parking_items
    
    result['net_reg_total'] = f"{net_reg_total:.3f}"

    return result, clean_text

def extract_text_from_pdf(uploaded_file):
    try:
        full_text = ""
        with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text: full_text += text
        
        data, raw_clean = parse_pdf_logic(full_text)
        return data, raw_clean 
    except Exception as e:
        return {"error": str(e)}, ""