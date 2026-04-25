# -*- coding: utf-8 -*-
"""
爱准（AIMT）+ 上海质检院（SQI）证书提取脚本
用于从爱准计量检测 / 上海质检院格式的校准证书 PDF 中提取证书信息

用法:
  python scripts/extract_aimt_sqi.py <pdf_path> [output.json]

依赖:
  pip install pdfplumber paddleocr paddlepaddle opencv-python pdf2image numpy pillow python-dateutil
"""
import sys
import os
sys.stdout.reconfigure(encoding='utf-8')
import re
import json
from pdf2image import convert_from_path
from paddleocr import PaddleOCR
from PIL import Image, ImageFile
ImageFile.LOAD_TRUNCATED_IMAGES = True
import numpy as np
from datetime import date, timedelta

try:
    from dateutil.relativedelta import relativedelta
    HAS_DATEUTIL = True
except ImportError:
    HAS_DATEUTIL = False


def ocr_page(pdf_path, page_num, dpi=200, ocr_engine=None):
    """OCR 指定页面，返回文字行列表"""
    pages = convert_from_path(pdf_path, dpi=dpi, first_page=page_num, last_page=page_num)
    if not pages:
        return []
    img = np.array(pages[0])
    result = ocr_engine.ocr(img, cls=True)
    lines = []
    if result and result[0]:
        for r in result[0]:
            text = r[1][0].strip()
            if text:
                lines.append(text)
    return lines


def add_12months_minus1(date_str):
    """计算 +12个月-1天 的有效期"""
    try:
        parts = date_str.split('.')
        d = date(int(parts[0]), int(parts[1]), int(parts[2]))
        if HAS_DATEUTIL:
            d2 = d + relativedelta(months=12) - timedelta(days=1)
        else:
            y, m = d.year + 1, d.month
            try:
                d2 = date(y, m, d.day) - timedelta(days=1)
            except ValueError:
                import calendar
                last_day = calendar.monthrange(y, m)[1]
                d2 = date(y, m, min(d.day, last_day)) - timedelta(days=1)
        return f"{d2.year}.{str(d2.month).zfill(2)}.{str(d2.day).zfill(2)}"
    except:
        return ""


KNOWN_BRANDS = {"JINTENG", "JUMO", "WIKA", "ZY", "AZ", "YOKOGAWA",
                "EMERSON", "SIEMENS", "ABB", "ROSEMOUNT",
                "Shanghai ZhenTai", "上海振太"}

BAD_WORDS_LOWER = {
    "approved", "checked", "calibrated", "certificate", "manufacturer",
    "address", "statement", "year", "month", "day", "page", "cnas",
    "room", "building", "jianyun", "pudong", "tofflon", "aizm",
    "上海", "东富龙", "爱准", "邦盟", "批准", "核验", "校准员",
    "地址", "服务", "电话", "传真", "投诉", "本证书", "未经",
    "部分采用", "再校", "sample", "contact", "client", "委托",
    "canasration", "lacmcnas", "iacmra", "jlac", "cnasl", "calibration"
}


def is_bad(text):
    t = text.lower().strip()
    for bw in BAD_WORDS_LOWER:
        if bw in t:
            return True
    return False


def is_aimt_cert(lines):
    """是爱准证书封面（AIMT格式）
    
    条件：含有 AIMT证书编号 + 含有仪器名称 + 含有机构名（上海爱准/Shanghai AIMT）
    不强制要求"第1页共N页"，因为部分封面OCR漏掉了该标识
    """
    joined = ' '.join(lines)
    has_aimt_cert = bool(re.search(r'AIMT\d{4}-[A-Z]-\d+', joined.replace(' ', '')))
    if not has_aimt_cert:
        return False
    # 必须有仪器名称才算封面（排除后续数据页）
    name_kws = ['Pressure gauge', 'Pressure transmitter', 'Temperature Probe',
                '压力表', '压力变送器', '温度探头']
    has_name = any(kw.lower() in joined.lower() for kw in name_kws)
    if not has_name:
        return False
    # 必须有机构名（爱准封面固定出现，数据页没有）
    has_org = '上海爱准' in joined or 'Shanghai AIMT' in joined or 'Shanghai AlMT' in joined
    if not has_org:
        return False
    # 数据页（第2/3页）通常有"Page 2 of"/"第2页共"字样，排除
    is_data_page = bool(re.search(r'Page [23] of|第[23]页共\d+页', joined))
    return not is_data_page


def is_sqi_cert(lines):
    """是上海质检院封面
    
    条件：含 SQI 机构名 + 有器具名称 + 有证书编号 + 非续页
    """
    joined = ' '.join(lines)
    has_sqi = '上海市质量监督检验技术研究院' in joined or 'Shanghai Institute of Quality' in joined
    if not has_sqi:
        return False
    # 必须是封面（有器具名称）：电容薄膜真空计
    has_name = '电容薄膜真空计' in joined or 'Capacitance Diaphragm Vacuum' in joined
    # 同时必须有证书编号行
    has_cert = bool(re.search(r'[A-Z]\d{5}[A-Z]\d{5}', joined.replace(' ', '')))
    # 防止第2/3页也被识别：第2/3页有 SQI/JL 标志或"续页"字样
    is_continuation = 'SQI/JL' in joined or '证书续页' in joined or 'Continued page' in joined
    return has_name and has_cert and not is_continuation


def extract_aimt_date(lines):
    """
    爱准证书日期提取
    版面有两个日期：
      - `2026年03月30日`（样品接收/处理日期）
      - `年04月02日` 或 `年04月03日`（校准完成日期）- 这才是 cal_date
    规律：校准完成日通常是月份较大或相同但日期靠后的那个
    """
    # 找所有完整日期行 YYYY年MM月DD日
    full_dates = []
    for l in lines:
        m = re.match(r'(\d{4})年(\d{1,2})月(\d{1,2})日', l.strip())
        if m:
            full_dates.append((m.group(1), m.group(2).zfill(2), m.group(3).zfill(2)))

    # 找 `年MM月DD日`（年份在前一行）
    partial_dates = []
    for i, l in enumerate(lines):
        m = re.match(r'^年(\d{1,2})月(\d{1,2})日$', l.strip())
        if m:
            # 找附近的年份
            for j in range(max(0, i-5), i):
                yr = re.match(r'^(\d{4})$', lines[j].strip())
                if yr:
                    partial_dates.append((yr.group(1), m.group(1).zfill(2), m.group(2).zfill(2)))
                    break
        # 也处理 `04月02日` 格式（没有年字）
        m2 = re.match(r'^(\d{1,2})月(\d{1,2})日$', l.strip())
        if m2:
            for j in range(max(0, i-5), i):
                yr = re.match(r'^(\d{4})$', lines[j].strip())
                if yr:
                    partial_dates.append((yr.group(1), m2.group(1).zfill(2), m2.group(2).zfill(2)))
                    break

    all_dates = full_dates + partial_dates

    if not all_dates:
        return ""

    # 取月份最大的那个（校准完成日），如果月份相同则取日期最大的
    cal_tuple = max(all_dates, key=lambda x: (int(x[1]), int(x[2])))
    return f"{cal_tuple[0]}.{cal_tuple[1]}.{cal_tuple[2]}"


def extract_aimt_cert_data(lines):
    """从爱准封面行列表提取所有字段"""
    cert = {}

    # 1. 证书编号
    joined = ' '.join(lines).replace(' ', '')
    m = re.search(r'AIMT\d{4}-[A-Z]-\d+', joined)
    cert['cert'] = m.group(0) if m else ""

    # 2. 仪器名称
    name_en, name_cn = "", ""
    en_kw = [("Temperature Probe (Double)", "温度探头（双支）"),
             ("Temperature Probe (Single)", "温度探头（单支）"),
             ("Pressure transmitter", "压力变送器"),
             ("Pressure gauge", "压力表")]
    for en, cn in en_kw:
        for l in lines:
            if en.lower() in l.lower() and len(l) < 50:
                name_en = en
                name_cn = cn
                break
        if name_en:
            break
    # 中文备用
    if not name_cn:
        for l in lines:
            if '压力表' in l and len(l) < 10: 
                name_cn = '压力表'; name_en = 'Pressure gauge'; break
            if '压力变送器' in l and len(l) < 10: 
                name_cn = '压力变送器'; name_en = 'Pressure transmitter'; break
            if '温度探头（双支）' in l and len(l) < 12: 
                name_cn = '温度探头（双支）'; name_en = 'Temperature Probe (Double)'; break
            if '温度探头（单支）' in l and len(l) < 12: 
                name_cn = '温度探头（单支）'; name_en = 'Temperature Probe (Single)'; break
    cert['name_en'] = name_en
    cert['name_cn'] = name_cn

    # 3. 管理编号（SN）- 优先项目管理编号格式
    sn = ""
    for l in lines:
        m2 = re.match(r'^(\d{4}-\d+[A-Za-z]+/\S+)$', l.strip())
        if m2:
            sn = m2.group(1)
            break
    # 备用：WIKA风格序列号（字母数字混合）
    if not sn:
        for l in lines:
            if is_bad(l): continue
            m3 = re.match(r'^[A-Z0-9]{8,15}$', l.strip())
            if m3:
                val = l.strip()
                if not re.match(r'^\d{4}$', val) and not re.match(r'^\d{5,}$', val):
                    sn = val
                    break
    # 再备用：Q开头序列号
    if not sn:
        for l in lines:
            m4 = re.match(r'^Q\d{7}$', l.strip())
            if m4:
                sn = l.strip()
                break
    cert['sn'] = sn

    # 4. 品牌
    brand = ""
    for l in lines:
        if l.strip() in KNOWN_BRANDS:
            brand = l.strip()
            break
    cert['brand_en'] = brand
    cert['brand_cn'] = ""

    # 5. 型号
    model = ""
    model_candidates = []
    for l in lines:
        c = l.strip()
        if not c or is_bad(c): continue
        # 规格范围格式：(-0.1~3.8)MPa、S-20/(-0.1~0.3)MPa
        if re.search(r'[（(][-\d\.~]+[)）](?:MPa|kPa|bar)', c) and len(c) < 60:
            model_candidates.append(c)
        # PT100/3Wire
        elif re.match(r'^PT100/\dWire$', c):
            model_candidates.append(c)
        # S-XX/型号
        elif re.search(r'^S-\d+/[（(]', c):
            model_candidates.append(c)
    if model_candidates:
        for mc in model_candidates:
            if re.search(r'^S-\d+/', mc):
                model = mc; break
        if not model:
            model = model_candidates[0]
    if not model:
        for i, l in enumerate(lines):
            if 'Model/Specification' in l or ('型号' in l and len(l) < 15):
                for delta in range(1, 5):
                    j = i + delta
                    if j >= len(lines): break
                    c = lines[j].strip()
                    if not c or is_bad(c): continue
                    if re.search(r'MPa|bar|Wire|Torr|PT100|RTD|[（(][-\d\.~]', c, re.I):
                        model = c; break
                break
    cert['model'] = model

    # 6. 日期
    cal_date = extract_aimt_date(lines)
    cert['cal_date'] = cal_date
    cert['due'] = add_12months_minus1(cal_date) if cal_date else ""
    cert['pid'] = ""

    return cert


def extract_sqi_cert_data(lines, extra_lines=None):
    """从上海质检院封面提取字段（extra_lines为后续页面内容，用于获取日期和有效期）"""
    cert = {}

    # 证书编号：J26439S00738 格式
    for l in lines:
        m = re.search(r'[A-Z]\d{5}[A-Z]\d{5}', l.replace(' ', ''))
        if m:
            cert['cert'] = m.group(0)
            break
    if 'cert' not in cert:
        cert['cert'] = ""

    # 名称
    cert['name_cn'] = '电容薄膜真空计'
    cert['name_en'] = 'Capacitance Diaphragm Vacuum Gauge'

    # 序列号
    sn = ""
    sn_idx = -1
    for i, l in enumerate(lines):
        if 'SerialNo.' in l.replace(' ', '') or 'Serial No.' in l or '出厂编号' in l:
            sn_idx = i; break
    if sn_idx >= 0:
        for delta in range(1, 6):
            j = sn_idx + delta
            if j >= len(lines): break
            c = lines[j].strip()
            if c and re.match(r'^[A-Z]\d{4}[A-Z]\d{3}[A-Z]$', c):
                sn = c; break
    if not sn:
        for l in lines:
            m2 = re.match(r'^[A-Z]\d{4}[A-Z]\d{3}[A-Z]$', l.strip())
            if m2:
                sn = l.strip(); break
    cert['sn'] = sn

    # 品牌
    cert['brand_en'] = "Shanghai ZhenTai"
    cert['brand_cn'] = "上海振太仪表"

    # 型号
    model = ""
    type_idx = -1
    for i, l in enumerate(lines):
        if 'Type/Specification' in l or '型号/规格' in l:
            type_idx = i; break
    if type_idx >= 0:
        for delta in range(1, 5):
            j = type_idx + delta
            if j >= len(lines): break
            c = lines[j].strip()
            if not c or is_bad(c): continue
            if re.search(r'CPDA|Torr|RTD|PT100|[A-Z]{2,}\d+', c) and len(c) < 40:
                model = c; break
    if not model:
        for l in lines:
            if re.search(r'CPDA|Torr', l) and not is_bad(l):
                model = l.strip(); break
    cert['model'] = model

    # 日期：优先从附加页面找"建议于YYYY年M月D日前复校"反推
    cal_date = ""
    # 先在封面找完整日期
    for l in lines:
        m3 = re.match(r'(\d{4})年(\d{1,2})月(\d{1,2})日', l.strip())
        if m3:
            cal_date = f"{m3.group(1)}.{m3.group(2).zfill(2)}.{m3.group(3).zfill(2)}"
            break
    # 如果封面没有完整日期，从附加页面找
    if not cal_date and extra_lines:
        year = ""
        for l in lines:
            m4 = re.match(r'^(\d{4})年?$', l.strip())
            if m4 and int(m4.group(1)) > 2000:
                year = m4.group(1); break
        # 从extra_lines中找"建议于YYYY年M月D日前复校"来反推
        if year or True:  # 直接尝试找建议复校日期
            for l in extra_lines:
                m7 = re.search(r'建议于(\d{4})年(\d{1,2})月(\d{1,2})日前', l)
                if m7:
                    due_y = int(m7.group(1))
                    due_m = int(m7.group(2))
                    due_d = int(m7.group(3))
                    due_date_obj = date(due_y, due_m, due_d)
                    if HAS_DATEUTIL:
                        cal_date_obj = due_date_obj - relativedelta(months=12) + timedelta(days=1)
                    else:
                        cal_date_obj = date(due_y - 1, due_m, due_d) + timedelta(days=1)
                    cal_date = f"{cal_date_obj.year}.{str(cal_date_obj.month).zfill(2)}.{str(cal_date_obj.day).zfill(2)}"
                    break

    # 有效期：直接取"建议于...前复校"的日期
    due_date = ""
    if extra_lines:
        for l in extra_lines:
            m7 = re.search(r'建议于(\d{4})年(\d{1,2})月(\d{1,2})日前', l)
            if m7:
                due_date = f"{m7.group(1)}.{m7.group(2).zfill(2)}.{m7.group(3).zfill(2)}"
                break
    if not due_date and cal_date:
        due_date = add_12months_minus1(cal_date)

    cert['cal_date'] = cal_date
    cert['due'] = due_date
    cert['pid'] = ""
    return cert


def main():
    if len(sys.argv) < 2:
        print("用法: python extract_aimt_sqi.py <pdf_path> [output.json]")
        print("示例: python extract_aimt_sqi.py certificates.pdf extracted_data.json")
        sys.exit(1)

    pdf_path = sys.argv[1]
    
    # 输出路径：默认与PDF同目录下的 extracted_data.json
    if len(sys.argv) >= 3:
        output_json = sys.argv[2]
    else:
        base = os.path.splitext(os.path.basename(pdf_path))[0]
        output_json = os.path.join(os.path.dirname(pdf_path) or '.', f"{base}_extracted.json")

    # 初始化 OCR 引擎
    print("初始化 PaddleOCR...")
    ocr_engine = PaddleOCR(lang='ch', use_angle_cls=True, show_log=False)

    # 获取总页数
    import pdfplumber
    with pdfplumber.open(pdf_path) as pdf:
        total_pages = len(pdf.pages)
    print(f"PDF: {pdf_path}")
    print(f"总页数: {total_pages}")

    # OCR 所有页面
    all_lines = {}
    for pg in range(1, total_pages + 1):
        print(f"OCR 第 {pg}/{total_pages} 页...", flush=True)
        all_lines[pg] = ocr_page(pdf_path, pg, dpi=200, ocr_engine=ocr_engine)

    print("\n开始提取证书...")
    certs = []
    cert_no = 0
    i = 1
    while i <= total_pages:
        lines = all_lines[i]

        if is_sqi_cert(lines):
            print(f"\n第 {i} 页 → 上海质检院证书封面")
            extra = []
            for j in range(i+1, min(i+4, total_pages+1)):
                extra.extend(all_lines[j])
            cert_data = extract_sqi_cert_data(lines, extra)
            cert_no += 1
            cert_data['no'] = str(cert_no)
            certs.append(cert_data)

        elif is_aimt_cert(lines):
            print(f"\n第 {i} 页 → 爱准证书封面")
            cert_data = extract_aimt_cert_data(lines)
            cert_no += 1
            cert_data['no'] = str(cert_no)
            certs.append(cert_data)
        else:
            print(f"第 {i} 页：跳过")
            i += 1
            continue

        # 打印提取结果
        print(f"  证书编号: {cert_data.get('cert', '?')}")
        print(f"  名称: {cert_data.get('name_en', '')} / {cert_data.get('name_cn', '')}")
        print(f"  SN: {cert_data.get('sn', '')}")
        print(f"  品牌: {cert_data.get('brand_en', '')}")
        print(f"  型号: {cert_data.get('model', '')}")
        print(f"  校准日期: {cert_data.get('cal_date', '')}")
        print(f"  有效期: {cert_data.get('due', '')}")
        i += 1

    # 输出 JSON
    with open(output_json, 'w', encoding='utf-8') as f:
        json.dump(certs, f, ensure_ascii=False, indent=2)

    print(f"\n\n共提取 {len(certs)} 份证书，已写入 {output_json}")
    print("\n===== 汇总 =====")
    print(f"{'序号':<4} {'证书编号':<25} {'名称':<20} {'SN':<22} {'品牌':<15} {'校准日期':<12} {'有效期'}")
    for c in certs:
        print(f"{c['no']:<4} {c.get('cert',''):<25} {c.get('name_cn',''):<20} "
              f"{c.get('sn',''):<22} {c.get('brand_en',''):<15} "
              f"{c.get('cal_date',''):<12} {c.get('due','')}")


if __name__ == '__main__':
    main()
