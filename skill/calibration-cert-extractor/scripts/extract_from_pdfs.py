#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import os
import re
import json
import sys
import datetime
import logging

# 关闭PaddleOCR冗余日志
logging.getLogger("ppocr").setLevel(logging.ERROR)

# 依赖兼容处理
try:
    import pdfplumber
except ImportError:
    print("⚠️  pdfplumber 未安装，WorkBuddy 会自动安装")

ocr = None
OCR_READY = False

# 初始化PaddleOCR（本地、离线）
try:
    from paddleocr import PaddleOCR
    from pdf2image import convert_from_path
    import cv2
    import numpy as np

    ocr = PaddleOCR(
        lang="ch",          # 中英文混合识别
        use_angle_cls=True, # 自动纠正倾斜文本
        show_log=False      # 关闭日志输出
    )
    OCR_READY = True
except ImportError:
    print("⚠️  PaddleOCR相关依赖未安装，扫描版PDF无法识别")


def pdf_ocr_text(pdf_path, page_index):
    """使用本地PaddleOCR识别PDF单页文本（仅扫描版PDF调用）"""
    if not OCR_READY:
        return ""
    try:
        # 将指定页转为图片（dpi=200，兼顾速度和识别率）
        images = convert_from_path(
            pdf_path,
            first_page=page_index + 1,
            last_page=page_index + 1,
            dpi=200
        )
        text = ""
        for img in images:
            # 转换为OpenCV格式
            img_cv = cv2.cvtColor(np.array(img), cv2.COLOR_RGB2BGR)
            # PaddleOCR识别
            res = ocr.ocr(img_cv, cls=True)
            # 拼接识别结果
            for line in res:
                for word in line:
                    text += word[1][0] + "\n"
        return text.strip()
    except Exception as e:
        print(f"⚠️  OCR识别失败：{e}")
        return ""


def extract_cert_from_text(text):
    """从单页文本中提取校准证书信息（兼容原模板字段）"""
    if not text or not text.strip():
        return None

    # 提取证书编号（核心字段）
    cert_match = re.search(r"证\s*书\s*编\s*号[:：]\s*(AIMT[\d\-A-Z]+)", text)
    if not cert_match:
        return None
    cert_no = cert_match.group(1)

    # 提取器具名称（中英文）
    name_match = re.search(r"([^\n]+)\s*\n计\s*量\s*器\s*具\s*名\s*称\s*\n([^\n]+)\s*\nName of instrument", text)
    name_cn = name_match.group(1).strip() if name_match else ""
    name_en = name_match.group(2).strip() if name_match else ""

    # 提取型号/规格
    model_match = re.search(r"型\s*号\s*/\s*规\s*格\s*\n?([^\n]+)\s*\n?Model/Specification", text)
    model = model_match.group(1).strip() if model_match else ""

    # 提取器具编号/管理编号（温度探头兼容）
    sn_match = re.search(r"器\s*具\s*编\s*号\s*\n?([^\n]+)\s*\n?Serial No", text)
    sn = sn_match.group(1).strip() if sn_match else ""
    asset_match = re.search(r"管\s*理\s*编\s*号\s*\n?([^\n]+)\s*\n?Asset No", text)
    asset_no = asset_match.group(1).strip() if asset_match else ""
    if "温度探头" in name_cn and asset_no and asset_no != "/":
        sn = asset_no

    # 提取制造单位
    brand_match = re.search(r"制\s*造\s*单\s*位\s*\n?([^\n]+)\s*\n?Manufacturer", text)
    brand = brand_match.group(1).strip() if brand_match else ""

    # 提取校准日期并计算有效期
    cal_date_match = re.search(r"校\s*准\s*日\s*期[:：]\s*(\d{4})年(\d{1,2})月(\d{1,2})日", text)
    cal_date = ""
    due_date = ""
    if cal_date_match:
        cal_date = f"{cal_date_match.group(1)}.{cal_date_match.group(2).zfill(2)}.{cal_date_match.group(3).zfill(2)}"
        try:
            cal_dt = datetime.datetime.strptime(cal_date, "%Y.%m.%d")
            due_dt = datetime.datetime(cal_dt.year + 1, cal_dt.month, cal_dt.day) - datetime.timedelta(days=1)
            due_date = due_dt.strftime("%Y.%m.%d")
        except:
            pass

    return {
        "name_en": name_en,
        "name_cn": name_cn,
        "model": model,
        "sn": sn if sn != "/" else "",
        "brand_en": brand,
        "cert": cert_no,
        "cal_date": cal_date,
        "due": due_date
    }


def extract_one_pdf(pdf_path):
    """处理单个PDF（文本/扫描版通用）"""
    results = []
    seen_certs = set()
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for idx, page in enumerate(pdf.pages):
                # 优先提取文本
                page_text = page.extract_text()
                # 文本为空则调用OCR
                if not page_text or len(page_text.strip()) < 20:
                    page_text = pdf_ocr_text(pdf_path, idx)
                if not page_text:
                    continue
                # 提取证书信息（避免重复证书）
                cert_data = extract_cert_from_text(page_text)
                if cert_data and cert_data["cert"] not in seen_certs:
                    seen_certs.add(cert_data["cert"])
                    results.append(cert_data)
    except Exception as e:
        print(f"⚠️  处理文件失败 {pdf_path}：{e}")
    return results


def main():
    if not OCR_READY:
        print("=" * 60)
        print("提示：扫描版PDF识别需要安装PaddleOCR相关依赖")
        print("WorkBuddy会自动安装，首次运行请稍等")
        print("=" * 60)

    output_file = sys.argv[1] if len(sys.argv) > 1 else "extracted_data.json"
    pdf_dir = sys.argv[2] if len(sys.argv) > 2 else "."

    # 遍历目录下所有PDF
    all_data = []
    pdf_files = [f for f in os.listdir(pdf_dir) if f.lower().endswith(".pdf")]
    pdf_files.sort()

    if not pdf_files:
        print(f"❌ 目录 {pdf_dir} 中未找到PDF文件")
        return

    for idx, pdf_file in enumerate(pdf_files, 1):
        pdf_path = os.path.join(pdf_dir, pdf_file)
        certs = extract_one_pdf(pdf_path)
        if certs:
            for cert in certs:
                cert["no"] = str(len(all_data) + 1)
                all_data.append(cert)
            print(f"✅ [{idx}/{len(pdf_files)}] {pdf_file} 提取 {len(certs)} 份证书")
        else:
            print(f"⚠️  [{idx}/{len(pdf_files)}] {pdf_file} 未识别到有效证书信息")

    # 保存提取结果
    with open(output_file, "w", encoding="utf-8") as f:
        json.dump(all_data, f, ensure_ascii=False, indent=2)
    print(f"\n📄 提取完成，共处理 {len(pdf_files)} 个文件，保存 {len(all_data)} 份证书到 {output_file}")


if __name__ == "__main__":
    main()