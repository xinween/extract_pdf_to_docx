---
requirements:
  - pdfplumber
  - python-docx
  - paddleocr
  - paddlepaddle
  - opencv-python
  - pdf2image
  - numpy
  - pillow
  - python-dateutil

name: calibration-cert-extractor
description: This skill should be used when the user wants to extract calibration certificate information from PDF files and output it to a docx template. Triggers include: "提取校准证书信息", "把PDF证书填到模板", "校准证书清单", "extract calibration certificate", "仪表校准".
---

# 仪表校准证书信息提取 Skill

Extract calibration certificate information from PDF files and populate into docx template.

## Quick Start

```bash
# Step 1: Extract data from PDFs (choose the right script for your cert org)
python scripts/extract_from_pdfs.py          # 邦盟检测集团 (BMG)
python scripts/extract_aimt_sqi.py <pdf>     # 爱准(AIMT) + 上海质检院(SQI)

# Step 2: Generate Word document
python scripts/create_calibration_doc.py template.docx output.docx extracted_data.json
```

## Workflow

### Step 1: Identify Calibration Organization & Select Script

| Organization | Script | Cert No. Format |
|---|---|---|
| **邦盟检测集团 (BMG)** | `extract_from_pdfs.py` | `HDYLXXXXXXX` / `HDWDXXXXXXX` |
| **爱准计量 (AIMT)** | `extract_aimt_sqi.py` | `AIMTYYYY-X-XXXXXX` |
| **上海质检院 (SQI)** | `extract_aimt_sqi.py` (same script) | `[A-Z]\d{5}[A-Z]\d{5}` |

**How to identify:** Check the first page of PDF for organization name or cert number format.

### Step 2: Find Template and PDF Files

1. Search for template file `*.docx` in workspace root (exclude `~$` temp files)
2. Search for all PDF files `*.pdf` in workspace
3. If no template found, ask user to provide one

### Step 3: Run Extraction

```bash
# For BMG certificates:
python scripts/extract_from_pdfs.py [output.json] [pdf_directory]

# For AIMT/SQI certificates:
python scripts/extract_aimt_sqi.py <pdf_path> [output.json]
```

### Step 4: Field Formatting Rules

See `references/field_rules.md` for detailed formatting rules.

**Required fields (in order):**

| Field | EN | CN | Notes |
|-------|----|----|-------|
| 序号 | No. | 序号 | Auto-detect Word numbering, skip if template has auto-num |
| 名称 | Instruments Name | 仪器仪表名称 | English on top, Chinese below (`\n`) |
| P&ID | P&ID No | P&ID编号 | Leave empty if not available |
| SN | Series No. | 序列号 | Use management number (管理编号) for temperature probes |
| Accuracy | Accuracy | 精度 | Optional column — only fill if template has this column |
| Range | Range | 量程 | Optional column — only fill if template has this column |
| 品牌 | Brand | 品牌/生产厂家 | English on top, Chinese below (`\n`) |
| 型号 | Model | 型号 | Include spec range e.g. `(-0.1~3.8)MPa` |
| 证书编号 | Calibration Certificate No. | 校准证书编号 | Full cert number as-is |
| 校准日期 | Cal. Date | 校准日期 | Format: YYYY.MM.DD |
| 有效期 | Due Date | 有效期至 | Priority: "建议复校日期" > cal_date+12m-1d |
| 备注 | Remark | 备注 | Always leave empty |

**Key rules:**
- **Bilingual text**: English on top, Chinese below, use `\n` (manual line break)
- **Date format**: YYYY.MM.DD
- **有效期至**: Priority to "建议于...前复校" date; otherwise calculate as cal_date + 12 months - 1 day
- **备注**: Always leave empty, never fill
- **温度探头**: Use management number (管理编号) instead of instrument serial number (器具编号)

### Step 5: Generate Output

1. Use `scripts/create_calibration_doc.py` to populate template
2. The script **auto-detects**:
   - Column mapping by header keywords (supports 10-col and 12-col templates)
   - Auto-numbered columns via `w:numId` XML detection → skips writing
3. Read template using docx library
4. Find the data table (typically the 4th table, index=3)
5. Clear existing data rows, keep headers
6. Insert extracted data with dynamic column mapping
7. Save output file

**Template compatibility:**

| Template Type | Columns | Has Accuracy? | Has Range? |
|---|---|---|---|
| Basic | 10 | ❌ | ❌ |
| Extended | 12 | ✅ | ✅ |

The script automatically maps columns by reading header text — no manual configuration needed.

### Step 6: Output Results

1. Display standard markdown table summary
2. Deliver the generated docx file
3. Also deliver the extracted JSON for review/debugging

## Scripts

| Script | Purpose | Usage |
|--------|---------|-------|
| `extract_from_pdfs.py` | BMG (邦盟) certificate extraction | `python extract_from_pdfs.py [output.json] [pdf_dir]` |
| `extract_aimt_sqi.py` | AIMT (爱准) + SQI (质检院) extraction | `python extract_aimt_sqi.py <pdf_path> [output.json]` |
| `create_calibration_doc.py` | Populate Word template with data | `python create_calibration_doc.py <template.docx> <output.docx> [data.json]` |

## References

- `references/field_rules.md`: Detailed field formatting rules
- `references/supported_orgs.md`: List of supported calibration organizations
