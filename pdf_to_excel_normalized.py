import os
import re
import sys

try:
    import pdfplumber
    import pandas as pd
except ModuleNotFoundError as e:
    print("缺少依赖库：", e.name)
    print("请先在命令行执行：")
    print(r"D:\python3.11\python.exe -m pip install pdfplumber pandas openpyxl")
    sys.exit(1)

pdf_folder = r"./pdfs"
template_excel = r"./骨密度检测结果-2024.xlsx"
output_excel = r"./output.xlsx"

try:
    sys.stdout.reconfigure(line_buffering=True)
except Exception:
    pass

def normalize_text(text: str) -> str:
    # 处理 PDF 提取后常见的“拆字/异体部件”问题
    replacements = {
        "\u3000": " ",
        "\xa0": " ",
        "⻣": "骨",
        "⾻": "骨",
        "⼈": "人",
        "⽇": "日",
        "⽣": "生",
        "⽤": "用",
        "⽐": "比",
        "⻅": "见",
        "⽅": "方",
        "⽴": "立",
        "⻘": "青",
        "⻛": "风",
        "⼥": "女",
        "桡⻣": "桡骨",
        "⻣质": "骨质",
        "⻣折": "骨折",
        "⻣密度": "骨密度",
    }
    for old, new in replacements.items():
        text = text.replace(old, new)
    return text

def first_group(pattern, text, flags=re.S):
    m = re.search(pattern, text, flags=flags)
    return m.group(1).strip() if m else ""

def extract_bmd_t_z(text):
    text = normalize_text(text)

    # 先找“桡骨33% 数值 数值 数值”
    m = re.search(
        r'桡骨33%\s*([\-]?\d+(?:\.\d+)?)\s+([\-]?\d+(?:\.\d+)?)\s+([\-]?\d+(?:\.\d+)?)',
        text,
        flags=re.S
    )
    if m:
        return m.group(1), m.group(2), m.group(3)

    # 备用：表头到桡骨33%整段
    m = re.search(
        r'区域\s*骨密度值.*?T值评分.*?Z值评分.*?桡骨33%\s*([\-]?\d+(?:\.\d+)?)\s+([\-]?\d+(?:\.\d+)?)\s+([\-]?\d+(?:\.\d+)?)',
        text,
        flags=re.S
    )
    if m:
        return m.group(1), m.group(2), m.group(3)

    # 再备用：逐行查找
    for line in text.splitlines():
        line = line.strip()
        if "桡骨33%" in line:
            nums = re.findall(r'-?\d+(?:\.\d+)?', line)
            filtered = []
            removed_33 = False
            for x in nums:
                if x == "33" and not removed_33:
                    removed_33 = True
                    continue
                filtered.append(x)
            if len(filtered) >= 3:
                return filtered[0], filtered[1], filtered[2]

    return "", "", ""

def extract_primary_fracture_prob(text):
    text = normalize_text(text)
    patterns = [
        r'原发性骨质疏松骨折\s*([\-]?\d+(?:\.\d+)?)\s*%',
        r'原发性骨质疏松骨折概率\s*([\-]?\d+(?:\.\d+)?)\s*%',
        r'主要骨质疏松性骨折\s*([\-]?\d+(?:\.\d+)?)\s*%',
        r'主要骨质疏松骨折\s*([\-]?\d+(?:\.\d+)?)\s*%',
    ]
    for p in patterns:
        v = first_group(p, text)
        if v:
            return v
    return ""

def extract_info_from_pdf(path):
    text_parts = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text() or ""
            text_parts.append(page_text)

    raw_text = "\n".join(text_parts)
    text = normalize_text(raw_text)

    result = {
        "编号（ID）": first_group(r'ID[:：]\s*([^\s]+)', text),
        "姓名": first_group(r'姓名[:：]\s*([^\s]+)', text),
        "骨密度值（g/cm²）": "",
        "T值": "",
        "Z值": "",
        "10年髋关节骨折概率%": first_group(r'髋关节\s*([\-]?\d+(?:\.\d+)?)\s*%', text),
        "原发性骨质疏松骨折概率": extract_primary_fracture_prob(text),
        "录入员姓名": "",
    }

    bmd, t, z = extract_bmd_t_z(text)
    result["骨密度值（g/cm²）"] = bmd
    result["T值"] = t
    result["Z值"] = z

    return result

if not os.path.exists(pdf_folder):
    print(f"未找到文件夹：{pdf_folder}")
    sys.exit(1)

pdf_files = []
for root, dirs, files in os.walk(pdf_folder):
    for file in files:
        if file.lower().endswith(".pdf"):
            pdf_files.append(os.path.join(root, file))

if not pdf_files:
    print("没有找到任何PDF文件。请确认PDF放在 pdfs 文件夹中。")
    sys.exit(1)

print(f"发现 {len(pdf_files)} 个PDF文件，开始处理...")

rows = []
failed = []

for i, path in enumerate(pdf_files, 1):
    print(f"[{i}/{len(pdf_files)}] 正在处理：{os.path.basename(path)}", flush=True)
    try:
        info = extract_info_from_pdf(path)
        rows.append(info)
    except Exception as e:
        failed.append((path, str(e)))

df = pd.DataFrame(rows)

wanted_cols = [
    "编号（ID）",
    "姓名",
    "骨密度值（g/cm²）",
    "T值",
    "Z值",
    "10年髋关节骨折概率%",
    "原发性骨质疏松骨折概率",
    "录入员姓名"
]

if os.path.exists(template_excel):
    try:
        template_df = pd.read_excel(template_excel)
        template_cols = list(template_df.columns)
        for col in template_cols:
            if col not in df.columns:
                df[col] = ""
        df = df[template_cols]
    except Exception as e:
        print("模板读取失败，将按默认列顺序输出：", e)
        for col in wanted_cols:
            if col not in df.columns:
                df[col] = ""
        df = df[wanted_cols]
else:
    for col in wanted_cols:
        if col not in df.columns:
            df[col] = ""
    df = df[wanted_cols]

df.to_excel(output_excel, index=False)

print(f"\n完成！共输出 {len(df)} 条记录：{output_excel}")

if failed:
    fail_txt = "failed_files.txt"
    with open(fail_txt, "w", encoding="utf-8") as f:
        for p, err in failed:
            f.write(f"{p}\t{err}\n")
    print(f"有 {len(failed)} 个文件处理失败，详情见：{fail_txt}")
else:
    print("所有文件处理成功。")
