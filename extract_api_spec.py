# pdf2api_excel_auto_index.py
import re, json, shutil
from pathlib import Path
from typing import List, Dict
import pandas as pd
import pdfplumber

# ===== 路徑設定 =====
ROOT = Path(__file__).resolve().parent
SPEC_DIR = (ROOT / "spec_input").resolve()        # 放 PDF 的資料夾
OUTPUT_DIR = (ROOT / "output").resolve()          # 輸出資料夾
OUTPUT_XLSX = OUTPUT_DIR / "API_upload_.xlsx"     # 最終輸出檔名


# ===== 目錄確保存在 =====
def ensure_dirs():
    """若缺少 spec_input / output 就自動建立"""
    SPEC_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


# ===== 清空 output =====
def clean_output(dir_: Path):
    """
    建立並清空 output 目錄（刪除舊檔/舊資料夾），
    但保留以 . 開頭的檔案（例如 .gitkeep、.gitignore）
    """
    dir_.mkdir(parents=True, exist_ok=True)
    for p in dir_.iterdir():
        # 跳過 .gitkeep、.gitignore 等隱藏檔
        if p.name.startswith("."):
            continue

        if p.is_file():
            try:
                p.unlink()
            except Exception:
                pass
        elif p.is_dir():
            shutil.rmtree(p, ignore_errors=True)


# ===== 讀取 PDF 文字 =====
def read_pdf_text(pdf: Path) -> str:
    """把整份 PDF 的可選文字串起來；做簡單全形→半形的標準化"""
    lines: List[str] = []
    with pdfplumber.open(pdf) as doc:
        for page in doc.pages:
            t = page.extract_text() or ""
            if t:
                t = (
                    t.replace("：", ":")
                    .replace("（", "(")
                    .replace("）", ")")
                    .replace("．", ".")
                    .replace("－", "-")
                )
                lines.append(t.strip())
    return "\n".join(lines)


# ===== 找出所有 API URL =====
def find_urls(text: str) -> List[str]:
    """從全文找出所有 'POST /...' 的 URL（依出現順序去重）"""
    urls, seen = [], set()
    for m in re.finditer(r"POST\s+(/\S+)", text):
        u = m.group(1)
        if u not in seen:
            seen.add(u)
            urls.append(u)
    return urls


# ===== 抓 URL 附近的多段 JSON =====
def json_blocks_near(text: str, center: int, radius=30000, max_blocks=12) -> List[str]:
    """以 URL 位置為中心，往前後掃文字，擷取多段『平衡大括號』的 JSON"""
    s = max(0, center - radius)
    e = min(len(text), center + radius)
    span = text[s:e]
    out, i = [], 0
    while len(out) < max_blocks:
        p = span.find("{", i)
        if p == -1:
            break
        depth, start = 0, p
        for j in range(p, len(span)):
            ch = span[j]
            if ch == "{":
                depth += 1
            elif ch == "}":
                depth -= 1
                if depth == 0:
                    raw = span[start : j + 1].strip()
                    if len(raw) > 50:
                        out.append(raw)
                    i = j + 1
                    break
        else:
            break
    return out


# ===== 從多段 JSON 中挑出 Input/Output =====
def pick_io(blocks: List[str]) -> (str, str):
    def is_resp(b):
        t = b.lower()
        return any(k in t for k in ["msgrshdr", "rspcode", "responsejson", "error"])

    def is_req(b):
        t = b.lower()
        return any(k in t for k in ["securitycontext", "custid", "userid", "data", "body"]) and not is_resp(b)

    inp = next((b for b in blocks if is_req(b)), "")
    out = next((b for b in blocks if is_resp(b)), "")

    if not inp and blocks:
        inp = max(blocks, key=len)
    if not out:
        rest = [b for b in blocks if b is not inp]
        out = max(rest, key=len) if rest else ""

    def norm(s):
        try:
            return json.dumps(json.loads(s), ensure_ascii=False, indent=2)
        except Exception:
            return s

    return norm(inp), norm(out)


# ===== 主程式 =====
def main():
    ensure_dirs()           # 👈 新增：先確保兩個資料夾存在
    clean_output(OUTPUT_DIR)

    rows: List[Dict[str, str]] = []
    idx = 1  # 全域累加 Index

    for pdf in sorted(SPEC_DIR.glob("*.pdf")):
        text = read_pdf_text(pdf)
        if not text.strip():
            rows.append(
                {
                    "Index": idx,
                    "FileName": pdf.name,
                    "URL": "",
                    "Method": "POST",
                    "Input（上行 JSON）": "",
                    "Response Code": "200",
                    "Output（下行 JSON）": "",
                }
            )
            idx += 1
            continue

        urls = find_urls(text)
        for url in urls:
            m = re.search(re.escape(url), text)
            blocks = json_blocks_near(text, m.start()) if m else []
            inp, out = pick_io(blocks)
            rows.append(
                {
                    "Index": idx,
                    "FileName": pdf.name,
                    "URL": url,
                    "Method": "POST",
                    "Input（上行 JSON）": inp,
                    "Response Code": "200",
                    "Output（下行 JSON）": out,
                }
            )
            idx += 1

    # 存成 Excel
    df = pd.DataFrame(
        rows,
        columns=[
            "Index",
            "FileName",
            "URL",
            "Method",
            "Input（上行 JSON）",
            "Response Code",
            "Output（下行 JSON）",
        ],
    )
    with pd.ExcelWriter(OUTPUT_XLSX, engine="xlsxwriter") as w:
        df.to_excel(w, index=False, sheet_name="API")
        ws = w.sheets["API"]
        ws.set_column(0, 0, 8)   # Index
        ws.set_column(1, 1, 32)  # FileName
        ws.set_column(2, 2, 46)  # URL
        ws.set_column(3, 3, 10)  # Method
        ws.set_column(4, 4, 90)  # Input
        ws.set_column(5, 5, 12)  # Response Code
        ws.set_column(6, 6, 90)  # Output

    print(f"✅ 完成：{OUTPUT_XLSX}")

if __name__ == "__main__":
    main()