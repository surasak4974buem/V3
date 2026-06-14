# -*- coding: utf-8 -*-
"""Replace Wikipedia garuda URL with embedded base64 local image + fix print popup"""
import base64, re
from pathlib import Path

BASE = Path(r"D:\Cowork\NKPmedparts2026V2")
GARUDA = BASE / "เอกสารและรูป" / "ครุฑเล็ก (ไม่มีพื้นขาว).png"

# Build data URI
with open(GARUDA, "rb") as f:
    b64 = base64.b64encode(f.read()).decode()
DATA_URI = f"data:image/png;base64,{b64}"

# Read HTML
html_path = BASE / "index.html"
html = html_path.read_text(encoding="utf-8")

# 1. Replace garuda img block (from <img src="https://upload... to closing </svg>)
old_block = (
    '<img src="https://upload.wikimedia.org/wikipedia/commons/thumb/a/a8/'
    'Garuda_Emblem_of_Thailand_%28Monochrome_2%29.svg/'
    '120px-Garuda_Emblem_of_Thailand_%28Monochrome_2%29.svg.png" \n'
    '                         alt="ตราครุฑ" \n'
    '                         style="width: 55px; height: auto;"\n'
    "                         onerror=\"this.style.display='none'; "
    "document.getElementById('garuda-svg-fallback').style.display='block';\">\n"
    '                    <svg id="garuda-svg-fallback" viewBox="0 0 100 100" '
    'style="display: none; width: 55px; height: 55px; fill: black;">\n'
    '                      <path d="M50 15c-1.5 4.5-5 8.5-8.5 12-3.5 3.5-7.5 5.5-12.5 6 1.5 2 4 4.5 7.5 6.5 3.5 2 6 2.5 8.5.5s4-5 5-10c1 5 2.5 8 5 10s5 1.5 8.5-.5c3.5-2 6-4.5 7.5-6.5-5-.5-9-2.5-12.5-6-3.5-3.5-7-7.5-8.5-12zm0 18c-3 6-7 11-12 15-5 4-11 7-18 9 4 1 8 1 12-1 4-2 7-5 10-9 2 4 5 7 9 9 4 2 8 2 12 1-7-2-13-5-18-9-5-4-9-9-12-15zm-28 26c-1 3-3 6-6 8 4-1 7-3 9-6-2-1-3-2-3-2zm56 0c0 0-1 1-3 2 2 3 5 5 9 6-3-2-5-5-6-8zM50 36c-2 4-5 8-8 11s-7 5-11 6c3.5 1 7.5.5 11-1.5 1.5 1 3 2.5 4.5 4.5 1.5-2 3-3.5 4.5-4.5 3.5 2 7.5 2.5 11 1.5-4-1-8-3-11-6s-6-7-8-11z" />\n'
    '                    </svg>'
)
new_block = f'<img src="{DATA_URI}" alt="ตราครุฑ" style="width: 60px; height: auto; display:block;">'

if old_block in html:
    html = html.replace(old_block, new_block)
    print("[1] garuda image block replaced with base64 PNG")
else:
    # Fallback: regex to replace just the src attribute of the img
    print("[1] exact match failed, trying src-only replacement …")
    html_new = re.sub(
        r'src="https://upload\.wikimedia\.org[^"]*Garuda[^"]*"',
        f'src="{DATA_URI}"',
        html
    )
    if html_new != html:
        html = html_new
        # Also remove the onerror and the fallback SVG if they exist
        html = re.sub(
            r"\s+onerror=\"this\.style\.display='none';[^\"]+\"",
            "",
            html,
        )
        html = re.sub(
            r'<svg id="garuda-svg-fallback"[^>]*>.*?</svg>',
            "",
            html,
            flags=re.DOTALL,
        )
        print("[1] garuda src replaced via regex")
    else:
        print("[1] WARNING: garuda not found at all, skip")

# 2. Fix print button: replace window.print() with popup window approach
old_print = (
    'document.getElementById("btn-print-memo").addEventListener("click", '
    '() => window.print());'
)
new_print = r"""document.getElementById("btn-print-memo").addEventListener("click", () => {
        const printArea = document.getElementById("pr-print-area");
        if (!printArea) return;
        const w = window.open("", "_blank", "width=900,height=1200,scrollbars=yes");
        if (!w) { alert("กรุณาอนุญาต Popup สำหรับหน้านี้ก่อนกดพิมพ์"); return; }
        w.document.write(`<!DOCTYPE html>
<html lang="th">
<head>
<meta charset="UTF-8">
<title>บันทึกข้อความ - พิมพ์</title>
<style>
  @import url('https://fonts.googleapis.com/css2?family=Sarabun:ital,wght@0,400;0,700;1,400&display=swap');
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { background: #f0f0f0; font-family: 'Sarabun', 'TH Sarabun New', sans-serif; }
  #pr-print-area {
    background: white; color: black;
    width: 210mm; min-height: 297mm;
    margin: 20px auto; padding: 40px 50px;
    box-shadow: 0 4px 20px rgba(0,0,0,.2);
    font-size: 15px; line-height: 1.6;
    position: relative;
  }
  @media print {
    @page { size: A4 portrait; margin: 0; }
    body { background: white !important; }
    #pr-print-area {
      box-shadow: none !important;
      margin: 0 !important;
      width: 100% !important;
      min-height: 100vh;
    }
    .no-print { display: none !important; }
  }
  .print-toolbar {
    text-align: center; padding: 12px; background: #003580;
    color: white; font-family: sans-serif; font-size: 14px;
  }
  .print-toolbar button {
    margin: 0 6px; padding: 6px 20px;
    background: white; color: #003580;
    border: none; border-radius: 4px;
    cursor: pointer; font-size: 14px; font-weight: bold;
  }
  table { border-collapse: collapse; }
  th, td { border: 1px solid black; padding: 6px 4px; }
</style>
</head>
<body>
<div class="print-toolbar no-print">
  <span>ตัวอย่างบันทึกข้อความ A4</span>
  &nbsp;&nbsp;
  <button onclick="window.print()">🖨️ พิมพ์</button>
  <button onclick="window.close()">✕ ปิด</button>
</div>
${printArea.outerHTML}
</body>
</html>`);
        w.document.close();
      });"""

if old_print in html:
    html = html.replace(old_print, new_print)
    print("[2] print button handler replaced with popup window approach")
else:
    print("[2] WARNING: old print handler not found exactly")
    # Try alternate
    alt = 'getElementById("btn-print-memo").addEventListener("click", () => window.print())'
    if alt in html:
        html = html.replace(
            'addEventListener("click", () => window.print())',
            new_print.split('addEventListener("click",')[1],
        )
        print("[2] print handler replaced via alt match")
    else:
        print("[2] FAILED to find print handler")

html_path.write_text(html, encoding="utf-8")
print("index.html saved.")
