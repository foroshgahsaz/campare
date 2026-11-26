#!/usr/bin/env python3
"""Bundle libs into single stock/stock-report.html (offline)."""

import base64
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parent
OUT = ROOT / "stock-report.html"
APP = ROOT / "app.js"

LIBS = {
    "XLSX": REPO / "xlsx.mini.min.js",
    "HTML2PDF": ROOT / "html2pdf.bundle.min.js",
    "FONT": ROOT / "Vazirmatn-Regular.woff2",
}

URLS = {
    "XLSX": "https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.mini.min.js",
    "HTML2PDF": "https://unpkg.com/html2pdf.js@0.10.2/dist/html2pdf.bundle.min.js",
    "FONT": "https://cdn.jsdelivr.net/gh/rastikerdar/vazirmatn@v33.003/fonts/webfonts/Vazirmatn-Regular.woff2",
}


def ensure_lib(key: str) -> Path:
    path = LIBS[key]
    if not path.is_file():
        path.parent.mkdir(parents=True, exist_ok=True)
        print(f"Downloading {key} …")
        urllib.request.urlretrieve(URLS[key], path)
    return path


def font_face_data_url() -> str:
    data = ensure_lib("FONT").read_bytes()
    b64 = base64.b64encode(data).decode("ascii")
    return (
        "@font-face {"
        'font-family: "Vazirmatn";'
        "src: url(data:font/woff2;base64," + b64 + ') format("woff2");'
        "font-weight: 400;"
        "font-style: normal;"
        "}"
    )


HTML_HEAD = """<!DOCTYPE html>
<html lang="fa" dir="rtl">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>گزارش موجودی — لیست محصول</title>
  <style>
{FONT_FACE}
    :root {{
      --bg: #0f1419;
      --surface: #1a2332;
      --border: #2d3a4d;
      --text: #e7ecf3;
      --muted: #8b9cb3;
      --accent: #10b981;
      --accent-hover: #059669;
      --danger: #ef4444;
      --warn: #f59e0b;
      --ok: #22c55e;
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      font-family: "Vazirmatn", Tahoma, sans-serif;
      background: var(--bg);
      color: var(--text);
      line-height: 1.55;
      min-height: 100vh;
    }}
    .wrap {{ max-width: 1200px; margin: 0 auto; padding: 1.25rem 1rem 2.5rem; }}
    h1 {{ font-size: 1.45rem; margin: 0 0 0.25rem; }}
  .sub {{ color: var(--muted); font-size: 0.88rem; margin-bottom: 1.25rem; }}
    .card {{
      background: var(--surface);
      border: 1px solid var(--border);
      border-radius: 12px;
      padding: 1.1rem 1.15rem;
      margin-bottom: 0.9rem;
    }}
    .card h2 {{ font-size: 1rem; margin: 0 0 0.75rem; }}
    .grid {{ display: grid; gap: 0.85rem; }}
    @media (min-width: 720px) {{
      .grid-2 {{ grid-template-columns: 1fr 1fr; }}
      .grid-3 {{ grid-template-columns: 1fr 1fr 1fr; }}
    }}
    label {{ display: block; font-size: 0.82rem; color: var(--muted); margin-bottom: 0.3rem; }}
    input, select {{
      width: 100%;
      padding: 0.5rem 0.6rem;
      border: 1px solid var(--border);
      border-radius: 8px;
      background: #0f1419;
      color: var(--text);
      font-family: inherit;
      font-size: 0.9rem;
    }}
    input[type="file"] {{
      border-style: dashed;
      padding: 0.55rem;
      cursor: pointer;
    }}
    input[type="file"]::file-selector-button {{
      margin-left: 0.5rem;
      padding: 0.35rem 0.7rem;
      border: none;
      border-radius: 6px;
      background: var(--accent);
      color: #fff;
      font-family: inherit;
      cursor: pointer;
    }}
    .meta {{ font-size: 0.78rem; color: var(--muted); margin-top: 0.25rem; min-height: 1.1rem; }}
    .actions {{ display: flex; flex-wrap: wrap; gap: 0.6rem; margin-top: 0.9rem; }}
    button {{
      padding: 0.6rem 1.1rem;
      border: none;
      border-radius: 8px;
      font-family: inherit;
      font-size: 0.92rem;
      cursor: pointer;
    }}
    button.primary {{ background: var(--accent); color: #fff; }}
    button.primary:hover:not(:disabled) {{ background: var(--accent-hover); }}
    button.secondary {{ background: transparent; border: 1px solid var(--border); color: var(--text); }}
    button:disabled {{ opacity: 0.45; cursor: not-allowed; }}
    .error {{ color: var(--danger); font-size: 0.88rem; margin-top: 0.5rem; }}
    .stats {{ display: flex; flex-wrap: wrap; gap: 0.6rem; margin-bottom: 0.75rem; }}
    .stat {{
      flex: 1; min-width: 130px;
      background: #0f1419;
      border: 1px solid var(--border);
      border-radius: 8px;
      padding: 0.55rem 0.75rem;
    }}
    .stat strong {{ display: block; font-size: 1.2rem; }}
    .stat span {{ font-size: 0.75rem; color: var(--muted); }}
    .stat.ok strong {{ color: var(--ok); }}
    .stat.warn strong {{ color: var(--warn); }}
    .table-wrap {{ overflow: auto; max-height: 50vh; border: 1px solid var(--border); border-radius: 8px; }}
    table {{ width: 100%; border-collapse: collapse; font-size: 0.82rem; }}
    th, td {{ padding: 0.45rem 0.55rem; border-bottom: 1px solid var(--border); text-align: right; white-space: nowrap; }}
    th {{ position: sticky; top: 0; background: #243044; z-index: 1; }}
    tr.row-missing td {{ background: rgba(245, 158, 11, 0.08); }}
    .hidden {{ display: none !important; }}
    .logo-preview {{ max-height: 56px; max-width: 160px; margin-top: 0.5rem; object-fit: contain; }}

    /* ناحیهٔ PDF — باید در viewport باشد تا html2canvas رندر کند */
    #pdfRoot {{
      position: fixed;
      left: 0;
      top: 0;
      width: 270mm;
      max-width: 100vw;
      background: #fff;
      color: #111;
      font-family: "Vazirmatn", Tahoma, sans-serif;
      direction: rtl;
      padding: 12mm 14mm 16mm;
      z-index: -1;
      opacity: 1;
      pointer-events: none;
      overflow: visible;
    }}
    #pdfRoot.pdf-exporting {{
      z-index: 100001;
    }}
    .pdf-export-mask {{
      position: fixed;
      inset: 0;
      z-index: 100000;
      background: rgba(15, 20, 25, 0.65);
      display: flex;
      align-items: center;
      justify-content: center;
      color: #fff;
      font-size: 1rem;
    }}
    .pdf-header {{
      display: flex;
      align-items: center;
      justify-content: space-between;
      border-bottom: 2px solid #10b981;
      padding-bottom: 8px;
      margin-bottom: 10px;
    }}
    .pdf-header img {{ max-height: 52px; max-width: 140px; object-fit: contain; }}
    .pdf-header h1 {{ font-size: 16px; margin: 0; color: #111; }}
    .pdf-header .date {{ font-size: 11px; color: #555; margin-top: 4px; }}
    .pdf-table {{ width: 100%; border-collapse: collapse; font-size: 10px; }}
    .pdf-table th, .pdf-table td {{
      border: 1px solid #ccc;
      padding: 4px 6px;
      text-align: right;
    }}
    .pdf-table th {{ background: #f0fdf4; color: #111; }}
    .pdf-footer {{
      margin-top: 12px;
      padding-top: 8px;
      border-top: 1px solid #ddd;
      font-size: 10px;
      color: #444;
      text-align: center;
    }}
  </style>
</head>
<body>
  <div class="wrap">
    <h1>گزارش موجودی بر اساس لیست محصول</h1>
    <p class="sub">فایل لیست محصول (قیمت) را با موجودی انبار بر اساس <strong>کد کالا</strong> تطبیق می‌دهد و خروجی PDF / Excel می‌دهد. همه‌چیز در مرورگر شما اجرا می‌شود.</p>

    <div class="card">
      <h2>مشخصات مجموعه (هدر و فوتر PDF)</h2>
      <div class="grid grid-3">
        <div>
          <label for="companyName">نام مجموعه</label>
          <input type="text" id="companyName" placeholder="مثال: فروشگاه …" />
        </div>
        <div>
          <label for="companyPhone">تلفن</label>
          <input type="text" id="companyPhone" placeholder="021-…" />
        </div>
        <div>
          <label for="reportTitle">عنوان گزارش</label>
          <input type="text" id="reportTitle" value="گزارش موجودی لیست محصول" />
        </div>
      </div>
      <div class="grid grid-2" style="margin-top:0.75rem;">
        <div>
          <label for="companyAddress">آدرس (فوتر PDF)</label>
          <input type="text" id="companyAddress" placeholder="آدرس کامل …" />
        </div>
        <div>
          <label for="logoFile">لوگو (هدر PDF)</label>
          <input type="file" id="logoFile" accept="image/png,image/jpeg,image/webp,image/svg+xml" />
          <img id="logoPreview" class="logo-preview hidden" alt="لوگو" />
        </div>
      </div>
    </div>

    <div class="card">
      <div class="grid grid-2">
        <div>
          <h2>لیست محصول / قیمت</h2>
          <label for="fileProduct">فایل اکسل</label>
          <input type="file" id="fileProduct" accept=".xlsx,.xls,.xlsm,.csv" />
          <div class="meta" id="metaProduct"></div>
          <label for="sheetProduct" style="margin-top:0.5rem;">برگه</label>
          <select id="sheetProduct" disabled><option>—</option></select>
          <label for="headerRowProduct" style="margin-top:0.45rem;">شماره سطر عنوان</label>
          <input type="number" id="headerRowProduct" value="1" min="1" />
          <label for="colProductCode" style="margin-top:0.45rem;">ستون کد کالا</label>
          <select id="colProductCode"></select>
        </div>
        <div>
          <h2>موجودی انبار</h2>
          <label for="fileInventory">فایل اکسل</label>
          <input type="file" id="fileInventory" accept=".xlsx,.xls,.xlsm,.csv" />
          <div class="meta" id="metaInventory"></div>
          <label for="sheetInventory" style="margin-top:0.5rem;">برگه</label>
          <select id="sheetInventory" disabled><option>—</option></select>
          <label for="headerRowInventory" style="margin-top:0.45rem;">شماره سطر عنوان</label>
          <input type="number" id="headerRowInventory" value="1" min="1" />
          <label for="colInventoryCode" style="margin-top:0.45rem;">ستون کد کالا</label>
          <select id="colInventoryCode"></select>
          <label for="colInventoryQty" style="margin-top:0.45rem;">ستون موجودی / مانده</label>
          <select id="colInventoryQty"></select>
        </div>
      </div>
      <div class="actions">
        <button type="button" class="primary" id="btnRun" disabled>ساخت گزارش</button>
        <button type="button" class="secondary" id="btnPdf" disabled>دانلود PDF</button>
        <button type="button" class="secondary" id="btnXlsx" disabled>دانلود Excel</button>
      </div>
      <div class="error" id="errBox"></div>
    </div>

    <div class="card hidden" id="resultCard">
      <div class="stats" id="stats"></div>
      <div class="table-wrap">
        <table>
          <thead id="previewHead"><tr></tr></thead>
          <tbody id="previewBody"></tbody>
        </table>
      </div>
    </div>
  </div>

  <div id="pdfRoot" aria-hidden="true">
    <div class="pdf-header">
      <div>
        <h1 id="pdfTitle">گزارش موجودی لیست محصول</h1>
        <div class="date" id="pdfDate"></div>
      </div>
      <img id="pdfLogo" class="hidden" alt="" />
    </div>
    <table class="pdf-table">
      <thead><tr id="pdfTableHead"></tr></thead>
      <tbody id="pdfTableBody"></tbody>
    </table>
    <div class="pdf-footer" id="pdfFooter"></div>
  </div>
"""


def main() -> None:
    font_css = font_face_data_url()
    head = HTML_HEAD.replace("{{", "{").replace("}}", "}").replace("{FONT_FACE}", font_css)

    xlsx_src = ensure_lib("XLSX").read_text(encoding="utf-8", errors="surrogateescape")
    html2pdf_src = ensure_lib("HTML2PDF").read_text(encoding="utf-8", errors="surrogateescape")
    app_src = APP.read_text(encoding="utf-8")

    out = (
        head
        + "\n  <script>\n"
        + xlsx_src
        + "\n  </script>\n  <script>\n"
        + html2pdf_src
        + "\n  </script>\n  <script>\n"
        + app_src
        + "\n  </script>\n</body>\n</html>\n"
    )
    OUT.write_text(out, encoding="utf-8")
    print(f"Wrote {OUT} ({OUT.stat().st_size:,} bytes)")


if __name__ == "__main__":
    main()