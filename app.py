# 勝康居家長照 - 服務紀錄核對系統（網頁版）v2
from flask import Flask, request, jsonify, render_template_string
import os, io, re, json, base64, tempfile
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB

CLAUDE_API_KEY = os.environ.get('CLAUDE_API_KEY', '')

HTML = '''<!DOCTYPE html>
<html lang="zh-TW">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>服務紀錄核對系統｜勝康居家長照</title>
<style>
* { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: "Microsoft JhengHei", sans-serif; background: #f5f6f8; color: #3c4043; }
header { background: #1B3A6B; color: white; padding: 16px 28px; display: flex; align-items: center; gap: 12px; }
header h1 { font-size: 17px; font-weight: 600; }
header p { font-size: 12px; opacity: 0.65; margin-top: 2px; }
.container { max-width: 900px; margin: 0 auto; padding: 24px 16px; }
.card { background: white; border-radius: 10px; border: 1px solid #d0d5dd; padding: 24px; margin-bottom: 20px; }
.card h2 { font-size: 15px; color: #1B3A6B; margin-bottom: 16px; display: flex; align-items: center; gap: 8px; }
.tag { background: #1B3A6B; color: white; font-size: 11px; padding: 2px 8px; border-radius: 4px; }
.upload-box { border: 2px dashed #d0d5dd; border-radius: 8px; padding: 24px; text-align: center; cursor: pointer; background: #f5f6f8; transition: all 0.2s; }
.upload-box:hover { border-color: #2E8B7A; background: #e8f5f2; }
.upload-box.loaded { border-color: #1a7a4a; background: #eafaf1; border-style: solid; }
.upload-box p { font-size: 13px; color: #9aa0a8; margin-top: 6px; }
.upload-box .fname { font-size: 14px; color: #1a7a4a; font-weight: 600; margin-top: 6px; }
input[type=file] { display: none; }
.btn { padding: 10px 24px; border: none; border-radius: 6px; font-size: 14px; font-family: inherit; cursor: pointer; font-weight: 600; }
.btn-primary { background: #2E8B7A; color: white; width: 100%; margin-top: 8px; font-size: 15px; padding: 14px; }
.btn-primary:hover { background: #24756a; }
.btn-primary:disabled { opacity: 0.5; cursor: not-allowed; }
.btn-export { background: #e67e22; color: white; margin-top: 12px; }
.progress { background: #e8eaed; border-radius: 6px; height: 8px; margin: 12px 0; overflow: hidden; display: none; }
.progress-bar { background: #2E8B7A; height: 100%; width: 0%; transition: width 0.3s; }
.log { background: #1a1a2e; color: #00ff88; border-radius: 8px; padding: 14px; font-size: 12px; font-family: monospace; min-height: 80px; max-height: 200px; overflow-y: auto; display: none; margin-top: 12px; }
.results { margin-top: 20px; }
.summary { display: flex; gap: 12px; margin-bottom: 16px; flex-wrap: wrap; }
.stat { flex: 1; min-width: 100px; padding: 12px; border-radius: 8px; text-align: center; }
.stat .num { font-size: 28px; font-weight: 700; }
.stat .lbl { font-size: 12px; margin-top: 2px; opacity: 0.8; }
.stat.ok { background: #eafaf1; color: #1a7a4a; }
.stat.ng { background: #fdecea; color: #c0392b; }
.stat.warn { background: #fffbe6; color: #9a6f00; }
.stat.total { background: #e8f5f2; color: #2E8B7A; }
table { width: 100%; border-collapse: collapse; font-size: 13px; }
th { background: #1B3A6B; color: white; padding: 9px 12px; text-align: center; font-weight: 500; }
th:first-child { text-align: left; }
td { padding: 8px 12px; border-bottom: 1px solid #e8eaed; text-align: center; }
td:first-child { text-align: left; font-weight: 500; }
tr.ok td { background: #eafaf1; }
tr.ng td { background: #fdecea; }
tr.warn td { background: #fffbe6; }
</style>
</head>
<body>
<header>
  <div>
    <h1>📋 服務紀錄核對系統</h1>
    <p>勝康居家長照機構｜行政作業輔助系統</p>
  </div>
</header>
<div class="container">
  <div class="card">
    <h2><span class="tag">步驟1</span>上傳照管家 Excel</h2>
    <div class="upload-box" id="excelBox" onclick="document.getElementById('excelInput').click()">
      <div style="font-size:32px">📊</div>
      <p>點擊選擇 .xls 或 .xlsx 檔案</p>
      <div class="fname" id="excelName"></div>
      <input type="file" id="excelInput" accept=".xls,.xlsx" onchange="handleExcel(this)">
    </div>
  </div>
  <div class="card">
    <h2><span class="tag">步驟2</span>上傳掃描 PDF</h2>
    <div class="upload-box" id="pdfBox" onclick="document.getElementById('pdfInput').click()">
      <div style="font-size:32px">📄</div>
      <p>點擊選擇掃描的服務紀錄單 PDF（可選多個）</p>
      <div class="fname" id="pdfName"></div>
      <input type="file" id="pdfInput" accept=".pdf" multiple onchange="handlePdf(this)">
    </div>
  </div>
  <div class="card">
    <h2><span class="tag">步驟3</span>開始比對</h2>
    <button class="btn btn-primary" id="startBtn" onclick="startVerify()" disabled>🔍 開始 AI 比對</button>
    <div class="progress" id="progress"><div class="progress-bar" id="progressBar"></div></div>
    <div class="log" id="log"></div>
  </div>
  <div class="card results" id="resultCard" style="display:none">
    <h2>📊 比對結果</h2>
    <div class="summary" id="summary"></div>
    <button class="btn btn-export" onclick="exportCSV()">📥 下載報告（CSV）</button>
    <table style="margin-top:16px">
      <tr><th>個案姓名</th><th>BA項目</th><th>照管家</th><th>紙本</th><th>結果</th></tr>
      <tbody id="resultBody"></tbody>
    </table>
  </div>
</div>
<script>
let excelFile = null, pdfFiles = [];
let reportData = [];

function handleExcel(input) {
  excelFile = input.files[0];
  document.getElementById('excelName').textContent = '✅ ' + excelFile.name;
  document.getElementById('excelBox').classList.add('loaded');
  checkReady();
}

function handlePdf(input) {
  pdfFiles = Array.from(input.files);
  document.getElementById('pdfName').textContent = '✅ ' + pdfFiles.length + ' 份 PDF';
  document.getElementById('pdfBox').classList.add('loaded');
  checkReady();
}

function checkReady() {
  document.getElementById('startBtn').disabled = !(excelFile && pdfFiles.length > 0);
}

function addLog(msg) {
  const log = document.getElementById('log');
  log.style.display = 'block';
  log.innerHTML += msg + '\\n';
  log.scrollTop = log.scrollHeight;
}

function setProgress(pct) {
  document.getElementById('progress').style.display = 'block';
  document.getElementById('progressBar').style.width = pct + '%';
}

async function startVerify() {
  document.getElementById('startBtn').disabled = true;
  document.getElementById('log').innerHTML = '';
  document.getElementById('resultCard').style.display = 'none';
  addLog('🚀 開始處理...');
  setProgress(10);

  const formData = new FormData();
  formData.append('excel', excelFile);
  pdfFiles.forEach(f => formData.append('pdfs', f));

  try {
    addLog('📤 上傳檔案中...');
    setProgress(20);
    const resp = await fetch('/verify', { method: 'POST', body: formData });
    const data = await resp.json();

    if (!resp.ok) {
      addLog('❌ 錯誤：' + (data.error || '未知錯誤'));
      document.getElementById('startBtn').disabled = false;
      return;
    }

    setProgress(100);
    data.logs.forEach(l => addLog(l));
    addLog('✅ 比對完成！');
    showResults(data.report);
  } catch(e) {
    addLog('❌ 連線錯誤：' + e.message);
    document.getElementById('startBtn').disabled = false;
  }
}

function showResults(report) {
  reportData = report;
  const ok   = report.filter(r => r.結果.includes('符合')).length;
  const ng   = report.filter(r => r.結果.includes('差異')).length;
  const warn = report.filter(r => r.結果.includes('未讀')).length;

  document.getElementById('summary').innerHTML = `
    <div class="stat total"><div class="num">${report.length}</div><div class="lbl">全部項目</div></div>
    <div class="stat ok"><div class="num">${ok}</div><div class="lbl">✅ 符合</div></div>
    <div class="stat ng"><div class="num">${ng}</div><div class="lbl">❌ 差異</div></div>
    <div class="stat warn"><div class="num">${warn}</div><div class="lbl">⚠ 待確認</div></div>
  `;

  const tbody = document.getElementById('resultBody');
  tbody.innerHTML = report.map(r => {
    const cls = r.結果.includes('符合') ? 'ok' : (r.結果.includes('差異') ? 'ng' : 'warn');
    return `<tr class="${cls}"><td>${r.個案姓名}</td><td>${r.BA項目}</td><td>${r.照管家次數}</td><td>${r.紙本次數}</td><td>${r.結果}</td></tr>`;
  }).join('');

  document.getElementById('resultCard').style.display = 'block';
  document.getElementById('startBtn').disabled = false;
  document.getElementById('resultCard').scrollIntoView({behavior:'smooth'});
}

function exportCSV() {
  const headers = ['個案姓名','BA項目','照管家次數','紙本次數','結果'];
  const rows = reportData.map(r => headers.map(h => r[h]).join(','));
  const csv = '\\ufeff' + headers.join(',') + '\\n' + rows.join('\\n');
  const blob = new Blob([csv], {type:'text/csv;charset=utf-8'});
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = '比對報告.csv';
  a.click();
}
</script>
</body>
</html>'''

@app.route('/')
def index():
    return render_template_string(HTML)

@app.route('/verify', methods=['POST'])
def verify():
    if not CLAUDE_API_KEY:
        return jsonify({'error': '未設定 Claude API Key'}), 500

    try:
        import anthropic
        import pandas as pd
        from pdf2image import convert_from_bytes
    except ImportError as e:
        return jsonify({'error': f'缺少套件：{e}'}), 500

    logs = []
    excel_file = request.files.get('excel')
    pdf_files  = request.files.getlist('pdfs')

    if not excel_file or not pdf_files:
        return jsonify({'error': '請上傳 Excel 和 PDF'}), 400

    # ── 讀 Excel ──────────────────────────────────────────────
    try:
        excel_bytes = excel_file.read()
        df = None
        for engine in ['xlrd', 'openpyxl']:
            try:
                df = pd.read_excel(io.BytesIO(excel_bytes), header=None, engine=engine)
                break
            except:
                continue
        if df is None:
            return jsonify({'error': 'Excel 讀取失敗，請確認檔案格式'}), 500

        header_row = -1
        for i, row in df.iterrows():
            vals = [str(v) for v in row.values]
            if '個案姓名' in vals and '次數' in vals:
                header_row = i
                break

        if header_row == -1:
            return jsonify({'error': 'Excel 找不到標題列（需含「個案姓名」和「次數」欄）'}), 500

        df.columns = df.iloc[header_row]
        df = df.iloc[header_row + 1:].reset_index(drop=True)

        excel_data = {}
        for _, row in df.iterrows():
            name  = str(row.get('個案姓名', '')).strip()
            code  = str(row.get('服務項目代碼', '')).strip()
            count = row.get('次數', 0)
            if not name or not code or name in ['合計', '總計', 'nan']:
                continue
            code = re.sub(r'\[補助\]|\[自費\]', '', code).strip().upper()
            try:
                count = int(float(count))
            except:
                count = 0
            if count <= 0:
                continue
            if name not in excel_data:
                excel_data[name] = {}
            excel_data[name][code] = excel_data[name].get(code, 0) + count

        known_names = set(excel_data.keys())
        logs.append(f'📊 Excel 讀取完成：{len(excel_data)} 位個案（{", ".join(sorted(known_names))}）')
    except Exception as e:
        return jsonify({'error': f'Excel 讀取失敗：{e}'}), 500

    # ── 處理 PDF ──────────────────────────────────────────────
    client   = anthropic.Anthropic(api_key=CLAUDE_API_KEY)
    pdf_results = {}   # {個案姓名: {BA碼: 累計次數}}
    known_str = '、'.join(sorted(known_names))

    def match_name(raw):
        """從已知名單找最相符的個案姓名"""
        raw = (raw or '').strip()
        if raw in known_names:
            return raw
        # 允許1字不同（模糊比對）
        for kn in known_names:
            if len(raw) == len(kn) and sum(a != b for a, b in zip(raw, kn)) <= 1:
                return kn
        return None

    for pdf_file in pdf_files:
        pdf_name = pdf_file.filename
        logs.append(f'📄 處理：{pdf_name}')

        # 從檔名預先猜測個案姓名（如「薛東發5月.pdf」→「薛東發」）
        stem = re.split(r'[\d_\-年月]', os.path.splitext(pdf_name)[0])[0]
        name_from_filename = match_name(stem)

        try:
            pdf_bytes = pdf_file.read()
            images = convert_from_bytes(pdf_bytes, dpi=150)
        except Exception as e:
            logs.append(f'  ✗ PDF 轉圖失敗：{e}')
            continue

        # 逐頁處理，current_name 在整份 PDF 內持續傳遞
        # 但每頁都讓 AI 重新辨識姓名，以確保多個案 PDF 正確切換
        current_name = name_from_filename  # 先用檔名當預設

        for page_num, image in enumerate(images, 1):
            try:
                buf = io.BytesIO()
                image.save(buf, format='JPEG', quality=80)
                img_b64 = base64.standard_b64encode(buf.getvalue()).decode('utf-8')

                prompt = f"""這是居家服務紀錄單的掃描圖片，請仔細辨識後，用 JSON 格式回傳以下資訊。

【判斷說明】
1. is_summary（布林值）：
   - true = 這頁表格最右側有一欄標題是「總計組數次數」，且有數字
   - false = 正面（沒有總計欄，只有日期欄位）

2. 姓名：表格左上角「姓名」欄旁邊寫的個案姓名。
   請從已知名單中選最接近的：【{known_str}】
   若辨識不出來請填 null。

3. 總計（只在 is_summary=true 時填寫）：
   讀取最右側「總計」欄每個 BA 項目旁的數字。
   只列出數字大於 0 的項目。
   注意手寫數字：8 和 3 容易混淆、1 和 7 容易混淆。

請只回傳 JSON，不要其他說明：
{{
  "is_summary": true 或 false,
  "姓名": "個案姓名或 null",
  "總計": {{"BA03": 18, "BA07": 6}}
}}"""

                response = client.messages.create(
                    model="claude-sonnet-4-6",
                    max_tokens=1500,
                    messages=[{"role": "user", "content": [
                        {"type": "image", "source": {"type": "base64", "media_type": "image/jpeg", "data": img_b64}},
                        {"type": "text", "text": prompt}
                    ]}]
                )

                # 記錄 stop_reason 幫助除錯
                stop_reason = response.stop_reason
                if not response.content:
                    logs.append(f'  第{page_num}頁 API 回傳空內容，stop_reason={stop_reason}')
                    continue

                raw_text = response.content[0].text.strip()
                logs.append(f'  第{page_num}頁 API原文（前80字）：{raw_text[:80]}')
                # 清除 markdown 包裝
                clean = re.sub(r'```json|```', '', raw_text).strip()
                data = json.loads(clean)

                is_summary  = bool(data.get('is_summary', False))
                name_ai     = data.get('姓名') or ''
                totals      = data.get('總計') or {}

                # 更新 current_name（每頁都嘗試，不只總計頁）
                matched = match_name(name_ai)
                if matched:
                    current_name = matched

                if not is_summary:
                    logs.append(f'  第{page_num}頁：正面（個案={current_name}，跳過）')
                    continue

                # 總計頁：把數字累加進 pdf_results
                valid_totals = {}
                for code, val in totals.items():
                    try:
                        n = int(str(val).strip())
                        if n > 0:
                            valid_totals[code.upper()] = n
                    except:
                        pass

                logs.append(f'  第{page_num}頁：總計頁，個案={current_name}，讀到{len(valid_totals)}項 → {valid_totals}')

                if current_name and valid_totals:
                    if current_name not in pdf_results:
                        pdf_results[current_name] = {}
                    for code, count in valid_totals.items():
                        pdf_results[current_name][code] = pdf_results[current_name].get(code, 0) + count

            except json.JSONDecodeError as e:
                logs.append(f'  第{page_num}頁 JSON 解析失敗：{e}｜原文：{raw_text[:100]}')
            except Exception as e:
                logs.append(f'  第{page_num}頁失敗：{e}')

    # ── 比對 ──────────────────────────────────────────────────
    report = []
    for case_name in sorted(set(excel_data) | set(pdf_results)):
        ei = {k.upper(): v for k, v in excel_data.get(case_name, {}).items()}
        pi = {k.upper(): v for k, v in pdf_results.get(case_name, {}).items()}
        for code in sorted(set(ei) | set(pi)):
            exp = ei.get(code, 0)
            act = pi.get(code)
            if act is None:
                status = '⚠ 紙本未讀到'
            elif act == exp:
                status = '✅ 符合'
            else:
                status = '❌ 差異'
            report.append({
                '個案姓名': case_name,
                'BA項目':   code,
                '照管家次數': exp,
                '紙本次數':  act if act is not None else '',
                '結果':      status
            })

    ok_cnt   = sum(1 for r in report if '符合' in r['結果'])
    ng_cnt   = sum(1 for r in report if '差異' in r['結果'])
    warn_cnt = sum(1 for r in report if '未讀' in r['結果'])
    logs.append(f'📋 比對完成：符合{ok_cnt}項 / 差異{ng_cnt}項 / 待確認{warn_cnt}項')

    return jsonify({'logs': logs, 'report': report})


if __name__ == '__main__':
    app.run(debug=False, host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))
