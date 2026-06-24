# 勝康居家長照 - 服務紀錄核對系統（網頁版）
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
  const csv = '\ufeff' + headers.join(',') + '\\n' + rows.join('\\n');
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

    # 讀 Excel
    try:
        excel_bytes = excel_file.read()
        for engine in ['xlrd','openpyxl']:
            try:
                df = pd.read_excel(io.BytesIO(excel_bytes), header=None, engine=engine)
                break
            except: continue

        header_row = -1
        for i, row in df.iterrows():
            if '個案姓名' in [str(v) for v in row.values] and '次數' in [str(v) for v in row.values]:
                header_row = i; break

        df.columns = df.iloc[header_row]
        df = df.iloc[header_row+1:].reset_index(drop=True)

        excel_data = {}
        for _, row in df.iterrows():
            name = str(row.get('個案姓名','')).strip()
            code = str(row.get('服務項目代碼','')).strip()
            count = row.get('次數', 0)
            if not name or not code or name in ['合計','總計','nan']: continue
            code = re.sub(r'\[補助\]|\[自費\]','',code).strip().upper()
            try: count = int(float(count))
            except: count = 0
            if name not in excel_data: excel_data[name] = {}
            excel_data[name][code] = excel_data[name].get(code,0) + count

        known_names = set(excel_data.keys())
        logs.append(f'📊 Excel 讀取完成：{len(excel_data)} 位個案（{", ".join(known_names)}）')
    except Exception as e:
        return jsonify({'error': f'Excel 讀取失敗：{e}'}), 500

    # 處理 PDF
    client = anthropic.Anthropic(api_key=CLAUDE_API_KEY)
    pdf_results = {}
    known_str = '、'.join(known_names)

    for pdf_file in pdf_files:
        pdf_name = pdf_file.filename
        logs.append(f'📄 處理：{pdf_name}')
        try:
            pdf_bytes = pdf_file.read()
            images = convert_from_bytes(pdf_bytes, dpi=150)
        except Exception as e:
            logs.append(f'  ✗ PDF 轉圖失敗：{e}')
            continue

        current_name = None
        for page_num, image in enumerate(images, 1):
            try:
                buf = io.BytesIO()
                image.save(buf, format='JPEG', quality=80)
                img_b64 = base64.standard_b64encode(buf.getvalue()).decode('utf-8')

                prompt = f"""這是一份居家服務紀錄單的掃描圖片。

請判斷並回傳以下資訊：

1. 這頁有沒有「總計組數次數」欄？
   - 表格最右側有一欄標題寫「總計組數次數」→ is_summary: true
   - 沒有這欄 → is_summary: false

2. 找「姓名」欄位旁邊的個案姓名，從已知名單選最相符的：{known_str}

3. 如果有總計欄，讀出每個BA項目的總計次數（只讀最右側總計欄）

請用以下 JSON 格式回傳：
{{
  "is_summary": true,
  "姓名": "從已知名單選一個",
  "總計": {{"BA03": 10, "BA07": 6}}
}}

注意：
- 姓名必須從已知名單 [{known_str}] 中選
- 只列出總計大於0的BA項目
- 手寫數字注意：8和5容易混淆
- is_summary 為 false 時總計填空的 {{}}"""

                response = client.messages.create(
                    model="claude-sonnet-4-6",
                    max_tokens=1000,
                    messages=[{"role":"user","content":[
                        {"type":"image","source":{"type":"base64","media_type":"image/jpeg","data":img_b64}},
                        {"type":"text","text":prompt}
                    ]}]
                )

                text = re.sub(r'```json|```','',response.content[0].text).strip()
                data = json.loads(text)
                is_summary = data.get('is_summary', False)
                name = (data.get('姓名','') or '').strip()
                totals = {k.upper(): int(v) for k,v in (data.get('總計',{}) or {}).items()
                          if str(v).isdigit() and int(v) > 0}

                if name in known_names:
                    current_name = name
                elif name:
                    for kn in known_names:
                        if len(name)==len(kn) and sum(a!=b for a,b in zip(name,kn))<=1:
                            current_name = kn; break

                if not is_summary:
                    logs.append(f'  第{page_num}頁：正面（跳過）')
                    continue

                logs.append(f'  第{page_num}頁：總計頁，個案={current_name}，{len(totals)}項')

                if current_name and totals:
                    if current_name not in pdf_results: pdf_results[current_name] = {}
                    for code, count in totals.items():
                        pdf_results[current_name][code] = pdf_results[current_name].get(code,0) + count

            except Exception as e:
                logs.append(f'  第{page_num}頁失敗：{e}')

    # 比對
    report = []
    excel_norm = {n:{k.upper():v for k,v in items.items()} for n,items in excel_data.items()}
    pdf_norm   = {n:{k.upper():v for k,v in items.items()} for n,items in pdf_results.items()}

    for case_name in sorted(set(excel_norm)|set(pdf_norm)):
        ei = excel_norm.get(case_name,{})
        pi = pdf_norm.get(case_name,{})
        for code in sorted(set(ei)|set(pi)):
            exp = ei.get(code,0)
            act = pi.get(code,None)
            status = '⚠ 紙本未讀到' if act is None else ('✅ 符合' if act==exp else '❌ 差異')
            report.append({'個案姓名':case_name,'BA項目':code,
                           '照管家次數':exp,'紙本次數':act if act is not None else '','結果':status})

    return jsonify({'logs': logs, 'report': report})

if __name__ == '__main__':
    app.run(debug=False, host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))
