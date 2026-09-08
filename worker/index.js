// Pembayran Region - Cloudflare Worker
// Serves frontend only; Excel parsing happens client-side via SheetJS

export default {
  async fetch(request) {
    const url = new URL(request.url);
    const path = url.pathname;

    // CORS headers
    const corsHeaders = {
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type',
    };

    if (request.method === 'OPTIONS') {
      return new Response(null, { headers: corsHeaders });
    }

    // Frontend
    return new Response(HTML_PAGE, {
      headers: { 'Content-Type': 'text/html;charset=UTF-8' },
    });
  },
};

// Sheet name to display name mapping
var SHEET_DISPLAY = {
  'email CDP': 'Email CDP',
  'rekrut': 'Rekrut',
  'BAPP': 'BAPP',
  'Vendor': 'Vendor',
  'COLA': 'COLA',
  'CS': 'CS',
  'Kontanan': 'Kontanan',
  'Perdin RO': 'Perdin RO',
  'LPJ Legal': 'LPJ Legal',
  'Lain-Lain': 'Lain-Lain',
  'CDP': 'CDP',
  'Olah': 'Olah',
  'Paste': 'Paste',
  'Konfirm Kebun': 'Konfirm Kebun',
  'Timesheet': 'Timesheet',
  'LEmbur': 'Lembur',
  'BOT': 'BOT',
  'Hit Token': 'Hit Token',
  'Sheet4': 'Sheet4',
  'WORKPLACE': 'Workplace',
  'PDR': 'PDR',
  'Sheet1': 'Sheet1',
  'M Rekening': 'M Rekening',
  'Rekap': 'Rekap',
};

var GITHUB_RAW_URL = 'https://raw.githubusercontent.com/Jiganyusi/Pembayran-Region/main/2026%20-%20Rekap%20Input%20CAMS.xlsx';

// HTML Page - all logic client-side using SheetJS from CDN
var HTML_PAGE = '<!DOCTYPE html>\n' +
'<html lang="id">\n' +
'<head>\n' +
'<meta charset="UTF-8">\n' +
'<meta name="viewport" content="width=device-width, initial-scale=1.0">\n' +
'<title>Pembayran Region</title>\n' +
'<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>\n' +
'<style>\n' +
'* { margin: 0; padding: 0; box-sizing: border-box; }\n' +
'body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; background: #0f172a; color: #e2e8f0; min-height: 100vh; }\n' +
'.header { background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%); border-bottom: 1px solid #334155; padding: 1rem 2rem; position: sticky; top: 0; z-index: 100; }\n' +
'.header h1 { font-size: 1.5rem; color: #f8fafc; font-weight: 700; }\n' +
'.header p { font-size: 0.85rem; color: #94a3b8; margin-top: 0.25rem; }\n' +
'.container { max-width: 1400px; margin: 0 auto; padding: 1.5rem; }\n' +
'.sheet-tabs { display: flex; flex-wrap: wrap; gap: 0.5rem; margin-bottom: 1.5rem; }\n' +
'.tab { padding: 0.5rem 1rem; background: #1e293b; border: 1px solid #334155; border-radius: 0.5rem; cursor: pointer; font-size: 0.85rem; color: #94a3b8; transition: all 0.2s; white-space: nowrap; }\n' +
'.tab:hover { background: #334155; color: #f8fafc; }\n' +
'.tab.active { background: #3b82f6; border-color: #3b82f6; color: #fff; }\n' +
'.controls { display: flex; flex-wrap: wrap; gap: 1rem; margin-bottom: 1rem; align-items: center; }\n' +
'.search-box { flex: 1; min-width: 250px; }\n' +
'.search-box input { width: 100%; padding: 0.6rem 1rem; background: #1e293b; border: 1px solid #334155; border-radius: 0.5rem; color: #e2e8f0; font-size: 0.9rem; }\n' +
'.search-box input:focus { outline: none; border-color: #3b82f6; }\n' +
'.filter-select { padding: 0.6rem 1rem; background: #1e293b; border: 1px solid #334155; border-radius: 0.5rem; color: #e2e8f0; font-size: 0.9rem; }\n' +
'.stats { display: flex; gap: 1.5rem; font-size: 0.85rem; color: #94a3b8; }\n' +
'.stats span { color: #3b82f6; font-weight: 600; }\n' +
'.table-container { overflow-x: auto; background: #1e293b; border: 1px solid #334155; border-radius: 0.75rem; max-height: 70vh; overflow-y: auto; }\n' +
'table { width: 100%; border-collapse: collapse; font-size: 0.85rem; }\n' +
'thead { position: sticky; top: 0; z-index: 10; }\n' +
'th { background: #334155; color: #f8fafc; padding: 0.75rem 1rem; text-align: left; font-weight: 600; white-space: nowrap; cursor: pointer; user-select: none; border-bottom: 2px solid #475569; }\n' +
'th:hover { background: #475569; }\n' +
'th .sort-icon { margin-left: 0.5rem; opacity: 0.5; }\n' +
'th.sorted .sort-icon { opacity: 1; color: #3b82f6; }\n' +
'td { padding: 0.6rem 1rem; border-bottom: 1px solid #1e293b; color: #cbd5e1; white-space: nowrap; max-width: 300px; overflow: hidden; text-overflow: ellipsis; }\n' +
'tr:hover td { background: #252f42; }\n' +
'.loading { text-align: center; padding: 3rem; color: #94a3b8; }\n' +
'.spinner { display: inline-block; width: 2rem; height: 2rem; border: 3px solid #334155; border-top-color: #3b82f6; border-radius: 50%; animation: spin 0.8s linear infinite; margin-bottom: 1rem; }\n' +
'@keyframes spin { to { transform: rotate(360deg); } }\n' +
'.error { background: #7f1d1d; border: 1px solid #991b1b; color: #fecaca; padding: 1rem; border-radius: 0.5rem; margin: 1rem 0; }\n' +
'.no-data { text-align: center; padding: 3rem; color: #64748b; }\n' +
'.badge { display: inline-block; padding: 0.15rem 0.5rem; border-radius: 9999px; font-size: 0.7rem; font-weight: 600; }\n' +
'.badge-ok { background: #064e3b; color: #6ee7b7; }\n' +
'.badge-waiting { background: #78350f; color: #fcd34d; }\n' +
'.badge-paid { background: #1e3a5f; color: #93c5fd; }\n' +
'</style>\n' +
'</head>\n' +
'<body>\n' +
'<div class="header">\n' +
'  <h1>📊 Pembayran Region</h1>\n' +
'  <p>Data Rekap Pembayaran - Eagle High Plantations</p>\n' +
'</div>\n' +
'<div class="container">\n' +
'  <div class="sheet-tabs" id="sheetTabs"></div>\n' +
'  <div class="controls">\n' +
'    <div class="search-box">\n' +
'      <input type="text" id="searchInput" placeholder="🔍 Cari data...">\n' +
'    </div>\n' +
'    <select class="filter-select" id="filterStatus">\n' +
'      <option value="">Semua Status</option>\n' +
'      <option value="OK">OK</option>\n' +
'      <option value="LUNAS">LUNAS</option>\n' +
'      <option value="Waiting">Waiting</option>\n' +
'    </select>\n' +
'    <div class="stats">\n' +
'      Total: <span id="totalRows">0</span> | \n' +
'      Filtered: <span id="filteredRows">0</span>\n' +
'    </div>\n' +
'  </div>\n' +
'  <div id="content">\n' +
'    <div class="loading"><div class="spinner"></div><br>Loading...</div>\n' +
'  </div>\n' +
'</div>\n' +
'\n' +
'<script>\n' +
'var GITHUB_RAW_URL = "' + GITHUB_RAW_URL + '";\n' +
'var SHEET_DISPLAY = ' + JSON.stringify(SHEET_DISPLAY) + ';\n' +
'var wb = null;\n' +
'var currentSheet = null;\n' +
'var sheetData = { headers: [], rows: [] };\n' +
'var sortCol = -1;\n' +
'var sortAsc = true;\n' +
'\n' +
'// Load Excel on page load\n' +
'window.addEventListener("load", function() {\n' +
'  loadExcel();\n' +
'});\n' +
'\n' +
'function loadExcel() {\n' +
'  var content = document.getElementById("content");\n' +
'  content.innerHTML = \'<div class="loading"><div class="spinner"></div><br>Loading Excel dari GitHub...</div>\';\n' +
'  \n' +
'  fetch(GITHUB_RAW_URL)\n' +
'    .then(function(r) {\n' +
'      if (!r.ok) throw new Error("Gagal download Excel: " + r.status);\n' +
'      return r.arrayBuffer();\n' +
'    })\n' +
'    .then(function(buf) {\n' +
'      wb = XLSX.read(new Uint8Array(buf), { type: "array" });\n' +
'      loadSheetTabs();\n' +
'    })\n' +
'    .catch(function(err) {\n' +
'      content.innerHTML = \'<div class="error">❌ \' + err.message + \'</div>\';\n' +
'    });\n' +
'}\n' +
'\n' +
'function loadSheetTabs() {\n' +
'  var sheets = wb.SheetNames.map(function(name) {\n' +
'    return { name: name, display: SHEET_DISPLAY[name] || name };\n' +
'  });\n' +
'  \n' +
'  var container = document.getElementById("sheetTabs");\n' +
'  container.innerHTML = sheets.map(function(s) {\n' +
'    return \'<div class="tab" data-sheet="\' + s.name + \'">\' + s.display + \'</div>\';\n' +
'  }).join("");\n' +
'  \n' +
'  var tabs = container.querySelectorAll(".tab");\n' +
'  for (var i = 0; i < tabs.length; i++) {\n' +
'    tabs[i].addEventListener("click", function() {\n' +
'      for (var j = 0; j < tabs.length; j++) tabs[j].classList.remove("active");\n' +
'      this.classList.add("active");\n' +
'      loadSheet(this.dataset.sheet);\n' +
'    });\n' +
'  }\n' +
'  \n' +
'  // Auto-select first sheet\n' +
'  if (tabs.length > 0) tabs[0].click();\n' +
'}\n' +
'\n' +
'function loadSheet(name) {\n' +
'  currentSheet = name;\n' +
'  var content = document.getElementById("content");\n' +
'  content.innerHTML = \'<div class="loading"><div class="spinner"></div><br>Loading \' + name + \'...</div>\';\n' +
'  \n' +
'  var ws = wb.Sheets[name];\n' +
'  if (!ws) {\n' +
'    content.innerHTML = \'<div class="error">❌ Sheet "\' + name + \'" tidak ditemukan</div>\';\n' +
'    return;\n' +
'  }\n' +
'  \n' +
'  var json = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });\n' +
'  \n' +
'  // Find header row (first non-empty row with multiple cells)\n' +
'  var headerRow = 0;\n' +
'  for (var i = 0; i < Math.min(json.length, 10); i++) {\n' +
'    var row = json[i].filter(function(c) { return c !== ""; });\n' +
'    if (row.length >= 3) {\n' +
'      headerRow = i;\n' +
'      break;\n' +
'    }\n' +
'  }\n' +
'  \n' +
'  sheetData.headers = json[headerRow].map(function(h, i) {\n' +
'    return (h != null && h.toString().trim()) ? h.toString().trim() : "Col" + (i + 1);\n' +
'  });\n' +
'  sheetData.rows = json.slice(headerRow + 1).filter(function(row) {\n' +
'    return row.some(function(c) { return c !== ""; });\n' +
'  });\n' +
'  sheetData.totalRows = sheetData.rows.length;\n' +
'  \n' +
'  renderTable();\n' +
'}\n' +
'\n' +
'function renderTable() {\n' +
'  var search = document.getElementById("searchInput").value.toLowerCase();\n' +
'  var statusFilter = document.getElementById("filterStatus").value;\n' +
'  \n' +
'  var rows = sheetData.rows.slice();\n' +
'  \n' +
'  // Search filter\n' +
'  if (search) {\n' +
'    rows = rows.filter(function(row) {\n' +
'      return row.some(function(cell) { return String(cell).toLowerCase().includes(search); });\n' +
'    });\n' +
'  }\n' +
'  \n' +
'  // Status filter\n' +
'  if (statusFilter) {\n' +
'    rows = rows.filter(function(row) {\n' +
'      return row.some(function(cell) { return String(cell).toUpperCase().includes(statusFilter.toUpperCase()); });\n' +
'    });\n' +
'  }\n' +
'  \n' +
'  document.getElementById("totalRows").textContent = sheetData.totalRows;\n' +
'  document.getElementById("filteredRows").textContent = rows.length;\n' +
'  \n' +
'  if (rows.length === 0) {\n' +
'    document.getElementById("content").innerHTML = \'<div class="no-data">📭 Data tidak ditemukan</div>\';\n' +
'    return;\n' +
'  }\n' +
'  \n' +
'  // Sort\n' +
'  if (sortCol >= 0) {\n' +
'    rows.sort(function(a, b) {\n' +
'      var va = a[sortCol] != null ? a[sortCol] : "";\n' +
'      var vb = b[sortCol] != null ? b[sortCol] : "";\n' +
'      var na = parseFloat(va);\n' +
'      var nb = parseFloat(vb);\n' +
'      if (!isNaN(na) && !isNaN(nb)) return sortAsc ? na - nb : nb - na;\n' +
'      return sortAsc ? String(va).localeCompare(String(vb)) : String(vb).localeCompare(String(va));\n' +
'    });\n' +
'  }\n' +
'  \n' +
'  var headers = sheetData.headers.map(function(h, i) {\n' +
'    var sorted = i === sortCol ? " sorted" : "";\n' +
'    var icon = i === sortCol ? (sortAsc ? "▲" : "▼") : "⇅";\n' +
'    return \'<th class="\' + sorted + \'" data-col="\' + i + \'">\' + h + \'<span class="sort-icon">\' + icon + \'</span></th>\';\n' +
'  }).join("");\n' +
'  \n' +
'  var body = rows.slice(0, 500).map(function(row) {\n' +
'    var cells = sheetData.headers.map(function(_, i) {\n' +
'      var val = row[i] != null ? row[i] : "";\n' +
'      var badge = "";\n' +
'      var s = String(val).toUpperCase();\n' +
'      if (s === "OK") badge = \'<span class="badge badge-ok">OK</span>\';\n' +
'      else if (s === "LUNAS") badge = \'<span class="badge badge-paid">LUNAS</span>\';\n' +
'      else if (s.indexOf("WAITING") >= 0) badge = \'<span class="badge badge-waiting">Waiting</span>\';\n' +
'      else if (val === "") val = \'<span style="color:#475569">-</span>\';\n' +
'      return \'<td>\' + (badge || val) + \'</td>\';\n' +
'    }).join("");\n' +
'    return \'<tr>\' + cells + \'</tr>\';\n' +
'  }).join("");\n' +
'  \n' +
'  var html = \'<div class="table-container"><table><thead><tr>\' + headers + \'</tr></thead><tbody>\' + body + \'</tbody></table></div>\';\n' +
'  if (rows.length > 500) {\n' +
'    html += \'<p style="text-align:center;padding:1rem;color:#64748b">Menampilkan 500 dari \' + rows.length + \' baris</p>\';\n' +
'  }\n' +
'  \n' +
'  document.getElementById("content").innerHTML = html;\n' +
'  \n' +
'  // Bind sort\n' +
'  var ths = document.querySelectorAll("th");\n' +
'  for (var i = 0; i < ths.length; i++) {\n' +
'    ths[i].addEventListener("click", function() {\n' +
'      var col = parseInt(this.dataset.col);\n' +
'      if (sortCol === col) sortAsc = !sortAsc;\n' +
'      else { sortCol = col; sortAsc = true; }\n' +
'      renderTable();\n' +
'    });\n' +
'  }\n' +
'}\n' +
'\n' +
'document.getElementById("searchInput").addEventListener("input", function() { renderTable(); });\n' +
'document.getElementById("filterStatus").addEventListener("change", function() { renderTable(); });\n' +
'</script>\n' +
'</body>\n' +
'</html>';
