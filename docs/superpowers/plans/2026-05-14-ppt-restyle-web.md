# PPT Restyle Web Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a single-page web app that converts PPT styles using Pyodide (WASM Python) in the browser, deployable to GitHub Pages with zero backend.

**Architecture:** Pure frontend SPA — user uploads .pptx, Pyodide runs the existing `scripts/ppt_restyle.py` in-browser via its virtual filesystem, outputs a styled .pptx for download. The Python script is written to Pyodide's FS and imported as a module, reusing the exact same conversion logic as the CLI version.

**Tech Stack:** HTML + CSS + JavaScript + Pyodide CDN + python-pptx (via micropip)

---

## File Structure

```
web/
├── index.html          # Single-page app (HTML + inline CSS/JS, ~400 lines)
└── template.pptx       # Fixed template (copied from templates/template.pptx, 18MB)
```

Runtime dependency: `scripts/ppt_restyle.py` (563 lines) — its content is embedded as a JS template literal in `index.html` and written to Pyodide's virtual filesystem at conversion time.

---

### Task 1: Complete UI with Mock Conversion

Build the full single-page app with all 5 UI states, drag-and-drop file upload, validation, responsive layout, and a mock conversion function to verify all state transitions before integrating Pyodide.

**Files:**
- Create: `web/index.html`

- [ ] **Step 1: Create `web/` directory**

```bash
mkdir -p web
```

- [ ] **Step 2: Create `web/index.html` with complete HTML + CSS + JS**

Write `web/index.html` with the following content. This is the full app shell — all CSS is inline in `<style>`, all JS is inline in `<script>`. The conversion function is a mock (setTimeout) that will be replaced in Task 2.

```html
<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>PPT Restyle</title>
  <style>
    *, *::before, *::after { margin: 0; padding: 0; box-sizing: border-box; }
    :root {
      --primary: #4F46E5;
      --primary-hover: #4338CA;
      --success: #059669;
      --success-bg: #ECFDF5;
      --error: #DC2626;
      --error-bg: #FEF2F2;
      --gray-50: #F8FAFC;
      --gray-100: #F1F5F9;
      --gray-200: #E2E8F0;
      --gray-300: #CBD5E1;
      --gray-400: #94A3B8;
      --gray-500: #64748B;
      --gray-900: #0F172A;
    }
    body {
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
      background: var(--gray-100);
      min-height: 100vh;
    }
    .container { display: flex; min-height: 100vh; }
    .info-panel {
      flex: 1; background: linear-gradient(135deg, #1E1B4B, #312E81);
      color: #fff; padding: 4rem 3rem;
      display: flex; flex-direction: column; justify-content: center;
    }
    .info-panel h1 { font-size: 2.5rem; font-weight: 700; margin-bottom: 0.5rem; }
    .info-panel .subtitle { font-size: 1.25rem; opacity: 0.8; margin-bottom: 2rem; }
    .info-panel .description { line-height: 1.6; opacity: 0.9; margin-bottom: 2rem; }
    .features { list-style: none; }
    .features li {
      padding: 0.5rem 0; padding-left: 1.5rem;
      position: relative; opacity: 0.9;
    }
    .features li::before {
      content: "\2713"; position: absolute; left: 0; color: #818CF8;
    }
    .action-panel {
      flex: 1; display: flex; align-items: center;
      justify-content: center; padding: 2rem;
    }
    .card {
      background: #fff; border-radius: 1rem; padding: 2.5rem;
      width: 100%; max-width: 440px; box-shadow: 0 1px 3px rgba(0,0,0,0.1);
    }
    .upload-zone {
      border: 2px dashed var(--gray-300); border-radius: 0.75rem;
      padding: 3rem 2rem; text-align: center; cursor: pointer;
      transition: border-color 0.2s, background 0.2s;
    }
    .upload-zone:hover, .upload-zone.dragover {
      border-color: var(--primary); background: #EEF2FF;
    }
    .upload-zone .icon { font-size: 2.5rem; margin-bottom: 1rem; color: var(--gray-400); }
    .upload-zone p { color: var(--gray-500); margin-bottom: 0.25rem; }
    .upload-zone .hint { font-size: 0.8rem; color: var(--gray-400); margin-top: 0.75rem; }
    .link { color: var(--primary); cursor: pointer; text-decoration: underline; }
    .btn {
      display: block; width: 100%; padding: 0.875rem; border: none;
      border-radius: 0.5rem; font-size: 1rem; font-weight: 600;
      cursor: pointer; margin-top: 1.5rem; transition: background 0.2s;
    }
    .btn:disabled { opacity: 0.5; cursor: not-allowed; }
    .btn-primary { background: var(--primary); color: #fff; }
    .btn-primary:hover:not(:disabled) { background: var(--primary-hover); }
    .btn-success { background: var(--success); color: #fff; }
    .btn-success:hover { background: #047857; }
    .file-info {
      display: flex; align-items: center; gap: 0.75rem;
      padding: 1rem; background: var(--gray-50); border-radius: 0.5rem;
    }
    .file-icon { font-size: 2rem; }
    .file-info .details { flex: 1; min-width: 0; }
    .file-name {
      font-weight: 600; color: var(--gray-900);
      overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
    }
    .file-size { font-size: 0.85rem; color: var(--gray-500); }
    .file-info .action { white-space: nowrap; font-size: 0.85rem; }
    .progress-container {
      height: 6px; background: var(--gray-200);
      border-radius: 3px; overflow: hidden; margin-top: 1.5rem;
    }
    .progress-bar {
      height: 100%; background: var(--primary);
      border-radius: 3px; animation: slide 1.5s ease-in-out infinite;
    }
    @keyframes slide {
      0% { width: 20%; margin-left: 0; }
      50% { width: 40%; margin-left: 30%; }
      100% { width: 20%; margin-left: 80%; }
    }
    .status-text {
      text-align: center; color: var(--gray-500);
      margin-top: 1rem; font-size: 0.9rem;
    }
    .result-icon {
      width: 4rem; height: 4rem; border-radius: 50%;
      display: flex; align-items: center; justify-content: center;
      font-size: 1.5rem; font-weight: 700; margin: 0 auto 1rem;
    }
    .result-icon.success { background: var(--success-bg); color: var(--success); }
    .result-icon.error { background: var(--error-bg); color: var(--error); }
    .result-title {
      text-align: center; font-size: 1.25rem;
      font-weight: 600; color: var(--gray-900); margin-bottom: 0.5rem;
    }
    .result-detail {
      text-align: center; color: var(--gray-500);
      font-size: 0.9rem; margin-bottom: 1rem;
    }
    .error-box {
      color: var(--error); font-size: 0.85rem;
      padding: 0.75rem; background: var(--error-bg);
      border-radius: 0.5rem; word-break: break-word; margin-bottom: 0.5rem;
    }
    .link-center {
      display: block; text-align: center; margin-top: 1rem;
      color: var(--primary); cursor: pointer;
      text-decoration: underline; font-size: 0.9rem;
    }
    @media (max-width: 768px) {
      .container { flex-direction: column; }
      .info-panel { padding: 2rem 1.5rem; }
      .info-panel h1 { font-size: 1.75rem; }
      .action-panel { padding: 1.5rem; }
      .card { padding: 1.5rem; }
    }
  </style>
</head>
<body>
  <div class="container">
    <div class="info-panel">
      <h1>PPT Restyle</h1>
      <p class="subtitle">一键转换 PPT 样式</p>
      <p class="description">上传你的 PPT 文件，自动应用精美模板样式。所有内容原样保留，仅替换背景和页面布局。</p>
      <ul class="features">
        <li>保留所有文字、图表、动画</li>
        <li>自动适配模板布局和比例</li>
        <li>浏览器端处理，文件不上传到服务器</li>
        <li>支持 .pptx 格式，最大 50MB</li>
      </ul>
    </div>
    <div class="action-panel">
      <div class="card">
        <div id="state-initial">
          <div class="upload-zone" id="upload-zone">
            <div class="icon">&#128193;</div>
            <p>拖拽 PPT 文件到此处</p>
            <p>或 <span class="link">点击选择文件</span></p>
            <p class="hint">支持 .pptx 格式，最大 50MB</p>
          </div>
        </div>
        <div id="state-selected" hidden>
          <div class="file-info">
            <span class="file-icon">&#128196;</span>
            <div class="details">
              <p class="file-name" id="file-name"></p>
              <p class="file-size" id="file-size"></p>
            </div>
            <span class="action link" id="btn-reselect">重新选择</span>
          </div>
          <button class="btn btn-primary" id="btn-start">开始转换</button>
        </div>
        <div id="state-converting" hidden>
          <div class="file-info">
            <span class="file-icon">&#128196;</span>
            <div class="details">
              <p class="file-name" id="converting-name"></p>
              <p class="file-size" id="converting-size"></p>
            </div>
          </div>
          <div class="progress-container"><div class="progress-bar"></div></div>
          <p class="status-text" id="status-text">正在加载转换引擎...</p>
        </div>
        <div id="state-complete" hidden>
          <div class="result-icon success">&#10003;</div>
          <p class="result-title">转换完成</p>
          <p class="result-detail" id="output-name"></p>
          <button class="btn btn-success" id="btn-download">下载文件</button>
          <span class="link-center" id="btn-another">继续转换下一个</span>
        </div>
        <div id="state-error" hidden>
          <div class="result-icon error">&#10005;</div>
          <p class="result-title">转换失败</p>
          <p class="error-box" id="error-message"></p>
          <button class="btn btn-primary" id="btn-retry">重新选择文件</button>
        </div>
        <input type="file" id="file-input" accept=".pptx" hidden>
      </div>
    </div>
  </div>
  <script>
    const $ = id => document.getElementById(id);
    const STATES = ['initial', 'selected', 'converting', 'complete', 'error'];
    let selectedFile = null;
    let downloadUrl = null;

    function setState(name) {
      STATES.forEach(s => $('state-' + s).hidden = (s !== name));
    }

    function formatSize(bytes) {
      if (bytes < 1024) return bytes + ' B';
      if (bytes < 1048576) return (bytes / 1024).toFixed(1) + ' KB';
      return (bytes / 1048576).toFixed(1) + ' MB';
    }

    function handleFile(file) {
      if (!file) return;
      if (!file.name.toLowerCase().endsWith('.pptx')) {
        return alert('请选择 .pptx 格式的文件');
      }
      if (file.size > 50 * 1024 * 1024) {
        return alert('文件大小不能超过 50MB');
      }
      selectedFile = file;
      $('file-name').textContent = file.name;
      $('file-size').textContent = formatSize(file.size);
      setState('selected');
    }

    function resetToInitial() {
      $('file-input').value = '';
      selectedFile = null;
      if (downloadUrl) { URL.revokeObjectURL(downloadUrl); downloadUrl = null; }
      setState('initial');
    }

    // Upload zone interactions
    const zone = $('upload-zone');
    zone.addEventListener('click', () => $('file-input').click());
    zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('dragover'); });
    zone.addEventListener('dragleave', () => zone.classList.remove('dragover'));
    zone.addEventListener('drop', e => {
      e.preventDefault(); zone.classList.remove('dragover');
      handleFile(e.dataTransfer.files[0]);
    });
    $('file-input').addEventListener('change', e => handleFile(e.target.files[0]));

    // Button handlers
    $('btn-reselect').addEventListener('click', resetToInitial);
    $('btn-retry').addEventListener('click', resetToInitial);
    $('btn-another').addEventListener('click', resetToInitial);
    $('btn-start').addEventListener('click', startConversion);

    // === MOCK CONVERSION (replaced with Pyodide in Task 2) ===
    async function startConversion() {
      $('converting-name').textContent = selectedFile.name;
      $('converting-size').textContent = formatSize(selectedFile.size);
      setState('converting');
      try {
        $('status-text').textContent = '正在加载转换引擎...';
        await new Promise(r => setTimeout(r, 1000));
        $('status-text').textContent = '正在安装依赖...';
        await new Promise(r => setTimeout(r, 1000));
        $('status-text').textContent = '正在转换 PPT...';
        await new Promise(r => setTimeout(r, 1000));

        const outputName = selectedFile.name.replace(/\.pptx$/i, '_styled.pptx');
        $('output-name').textContent = outputName;
        $('btn-download').onclick = () => alert('Mock download: ' + outputName);
        setState('complete');
      } catch (err) {
        $('error-message').textContent = err.message || String(err);
        setState('error');
      }
    }
  </script>
</body>
</html>
```

- [ ] **Step 3: Open in browser and verify all 5 states**

```bash
open web/index.html
```

Verify checklist:
- Left panel: title "PPT Restyle", subtitle, description, 4 feature items with checkmarks
- Right panel: dashed upload zone with folder icon
- **Initial state:** upload zone visible, no button
- **Drag-and-drop:** dragging file over zone → border turns blue, background tints
- **Click upload:** clicking zone opens file picker
- **Validation:** selecting a `.pdf` → alert "请选择 .pptx 格式的文件"
- **File selected:** selecting a valid `.pptx` → file name + size + "重新选择" link + blue "开始转换" button
- **Converting:** click "开始转换" → progress bar animates, status text cycles through 3 stages
- **Complete:** mock success → green checkmark + "转换完成" + download button + "继续转换下一个"
- **Reset:** "继续转换下一个" or "重新选择" → back to initial state
- **Responsive:** resize to <768px → single column (info above, action below)

- [ ] **Step 4: Commit**

```bash
git add web/index.html
git commit -m "feat: add PPT Restyle web UI with all states and mock conversion"
```

---

### Task 2: Pyodide Integration and Real Conversion

Replace the mock conversion with actual Pyodide execution. The Python script is written to Pyodide's virtual filesystem and imported as a module.

**Files:**
- Copy: `templates/template.pptx` → `web/template.pptx`
- Modify: `web/index.html` (replace `<script>` block)
- Reference: `scripts/ppt_restyle.py` (content embedded as JS string)

- [ ] **Step 1: Copy template file to web/**

```bash
cp templates/template.pptx web/template.pptx
```

- [ ] **Step 2: Replace the `<script>` block in `web/index.html`**

Replace everything between `<script>` and `</script>` with the following. Key changes from the mock version:
1. `PPT_RESTYLE_SCRIPT` constant holds the full Python source
2. `loadPyodideEngine()` dynamically loads Pyodide from CDN
3. `ensureEngine()` handles one-time Pyodide + micropip setup
4. `startConversion()` uses real file I/O through Pyodide's virtual filesystem

```javascript
    const $ = id => document.getElementById(id);
    const STATES = ['initial', 'selected', 'converting', 'complete', 'error'];
    let selectedFile = null;
    let downloadUrl = null;
    let pyodide = null;
    let engineReady = false;

    // Full content of scripts/ppt_restyle.py embedded as a JS string.
    // Read scripts/ppt_restyle.py and paste its complete content here.
    // Safe to use template literal: the Python source contains zero backticks
    // and zero ${ sequences.
    const PPT_RESTYLE_SCRIPT = `
<PASTE FULL CONTENT OF scripts/ppt_restyle.py HERE — all 563 lines>
`;

    function setState(name) {
      STATES.forEach(s => $('state-' + s).hidden = (s !== name));
    }

    function formatSize(bytes) {
      if (bytes < 1024) return bytes + ' B';
      if (bytes < 1048576) return (bytes / 1024).toFixed(1) + ' KB';
      return (bytes / 1048576).toFixed(1) + ' MB';
    }

    function handleFile(file) {
      if (!file) return;
      if (!file.name.toLowerCase().endsWith('.pptx')) {
        return alert('请选择 .pptx 格式的文件');
      }
      if (file.size > 50 * 1024 * 1024) {
        return alert('文件大小不能超过 50MB');
      }
      selectedFile = file;
      $('file-name').textContent = file.name;
      $('file-size').textContent = formatSize(file.size);
      setState('selected');
    }

    function resetToInitial() {
      $('file-input').value = '';
      selectedFile = null;
      if (downloadUrl) { URL.revokeObjectURL(downloadUrl); downloadUrl = null; }
      setState('initial');
    }

    // Upload zone interactions
    const zone = $('upload-zone');
    zone.addEventListener('click', () => $('file-input').click());
    zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('dragover'); });
    zone.addEventListener('dragleave', () => zone.classList.remove('dragover'));
    zone.addEventListener('drop', e => {
      e.preventDefault(); zone.classList.remove('dragover');
      handleFile(e.dataTransfer.files[0]);
    });
    $('file-input').addEventListener('change', e => handleFile(e.target.files[0]));

    // Button handlers
    $('btn-reselect').addEventListener('click', resetToInitial);
    $('btn-retry').addEventListener('click', resetToInitial);
    $('btn-another').addEventListener('click', resetToInitial);
    $('btn-start').addEventListener('click', startConversion);

    // Pyodide engine loading — on-demand from CDN
    async function loadPyodideEngine() {
      return new Promise((resolve, reject) => {
        const script = document.createElement('script');
        script.src = 'https://cdn.jsdelivr.net/pyodide/v0.26.4/full/pyodide.js';
        script.onload = async () => {
          try { resolve(await loadPyodide()); }
          catch (e) { reject(e); }
        };
        script.onerror = () => reject(new Error('Pyodide 加载失败，请检查网络连接'));
        document.head.appendChild(script);
      });
    }

    // One-time setup: load Pyodide + install python-pptx + write script file
    async function ensureEngine() {
      if (engineReady) return;

      $('status-text').textContent = '正在加载转换引擎...';
      pyodide = await loadPyodideEngine();

      $('status-text').textContent = '正在安装依赖...';
      await pyodide.loadPackage('micropip');
      await pyodide.runPythonAsync(`
import micropip
await micropip.install('python-pptx')
      `);

      pyodide.FS.writeFile('ppt_restyle.py', PPT_RESTYLE_SCRIPT);
      engineReady = true;
    }

    async function startConversion() {
      $('converting-name').textContent = selectedFile.name;
      $('converting-size').textContent = formatSize(selectedFile.size);
      setState('converting');

      try {
        await ensureEngine();

        $('status-text').textContent = '正在转换 PPT...';

        // Write source file to virtual FS
        const sourceBuffer = await selectedFile.arrayBuffer();
        pyodide.FS.writeFile('/tmp/source.pptx', new Uint8Array(sourceBuffer));

        // Fetch and write template to virtual FS
        const templateResp = await fetch('template.pptx');
        if (!templateResp.ok) throw new Error('模板文件加载失败，请刷新重试');
        const templateBuffer = await templateResp.arrayBuffer();
        pyodide.FS.writeFile('/tmp/template.pptx', new Uint8Array(templateBuffer));

        // Run conversion
        await pyodide.runPythonAsync(`
from ppt_restyle import restyle
restyle('/tmp/source.pptx', '/tmp/template.pptx', '/tmp/output.pptx')
        `);

        // Read output and create download blob
        const outputData = pyodide.FS.readFile('/tmp/output.pptx');
        const blob = new Blob([outputData], {
          type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
        });

        if (downloadUrl) URL.revokeObjectURL(downloadUrl);
        downloadUrl = URL.createObjectURL(blob);

        const outputName = selectedFile.name.replace(/\.pptx$/i, '_styled.pptx');
        $('output-name').textContent = outputName;
        $('btn-download').onclick = () => {
          const a = document.createElement('a');
          a.href = downloadUrl;
          a.download = outputName;
          document.body.appendChild(a);
          a.click();
          document.body.removeChild(a);
        };

        setState('complete');
      } catch (err) {
        let msg = err.message || String(err);
        if (msg.includes('FileNotFoundError')) msg = '文件读取失败，请重试';
        else if (msg.includes('ValueError') && msg.includes('模板'))
          msg = '模板文件损坏，请刷新页面重试';
        $('error-message').textContent = msg;
        setState('error');
      }
    }
```

**Important:** In the `PPT_RESTYLE_SCRIPT` constant, paste the **complete** content of `scripts/ppt_restyle.py` (all 563 lines). The raw Python source is safe inside a JS template literal — verified zero backticks and zero `${` patterns.

- [ ] **Step 3: Start a local server and test real conversion**

`fetch('template.pptx')` requires HTTP (not `file://`), so serve with a local server:

```bash
cd /path/to/project/web && python3 -m http.server 8080
```

Open `http://localhost:8080` in browser. Upload a real `.pptx` file and verify:
1. "正在加载转换引擎..." appears — Pyodide downloads (~15MB, first time only)
2. "正在安装依赖..." appears — micropip installs python-pptx + lxml
3. "正在转换 PPT..." appears — restyle() executes
4. Green success screen with download button
5. Click "下载文件" → downloads `{name}_styled.pptx`
6. Open downloaded file in PowerPoint/Keynote → template styling applied, content preserved

- [ ] **Step 4: Commit**

```bash
git add web/index.html web/template.pptx
git commit -m "feat: integrate Pyodide for real PPT conversion in browser"
```

---

### Task 3: End-to-End Verification

Test the complete workflow including error paths, engine caching, and responsive layout.

**Files:** None (testing only, commit if fixes needed)

- [ ] **Step 1: Test happy path**

```bash
cd /path/to/project/web && python3 -m http.server 8080
```

At `http://localhost:8080`:
1. Drag a `.pptx` file onto upload zone → file info appears
2. Click "开始转换" → progress through all stages → success
3. Click "下载文件" → `.pptx` downloads
4. Open in PowerPoint/Keynote → template styling applied, all content preserved
5. Click "继续转换下一个" → returns to initial state

- [ ] **Step 2: Test validation and error cases**

1. Upload a `.pdf` → alert "请选择 .pptx 格式的文件"
2. Upload a `.txt` renamed to `.doc` → alert (not `.pptx`)
3. Click "重新选择" in file-selected state → returns to initial
4. Click "重新选择文件" in error state → returns to initial

- [ ] **Step 3: Test engine caching (second conversion)**

After one successful conversion:
1. Click "继续转换下一个"
2. Upload another `.pptx`
3. Click "开始转换"
4. Verify status skips straight to "正在转换 PPT..." (engine already loaded)
5. Conversion completes faster than first time

- [ ] **Step 4: Test responsive layout**

Open browser DevTools → toggle device toolbar:
1. Mobile (375px) → single column: info panel above, action panel below
2. Tablet (768px) → switches to two-column at this breakpoint
3. Desktop (1200px) → comfortable two-column layout

- [ ] **Step 5: Commit fixes (if any)**

If any issues were found and fixed in Steps 1-4:

```bash
git add web/index.html
git commit -m "fix: address issues from E2E testing"
```
