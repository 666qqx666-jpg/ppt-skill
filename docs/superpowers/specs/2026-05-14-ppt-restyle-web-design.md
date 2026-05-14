# PPT Restyle Web 前端设计文档

## 概述

一个纯前端单页应用，允许用户上传 PPT 文件，在浏览器端通过 Pyodide 运行现有 `ppt_restyle.py` 脚本进行样式转换，转换完成后提供下载。部署到 GitHub Pages，无需后端服务器。

**技术栈：** HTML + CSS + JavaScript + Pyodide (python-pptx)
**部署方式：** GitHub Pages（纯静态文件）

## 架构

### 核心思路

通过 Pyodide（WebAssembly 版 Python 运行时）在浏览器中运行现有的 `ppt_restyle.py`，实现与 CLI 版本完全一致的转换逻辑，无需维护两套代码。

### 文件结构

```
web/
├── index.html          # 单页应用（HTML + 内联 CSS/JS）
└── template.pptx       # 固定模板文件（从 templates/template.pptx 复制）
```

部署时只需将 `web/` 目录的内容推送到 GitHub Pages 分支。

### 数据流

```
用户浏览器
  │
  ├─ 1. 用户拖拽/选择 .pptx 文件
  │     → FileReader API 读取为 ArrayBuffer
  │
  ├─ 2. 点击"开始转换"
  │     → 首次：加载 Pyodide（CDN ~15MB）+ 安装 python-pptx
  │     → 后续：直接使用已加载的 Pyodide
  │
  ├─ 3. 写入 Pyodide 虚拟文件系统
  │     → /tmp/source.pptx（用户上传的文件）
  │     → /tmp/template.pptx（通过 fetch 获取的模板文件）
  │
  ├─ 4. 执行 Python 转换
  │     → 调用 ppt_restyle.restyle('/tmp/source.pptx', '/tmp/template.pptx', '/tmp/output.pptx')
  │
  └─ 5. 读取结果
        → 从虚拟文件系统读取 /tmp/output.pptx
        → 生成 Blob → 创建下载链接
```

### Pyodide 集成

**加载时机：** 用户点击"开始转换"时按需加载，避免页面初始加载过慢。

**依赖安装：**
```python
import micropip
await micropip.install('python-pptx')
```

**Python 脚本加载：** 将 `scripts/ppt_restyle.py` 的内容以字符串形式内联到 JavaScript 中，通过 `pyodide.runPythonAsync()` 执行。

**缓存策略：** Pyodide 加载后保存引用，后续转换直接复用，无需重新加载。浏览器自身的 HTTP 缓存也会缓存 Pyodide 的 WASM 文件。

## 页面设计

### 布局

左右分栏布局：
- **左侧（信息区）：** 标题、功能说明、特性列表
- **右侧（操作区）：** 上传区域、操作按钮、状态反馈

移动端自动切换为上下单栏布局。

### UI 状态

页面有 5 个状态，右侧操作区根据状态切换显示内容：

| 状态 | 触发条件 | 右侧显示 |
|------|---------|----------|
| **初始** | 页面加载 | 虚线上传区（拖拽/点击）+ 灰色禁用按钮 |
| **文件已选择** | 选择文件后 | 文件名 + 大小 + "重新选择"链接 + 蓝色激活按钮 |
| **转换中** | 点击转换 | 文件信息 + 进度条 + 状态文字（加载引擎/转换中） |
| **转换完成** | 转换成功 | 绿色成功提示 + 文件名 + 下载按钮 + "继续转换"链接 |
| **转换失败** | 转换出错 | 红色错误提示 + 错误信息 + "重新选择文件"按钮 |

### 转换中的进度提示

由于无法精确追踪 Pyodide 内部进度，使用阶段性文字提示：
1. "正在加载转换引擎..."（Pyodide 加载阶段）
2. "正在安装依赖..."（micropip 安装阶段）
3. "正在转换 PPT..."（restyle 执行阶段）

配合一个不确定进度的动画条。

### 文件验证

上传时在前端做基本校验：
- 文件扩展名必须为 `.pptx`
- 文件大小上限 50MB（防止浏览器内存溢出）

## 关键实现细节

### Python 脚本内联

将 `ppt_restyle.py` 的源码作为 JavaScript 字符串嵌入 `index.html`。转换时通过 Pyodide 执行：

```javascript
await pyodide.runPythonAsync(PPT_RESTYLE_SCRIPT);
await pyodide.runPythonAsync(`
from ppt_restyle import restyle
restyle('/tmp/source.pptx', '/tmp/template.pptx', '/tmp/output.pptx')
`);
```

### 文件系统操作

```javascript
// 写入源文件
pyodide.FS.writeFile('/tmp/source.pptx', new Uint8Array(sourceBuffer));

// 写入模板
const templateBuffer = await fetch('template.pptx').then(r => r.arrayBuffer());
pyodide.FS.writeFile('/tmp/template.pptx', new Uint8Array(templateBuffer));

// 读取结果
const outputData = pyodide.FS.readFile('/tmp/output.pptx');
const blob = new Blob([outputData], { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' });
```

### 下载实现

```javascript
const url = URL.createObjectURL(blob);
const a = document.createElement('a');
a.href = url;
a.download = `${originalName}_styled.pptx`;
// 用户点击下载按钮时触发
```

### 错误处理

捕获 Python 端抛出的异常，提取错误信息显示给用户：
- `FileNotFoundError` → "文件读取失败"
- `ValueError`（模板验证失败）→ "模板文件损坏，请刷新重试"
- 其他异常 → 显示原始错误信息 + 建议检查文件格式

## 响应式设计

- **桌面（≥768px）：** 左右分栏，各占 50%
- **移动端（<768px）：** 上下单栏，信息区在上，操作区在下

## 已知限制

- 首次转换需下载 Pyodide（~15MB），后续使用浏览器缓存
- 大文件（>20MB）可能导致浏览器内存压力
- 不支持 `.ppt`（旧格式），仅支持 `.pptx`
- 离线不可用（Pyodide 从 CDN 加载）

## 依赖

- Pyodide（CDN 加载，无需本地安装）
- python-pptx（通过 micropip 在浏览器端安装）
- 无其他前端框架依赖，纯原生 HTML/CSS/JS
