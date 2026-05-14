# PPT 样式重塑 实现计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 实现一个 Claude Code skill + Python 脚本，将用户 PPT 的内容迁移到模板 PPT 的样式中。

**Architecture:** 以模板文件为基底（shutil.copy），在副本内复制模板幻灯片，再将源 PPT 内容按坐标映射迁移进去。单文件脚本 `ppt_restyle.py`，通过 Skill 定义触发。

**Tech Stack:** Python 3.8+, python-pptx, lxml, pytest

---

## 文件结构

```
/Users/qqx/my_code_cursor/lq/
├── scripts/
│   └── ppt_restyle.py        # 核心脚本（所有逻辑）
├── tests/
│   ├── conftest.py            # 测试夹具（程序化生成测试 PPT）
│   └── test_ppt_restyle.py    # 所有测试
├── templates/                  # 模板存放目录
├── output/                     # 输出目录
├── requirements.txt
└── pytest.ini

~/.claude/skills/ppt-restyle/
├── SKILL.md
└── references/
    └── restyle-guide.md
```

---

### Task 1: 项目脚手架

**Files:**
- Create: `/Users/qqx/my_code_cursor/lq/requirements.txt`
- Create: `/Users/qqx/my_code_cursor/lq/pytest.ini`

- [ ] **Step 1: 创建目录结构**

```bash
mkdir -p /Users/qqx/my_code_cursor/lq/{scripts,tests,templates,output}
```

- [ ] **Step 2: 创建 requirements.txt**

```
python-pptx>=0.6.21
pytest>=7.0
```

- [ ] **Step 3: 创建 pytest.ini**

```ini
[pytest]
testpaths = tests
python_files = test_*.py
python_functions = test_*
```

- [ ] **Step 4: 安装依赖**

```bash
pip install python-pptx pytest
```

Expected: 安装成功，无报错

- [ ] **Step 5: 验证安装**

```bash
python3 -c "from pptx import Presentation; print('OK')"
```

Expected: 输出 `OK`

---

### Task 2: 测试夹具

**Files:**
- Create: `/Users/qqx/my_code_cursor/lq/tests/conftest.py`

- [ ] **Step 1: 创建 conftest.py**

```python
import pytest
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor


@pytest.fixture
def template_pptx(tmp_path):
    """生成一个标准模板 PPT：封面页 + 内容页。"""
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    # 第1页：封面
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.placeholders[0].text = "模板标题"
    bg = slide.background.fill
    bg.solid()
    bg.fore_color.rgb = RGBColor(0x00, 0x33, 0x66)

    # 第2页：内容页
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.placeholders[0].text = "页面标题"
    slide.placeholders[1].text = "页面内容"
    bg = slide.background.fill
    bg.solid()
    bg.fore_color.rgb = RGBColor(0xF0, 0xF0, 0xF0)

    path = tmp_path / "template.pptx"
    prs.save(str(path))
    return str(path)


@pytest.fixture
def source_pptx(tmp_path):
    """生成一个包含混合内容的源 PPT：封面 + 文本页 + 表格页。"""
    prs = Presentation()

    # 第1页：封面
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.placeholders[0].text = "我的演示文稿"

    # 第2页：文本内容
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.placeholders[0].text = "第一节标题"
    slide.placeholders[1].text = "这是第一节的内容"
    txBox = slide.shapes.add_textbox(Inches(1), Inches(3), Inches(5), Inches(1))
    txBox.text_frame.text = "额外文本框内容"

    # 第3页：表格
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.placeholders[0].text = "数据表格"
    table_shape = slide.shapes.add_table(
        2, 3, Inches(1), Inches(2), Inches(8), Inches(2)
    )
    tbl = table_shape.table
    for row_idx, row_data in enumerate([["A", "B", "C"], ["1", "2", "3"]]):
        for col_idx, val in enumerate(row_data):
            tbl.cell(row_idx, col_idx).text = val

    path = tmp_path / "source.pptx"
    prs.save(str(path))
    return str(path)


@pytest.fixture
def source_single_slide_pptx(tmp_path):
    """仅有封面的源 PPT。"""
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.placeholders[0].text = "仅封面"
    path = tmp_path / "single.pptx"
    prs.save(str(path))
    return str(path)
```

- [ ] **Step 2: 验证夹具可用**

```bash
cd /Users/qqx/my_code_cursor/lq && python3 -c "
from tests.conftest import *
import tempfile, pathlib
tmp = pathlib.Path(tempfile.mkdtemp())
# 调用 fixture 函数（不通过 pytest，直接验证）
from pptx import Presentation
p = template_pptx(tmp)
prs = Presentation(p)
assert len(prs.slides) == 2
print('夹具 OK')
"
```

Expected: `夹具 OK`

- [ ] **Step 3: 提交**

```bash
git add scripts/ tests/ templates/ output/ requirements.txt pytest.ini
git commit -m "chore: 项目脚手架和测试夹具"
```

---

### Task 3: 模板校验（TDD）

**Files:**
- Create: `/Users/qqx/my_code_cursor/lq/scripts/ppt_restyle.py`
- Modify: `/Users/qqx/my_code_cursor/lq/tests/test_ppt_restyle.py`

- [ ] **Step 1: 写失败的测试**

创建 `tests/test_ppt_restyle.py`：

```python
import pytest
from pptx import Presentation
from pptx.util import Inches

from scripts.ppt_restyle import validate_template


class TestValidateTemplate:
    def test_valid_template(self, template_pptx):
        prs = Presentation(template_pptx)
        validate_template(prs)

    def test_rejects_single_slide(self, tmp_path):
        prs = Presentation()
        prs.slides.add_slide(prs.slide_layouts[0])
        path = tmp_path / "bad.pptx"
        prs.save(str(path))
        with pytest.raises(ValueError, match="2 页"):
            validate_template(Presentation(str(path)))

    def test_rejects_three_slides(self, tmp_path):
        prs = Presentation()
        for _ in range(3):
            prs.slides.add_slide(prs.slide_layouts[0])
        path = tmp_path / "bad.pptx"
        prs.save(str(path))
        with pytest.raises(ValueError, match="2 页"):
            validate_template(Presentation(str(path)))
```

- [ ] **Step 2: 运行测试，确认失败**

```bash
cd /Users/qqx/my_code_cursor/lq && python3 -m pytest tests/test_ppt_restyle.py::TestValidateTemplate -v
```

Expected: FAIL — `ModuleNotFoundError` 或 `ImportError`

- [ ] **Step 3: 实现最小代码**

创建 `scripts/ppt_restyle.py`（初始内容）：

```python
import sys
import os
import shutil
from copy import deepcopy
from io import BytesIO
from pathlib import Path

from pptx import Presentation
from pptx.util import Emu
from pptx.enum.shapes import MSO_SHAPE_TYPE

NSMAP = {
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}

SHAPE_TAGS = {'sp', 'pic', 'graphicFrame', 'grpSp', 'cxnSp'}


def _find_title_placeholder(slide):
    for ph in slide.placeholders:
        if ph.placeholder_format.idx == 0:
            return ph
    return None


def _find_content_placeholder(slide):
    for ph in slide.placeholders:
        if ph.placeholder_format.idx == 1:
            return ph
    return None


def validate_template(prs):
    if len(prs.slides) != 2:
        raise ValueError(f"模板必须恰好有 2 页幻灯片，当前有 {len(prs.slides)} 页")

    cover = prs.slides[0]
    if not _find_title_placeholder(cover):
        raise ValueError("模板封面页（第1页）缺少标题占位符")

    content = prs.slides[1]
    if not _find_title_placeholder(content):
        raise ValueError("模板内容页（第2页）缺少标题占位符")
    if not _find_content_placeholder(content):
        raise ValueError("模板内容页（第2页）缺少内容占位符")
```

- [ ] **Step 4: 运行测试，确认通过**

```bash
cd /Users/qqx/my_code_cursor/lq && python3 -m pytest tests/test_ppt_restyle.py::TestValidateTemplate -v
```

Expected: 3 passed

- [ ] **Step 5: 提交**

```bash
git add scripts/ppt_restyle.py tests/test_ppt_restyle.py
git commit -m "feat: 模板校验逻辑"
```

---

### Task 4: 幻灯片复制与删除（TDD）

**Files:**
- Modify: `/Users/qqx/my_code_cursor/lq/scripts/ppt_restyle.py`
- Modify: `/Users/qqx/my_code_cursor/lq/tests/test_ppt_restyle.py`

- [ ] **Step 1: 写失败的测试**

在 `tests/test_ppt_restyle.py` 追加：

```python
from scripts.ppt_restyle import duplicate_slide, remove_slide


class TestSlideOperations:
    def test_duplicate_slide_adds_one(self, template_pptx):
        prs = Presentation(template_pptx)
        assert len(prs.slides) == 2
        duplicate_slide(prs, 1)
        assert len(prs.slides) == 3

    def test_duplicate_preserves_shapes(self, template_pptx):
        prs = Presentation(template_pptx)
        original = prs.slides[1]
        original_count = len(original.shapes)
        new_slide = duplicate_slide(prs, 1)
        assert len(new_slide.shapes) == original_count

    def test_remove_slide(self, template_pptx):
        prs = Presentation(template_pptx)
        duplicate_slide(prs, 1)
        assert len(prs.slides) == 3
        remove_slide(prs, 2)
        assert len(prs.slides) == 2
```

- [ ] **Step 2: 运行测试，确认失败**

```bash
cd /Users/qqx/my_code_cursor/lq && python3 -m pytest tests/test_ppt_restyle.py::TestSlideOperations -v
```

Expected: FAIL — `ImportError`

- [ ] **Step 3: 实现**

在 `scripts/ppt_restyle.py` 追加：

```python
def _tag_local(element):
    tag = element.tag
    return tag.split('}')[-1] if '}' in tag else tag


def _remap_rids(element, rId_map):
    r_ns = NSMAP['r']
    attrs = [f'{{{r_ns}}}embed', f'{{{r_ns}}}link', f'{{{r_ns}}}id']
    for el in element.iter():
        for attr in attrs:
            val = el.get(attr)
            if val and val in rId_map:
                el.set(attr, rId_map[val])


def _copy_background(src_slide, dst_slide):
    src_cSld = src_slide._element.find('p:cSld', NSMAP)
    dst_cSld = dst_slide._element.find('p:cSld', NSMAP)
    if src_cSld is None or dst_cSld is None:
        return
    src_bg = src_cSld.find('p:bg', NSMAP)
    if src_bg is not None:
        dst_bg = dst_cSld.find('p:bg', NSMAP)
        if dst_bg is not None:
            dst_cSld.remove(dst_bg)
        dst_cSld.insert(0, deepcopy(src_bg))


def duplicate_slide(prs, slide_index):
    src_slide = prs.slides[slide_index]
    new_slide = prs.slides.add_slide(src_slide.slide_layout)

    dst_spTree = new_slide.shapes._spTree

    for child in list(dst_spTree):
        if _tag_local(child) in SHAPE_TAGS:
            dst_spTree.remove(child)

    rId_map = {}
    for rId, rel in src_slide.part.rels.items():
        if not rel.is_external:
            new_rId = new_slide.part.relate_to(rel.target_part, rel.reltype)
            rId_map[rId] = new_rId

    for child in src_slide.shapes._spTree:
        if _tag_local(child) in SHAPE_TAGS:
            new_child = deepcopy(child)
            _remap_rids(new_child, rId_map)
            dst_spTree.append(new_child)

    _copy_background(src_slide, new_slide)
    return new_slide


def remove_slide(prs, slide_index):
    rId_lst = prs.slides._sldIdLst
    sldId = rId_lst[slide_index]
    rId = sldId.get(f'{{{NSMAP["r"]}}}id')
    prs.part.drop_rel(rId)
    rId_lst.remove(sldId)
```

- [ ] **Step 4: 运行测试，确认通过**

```bash
cd /Users/qqx/my_code_cursor/lq && python3 -m pytest tests/test_ppt_restyle.py::TestSlideOperations -v
```

Expected: 3 passed

- [ ] **Step 5: 提交**

```bash
git add scripts/ppt_restyle.py tests/test_ppt_restyle.py
git commit -m "feat: 幻灯片复制与删除"
```

---

### Task 5: 标题操作（TDD）

**Files:**
- Modify: `/Users/qqx/my_code_cursor/lq/scripts/ppt_restyle.py`
- Modify: `/Users/qqx/my_code_cursor/lq/tests/test_ppt_restyle.py`

- [ ] **Step 1: 写失败的测试**

在 `tests/test_ppt_restyle.py` 追加：

```python
from scripts.ppt_restyle import get_title_text, set_title_text


class TestTitleOperations:
    def test_get_title_text(self, template_pptx):
        prs = Presentation(template_pptx)
        assert get_title_text(prs.slides[0]) == "模板标题"

    def test_get_title_text_empty(self, tmp_path):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank layout
        assert get_title_text(slide) == ""

    def test_set_title_text(self, template_pptx):
        prs = Presentation(template_pptx)
        set_title_text(prs.slides[0], "新标题")
        assert get_title_text(prs.slides[0]) == "新标题"
```

- [ ] **Step 2: 运行测试，确认失败**

```bash
cd /Users/qqx/my_code_cursor/lq && python3 -m pytest tests/test_ppt_restyle.py::TestTitleOperations -v
```

Expected: FAIL

- [ ] **Step 3: 实现**

在 `scripts/ppt_restyle.py` 追加：

```python
def get_title_text(slide):
    ph = _find_title_placeholder(slide)
    if ph and ph.has_text_frame:
        return ph.text_frame.text
    for shape in slide.shapes:
        if shape.has_text_frame and shape.name.lower().startswith('title'):
            return shape.text_frame.text
    return ""


def set_title_text(slide, text):
    ph = _find_title_placeholder(slide)
    if not ph or not ph.has_text_frame:
        return
    tf = ph.text_frame
    if tf.paragraphs:
        para = tf.paragraphs[0]
        if para.runs:
            para.runs[0].text = text
            for run in para.runs[1:]:
                run._r.getparent().remove(run._r)
        else:
            para.text = text
        for p in tf.paragraphs[1:]:
            p._p.getparent().remove(p._p)
    else:
        tf.text = text
```

- [ ] **Step 4: 运行测试，确认通过**

```bash
cd /Users/qqx/my_code_cursor/lq && python3 -m pytest tests/test_ppt_restyle.py::TestTitleOperations -v
```

Expected: 3 passed

- [ ] **Step 5: 提交**

```bash
git add scripts/ppt_restyle.py tests/test_ppt_restyle.py
git commit -m "feat: 标题读取与设置"
```

---

### Task 6: 内容分析与坐标映射（TDD）

**Files:**
- Modify: `/Users/qqx/my_code_cursor/lq/scripts/ppt_restyle.py`
- Modify: `/Users/qqx/my_code_cursor/lq/tests/test_ppt_restyle.py`

- [ ] **Step 1: 写失败的测试**

在 `tests/test_ppt_restyle.py` 追加：

```python
from pptx.util import Inches
from scripts.ppt_restyle import (
    get_content_shapes, compute_bounding_box,
    compute_mapping, map_position
)


class TestContentAnalysis:
    def test_get_content_shapes_excludes_title(self, source_pptx):
        prs = Presentation(source_pptx)
        slide = prs.slides[1]  # 文本页
        shapes = get_content_shapes(slide)
        titles = [s for s in shapes if get_title_text(slide) and
                  hasattr(s, 'placeholder_format') and
                  s.placeholder_format is not None and
                  s.placeholder_format.idx == 0]
        assert len(titles) == 0

    def test_get_content_shapes_includes_textbox(self, source_pptx):
        prs = Presentation(source_pptx)
        slide = prs.slides[1]
        shapes = get_content_shapes(slide)
        texts = [s.text_frame.text for s in shapes if s.has_text_frame]
        assert "额外文本框内容" in texts

    def test_compute_bounding_box(self, source_pptx):
        prs = Presentation(source_pptx)
        shapes = get_content_shapes(prs.slides[1])
        bbox = compute_bounding_box(shapes)
        assert bbox is not None
        left, top, right, bottom = bbox
        assert right > left
        assert bottom > top

    def test_compute_bounding_box_empty(self):
        assert compute_bounding_box([]) is None


class TestCoordinateMapping:
    def test_same_size_scale_one(self):
        src = (0, 0, 1000, 1000)
        dst = (0, 0, 1000, 1000)
        mapping = compute_mapping(src, dst)
        scale = mapping[0]
        assert scale == pytest.approx(1.0)

    def test_half_size_scale(self):
        src = (0, 0, 2000, 2000)
        dst = (0, 0, 1000, 1000)
        mapping = compute_mapping(src, dst)
        scale = mapping[0]
        assert scale == pytest.approx(0.5)

    def test_map_position_identity(self):
        mapping = (1.0, 100, 100, 100, 100)
        result = map_position(200, 200, 50, 50, mapping)
        assert result == (200, 200, 50, 50)

    def test_map_position_with_scale(self):
        mapping = (0.5, 0, 0, 0, 0)
        result = map_position(200, 400, 100, 100, mapping)
        assert result == (100, 200, 50, 50)
```

- [ ] **Step 2: 运行测试，确认失败**

```bash
cd /Users/qqx/my_code_cursor/lq && python3 -m pytest tests/test_ppt_restyle.py::TestContentAnalysis tests/test_ppt_restyle.py::TestCoordinateMapping -v
```

Expected: FAIL

- [ ] **Step 3: 实现**

在 `scripts/ppt_restyle.py` 追加：

```python
def get_content_shapes(slide):
    title_ph = _find_title_placeholder(slide)
    title_el = title_ph._element if title_ph else None

    content = []
    for shape in slide.shapes:
        if shape._element is title_el:
            continue
        if hasattr(shape, 'placeholder_format') and shape.placeholder_format is not None:
            idx = shape.placeholder_format.idx
            if idx in (0, 10, 11, 12):
                continue
        content.append(shape)
    return content


def compute_bounding_box(shapes):
    if not shapes:
        return None
    left = min(s.left for s in shapes)
    top = min(s.top for s in shapes)
    right = max(s.left + s.width for s in shapes)
    bottom = max(s.top + s.height for s in shapes)
    return (left, top, right, bottom)


def compute_mapping(src_bounds, dst_bounds):
    src_left, src_top, src_right, src_bottom = src_bounds
    dst_left, dst_top, dst_right, dst_bottom = dst_bounds

    src_w = src_right - src_left
    src_h = src_bottom - src_top
    dst_w = dst_right - dst_left
    dst_h = dst_bottom - dst_top

    if src_w == 0 or src_h == 0:
        return (1.0, dst_left, dst_top, src_left, src_top)

    scale = min(dst_w / src_w, dst_h / src_h)
    return (scale, dst_left, dst_top, src_left, src_top)


def map_position(left, top, width, height, mapping):
    scale, dst_left, dst_top, src_left, src_top = mapping
    new_left = int((left - src_left) * scale + dst_left)
    new_top = int((top - src_top) * scale + dst_top)
    new_width = int(width * scale)
    new_height = int(height * scale)
    return (new_left, new_top, new_width, new_height)
```

- [ ] **Step 4: 运行测试，确认通过**

```bash
cd /Users/qqx/my_code_cursor/lq && python3 -m pytest tests/test_ppt_restyle.py::TestContentAnalysis tests/test_ppt_restyle.py::TestCoordinateMapping -v
```

Expected: 8 passed

- [ ] **Step 5: 提交**

```bash
git add scripts/ppt_restyle.py tests/test_ppt_restyle.py
git commit -m "feat: 内容分析与坐标映射"
```

---

### Task 7: 元素迁移（TDD）

**Files:**
- Modify: `/Users/qqx/my_code_cursor/lq/scripts/ppt_restyle.py`
- Modify: `/Users/qqx/my_code_cursor/lq/tests/test_ppt_restyle.py`

- [ ] **Step 1: 写失败的测试**

在 `tests/test_ppt_restyle.py` 追加：

```python
from scripts.ppt_restyle import migrate_content, get_content_area_bounds


class TestShapeMigration:
    def test_migrate_textbox(self, template_pptx, source_pptx):
        tpl = Presentation(template_pptx)
        src = Presentation(source_pptx)
        dst_bounds = get_content_area_bounds(tpl.slides[1])

        dst_slide = duplicate_slide(tpl, 1)
        migrate_content(dst_slide, src.slides[1], dst_bounds)

        texts = [s.text_frame.text for s in dst_slide.shapes if s.has_text_frame]
        assert any("额外文本框内容" in t for t in texts)

    def test_migrate_table(self, template_pptx, source_pptx):
        tpl = Presentation(template_pptx)
        src = Presentation(source_pptx)
        dst_bounds = get_content_area_bounds(tpl.slides[1])

        dst_slide = duplicate_slide(tpl, 1)
        migrate_content(dst_slide, src.slides[2], dst_bounds)

        tables = [s for s in dst_slide.shapes if s.has_table]
        assert len(tables) >= 1
        assert tables[0].table.cell(0, 0).text == "A"

    def test_migrate_empty_slide_no_crash(self, template_pptx, tmp_path):
        tpl = Presentation(template_pptx)
        empty_prs = Presentation()
        empty_slide = empty_prs.slides.add_slide(empty_prs.slide_layouts[6])
        dst_bounds = get_content_area_bounds(tpl.slides[1])

        dst_slide = duplicate_slide(tpl, 1)
        migrate_content(dst_slide, empty_slide, dst_bounds)
```

- [ ] **Step 2: 运行测试，确认失败**

```bash
cd /Users/qqx/my_code_cursor/lq && python3 -m pytest tests/test_ppt_restyle.py::TestShapeMigration -v
```

Expected: FAIL

- [ ] **Step 3: 实现**

在 `scripts/ppt_restyle.py` 追加：

```python
def get_content_area_bounds(slide):
    ph = _find_content_placeholder(slide)
    if ph:
        return (ph.left, ph.top, ph.left + ph.width, ph.top + ph.height)
    raise ValueError("无法确定内容区域边界")


def migrate_content(dst_slide, src_slide, dst_bounds):
    content_shapes = get_content_shapes(src_slide)
    if not content_shapes:
        return

    content_ph = _find_content_placeholder(dst_slide)
    if content_ph:
        content_ph._element.getparent().remove(content_ph._element)

    src_bounds = compute_bounding_box(content_shapes)
    if src_bounds is None:
        return

    mapping = compute_mapping(src_bounds, dst_bounds)

    for shape in content_shapes:
        _migrate_single_shape(dst_slide, shape, mapping, src_slide)


def _migrate_single_shape(dst_slide, shape, mapping, src_slide):
    if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
        _migrate_picture(dst_slide, shape, mapping)
    elif hasattr(shape, 'has_chart') and shape.has_chart:
        _migrate_chart(dst_slide, shape, mapping, src_slide)
    else:
        _migrate_generic_shape(dst_slide, shape, mapping)


def _migrate_picture(dst_slide, shape, mapping):
    new_left, new_top, new_width, new_height = map_position(
        shape.left, shape.top, shape.width, shape.height, mapping
    )
    image_stream = BytesIO(shape.image.blob)
    dst_slide.shapes.add_picture(
        image_stream, new_left, new_top, new_width, new_height
    )


def _migrate_generic_shape(dst_slide, shape, mapping):
    new_el = deepcopy(shape._element)
    new_left, new_top, new_width, new_height = map_position(
        shape.left, shape.top, shape.width, shape.height, mapping
    )
    xfrm = new_el.find('.//a:xfrm', NSMAP)
    if xfrm is not None:
        off = xfrm.find('a:off', NSMAP)
        if off is not None:
            off.set('x', str(new_left))
            off.set('y', str(new_top))
        ext = xfrm.find('a:ext', NSMAP)
        if ext is not None:
            ext.set('cx', str(new_width))
            ext.set('cy', str(new_height))
    dst_slide.shapes._spTree.append(new_el)


def _migrate_chart(dst_slide, shape, mapping, src_slide):
    new_el = deepcopy(shape._element)
    new_left, new_top, new_width, new_height = map_position(
        shape.left, shape.top, shape.width, shape.height, mapping
    )
    xfrm = new_el.find('.//a:xfrm', NSMAP)
    if xfrm is not None:
        off = xfrm.find('a:off', NSMAP)
        if off is not None:
            off.set('x', str(new_left))
            off.set('y', str(new_top))
        ext = xfrm.find('a:ext', NSMAP)
        if ext is not None:
            ext.set('cx', str(new_width))
            ext.set('cy', str(new_height))

    chart_ns = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
    chart_ref = new_el.find(f'.//{{{chart_ns}}}chart')
    if chart_ref is not None:
        old_rId = chart_ref.get(f'{{{NSMAP["r"]}}}id')
        if old_rId and old_rId in src_slide.part.rels:
            rel = src_slide.part.rels[old_rId]
            new_rId = dst_slide.part.relate_to(rel.target_part, rel.reltype)
            chart_ref.set(f'{{{NSMAP["r"]}}}id', new_rId)

    dst_slide.shapes._spTree.append(new_el)
```

- [ ] **Step 4: 运行测试，确认通过**

```bash
cd /Users/qqx/my_code_cursor/lq && python3 -m pytest tests/test_ppt_restyle.py::TestShapeMigration -v
```

Expected: 3 passed

- [ ] **Step 5: 提交**

```bash
git add scripts/ppt_restyle.py tests/test_ppt_restyle.py
git commit -m "feat: 元素迁移（文本/表格/图片/图表）"
```

---

### Task 8: 主流程编排与 CLI（TDD）

**Files:**
- Modify: `/Users/qqx/my_code_cursor/lq/scripts/ppt_restyle.py`
- Modify: `/Users/qqx/my_code_cursor/lq/tests/test_ppt_restyle.py`

- [ ] **Step 1: 写失败的测试**

在 `tests/test_ppt_restyle.py` 追加：

```python
from scripts.ppt_restyle import restyle


class TestRestyle:
    def test_full_restyle(self, template_pptx, source_pptx, tmp_path):
        output_path = str(tmp_path / "output.pptx")
        restyle(source_pptx, template_pptx, output_path)

        result = Presentation(output_path)
        # 封面 + 2 内容页 = 3 页
        assert len(result.slides) == 3

    def test_cover_title_replaced(self, template_pptx, source_pptx, tmp_path):
        output_path = str(tmp_path / "output.pptx")
        restyle(source_pptx, template_pptx, output_path)

        result = Presentation(output_path)
        assert get_title_text(result.slides[0]) == "我的演示文稿"

    def test_content_title_replaced(self, template_pptx, source_pptx, tmp_path):
        output_path = str(tmp_path / "output.pptx")
        restyle(source_pptx, template_pptx, output_path)

        result = Presentation(output_path)
        assert get_title_text(result.slides[1]) == "第一节标题"
        assert get_title_text(result.slides[2]) == "数据表格"

    def test_single_slide_source(self, template_pptx, source_single_slide_pptx, tmp_path):
        output_path = str(tmp_path / "output.pptx")
        restyle(source_single_slide_pptx, template_pptx, output_path)

        result = Presentation(output_path)
        assert len(result.slides) == 1
        assert get_title_text(result.slides[0]) == "仅封面"

    def test_missing_source_raises(self, template_pptx, tmp_path):
        with pytest.raises(FileNotFoundError):
            restyle("/nonexistent.pptx", template_pptx, str(tmp_path / "out.pptx"))

    def test_missing_template_raises(self, source_pptx, tmp_path):
        with pytest.raises(FileNotFoundError):
            restyle(source_pptx, "/nonexistent.pptx", str(tmp_path / "out.pptx"))
```

- [ ] **Step 2: 运行测试，确认失败**

```bash
cd /Users/qqx/my_code_cursor/lq && python3 -m pytest tests/test_ppt_restyle.py::TestRestyle -v
```

Expected: FAIL

- [ ] **Step 3: 实现**

在 `scripts/ppt_restyle.py` 追加：

```python
def restyle(source_path, template_path, output_path):
    if not os.path.exists(source_path):
        raise FileNotFoundError(f"源文件不存在: {source_path}")
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"模板文件不存在: {template_path}")

    os.makedirs(os.path.dirname(output_path) or '.', exist_ok=True)

    template_prs = Presentation(template_path)
    validate_template(template_prs)
    content_bounds = get_content_area_bounds(template_prs.slides[1])

    src_prs = Presentation(source_path)
    if len(src_prs.slides) == 0:
        raise ValueError("源 PPT 没有任何幻灯片")

    shutil.copy2(template_path, output_path)
    prs = Presentation(output_path)

    cover_title = get_title_text(src_prs.slides[0])
    set_title_text(prs.slides[0], cover_title)

    template_slide_index = 1
    has_content_slides = False

    for i in range(1, len(src_prs.slides)):
        src_slide = src_prs.slides[i]
        title = get_title_text(src_slide)
        shapes = get_content_shapes(src_slide)

        if not shapes and not title:
            continue

        has_content_slides = True
        new_slide = duplicate_slide(prs, template_slide_index)
        set_title_text(new_slide, title)
        migrate_content(new_slide, src_slide, content_bounds)

    remove_slide(prs, template_slide_index)

    if not has_content_slides and len(src_prs.slides) == 1:
        pass  # 仅保留封面

    prs.save(output_path)
    return output_path


def main():
    if len(sys.argv) < 2:
        print("用法: python ppt_restyle.py <源PPT路径> [模板路径] [输出路径]")
        sys.exit(1)

    source = sys.argv[1]
    base_dir = Path(__file__).parent.parent

    template = (
        sys.argv[2] if len(sys.argv) > 2
        else str(base_dir / 'templates' / 'template.pptx')
    )

    if len(sys.argv) > 3:
        output = sys.argv[3]
    else:
        source_name = Path(source).stem
        output = str(base_dir / 'output' / f'{source_name}_styled.pptx')

    try:
        result = restyle(source, template, output)
        print(f"转换完成: {result}")
    except Exception as e:
        print(f"转换失败: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == '__main__':
    main()
```

- [ ] **Step 4: 运行测试，确认通过**

```bash
cd /Users/qqx/my_code_cursor/lq && python3 -m pytest tests/test_ppt_restyle.py::TestRestyle -v
```

Expected: 6 passed

- [ ] **Step 5: 运行全部测试**

```bash
cd /Users/qqx/my_code_cursor/lq && python3 -m pytest tests/ -v
```

Expected: 全部通过（预计 19 个测试）

- [ ] **Step 6: 提交**

```bash
git add scripts/ppt_restyle.py tests/test_ppt_restyle.py
git commit -m "feat: 主流程编排与 CLI 入口"
```

---

### Task 9: Skill 定义文件

**Files:**
- Create: `~/.claude/skills/ppt-restyle/SKILL.md`
- Create: `~/.claude/skills/ppt-restyle/references/restyle-guide.md`

- [ ] **Step 1: 创建目录**

```bash
mkdir -p ~/.claude/skills/ppt-restyle/references
```

- [ ] **Step 2: 创建 SKILL.md**

```markdown
---
name: ppt-restyle
description: 将用户 PPT 的内容迁移到模板 PPT 的样式中。接收源 PPT 文件路径，保持所有内容不变，替换为模板的背景和页面布局。
argument-hint: "<源PPT文件路径>"
metadata:
  author: user
  version: "1.0.0"
---

# PPT 样式重塑

将用户 PPT 的全部内容保持不变，套用模板 PPT 的视觉样式。

<args>$ARGUMENTS</args>

## 使用场景

- 统一多份 PPT 的视觉风格
- 快速将内容适配到公司/团队模板
- 批量重新设计演示文稿样式

## 工作流程

1. 用户提供源 PPT 文件路径作为参数
2. 从 `/Users/qqx/my_code_cursor/lq/templates/template.pptx` 读取模板
3. 执行转换脚本
4. 输出到 `/Users/qqx/my_code_cursor/lq/output/` 目录

## 执行步骤

1. 验证源文件存在
2. 验证模板文件存在（`/Users/qqx/my_code_cursor/lq/templates/template.pptx`）
3. 执行转换：

```bash
python3 /Users/qqx/my_code_cursor/lq/scripts/ppt_restyle.py "$ARGUMENTS"
```

4. 成功后告知用户输出文件路径
5. 若失败，展示错误信息并建议修复方案

## 模板要求

模板 PPT 必须恰好有 2 页：
- **第1页（封面）：** 包含一个标题占位符（placeholder idx=0）
- **第2页（内容页）：** 包含标题占位符（idx=0）和内容占位符（idx=1）

## 参考文档

| 主题 | 文件 |
|------|------|
| 实现指南 | `references/restyle-guide.md` |
```

- [ ] **Step 3: 创建 references/restyle-guide.md**

```markdown
# PPT 样式重塑 - 实现指南

## 技术架构

- **技术栈：** Python 3.8+, python-pptx, lxml
- **方案：** 模板克隆 + 占位符替换
- **核心脚本：** `/Users/qqx/my_code_cursor/lq/scripts/ppt_restyle.py`

## 处理流程

1. 复制模板文件作为基底
2. 封面页：替换标题占位符文本
3. 内容页：对源 PPT 每一页，克隆模板第2页，迁移标题和所有内容元素
4. 删除原始模板内容页
5. 保存输出

## 元素迁移策略

| 类型 | 方式 |
|------|------|
| 文本框 | XML 深拷贝 + 坐标调整 |
| 图片 | 提取二进制数据 + add_picture() |
| 表格 | XML 深拷贝 + 坐标调整 |
| 图表 | XML 深拷贝 + chart part 关系迁移 |

## 坐标映射

所有内容元素按等比缩放映射到模板内容区域，保持宽高比（取 `min(scale_x, scale_y)`）。

## 已知限制

- SmartArt 可能降级为静态图形
- 不支持嵌入视频
- 动画和切换效果不保留
- 图表外部数据链接转为内嵌数据

## 调试

```bash
# 直接运行脚本
python3 /Users/qqx/my_code_cursor/lq/scripts/ppt_restyle.py <源文件> [模板路径] [输出路径]

# 运行测试
cd /Users/qqx/my_code_cursor/lq && python3 -m pytest tests/ -v
```
```

- [ ] **Step 4: 验证 skill 文件**

```bash
cat ~/.claude/skills/ppt-restyle/SKILL.md | head -5
ls ~/.claude/skills/ppt-restyle/references/
```

Expected: 文件存在且内容正确

- [ ] **Step 5: 提交**

```bash
cd /Users/qqx/my_code_cursor/lq
git add -A
git commit -m "feat: ppt-restyle skill 定义文件"
```

---

## 执行清单摘要

| Task | 描述 | 预计耗时 |
|------|------|----------|
| 1 | 项目脚手架 | 2 min |
| 2 | 测试夹具 | 3 min |
| 3 | 模板校验（TDD） | 5 min |
| 4 | 幻灯片复制与删除（TDD） | 5 min |
| 5 | 标题操作（TDD） | 3 min |
| 6 | 内容分析与坐标映射（TDD） | 5 min |
| 7 | 元素迁移（TDD） | 5 min |
| 8 | 主流程编排与 CLI（TDD） | 5 min |
| 9 | Skill 定义文件 | 3 min |
| **合计** | | **~36 min** |
