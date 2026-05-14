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


def get_content_shapes(slide):
    title_ph = _find_title_placeholder(slide)
    title_el = title_ph._element if title_ph else None

    content = []
    for shape in slide.shapes:
        if shape._element is title_el:
            continue
        try:
            ph_fmt = shape.placeholder_format
        except ValueError:
            ph_fmt = None
        if ph_fmt is not None:
            if ph_fmt.idx in (0, 10, 11, 12):
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
