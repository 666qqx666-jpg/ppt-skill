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
