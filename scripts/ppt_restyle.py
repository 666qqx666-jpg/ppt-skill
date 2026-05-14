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
