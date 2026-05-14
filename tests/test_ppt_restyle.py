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


from scripts.ppt_restyle import (
    get_content_shapes, compute_bounding_box,
    compute_mapping, map_position
)


class TestContentAnalysis:
    def test_get_content_shapes_excludes_title(self, source_pptx):
        prs = Presentation(source_pptx)
        slide = prs.slides[1]  # 文本页
        shapes = get_content_shapes(slide)

        def _ph_idx(s):
            try:
                ph = s.placeholder_format
                return ph.idx if ph is not None else None
            except ValueError:
                return None

        titles = [s for s in shapes if get_title_text(slide) and _ph_idx(s) == 0]
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
