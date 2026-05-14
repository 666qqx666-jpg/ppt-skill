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
