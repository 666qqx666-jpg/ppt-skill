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
