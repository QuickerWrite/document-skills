import unittest
from io import BytesIO
from pptx import Presentation
from pptx.util import Inches
from document_skills.pptx_patch import inspect_pptx, patch_pptx
from document_skills.render import _guard_ooxml
class WorkerTests(unittest.TestCase):
    def test_patch_without_server_import(self):
        p=Presentation();s=p.slides.add_slide(p.slide_layouts[6]);shape=s.shapes.add_textbox(Inches(1),Inches(1),Inches(4),Inches(1));shape.text='Before'
        b=BytesIO();p.save(b)
        result=patch_pptx(b.getvalue(),[{'slide_number':1,'element_id':shape.shape_id,'changes':{'text':'After'}}])
        self.assertEqual(inspect_pptx(result)['slides'][0]['elements'][0]['text'],'After')
    def test_bad_archive_rejected(self):
        with self.assertRaises(ValueError): _guard_ooxml(b'not office')
