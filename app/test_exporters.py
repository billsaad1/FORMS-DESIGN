import os
import unittest
from app.docx_generator import generate_form1_docx, generate_form2_docx, generate_form3_docx
from app.excel_generator import generate_form1_xlsx, generate_form2_xlsx, generate_form3_xlsx

class TestExporters(unittest.TestCase):

    def setUp(self):
        # Create a temp directory for test output files
        self.output_dir = "test_outputs"
        os.makedirs(self.output_dir, exist_ok=True)

    def test_word_export_form1(self):
        filepath = os.path.join(self.output_dir, "test_form1.docx")
        data = generate_form1_docx()
        with open(filepath, 'wb') as f:
            f.write(data)
        self.assertTrue(os.path.exists(filepath))
        self.assertGreater(os.path.getsize(filepath), 0)

    def test_word_export_form2(self):
        filepath = os.path.join(self.output_dir, "test_form2.docx")
        data = generate_form2_docx()
        with open(filepath, 'wb') as f:
            f.write(data)
        self.assertTrue(os.path.exists(filepath))
        self.assertGreater(os.path.getsize(filepath), 0)

    def test_word_export_form3(self):
        filepath = os.path.join(self.output_dir, "test_form3.docx")
        data = generate_form3_docx()
        with open(filepath, 'wb') as f:
            f.write(data)
        self.assertTrue(os.path.exists(filepath))
        self.assertGreater(os.path.getsize(filepath), 0)

    def test_excel_export_form1(self):
        filepath = os.path.join(self.output_dir, "test_form1.xlsx")
        data = generate_form1_xlsx()
        with open(filepath, 'wb') as f:
            f.write(data)
        self.assertTrue(os.path.exists(filepath))
        self.assertGreater(os.path.getsize(filepath), 0)

    def test_excel_export_form2(self):
        filepath = os.path.join(self.output_dir, "test_form2.xlsx")
        data = generate_form2_xlsx()
        with open(filepath, 'wb') as f:
            f.write(data)
        self.assertTrue(os.path.exists(filepath))
        self.assertGreater(os.path.getsize(filepath), 0)

    def test_excel_export_form3(self):
        filepath = os.path.join(self.output_dir, "test_form3.xlsx")
        data = generate_form3_xlsx()
        with open(filepath, 'wb') as f:
            f.write(data)
        self.assertTrue(os.path.exists(filepath))
        self.assertGreater(os.path.getsize(filepath), 0)

if __name__ == "__main__":
    unittest.main()
