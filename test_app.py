from io import BytesIO
import os
import unittest
from unittest.mock import patch
from streamlit.testing.v1 import AppTest
from test_quote_service import upload_bytes


class AppTests(unittest.TestCase):
    def test_upload_quote_download_and_no_duplicate_submit(self):
        upload = BytesIO(upload_bytes())
        upload.name = '测试模板.xlsx'
        answer = {'options': [{'carrierCode': 'TEST', 'carrierName': 'Test carrier', 'totalAmount': 123.45}], 'quoteNo': 'TEST', 'freightClass': '100'}
        with patch.dict(os.environ, {'FIMILE_API_TOKEN': 'test-only'}), patch('streamlit.file_uploader', return_value=upload), patch('quote_service.request_quote', return_value=answer) as request:
            app = AppTest.from_file('app.py').run(timeout=20)
            self.assertFalse(app.exception)
            self.assertEqual(app.button[0].label, '开始询价')
            app.button[0].click().run(timeout=20)
            self.assertFalse(app.exception)
            self.assertEqual(request.call_count, 1)
            self.assertEqual(app.session_state['results'][2]['options'][0]['totalAmount'], 123.45)
            self.assertTrue(app.button[0].disabled)
            app.run(timeout=20)
            self.assertEqual(request.call_count, 1)


if __name__ == '__main__':
    unittest.main()
