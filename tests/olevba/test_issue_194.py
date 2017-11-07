""" Tests if olevba detects encrypted documents
"""

import unittest, sys, os

from os.path import join
from tests.test_utils import DATA_BASE_DIR

if sys.version_info[0] <= 2:
    from oletools import olevba
else:
    from oletools import olevba3 as olevba

class TestEncryptedDocumentDetection(unittest.TestCase):
    def test_encrypted_document_detection(self):
        """ Run olevba and check if the return code indicates encryption """
        filename = join(DATA_BASE_DIR, 'basic/encrypted.docx')

        try:
            return_code = olevba.main([filename])
        except SystemExit as se:
            return_code = se.code

        self.assertEqual(return_code, olevba.RETURN_ENCRYPTED_OOXML)

# just in case somebody calls this file as a script
if __name__ == '__main__':
    unittest.main()
