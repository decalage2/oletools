"""
Test if oleid detects encrypted documents
"""

import unittest
from os.path import join
from tests.test_utils import DATA_BASE_DIR

from oletools import oleid


class TestEncryptedDocumentDetection(unittest.TestCase):
    def test_encrypted_document_detection(self):
        """ Run oleid and check if the document is flagged as encrypted """
        filename = join(DATA_BASE_DIR, 'basic', 'encrypted.docx')

        oleid_instance = oleid.OleID(filename)
        indicators = oleid_instance.check()

        is_encrypted = next(i.value for i in indicators if i.id == 'encrypted')

        self.assertEqual(is_encrypted, True)


# just in case somebody calls this file as a script
if __name__ == '__main__':
    unittest.main()