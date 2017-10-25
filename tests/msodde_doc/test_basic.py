""" Test some basic behaviour of msodde.py

Ensure that
- doc and docx are read without error
- garbage returns error return status
- dde-links are found where appropriate
"""

import unittest
from oletools import msodde
import shlex
from os.path import join, dirname, normpath

BASE_DIR = normpath(join(dirname(__file__), '..', 'test-data'))


class TestReturnCode(unittest.TestCase):

    def test_valid_doc(self):
        """ check that a valid doc file leads to 0 exit status """
        print(join(BASE_DIR, 'msodde-doc/test_document.doc'))
        self.do_test_validity(join(BASE_DIR, 'msodde-doc/test_document.doc'))

    def test_valid_docx(self):
        """ check that a valid docx file leads to 0 exit status """
        self.do_test_validity(join(BASE_DIR, 'msodde-doc/test_document.docx'))

    def test_invalid_none(self):
        """ check that no file argument leads to non-zero exit status """
        self.do_test_validity('', True)

    def test_invalid_empty(self):
        """ check that empty file argument leads to non-zero exit status """
        self.do_test_validity(join(BASE_DIR, 'basic/empty'), True)

    def test_invalid_text(self):
        """ check that text file argument leads to non-zero exit status """
        self.do_test_validity(join(BASE_DIR, 'basic/text'), True)

    def do_test_validity(self, args, expect_error=False):
        """ helper for test_valid_doc[x] """
        args = shlex.split(args)
        return_code = -1
        have_exception = False
        try:
            return_code = msodde.main(args)
        except Exception:
            have_exception = True
        except SystemExit as se:     # sys.exit() was called
            return_code = se.code
            if se.code is None:
                return_code = 0

        self.assertEqual(expect_error, have_exception or (return_code != 0))


if __name__ == '__main__':
    unittest.main()
