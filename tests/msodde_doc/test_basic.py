""" Test some basic behaviour of msodde.py

Ensure that
- doc and docx are read without error
- garbage returns error return status
- dde-links are found where appropriate
"""

from __future__ import print_function

import unittest
from oletools import msodde
import shlex
from os.path import join, dirname, normpath
import sys

# python 2/3 version conflict:
if sys.version_info.major <= 2:
    from StringIO import StringIO
    #from io import BytesIO as StringIO - try if print() gives UnicodeError
else:
    from io import StringIO


# base directory for test input
BASE_DIR = normpath(join(dirname(__file__), '..', 'test-data'))


class TestReturnCode(unittest.TestCase):

    def test_valid_doc(self):
        """ check that a valid doc file leads to 0 exit status """
        for filename in ('dde-test-from-office2003', 'dde-test-from-office2016',
                         'harmless-clean'):
            self.do_test_validity(join(BASE_DIR, 'msodde-doc',
                                       filename + '.doc'))

    def test_valid_docx(self):
        """ check that a valid docx file leads to 0 exit status """
        for filename in 'dde-test', 'harmless-clean':
            self.do_test_validity(join(BASE_DIR, 'msodde-doc',
                                       filename + '.docx'))

    def test_valid_docm(self):
        """ check that a valid docm file leads to 0 exit status """
        for filename in 'dde-test', 'harmless-clean':
            self.do_test_validity(join(BASE_DIR, 'msodde-doc',
                                       filename + '.docm'))

    def test_invalid_other(self):
        """ check that xml do not work yet """
        for extn in '-2003.xml', '.xml':
            self.do_test_validity(join(BASE_DIR, 'msodde-doc',
                                       'harmless-clean' + extn), True)

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

        self.assertEqual(expect_error, have_exception or (return_code != 0),
                         msg='Args={0}, expect={1}, exc={2}, return={3}'
                             .format(args, expect_error, have_exception,
                                     return_code))


class OutputCapture:
    """ context manager that captures stdout """

    def __init__(self):
        self.output = StringIO()   # in py2, this actually is BytesIO

    def __enter__(self):
        sys.stdout = self.output
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        sys.stdout = sys.__stdout__    # re-set to original

        if exc_type:   # there has been an error
            print('Got error during output capture!')
            print('Print captured output and re-raise:')
            for line in self.output.getvalue().splitlines():
                print(line.rstrip())  # print output before re-raising

    def __iter__(self):
        for line in self.output.getvalue().splitlines():
            yield line.rstrip()   # remove newline at end of line


class TestDdeInDoc(unittest.TestCase):

    def get_dde_from_output(self, capturer):
        """ helper to read dde links from captured output """
        have_start_line = False
        result = []
        for line in capturer:
            if not line.strip():
                continue   # skip empty lines
            if have_start_line:
                result.append(line)
            elif line == 'DDE Links:':
                have_start_line = True

        self.assertTrue(have_start_line) # ensure output was complete
        return result

    def test_with_dde(self):
        """ check that dde links appear on stdout """
        with OutputCapture() as capturer:
            msodde.main([join(BASE_DIR, 'msodde-doc',
                              'dde-test-from-office2003.doc')])
        self.assertNotEqual(len(self.get_dde_from_output(capturer)), 0,
                            msg='Found no dde links in output for doc file')

    def test_no_dde(self):
        """ check that no dde links appear on stdout """
        with OutputCapture() as capturer:
            msodde.main([join(BASE_DIR, 'msodde-doc', 'harmless-clean.doc')])
        self.assertEqual(len(self.get_dde_from_output(capturer)), 0,
                         msg='Found dde links in output for doc file')


if __name__ == '__main__':
    unittest.main()
