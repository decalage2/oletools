""" Test some basic behaviour of msodde.py

Ensure that
- doc and docx are read without error
- garbage returns error return status
- dde-links are found where appropriate
"""

from __future__ import print_function

import unittest
from platform import python_implementation
import sys
import os
from os.path import join, basename
from oletools import msodde
from oletools.crypto import \
    WrongEncryptionPassword, CryptoLibNotImported, check_msoffcrypto
from tests.test_utils import call_and_capture, decrypt_sample,\
    DATA_BASE_DIR as BASE_DIR


# Check whether we run with PyPy on windows because that causes trouble
# when using the :py:func:`tests.test_utils.decrypt_sample`.
#
# :return: `(do_skip, explanation)` where `do_skip` is `True` iff running
# PyPy on Windows; `explanation` is a simple text string
SKIP_PYPY_WIN = (
    python_implementation().lower().startswith('pypy')
            and sys.platform.lower().startswith('win'),
    "On PyPy there is a problem with deleting temp files for decrypt_sample"
)


class TestReturnCode(unittest.TestCase):
    """ check return codes and exception behaviour (not text output) """
    @unittest.skipIf(*SKIP_PYPY_WIN)
    def test_valid_doc(self):
        """ check that a valid doc file leads to 0 exit status """
        for filename in (
                'harmless-clean.doc',
                'dde-test-from-office2003.doc.zip',
                'dde-test-from-office2016.doc.zip',
                'dde-test-from-office2013-utf_16le-korean.doc.zip',
        ):
            with decrypt_sample(join('msodde', filename)) as temp_name:
                self.do_test_validity(temp_name)

    def test_valid_docx(self):
        """ check that a valid docx file leads to 0 exit status """
        for filename in 'dde-test', 'harmless-clean':
            self.do_test_validity(join(BASE_DIR, 'msodde',
                                       filename + '.docx'))

    def test_valid_docm(self):
        """ check that a valid docm file leads to 0 exit status """
        for filename in 'dde-test', 'harmless-clean':
            self.do_test_validity(join(BASE_DIR, 'msodde',
                                       filename + '.docm'))

    @unittest.skipIf(*SKIP_PYPY_WIN)
    def test_valid_xml(self):
        """ check that xml leads to 0 exit status """
        for filename in (
                'harmless-clean-2003.xml',
                'dde-in-excel2003.xml',
                'dde-in-word2003.xml.zip',
                'dde-in-word2007.xml.zip'
        ):
            with decrypt_sample(join('msodde', filename)) as temp_name:
                self.do_test_validity(temp_name)

    def test_invalid_none(self):
        """ check that no file argument leads to non-zero exit status """
        if sys.hexversion > 0x03030000:   # version 3.3 and higher
            # different errors probably depending on whether msoffcryto is
            # available or not
            expect_error = (AttributeError, FileNotFoundError)
        else:
            expect_error = (AttributeError, IOError)
        self.do_test_validity('', expect_error)

    def test_invalid_empty(self):
        """ check that empty file argument leads to non-zero exit status """
        self.do_test_validity(join(BASE_DIR, 'basic', 'empty'), Exception)

    def test_invalid_text(self):
        """ check that text file argument leads to non-zero exit status """
        self.do_test_validity(join(BASE_DIR, 'basic', 'text'), Exception)

    def test_encrypted(self):
        """
        check that encrypted files lead to non-zero exit status

        Currently, only the encryption applied by Office 2010 (CryptoApi RC4
        Encryption) is tested.
        """
        CRYPT_DIR = join(BASE_DIR, 'encrypted')
        have_crypto = check_msoffcrypto()
        for filename in os.listdir(CRYPT_DIR):
            if have_crypto and 'standardpassword' in filename:
                # these are automagically decrypted
                self.do_test_validity(join(CRYPT_DIR, filename))
            elif have_crypto:
                self.do_test_validity(join(CRYPT_DIR, filename),
                                      WrongEncryptionPassword)
            else:
                self.do_test_validity(join(CRYPT_DIR, filename),
                                      CryptoLibNotImported)

    def do_test_validity(self, filename, expect_error=None):
        """ helper for test_[in]valid_* """
        found_error = None
        try:
            msodde.process_maybe_encrypted(filename,
                field_filter_mode=msodde.FIELD_FILTER_BLACKLIST)
        except Exception as exc:
            found_error = exc

        if expect_error and not found_error:
            self.fail('Expected {} but msodde finished without errors for {}'
                      .format(expect_error, filename))
        elif not expect_error and found_error:
            self.fail('Unexpected error {} from msodde for {}'
                      .format(found_error, filename))
        elif expect_error and not isinstance(found_error, expect_error):
            self.fail('Wrong kind of error {} from msodde for {}, expected {}'
                      .format(type(found_error), filename, expect_error))


@unittest.skipIf(not check_msoffcrypto(),
                 'Module msoffcrypto not installed for {}'
                 .format(basename(sys.executable)))
class TestErrorOutput(unittest.TestCase):
    """msodde does not specify error by return code but text output."""

    def test_crypt_output(self):
        """Check for helpful error message when failing to decrypt."""
        for suffix in 'doc', 'docm', 'docx', 'ppt', 'pptm', 'pptx', 'xls', \
                'xlsb', 'xlsm', 'xlsx':
            example_file = join(BASE_DIR, 'encrypted', 'encrypted.' + suffix)
            output, ret_code = call_and_capture('msodde', [example_file, ],
                                                accept_nonzero_exit=True)
            self.assertEqual(ret_code, 1)
            self.assertIn('passwords could not decrypt office file', output,
                          msg='Unexpected output: {}'.format(output.strip()))


class TestDdeLinks(unittest.TestCase):
    """ capture output of msodde and check dde-links are found correctly """

    @staticmethod
    def get_dde_from_output(output):
        """ helper to read dde links from captured output
        """
        return [o for o in output.splitlines()]

    @unittest.skipIf(*SKIP_PYPY_WIN)
    def test_with_dde(self):
        """ check that dde links appear on stdout """
        filename = 'dde-test-from-office2003.doc.zip'
        with decrypt_sample(join('msodde', filename)) as temp_file:
            output = msodde.process_maybe_encrypted(temp_file,
                field_filter_mode=msodde.FIELD_FILTER_BLACKLIST)
            self.assertNotEqual(len(self.get_dde_from_output(output)), 0,
                                msg='Found no dde links in output of ' + filename)

    def test_no_dde(self):
        """ check that no dde links appear on stdout """
        filename = 'harmless-clean.doc'
        output = msodde.process_maybe_encrypted(
            join(BASE_DIR, 'msodde', filename),
            field_filter_mode=msodde.FIELD_FILTER_BLACKLIST)
        self.assertEqual(len(self.get_dde_from_output(output)), 0,
                         msg='Found dde links in output of ' + filename)

    @unittest.skipIf(*SKIP_PYPY_WIN)
    def test_with_dde_utf16le(self):
        """ check that dde links appear on stdout """
        filename = 'dde-test-from-office2013-utf_16le-korean.doc.zip'
        with decrypt_sample(join('msodde', filename)) as temp_file:
            output = msodde.process_maybe_encrypted(temp_file,
                field_filter_mode=msodde.FIELD_FILTER_BLACKLIST)
            self.assertNotEqual(len(self.get_dde_from_output(output)), 0,
                                msg='Found no dde links in output of ' + filename)

    def test_excel(self):
        """ check that dde links are found in excel 2007+ files """
        expect = ['cmd /c calc.exe', ]
        for extn in 'xlsx', 'xlsm', 'xlsb':
            output = msodde.process_maybe_encrypted(
                join(BASE_DIR, 'msodde', 'dde-test.' + extn),
                field_filter_mode=msodde.FIELD_FILTER_BLACKLIST)

            self.assertEqual(expect, self.get_dde_from_output(output),
                             msg='unexpected output for dde-test.{0}: {1}'
                                 .format(extn, output))

    @unittest.skipIf(*SKIP_PYPY_WIN)
    def test_xml(self):
        """ check that dde in xml from word / excel is found """
        for filename in ('dde-in-excel2003.xml',
                         'dde-in-word2003.xml.zip',
                         'dde-in-word2007.xml.zip'):
            with decrypt_sample(join('msodde', filename)) as temp_file:
                output = msodde.process_maybe_encrypted(temp_file,
                    field_filter_mode=msodde.FIELD_FILTER_BLACKLIST)
                links = self.get_dde_from_output(output)
                self.assertEqual(len(links), 1, 'found {0} dde-links in {1}'
                                                .format(len(links), filename))
                self.assertTrue('cmd' in links[0], 'no "cmd" in dde-link for {0}'
                                                   .format(filename))
                self.assertTrue('calc' in links[0], 'no "calc" in dde-link for {0}'
                                                    .format(filename))

    def test_clean_rtf_blacklist(self):
        """ find a lot of hyperlinks in rtf spec """
        filename = 'RTF-Spec-1.7.rtf'
        output = msodde.process_maybe_encrypted(
            join(BASE_DIR, 'msodde', filename),
            field_filter_mode=msodde.FIELD_FILTER_BLACKLIST)
        self.assertEqual(len(self.get_dde_from_output(output)), 1413)

    def test_clean_rtf_ddeonly(self):
        """ find no dde links in rtf spec """
        filename = 'RTF-Spec-1.7.rtf'
        output = msodde.process_maybe_encrypted(
            join(BASE_DIR, 'msodde', filename),
            field_filter_mode=msodde.FIELD_FILTER_DDE)
        self.assertEqual(len(self.get_dde_from_output(output)), 0,
                         msg='Found dde links in output of ' + filename)


if __name__ == '__main__':
    unittest.main()
