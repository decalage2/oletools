"""
Test basic functionality of olevba[3]
"""

import unittest
import os
from os.path import join, splitext
import re
import json

# Directory with test data, independent of current working directory
from tests.test_utils import DATA_BASE_DIR, call_and_capture


class TestOlevbaBasic(unittest.TestCase):
    """Tests olevba basic functionality"""

    def test_text_behaviour(self):
        """Test behaviour of olevba when presented with pure text file."""
        self.do_test_behaviour('text')

    def test_empty_behaviour(self):
        """Test behaviour of olevba when presented with pure text file."""
        self.do_test_behaviour('empty')

    def do_test_behaviour(self, filename):
        """Helper for test_{text,empty}_behaviour."""
        input_file = join(DATA_BASE_DIR, 'basic', filename)
        output, _ = call_and_capture('olevba', args=(input_file, ))

        # check output
        self.assertTrue(re.search(r'^Type:\s+Text\s*$', output, re.MULTILINE),
                        msg='"Type: Text" not found in output:\n' + output)
        self.assertTrue(re.search(r'^No suspicious .+ found.$', output,
                                  re.MULTILINE),
                        msg='"No suspicous...found" not found in output:\n' + \
                            output)
        self.assertNotIn('error', output.lower())

        # check warnings
        for line in output.splitlines():
            if line.startswith('WARNING ') and 'encrypted' in line:
                continue   # encryption warnings are ok
            elif 'warn' in line.lower():
                raise self.fail('Found "warn" in output line: "{}"'
                                .format(line.rstrip()))
        # TODO: I disabled this test because we do not log "not encrypted" as warning anymore
        # to avoid other issues.
        # If we really want to test this, then the test should be run with log level INFO:
        # self.assertIn('not encrypted', output)

    def test_rtf_behaviour(self):
        """Test behaviour of olevba when presented with an rtf file."""
        input_file = join(DATA_BASE_DIR, 'msodde', 'RTF-Spec-1.7.rtf')
        output, ret_code = call_and_capture('olevba', args=(input_file, ),
                                            accept_nonzero_exit=True)

        # check that return code is olevba.RETURN_OPEN_ERROR
        self.assertEqual(ret_code, 5)

        # check output:
        self.assertIn('FileOpenError', output)
        self.assertIn('is RTF', output)
        self.assertIn('rtfobj', output)
        # TODO: I disabled this test because we do not log "not encrypted" as warning anymore
        # to avoid other issues.
        # If we really want to test this, then the test should be run with log level INFO:
        # self.assertIn('not encrypted', output)

        # check warnings
        for line in output.splitlines():
            if line.startswith('WARNING ') and 'encrypted' in line:
                continue   # encryption warnings are ok
            elif 'warn' in line.lower():
                raise self.fail('Found "warn" in output line: "{}"'
                                .format(line.rstrip()))

    @unittest.skipIf('OLETOOLS_TEST_SKIP_SLOW' in os.environ and os.environ['OLETOOLS_TEST_SKIP_SLOW'] == '1',
                     "Skip slower tests")
    def test_crypt_return(self):
        """
        Test that encrypted files give a certain return code.

        Currently, only the encryption applied by Office 2010 (CryptoApi RC4
        Encryption) is tested.
        """
        CRYPT_DIR = join(DATA_BASE_DIR, 'encrypted')
        CRYPT_RETURN_CODE = 9
        ADD_ARGS = [], ['-d', ], ['-a', ], ['-j', ], ['-t', ]   # only 1st file
        EXCEPTIONS = ['autostart-encrypt-standardpassword.xls',   # These ...
                      'autostart-encrypt-standardpassword.xlsm',  # files ...
                      'autostart-encrypt-standardpassword.xlsb',  # are ...
                      'dde-test-encrypt-standardpassword.xls',    # automati...
                      'dde-test-encrypt-standardpassword.xlsx',   # ...cally...
                      'dde-test-encrypt-standardpassword.xlsm',   # decrypted.
                      'dde-test-encrypt-standardpassword.xlsb']
        for filename in os.listdir(CRYPT_DIR):
            if filename in EXCEPTIONS:
                continue
            full_name = join(CRYPT_DIR, filename)
            for args in ADD_ARGS:
                _, ret_code = call_and_capture('olevba',
                                               args=[full_name, ] + args,
                                               accept_nonzero_exit=True)
                self.assertEqual(ret_code, CRYPT_RETURN_CODE,
                                 msg='Wrong return code {} for args {}'\
                                     .format(ret_code, args + [filename, ]))

                # test only first file with all arg combinations, others just
                # without arg (test takes too long otherwise)
                ADD_ARGS = ([], )

    def test_xlm(self):
        """Test that xlm macros are found."""
        XLM_DIR = join(DATA_BASE_DIR, 'excel4-macros')
        ADD_ARGS = ['-j']

        for filename in os.listdir(XLM_DIR):
            full_name = join(XLM_DIR, filename)
            suffix = splitext(filename)[1]
            out_str, ret_code = call_and_capture('olevba',
                                                 args=[full_name, ] + ADD_ARGS,
                                                 accept_nonzero_exit=True)
            output = json.loads(out_str)
            self.assertGreaterEqual(len(output), 2)
            self.assertEqual(output[0]['type'], 'MetaInformation')
            self.assertEqual(output[0]['script_name'], 'olevba')
            for entry in output[1:]:
                if entry['type'] in ('msg', 'warning'):
                    continue    # ignore messages
                result = entry
                break
            self.assertTrue(result['json_conversion_successful'])
            if suffix in ('.xlsb', '.xltm', '.xlsm'):
                # TODO: cannot extract xlm macros for these types yet
                self.assertEqual(result['macros'], [])
            else:
                code = result['macros'][0]['code']
                if suffix == '.slk':
                    self.assertIn('Excel 4 macros extracted', code)
                else:
                    self.assertIn('Excel 4.0 macro sheet', code)
                self.assertIn('Auto_Open', code)
                if 'excel5' not in filename:    # TODO: is not found in excel5
                    self.assertIn('ALERT(', code)
                self.assertIn('HALT()', code)

                self.assertIn(len(result['analysis']), (2, 3))
                types = [entry['type'] for entry in result['analysis']]
                keywords = [entry['keyword'] for entry in result['analysis']]
                self.assertIn('Auto_Open', keywords)
                self.assertIn('XLM macro', keywords)
                self.assertIn('AutoExec', types)
                self.assertIn('Suspicious', types)

    def test_dir_stream_record_project_compat_version(self):
        """Test PROJECTCOMPATVERSION record on dir stream with a ppt file."""
        input_file = join(DATA_BASE_DIR, 'olevba', 'sample_with_vba.ppt')
        output, ret_code = call_and_capture('olevba', args=(input_file, "--loglevel", "debug"))

        # check return code
        self.assertEqual(ret_code, 0)

        # not expected string:
        self.assertNotIn('invalid value for PROJECTLCID_Id expected 0002 got', output)
        self.assertNotIn('Error in _extract_vba', output)

        # compat version in debug mode:
        self.assertIn('compat version: 2', output)

        # vba contents:
        self.assertIn('Sub Action_Click()\n  MsgBox "The action button clicked!"\nEnd Sub', output)


# just in case somebody calls this file as a script
if __name__ == '__main__':
    unittest.main()
