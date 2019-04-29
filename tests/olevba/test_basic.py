"""
Test basic functionality of olevba[3]
"""

import unittest
import sys
import os
from os.path import join
from contextlib import contextmanager
try:
    from cStringIO import StringIO
except ImportError:   # py3:
    from io import StringIO
if sys.version_info.major <= 2:
    from oletools import olevba
else:
    from oletools import olevba3 as olevba

# Directory with test data, independent of current working directory
from tests.test_utils import DATA_BASE_DIR


@contextmanager
def capture_output():
    """
    Temporarily replace stdout/stderr with buffers to capture output.

    Once we only support python>=3.4: this is already built into python as
    :py:func:`contextlib.redirect_stdout`.

    Not quite sure why, but seems to only work once per test function ...
    """
    orig_stdout = sys.stdout
    orig_stderr = sys.stderr

    try:
        sys.stdout = StringIO()
        sys.stderr = StringIO()
        yield sys.stdout, sys.stderr

    finally:
        sys.stdout = orig_stdout
        sys.stderr = orig_stderr


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
        ret_code = -1

        # run olevba, capturing its output and return code
        with capture_output() as (stdout, stderr):
            with self.assertRaises(SystemExit) as raise_context:
                olevba.main([input_file, ])
            ret_code = raise_context.exception.code

        # check that return code is 0
        self.assertEqual(ret_code, 0)

        # check there are only warnings in stderr
        stderr = stderr.getvalue()
        skip_line = False
        for line in stderr.splitlines():
            if skip_line:
                skip_line = False
                continue
            self.assertTrue(line.startswith('WARNING ') or
                            'ResourceWarning' in line,
                            msg='Line "{}" in stderr is unexpected for {}'\
                                .format(line.rstrip(), filename))
            if 'ResourceWarning' in line:
                skip_line = True
        self.assertIn('not encrypted', stderr)

        # check stdout
        stdout = stdout.getvalue().lower()
        self.assertIn(input_file.lower(), stdout)
        self.assertIn('type: text', stdout)
        self.assertIn('no suspicious', stdout)
        self.assertNotIn('error', stdout)
        self.assertNotIn('warn', stdout)

    def test_rtf_behaviour(self):
        """Test behaviour of olevba when presented with an rtf file."""
        input_file = join(DATA_BASE_DIR, 'msodde', 'RTF-Spec-1.7.rtf')
        ret_code = -1

        # run olevba, capturing its output and return code
        with capture_output() as (stdout, stderr):
            with self.assertRaises(SystemExit) as raise_context:
                olevba.main([input_file, ])
            ret_code = raise_context.exception.code

        # check that return code is olevba.RETURN_OPEN_ERROR
        self.assertEqual(ret_code, 5)
        stdout = stdout.getvalue().lower()
        self.assertNotIn('error', stdout)
        self.assertNotIn('warn', stdout)

        stderr = stderr.getvalue().lower()
        self.assertIn('fileopenerror', stderr)
        self.assertIn('is rtf', stderr)
        self.assertIn('rtfobj.py', stderr)
        self.assertIn('not encrypted', stderr)

    def test_crypt_return(self):
        """
        Tests that encrypted files give a certain return code.

        Currently, only the encryption applied by Office 2010 (CryptoApi RC4
        Encryption) is tested.
        """
        CRYPT_DIR = join(DATA_BASE_DIR, 'encrypted')
        CRYPT_RETURN_CODE = 9
        ADD_ARGS = [], ['-d', ], ['-a', ], ['-j', ], ['-t', ]
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
                try:
                    olevba.main(args + [full_name, ])
                    self.fail('Olevba should have exited')
                except SystemExit as sys_exit:
                    ret_code = sys_exit.code or 0   # sys_exit.code can be None
                self.assertEqual(ret_code, CRYPT_RETURN_CODE,
                                 msg='Wrong return code {} for args {}'\
                                     .format(ret_code, args + [filename, ]))


# just in case somebody calls this file as a script
if __name__ == '__main__':
    unittest.main()
