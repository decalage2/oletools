"""Check decryption of files from msodde works."""

import sys
import unittest
from os.path import basename, join as pjoin

from tests.test_utils import DATA_BASE_DIR, call_and_capture

from oletools import crypto


@unittest.skipIf(not crypto.check_msoffcrypto(),
                 'Module msoffcrypto not installed for {}'
                 .format(basename(sys.executable)))
class MsoddeCryptoTest(unittest.TestCase):
    """Test integration of decryption in msodde."""

    def test_standard_password(self):
        """Check dde-link is found in xls[mb] sample files."""
        for suffix in 'xls', 'xlsx', 'xlsm', 'xlsb':
            example_file = pjoin(DATA_BASE_DIR, 'encrypted',
                                 'dde-test-encrypt-standardpassword.' + suffix)
            output, _ = call_and_capture('msodde', [example_file, ])
            self.assertIn('\nDDE Links:\ncmd /c calc.exe\n', output,
                          msg='Unexpected output {!r} for {}'
                              .format(output, suffix))

    # TODO: add more, in particular a sample with a "proper" password


if __name__ == '__main__':
    unittest.main()
