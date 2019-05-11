"""Check decryption of files from msodde works."""

import sys
import unittest
from os.path import basename, join as pjoin

from tests.test_utils import DATA_BASE_DIR

from oletools import crypto
from oletools import msodde


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
            link_text = msodde.process_maybe_encrypted(example_file)
            self.assertEqual(link_text, 'cmd /c calc.exe',
                             msg='Unexpected output {!r} for {}'
                                 .format(link_text, suffix))


if __name__ == '__main__':
    unittest.main()
