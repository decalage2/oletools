"""Check decryption of files from olevba works."""

import sys
import unittest
from os.path import basename, join as pjoin
import json
from collections import OrderedDict

from tests.test_utils import DATA_BASE_DIR, call_and_capture

from oletools import crypto


@unittest.skipIf(not crypto.check_msoffcrypto(),
                 'Module msoffcrypto not installed for {}'
                 .format(basename(sys.executable)))
class OlevbaCryptoWriteProtectTest(unittest.TestCase):
    """
    Test documents that are 'write-protected' through encryption.

    Excel has a way to 'write-protect' documents by encrypting them with a
    hard-coded standard password. When looking at the file-structure you see
    an OLE-file with streams `EncryptedPackage`, `StrongEncryptionSpace`, and
    `EncryptionInfo`. Contained in the first is the actual file.  When opening
    such a file in excel, it is decrypted without the user noticing.

    Olevba should detect such encryption, try to decrypt with the standard
    password and look for VBA code in the decrypted file.

    All these tests are skipped if the module `msoffcrypto-tools` is not
    installed.
    """
    def test_autostart(self):
        """Check that autostart macro is found in xls[mb] sample file."""
        for suffix in 'xlsm', 'xlsb':
            example_file = pjoin(
                DATA_BASE_DIR, 'encrypted',
                'autostart-encrypt-standardpassword.' + suffix)
            output, _ = call_and_capture('olevba', args=('-j', example_file),
                                         exclude_stderr=True)
            data = json.loads(output, object_pairs_hook=OrderedDict)
            # debug: json.dump(data, sys.stdout, indent=4)
            self.assertEqual(len(data), 3)
            self.assertIn('script_name', data[0])
            self.assertIn('version', data[0])
            self.assertEqual(data[0]['type'], 'MetaInformation')
            self.assertEqual(data[1]['container'], None)
            self.assertEqual(data[1]['file'], example_file)
            self.assertEqual(data[1]['analysis'], None)
            self.assertEqual(data[1]['macros'], [])
            self.assertEqual(data[1]['type'], 'OLE')
            self.assertEqual(data[2]['container'], example_file)
            self.assertNotEqual(data[2]['file'], example_file)
            self.assertEqual(data[2]['type'], "OpenXML")
            analysis = data[2]['analysis']
            self.assertEqual(analysis[0]['type'], 'AutoExec')
            self.assertEqual(analysis[0]['keyword'], 'Auto_Open')
            macros = data[2]['macros']
            self.assertEqual(macros[0]['vba_filename'], 'Modul1.bas')
            self.assertIn('Sub Auto_Open()', macros[0]['code'])


if __name__ == '__main__':
    unittest.main()
