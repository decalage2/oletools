"""
Test basic functionality of olevba[3]
"""

import unittest
import sys
if sys.version_info.major <= 2:
    from oletools import olevba
else:
    from oletools import olevba3 as olevba
import os
from os.path import join

# Directory with test data, independent of current working directory
from tests.test_utils import DATA_BASE_DIR


class TestOlevbaBasic(unittest.TestCase):
    """Tests olevba basic functionality"""

    def test_crypt_return(self):
        """
        Tests that encrypted files give a certain return code.

        Currently, only the encryption applied by Office 2010 (CryptoApi RC4
        Encryption) is tested.
        """
        CRYPT_DIR = join(DATA_BASE_DIR, 'encrypted')
        CRYPT_RETURN_CODE = 9
        ADD_ARGS = [], ['-d', ], ['-a', ], ['-j', ], ['-t', ]
        for filename in os.listdir(CRYPT_DIR):
            full_name = join(CRYPT_DIR, filename)
            for args in ADD_ARGS:
                try:
                    ret_code = olevba.main(args + [full_name, ])
                except SystemExit as se:
                    ret_code = se.code or 0   # se.code can be None
                self.assertEqual(ret_code, CRYPT_RETURN_CODE,
                                 msg='Wrong return code {} for args {}'
                                     .format(ret_code, args + [filename, ]))


# just in case somebody calls this file as a script
if __name__ == '__main__':
    unittest.main()
