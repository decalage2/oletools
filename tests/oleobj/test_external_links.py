""" Test that oleobj detects external links in relationships files.
"""

import unittest
import os
from os import path

# Directory with test data, independent of current working directory
from tests.test_utils import DATA_BASE_DIR, call_and_capture
from oletools import oleobj

BASE_DIR = path.join(DATA_BASE_DIR, 'oleobj', 'external_link')


class TestExternalLinks(unittest.TestCase):
    def test_external_links(self):
        """
        loop through sample files asserting that external links are found
        """

        for dirpath, _, filenames in os.walk(BASE_DIR):
            for filename in filenames:
                file_path = path.join(dirpath, filename)
                if not path.isfile(file_path):
                    continue

                output, ret_val = call_and_capture('oleobj', ['--nodump', file_path, ],
                                                   accept_nonzero_exit=True)
                self.assertEqual(ret_val, oleobj.RETURN_FOUND_EXTERNAL,
                                 msg='Wrong return value {} for {}. Output:\n{}'
                                     .format(ret_val, filename, output))
                found_relationship = False
                for line in output.splitlines():
                    if line.startswith('Found relationship'):
                        found_relationship = True
                        break
                self.assertTrue(found_relationship)


# just in case somebody calls this file as a script
if __name__ == '__main__':
    unittest.main()
