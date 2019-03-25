""" Test ppt_parser and ppt_record_parser """

import unittest
import os
from os.path import join, splitext

# Directory with test data, independent of current working directory
from tests.test_utils import DATA_BASE_DIR

from oletools import ppt_record_parser
# ppt_parser not tested yet


class TestBasic(unittest.TestCase):
    """ test basic functionality of ppt parsing """

    def test_is_ppt(self):
        """ test ppt_record_parser.is_ppt(filename) """
        exceptions = ['encrypted.ppt', ]     # actually is ppt but embedded
        for base_dir, _, files in os.walk(DATA_BASE_DIR):
            for filename in files:
                if filename in exceptions:
                    continue
                full_name = join(base_dir, filename)
                extn = splitext(filename)[1]
                if extn in ('.ppt', '.pps', '.pot'):
                    self.assertTrue(ppt_record_parser.is_ppt(full_name),
                                    msg='{0} not recognized as ppt file'
                                        .format(full_name))
                else:
                    self.assertFalse(ppt_record_parser.is_ppt(full_name),
                                     msg='{0} erroneously recognized as ppt'
                                         .format(full_name))


# just in case somebody calls this file as a script
if __name__ == '__main__':
    unittest.main()
