""" Test rtfobj.is_rtf """

from __future__ import print_function

import unittest
from os.path import join
from os import walk

from oletools.rtfobj import is_rtf, RTF_MAGIC

# Directory with test data, independent of current working directory
from tests.test_utils import DATA_BASE_DIR


class TestIsRtf(unittest.TestCase):
    """ Tests rtfobj.is_rtf """

    def test_bytearray(self):
        """ test that is_rtf works with bytearray """
        self.assertTrue(is_rtf(bytearray(RTF_MAGIC + b'asdfasdfasdfasdfasdf')))
        self.assertFalse(is_rtf(bytearray(RTF_MAGIC.upper() + b'asdfasdasdff')))
        self.assertFalse(is_rtf(bytearray(b'asdfasdfasdfasdfasdfasdfsdfsdfa')))

    def test_bytes(self):
        """ test that is_rtf works with bytearray """
        self.assertTrue(is_rtf(RTF_MAGIC + b'asasdffdfasdfasdfasdfasdf', True))
        self.assertFalse(is_rtf(RTF_MAGIC.upper() + b'asdffasdfasdasdff', True))
        self.assertFalse(is_rtf(b'asdfasdfasdfasdfasdfasdasdfffsdfsdfa', True))

    def test_tuple(self):
        """ test that is_rtf works with byte tuples """
        data = tuple(byte_char for byte_char in RTF_MAGIC + b'asdfasfadfdfsdf')
        self.assertTrue(is_rtf(data))

        data = tuple(byte_char for byte_char in RTF_MAGIC.upper() + b'asfasdf')
        self.assertFalse(is_rtf(data))

        data = tuple(byte_char for byte_char in b'asdfasfassdfsdsfeereasdfwdf')
        self.assertFalse(is_rtf(data))

    def test_iterable(self):
        """ test that is_rtf works with byte iterables """
        data = (byte_char for byte_char in RTF_MAGIC + b'asdfasfasasdfasdfddf')
        self.assertTrue(is_rtf(data))

        data = (byte_char for byte_char in RTF_MAGIC.upper() + b'asdfassfasdf')
        self.assertFalse(is_rtf(data))

        data = (byte_char for byte_char in b'asdfasfasasdfasdfasdfsdfdwerwedf')
        self.assertFalse(is_rtf(data))

    def test_files(self):
        """ test on real files """
        for base_dir, _, files in walk(DATA_BASE_DIR):
            for filename in files:
                full_path = join(base_dir, filename)
                expect = filename.endswith('.rtf')
                self.assertEqual(is_rtf(full_path), expect,
                                 'is_rtf({0}) did not return {1}'
                                 .format(full_path, expect))
                with open(full_path, 'rb') as handle:
                    self.assertEqual(is_rtf(handle), expect,
                                     'is_rtf(open({0})) did not return {1}'
                                     .format(full_path, expect))


# just in case somebody calls this file as a script
if __name__ == '__main__':
    unittest.main()
