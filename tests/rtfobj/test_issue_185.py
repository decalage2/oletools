import unittest, sys, os

from tests.test_utils import testdata_reader
from oletools import rtfobj

class TestRtfObjIssue185(unittest.TestCase):
    def test_skip_space_after_bin_control_word(self):
        data = testdata_reader.read_encrypted('rtfobj/issue_185.rtf.zip')
        rtfp = rtfobj.RtfObjParser(data)
        rtfp.parse()
        objects = rtfp.objects

        self.assertTrue(len(objects) == 1)

if __name__ == '__main__':
    unittest.main()
