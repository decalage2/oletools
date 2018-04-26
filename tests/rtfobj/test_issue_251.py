import unittest, sys, os

from tests.test_utils import testdata_reader
from oletools import rtfobj

class TestRtfObjIssue251(unittest.TestCase):
    def test_bin_no_param(self):
        data = testdata_reader.read('rtfobj/issue_251.rtf')
        rtfp = rtfobj.RtfObjParser(data)
        rtfp.parse()
        objects = rtfp.objects

        self.assertTrue(len(objects) == 1)

if __name__ == '__main__':
    unittest.main()
