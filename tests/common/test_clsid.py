import unittest

from oletools.common.clsid import KNOWN_CLSIDS


class TestCommonClsid(unittest.TestCase):

    def test_known_clsids_uppercase(self):
        for k, v in KNOWN_CLSIDS.items():
            k_upper = k.upper()
            self.assertEqual(k, k_upper)
