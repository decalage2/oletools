"""Tests for CVE-2025-21298 detection in rtfobj (issue #883)."""

import struct
import unittest
import subprocess
import sys
from os.path import join, abspath, dirname

from tests.test_utils import testdata_reader
from oletools import rtfobj
from oletools import oleobj


class TestRtfObjIssue883(unittest.TestCase):
    """CVE-2025-21298 malformed StaticDib OLE object detection."""

    @classmethod
    def setUpClass(cls):
        cls.poc_data = testdata_reader.read(
            join('rtfobj', 'cve-2025-21298-poc.rtf'))

    def _parse_objects(self, data):
        rtfp = rtfobj.RtfObjParser(data)
        rtfp.parse()
        return rtfp.objects

    def test_poc_matches_indicator(self):
        objects = self._parse_objects(self.poc_data)
        self.assertEqual(len(objects), 1)
        self.assertTrue(rtfobj.is_cve_2025_21298_indicator(objects[0]))

    def test_poc_object_attributes(self):
        obj = self._parse_objects(self.poc_data)[0]
        self.assertEqual(obj.class_name, b'StaticDib')
        self.assertIsNone(obj.clsid)
        self.assertEqual(obj.oledata_size, 4)

    def test_rtfobj_cli_shows_cve_warning(self):
        poc_path = abspath(join(
            dirname(__file__), '..', 'test-data', 'rtfobj',
            'cve-2025-21298-poc.rtf'))
        result = subprocess.run(
            [sys.executable, '-m', 'oletools.rtfobj', poc_path],
            capture_output=True,
            text=True,
            check=False,
        )
        self.assertEqual(result.returncode, 0, msg=result.stderr)
        self.assertIn('CVE-2025-21298', result.stdout)

    def test_issue_251_not_flagged(self):
        data = testdata_reader.read(join('rtfobj', 'issue_251.rtf'))
        objects = self._parse_objects(data)
        for obj in objects:
            self.assertFalse(rtfobj.is_cve_2025_21298_indicator(obj))

    def _make_staticdib_obj(self, class_name=b'StaticDib', oledata=b'\x00' * 4,
                            clsid=None, oledata_size=None):
        obj = rtfobj.RtfObject()
        obj.is_ole = True
        obj.format_id = oleobj.OleObject.TYPE_EMBEDDED
        obj.class_name = class_name
        obj.clsid = clsid
        obj.oledata = oledata
        obj.oledata_size = oledata_size if oledata_size is not None else len(oledata)
        return obj

    def test_class_name_null_suffix(self):
        obj = self._make_staticdib_obj(class_name=b'StaticDib\x00')
        self.assertTrue(rtfobj.is_cve_2025_21298_indicator(obj))

    def test_class_name_case_insensitive(self):
        obj = self._make_staticdib_obj(class_name=b'staticdib')
        self.assertTrue(rtfobj.is_cve_2025_21298_indicator(obj))

    def test_padded_null_payload_still_flagged(self):
        obj = self._make_staticdib_obj(oledata=b'\x00' * 20)
        self.assertTrue(rtfobj.is_cve_2025_21298_indicator(obj))

    def test_malformed_with_clsid_still_flagged(self):
        obj = self._make_staticdib_obj(clsid='00000000-0000-0000-0000-000000000000')
        self.assertTrue(rtfobj.is_cve_2025_21298_indicator(obj))

    def test_plausible_dib_header_not_flagged(self):
        # Minimal BITMAPINFOHEADER (40 bytes) + one pixel row is plausible DIB data
        bi = struct.pack('<I', 40) + b'\x00' * 36
        obj = self._make_staticdib_obj(oledata=bi)
        self.assertFalse(rtfobj.is_cve_2025_21298_indicator(obj))

    def test_plausible_bmp_not_flagged(self):
        bmp = b'BM' + struct.pack('<IHHI', 62, 0, 0, 54) + b'\x00' * 48
        obj = self._make_staticdib_obj(oledata=bmp)
        self.assertFalse(rtfobj.is_cve_2025_21298_indicator(obj))


if __name__ == '__main__':
    unittest.main()
