"""
Test the MSO/ActiveMime extraction (mso_file_extract) in olevba
"""

import struct
import unittest
import zlib

from oletools.olevba import mso_file_extract

# zlib stream offset that Word writes into standard ActiveMime headers:
# the u16 hint at 0x1E is 4, and the stream starts at hint + 46 = 0x32
STD_OFFSET = 0x32


def make_activemime(payload, zlib_offset=STD_OFFSET, hint=None):
    """
    Build a synthetic ActiveMime container for tests.

    :param payload: bytes to store (compressed with zlib)
    :param zlib_offset: offset of the zlib stream within the container
    :param hint: value written to the u16 offset hint at 0x1E
                 (olevba reads the stream at hint + 46); defaults to the
                 value matching zlib_offset. Pass a wrong value to simulate
                 files whose header hint does not match reality.
    """
    if hint is None:
        hint = zlib_offset - 46
    header = bytearray(b'ActiveMime' + b'\x1f\xa6\x8f\x0c'
                       + b'\x00' * (zlib_offset - 14))
    struct.pack_into('<H', header, 0x1E, hint)
    return bytes(header) + zlib.compress(payload)


class TestMsoFileExtract(unittest.TestCase):
    """Tests the extraction of MSO/ActiveMime containers"""

    def test_standard_header(self):
        """Standard container: hint at 0x1E points to the zlib stream at 0x32."""
        payload = b'\xD0\xCF\x11\xE0' + b'A' * 64
        data = make_activemime(payload)
        self.assertEqual(mso_file_extract(data), payload)

    def test_bruteforce_fallback(self):
        """
        Regression test: when the header hint and the known fixed offsets do
        not match the actual zlib stream location, mso_file_extract falls back
        to brute-forcing 0x78 bytes. That fallback used a str regex pattern on
        bytes data and raised "TypeError: cannot use a string pattern on a
        bytes-like object" on Python 3, instead of recovering the data.
        """
        payload = b'\xD0\xCF\x11\xE0' + b'B' * 64
        # zlib at offset 33 (past the hint field at 0x1E), header hint
        # deliberately bogus (0xFFFF)
        data = make_activemime(payload, zlib_offset=33, hint=0xFFFF)
        self.assertEqual(mso_file_extract(data), payload)


if __name__ == '__main__':
    unittest.main()
